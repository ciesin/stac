#!/usr/bin/env python3
"""Fill missing CIESIN item metadata from each row's DOIs and listed XML.

The input is a previously generated ciesin_items_*.xlsx workbook. Existing
nonblank cells are preserved. The updated Items workbook includes an Update
details sheet showing original values, candidate values, final values, and the
action for each item field, along with an Assets workbook in the output folder.

Requires openpyxl. Run from the project root, for example:
  python scripts/update_ciesin_items_workbook.py --input outputs/stac_workbooks/ciesin_items_20261020T120000Z.xlsx
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

import build_ciesin_items_and_assets_workbooks as builder
import classify_ciesin_asset_subtypes as asset_subtypes
import dedupe_harvard_doi_record_assets as harvard_dedupe


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_XML_DIR = ROOT / "inputs" / "xmls"
ITEM_SHEET_DEFAULT = "Items"
COMPLETENESS_FIELDS = (
    "is_latest_version", "item_title", "resource_type", "inventory_stac_category", "tags",
    "bbox_west", "bbox_south", "bbox_east", "bbox_north",
    "temporal_start", "temporal_end", "license", "description", "providers",
    "citation", "version", "additional_resources", "purpose", "authors",
    "publishers", "place", "doi", "publication_date",
)


def cell_is_blank(value: Any) -> bool:
    return value is None or isinstance(value, str) and not value.strip()


def infer_assets_input(items_path: Path) -> Path | None:
    name = items_path.name
    if name.startswith("ciesin_items_updated_"):
        candidate = items_path.with_name(name.replace("ciesin_items_updated_", "ciesin_assets_updated_", 1))
    elif name.startswith("ciesin_items_"):
        candidate = items_path.with_name(name.replace("ciesin_items_", "ciesin_assets_", 1))
    else:
        return None
    return candidate if candidate.is_file() else None


def read_asset_rows(path: Path | None) -> list[dict[str, Any]]:
    if path is None or not path.is_file():
        return []
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if "Assets" not in workbook.sheetnames:
            return []
        sheet = workbook["Assets"]
        headers = [builder.clean(cell.value) for cell in sheet[1]]
        positions = {name: index for index, name in enumerate(headers) if name}
        if "href" not in positions:
            return []
        assets: list[dict[str, Any]] = []
        for values in sheet.iter_rows(min_row=2, values_only=True):
            row = {name: values[index] if index < len(values) else None
                   for name, index in positions.items()}
            if builder.clean(row.get("href")):
                assets.append(row)
        return assets
    finally:
        workbook.close()


def match_assets_to_rows(rows: list[dict[str, str]], assets: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    by_dataset_id: dict[str, list[dict[str, Any]]] = {}
    by_doi: dict[str, list[dict[str, Any]]] = {}
    for asset in assets:
        dataset_id = builder.clean(asset.get("dataset_id")).casefold()
        if dataset_id:
            by_dataset_id.setdefault(dataset_id, []).append(asset)
        dois = builder.doi_keys(asset.get("dataset_doi")) | builder.doi_keys(dataset_id)
        for doi in dois:
            by_doi.setdefault(doi, []).append(asset)

    matched: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        dataset_id = builder.clean(row.get("dataset_id")).casefold()
        row_dois = set().union(*(builder.doi_keys(row.get(column, ""))
                                 for column in (*builder.DOI_COLUMNS, "doi")))
        row_assets = list(by_dataset_id.get(dataset_id, [])) if dataset_id else []
        if not row_assets:
            row_assets = [asset for doi in row_dois for asset in by_doi.get(doi, [])]
        unique = {builder.clean(asset.get("href")): asset for asset in row_assets if builder.clean(asset.get("href"))}
        matched[row["_excel_row"]] = list(unique.values())
    return matched


def remove_duplicate_harvard_record_assets(assets: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], int]:
    """Drop Harvard landing-page assets when the same dataset DOI asset exists."""
    doi_assets: dict[tuple[str, str], str] = {}
    for asset in assets:
        doi = harvard_dedupe.doi_asset_doi(asset)
        if doi:
            for dataset_key in harvard_dedupe.dataset_keys(asset):
                doi_assets[(dataset_key, doi)] = builder.clean(asset.get("href"))

    kept: list[dict[str, Any]] = []
    removed = 0
    for asset in assets:
        if harvard_dedupe.is_harvard_record_page(asset):
            doi = harvard_dedupe.record_page_doi(asset)
            href = next((doi_assets[(key, doi)] for key in harvard_dedupe.dataset_keys(asset)
                         if doi and (key, doi) in doi_assets), "")
            if href:
                removed += 1
                continue
        kept.append(asset)
    return kept, removed


def merge_asset_rows(existing: list[dict[str, Any]], discovered: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], int]:
    """Keep prior assets, append new ones, and enrich exact URL matches."""
    merged: list[dict[str, Any]] = []
    positions: dict[tuple[str, str], int] = {}
    for candidate in [*existing, *discovered]:
        href = builder.clean(candidate.get("href"))
        if not href:
            continue
        dataset_id = builder.clean(candidate.get("dataset_id")).casefold()
        if not dataset_id:
            dataset_id = ";".join(sorted(builder.doi_keys(candidate.get("dataset_doi"))))
        key = (dataset_id, href.casefold())
        if key not in positions:
            positions[key] = len(merged)
            merged.append(dict(candidate))
            continue

        prior = merged[positions[key]]
        for field, value in candidate.items():
            if value in (None, "") or field.startswith("_"):
                continue
            if prior.get(field) in (None, ""):
                prior[field] = value
        # Replace generic labels and weak default media types when a fresh
        # repository lookup returned the actual filename/type.
        old_title = builder.clean(prior.get("title"))
        new_title = builder.clean(candidate.get("title"))
        if (new_title and new_title != old_title
                and (re.fullmatch(r"Academic Commons resource \d+", old_title, re.I)
                     or old_title.casefold() in {"download", "file", "about", "resource"})):
            prior["title"] = new_title
            prior["asset_key"] = candidate.get("asset_key") or prior.get("asset_key")
        old_type = builder.clean(prior.get("media_type")).casefold()
        new_type = builder.clean(candidate.get("media_type"))
        if new_type and (not old_type or old_type in {"text/html", "application/octet-stream"}
                         and new_type not in {"text/html", "application/octet-stream"}):
            prior["media_type"] = new_type
        for field in ("description", "roles", "source", "source_identifier", "file_size_bytes",
                      "checksum_md5", "created", "updated", "asset_match_basis"):
            if candidate.get(field) not in (None, ""):
                prior[field] = candidate[field]
    return merged, len(existing) + len(discovered) - len(merged)


def missing_completion_fields(row: dict[str, str], has_asset: bool,
                             row_assets: list[dict[str, Any]] | None = None) -> list[str]:
    missing = [field for field in COMPLETENESS_FIELDS if cell_is_blank(row.get(field))]
    if not has_asset:
        missing.append("assets")
    # A DOI landing page alone is not the downloadable file. Revisit rows with
    # an Academic Commons DOI until at least one repository file asset has been
    # captured; this also discovers citation_pdf_url links omitted by the API.
    if builder.doi_list(row.get("columbia_ac_doi", "")):
        has_ac_file = any(
            "columbia academic commons file" in builder.clean(asset.get("source")).casefold()
            for asset in (row_assets or [])
        )
        if not has_ac_file:
            missing.append("Columbia Academic Commons file assets")
    return missing


def read_item_rows(sheet: Any) -> tuple[list[str], list[dict[str, str]], dict[str, int]]:
    header_values = [builder.clean(cell.value) for cell in sheet[1]]
    headers = [header for header in header_values if header]
    header_to_col = {builder.clean(sheet.cell(1, col).value): col for col in range(1, sheet.max_column + 1)
                     if builder.clean(sheet.cell(1, col).value)}
    required = ["item_title", *builder.DOI_COLUMNS, "inventory_xml_location"]
    missing = [column for column in required if column not in header_to_col]
    if missing:
        raise ValueError(f"Items sheet is missing required columns: {', '.join(missing)}")

    rows: list[dict[str, str]] = []
    row_numbers: list[int] = []
    for excel_row in range(2, sheet.max_row + 1):
        row: dict[str, str] = {}
        for name, col in header_to_col.items():
            row[name] = builder.clean(sheet.cell(excel_row, col).value)
        if any(row.get(column) for column in ["item_title", *builder.DOI_COLUMNS, "inventory_xml_location"]):
            row["_excel_row"] = str(excel_row)
            rows.append(row)
            row_numbers.append(excel_row)
    return headers, rows, {row["_excel_row"]: row_number for row, row_number in zip(rows, row_numbers)}


def ensure_item_columns(sheet: Any) -> dict[str, int]:
    positions: dict[str, int] = {}
    for col in range(1, sheet.max_column + 1):
        value = builder.clean(sheet.cell(1, col).value)
        if value:
            positions[value] = col
    for column in builder.ITEM_COLUMNS:
        if column in positions:
            continue
        col = sheet.max_column + 1
        cell = sheet.cell(1, col, column)
        if col > 1:
            source = sheet.cell(1, col - 1)
            if source.has_style:
                cell._style = source._style
            if source.number_format:
                cell.number_format = source.number_format
        if not cell.has_style:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
        positions[column] = col
    return positions


def fill_missing_cells(sheet: Any, rows: list[dict[str, str]], items: list[dict[str, Any]],
                       row_numbers: dict[str, int], positions: dict[str, int],
                       skip_row_ids: set[str]) -> Counter[str]:
    filled: Counter[str] = Counter()
    for row, candidate in zip(rows, items):
        if row["_excel_row"] in skip_row_ids:
            continue
        excel_row = row_numbers[row["_excel_row"]]
        for column in builder.ITEM_COLUMNS:
            value = candidate.get(column)
            if column not in positions or value is None or isinstance(value, str) and not value.strip():
                continue
            cell = sheet.cell(excel_row, positions[column])
            if cell_is_blank(cell.value):
                cell.value = value
                filled[column] += 1
    return filled


def add_update_log(workbook: Any, logs: list[tuple[str, str]]) -> None:
    base_name = "Update log"
    name = base_name
    suffix = 2
    while name in workbook.sheetnames:
        name = f"{base_name} {suffix}"
        suffix += 1
    sheet = workbook.create_sheet(name)
    sheet.append(["Metric", "Value"])
    for record in logs:
        sheet.append(list(record))
    builder.style_sheet(sheet)


def add_update_details(workbook: Any, rows: list[dict[str, str]],
                      candidates: list[dict[str, Any]], sheet: Any,
                      positions: dict[str, int], row_numbers: dict[str, int],
                      original_values: dict[str, dict[str, str]],
                      complete_row_ids: set[str], assets_by_row: dict[str, list[dict[str, Any]]],
                      offline: bool) -> None:
    """Record the old, proposed, and final value for every populated/candidate field."""
    base_name = "Update details"
    name = base_name
    suffix = 2
    while name in workbook.sheetnames:
        name = f"{base_name} {suffix}"
        suffix += 1
    details = workbook.create_sheet(name)
    details.append([
        "source_row", "row_status", "missing_fields", "dataset_id", "item_title", "field",
        "value_before", "candidate_value", "value_after", "action",
    ])
    for row, candidate in zip(rows, candidates):
        excel_row = row_numbers[row["_excel_row"]]
        row_assets = assets_by_row.get(row["_excel_row"], [])
        missing_fields = missing_completion_fields(row, bool(row_assets), row_assets)
        if row["_excel_row"] in complete_row_ids:
            row_status = "complete; DOI search skipped"
        else:
            row_status = "incomplete; offline mode" if offline else "incomplete; searched"
        missing_text = "; ".join(missing_fields)
        dataset_id = builder.clean(candidate.get("dataset_id") or row.get("dataset_id"))
        title = builder.clean(candidate.get("item_title") or row.get("item_title"))
        for field in builder.ITEM_COLUMNS:
            candidate_value = builder.clean(candidate.get(field))
            column = positions.get(field)
            before = original_values.get(row["_excel_row"], {}).get(field, "")
            after = builder.clean(sheet.cell(excel_row, column).value) if column else before
            if row["_excel_row"] in complete_row_ids:
                action = "skipped (complete)"
            elif before:
                action = "preserved"
            elif candidate_value:
                action = "filled"
            else:
                action = "no value found"
            if not (before or candidate_value or after):
                continue
            details.append([
                excel_row, row_status, missing_text, dataset_id, title, field,
                before, candidate_value, after, action,
            ])
    builder.style_sheet(details)
    details.freeze_panes = "A2"
    details.auto_filter.ref = details.dimensions


def copy_updated_values_to_contexts(contexts: list[dict[str, Any]], rows: list[dict[str, str]],
                                    sheet: Any, positions: dict[str, int],
                                    row_numbers: dict[str, int]) -> None:
    for context, row in zip(contexts, rows):
        excel_row = row_numbers[row["_excel_row"]]
        for column in ("dataset_id", "doi", "item_title", *builder.DOI_COLUMNS):
            col = positions.get(column)
            if col:
                value = builder.clean(sheet.cell(excel_row, col).value)
                if value:
                    context["item"][column] = value


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, type=Path,
                        help="Previously generated ciesin_items_*.xlsx workbook")
    parser.add_argument("--assets-input", type=Path,
                        help="Paired Assets workbook; inferred from the Items filename when omitted")
    parser.add_argument("--sheet", default=ITEM_SHEET_DEFAULT)
    parser.add_argument("--xml-dir", type=Path, default=DEFAULT_XML_DIR)
    parser.add_argument("--output-dir", type=Path,
                        help="Defaults to the input workbook's directory")
    parser.add_argument("--pause-seconds", type=float, default=1.5)
    parser.add_argument("--offline", action="store_true", help="Skip source API lookups; fill from listed XML only")
    args = parser.parse_args()
    if args.pause_seconds < 0:
        parser.error("--pause-seconds cannot be negative")

    input_path = args.input.resolve()
    xml_dir = args.xml_dir.resolve()
    output_dir = (args.output_dir or input_path.parent).resolve()
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    items_path = output_dir / f"ciesin_items_updated_{stamp}.xlsx"
    assets_path = output_dir / f"ciesin_assets_updated_{stamp}.xlsx"

    if not input_path.is_file():
        print(f"ERROR input workbook not found: {input_path}", file=sys.stderr)
        return 1
    assets_input_path = args.assets_input.resolve() if args.assets_input else infer_assets_input(input_path)
    existing_assets = read_asset_rows(assets_input_path)
    try:
        workbook = load_workbook(input_path)
        if args.sheet not in workbook.sheetnames:
            raise ValueError(f"Worksheet {args.sheet!r} not found. Available: {', '.join(workbook.sheetnames)}")
        sheet = workbook[args.sheet]
        _, rows, row_numbers = read_item_rows(sheet)
        positions = ensure_item_columns(sheet)
        original_values = {
            row["_excel_row"]: {
                field: builder.clean(sheet.cell(row_numbers[row["_excel_row"]], positions[field]).value)
                for field in builder.ITEM_COLUMNS if field in positions
            }
            for row in rows
        }
        assets_by_row = match_assets_to_rows(rows, existing_assets)
    except Exception as exc:
        print(f"ERROR reading item workbook: {exc}", file=sys.stderr)
        return 1

    logs: list[tuple[str, str]] = [
        ("Input workbook", str(input_path)), ("Input worksheet", args.sheet),
        ("Rows processed", str(len(rows))), ("XML directory", str(xml_dir)),
        ("Run timestamp UTC", datetime.now(timezone.utc).isoformat(timespec="seconds")),
    ]
    incomplete_rows = [row for row in rows
                       if missing_completion_fields(row, bool(assets_by_row.get(row["_excel_row"])),
                                                    assets_by_row.get(row["_excel_row"], []))]
    complete_row_ids = {row["_excel_row"] for row in rows
                        if not missing_completion_fields(row, bool(assets_by_row.get(row["_excel_row"])),
                                                         assets_by_row.get(row["_excel_row"], []))}
    logs.extend([
        ("Input Assets workbook", str(assets_input_path) if assets_input_path else "not found; no rows qualify as complete"),
        ("Existing asset rows read", str(len(existing_assets))),
        ("Complete rows skipped for DOI lookup", str(len(complete_row_ids))),
        ("Incomplete rows sent for DOI lookup", str(len(incomplete_rows))),
    ])
    if args.offline:
        remote_by_row: dict[str, list[dict[str, Any]]] = {}
        logs.append(("Per-row DOI lookups", "skipped by --offline"))
    else:
        remote_by_row, lookup_log = builder.lookup_inventory_dois(incomplete_rows, args.pause_seconds)
        logs.extend(lookup_log)

    candidates, contexts, _summary = builder.make_items(rows, xml_dir, remote_by_row, logs)
    filled = fill_missing_cells(sheet, rows, candidates, row_numbers, positions, complete_row_ids)
    logs.extend((f"Blank cells filled: {column}", str(count)) for column, count in sorted(filled.items()))
    logs.append(("Total blank cells filled", str(sum(filled.values()))))

    # Point assets at the final item identifiers and values from the updated
    # workbook, while leaving every populated input cell untouched.
    copy_updated_values_to_contexts(contexts, rows, sheet, positions, row_numbers)
    searchable_contexts = [context for context in contexts
                           if context["input"].get("_excel_row") not in complete_row_ids]
    discovered_assets = builder.assemble_assets(searchable_contexts, logs)
    assets, duplicate_asset_rows = merge_asset_rows(existing_assets, discovered_assets)
    assets, harvard_record_rows_removed = remove_duplicate_harvard_record_assets(assets)
    for asset in assets:
        if not builder.clean(asset.get("asset_subtype")):
            asset["asset_subtype"] = asset_subtypes.classify_asset(asset)
    logs.extend([
        ("Existing asset rows preserved", str(len(existing_assets))),
        ("New asset rows discovered", str(len(discovered_assets))),
        ("Exact dataset/URL asset duplicates merged", str(duplicate_asset_rows)),
        ("Duplicate Harvard record pages removed (DOI link retained)", str(harvard_record_rows_removed)),
    ])
    logs.extend([("Updated Items workbook", str(items_path)), ("Assets workbook", str(assets_path))])

    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        add_update_log(workbook, logs)
        add_update_details(workbook, rows, candidates, sheet, positions, row_numbers,
                           original_values, complete_row_ids, assets_by_row, args.offline)
        workbook.save(items_path)
        builder.save_workbook(assets_path, "Assets", builder.ASSET_COLUMNS, assets, logs)
    except Exception as exc:
        print(f"ERROR writing output workbooks: {exc}", file=sys.stderr)
        return 1

    print(json.dumps({
        "updated_items_workbook": str(items_path),
        "assets_workbook": str(assets_path),
        "rows_processed": len(rows),
        "doi_lookups_matched": sum(len(records) for records in remote_by_row.values()),
        "blank_cells_filled": sum(filled.values()),
        "assets_written": len(assets),
    }, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
