#!/usr/bin/env python3
"""Remove duplicate Harvard Dataverse record-page assets when a DOI link exists.

For each dataset, this keeps the https://doi.org/<DOI> asset and removes a
matching "Harvard Dataverse record" page asset. Other assets, including
Harvard Dataverse data files, are preserved. The input workbook is not changed.

Requires openpyxl. Example:
  python scripts/dedupe_harvard_doi_record_assets.py \
    --input outputs/stac_workbooks/ciesin_assets_updated.xlsx
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, unquote, urlparse

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


DOI_RE = re.compile(r"10\.\d{4,9}/[^\s\"<>]+", re.IGNORECASE)
TRAILING = ".,;:)]}"
REQUIRED_COLUMNS = {"dataset_id", "dataset_doi", "href", "source", "source_identifier"}


def clean(value: Any) -> str:
    return "" if value is None else re.sub(r"\s+", " ", str(value)).strip()


def dois(value: Any) -> set[str]:
    text = unquote(clean(value))
    text = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^doi:\s*", "", text, flags=re.IGNORECASE)
    return {match.rstrip(TRAILING).casefold() for match in DOI_RE.findall(text)}


def dataset_keys(row: dict[str, Any]) -> set[str]:
    keys: set[str] = set()
    dataset_id = clean(row.get("dataset_id")).casefold()
    if dataset_id:
        keys.add("id:" + dataset_id)
    for doi in dois(row.get("dataset_doi")):
        keys.add("doi:" + doi)
    return keys


def doi_asset_doi(row: dict[str, Any]) -> str:
    href = clean(row.get("href"))
    host = (urlparse(href).hostname or "").casefold()
    if host not in {"doi.org", "dx.doi.org"}:
        return ""
    found = dois(urlparse(href).path)
    return next(iter(found), "")


def record_page_doi(row: dict[str, Any]) -> str:
    """Get the DOI from a Harvard record asset's source ID or persistentId URL."""
    found = dois(row.get("source_identifier"))
    if found:
        return next(iter(found))
    parsed = urlparse(clean(row.get("href")))
    for value in parse_qs(parsed.query).values():
        found.update(dois(" ".join(value)))
    return next(iter(found), "")


def is_harvard_record_page(row: dict[str, Any]) -> bool:
    source = clean(row.get("source")).casefold()
    parsed = urlparse(clean(row.get("href")))
    host = (parsed.hostname or "").casefold()
    path = parsed.path.casefold()
    query = parse_qs(parsed.query)
    is_harvard_host = host == "dataverse.harvard.edu" or host.endswith(".dataverse.harvard.edu")
    # Dataverse record links commonly use dataset.xhtml?persistentId=doi:...
    # Some source labels say only "Harvard Dataverse DOI landing page", so
    # recognize the URL shape as well as descriptive source labels.
    is_dataset_record_url = (
        path.endswith("/dataset.xhtml") and
        any(key.casefold() in {"persistentid", "persistent_id"} for key in query)
    ) or (path.startswith("/dataset/") and is_harvard_host)
    is_labeled_record = "harvard dataverse" in source and "record" in source
    return is_harvard_host and (is_dataset_record_url or is_labeled_record)


def style_sheet(sheet: Any) -> None:
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for cells in sheet.columns:
        lengths = [len(str(cell.value or "")) for cell in list(cells)[:200]]
        sheet.column_dimensions[get_column_letter(cells[0].column)].width = min(
            max(max(lengths, default=10) + 2, 12), 55
        )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, type=Path, help="Assets workbook to clean")
    parser.add_argument("--output", type=Path,
                        help="Output workbook; defaults to <input name>_deduped.xlsx")
    parser.add_argument("--sheet", default="Assets", help="Assets worksheet name")
    args = parser.parse_args()

    input_path = args.input.resolve()
    if not input_path.is_file():
        print(f"ERROR: input workbook not found: {input_path}", file=sys.stderr)
        return 1
    output_path = args.output.resolve() if args.output else input_path.with_name(
        f"{input_path.stem}_deduped{input_path.suffix}"
    )
    if output_path == input_path:
        print("ERROR: output path must differ from input path", file=sys.stderr)
        return 1

    try:
        workbook = load_workbook(input_path)
        if args.sheet not in workbook.sheetnames:
            raise ValueError(f"Worksheet {args.sheet!r} not found; available: {', '.join(workbook.sheetnames)}")
        sheet = workbook[args.sheet]
        headers = [clean(cell.value) for cell in sheet[1]]
        positions = {name: index for index, name in enumerate(headers, start=1) if name}
        missing = REQUIRED_COLUMNS - positions.keys()
        if missing:
            raise ValueError(f"Assets sheet is missing columns: {', '.join(sorted(missing))}")

        records: list[dict[str, Any]] = []
        for row_number in range(2, sheet.max_row + 1):
            values = {name: sheet.cell(row_number, column).value
                      for name, column in positions.items()}
            values["_row_number"] = row_number
            values["_dataset_keys"] = dataset_keys(values)
            records.append(values)

        doi_assets: dict[tuple[str, str], str] = {}
        for row in records:
            doi = doi_asset_doi(row)
            if not doi:
                continue
            for key in row["_dataset_keys"]:
                doi_assets[(key, doi)] = clean(row.get("href"))

        remove_rows: list[int] = []
        report: list[list[Any]] = []
        for row in records:
            if not is_harvard_record_page(row):
                continue
            doi = record_page_doi(row)
            if not doi:
                continue
            doi_href = next((doi_assets[(key, doi)] for key in row["_dataset_keys"]
                             if (key, doi) in doi_assets), "")
            if not doi_href:
                continue
            remove_rows.append(row["_row_number"])
            report.append([
                row["_row_number"], clean(row.get("dataset_id")),
                clean(row.get("item_title")), doi,
                clean(row.get("href")), doi_href,
            ])

        for row_number in sorted(remove_rows, reverse=True):
            sheet.delete_rows(row_number, 1)
        style_sheet(sheet)

        report_name = "DOI deduplication"
        if report_name in workbook.sheetnames:
            del workbook[report_name]
        report_sheet = workbook.create_sheet(report_name)
        report_sheet.append([
            "removed_row_in_input", "dataset_id", "item_title", "matched_doi",
            "removed_harvard_record_href", "kept_doi_href",
        ])
        for row in report:
            report_sheet.append(row)
        report_sheet.append([])
        report_sheet.append(["Metric", "Value"])
        report_sheet.append(["Input workbook", str(input_path)])
        report_sheet.append(["Run timestamp UTC", datetime.now(timezone.utc).isoformat(timespec="seconds")])
        report_sheet.append(["Harvard record pages removed", len(remove_rows)])
        style_sheet(report_sheet)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(output_path)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"Output workbook: {output_path}")
    print(f"Duplicate Harvard record-page rows removed: {len(remove_rows)}")
    print("Matching DOI links and all nonduplicate file assets were preserved.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
