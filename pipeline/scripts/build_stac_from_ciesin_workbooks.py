#!/usr/bin/env python3
"""Build a static STAC catalog from the CIESIN Items and Assets workbooks.

By default, reads the paired workbooks from outputs/stac_workbooks and writes
STAC JSON under outputs/stac_from_workbooks/stac. Items are grouped into STAC
Collections using inventory_stac_category, and Assets are attached by dataset_id
with DOI matching as a fallback.

Requires openpyxl. Run from the project root or pass --items, --assets, and
--output-dir explicitly.
"""

from __future__ import annotations

import argparse
import json
import math
import mimetypes
import re
import shutil
import sys
import urllib.parse
from collections import defaultdict
from datetime import date, datetime, time, timezone
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

import classify_ciesin_asset_subtypes as asset_subtypes


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_ITEMS = ROOT / "outputs" / "stac_workbooks" / "ciesin_items_updated_20261021T235850Z.xlsx"
DEFAULT_ASSETS = ROOT / "outputs" / "stac_workbooks" / "ciesin_assets_updated_20261021T235850Z.xlsx"
DEFAULT_OUTPUT = ROOT / "outputs" / "stac_from_workbooks"

STAC_VERSION = "1.1.0"
ROOT_ID = "ciesin-sedac"
SCIENTIFIC_EXTENSION = "https://stac-extensions.github.io/scientific/v1.0.0/schema.json"
FILE_EXTENSION = "https://stac-extensions.github.io/file/v2.1.0/schema.json"


def clean(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return re.sub(r"\s+", " ", str(value)).strip()


def read_sheet(path: Path, sheet_name: str) -> list[dict[str, Any]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Worksheet {sheet_name!r} not found in {path.name}. Available: {', '.join(workbook.sheetnames)}")
        sheet = workbook[sheet_name]
        header_values = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
        headers = [clean(value) for value in header_values]
        rows: list[dict[str, Any]] = []
        for values in sheet.iter_rows(min_row=2, values_only=True):
            row = {header: values[index] if index < len(values) else None
                   for index, header in enumerate(headers) if header}
            if any(clean(value) for key, value in row.items() if key not in {"Update log", "Update details"}):
                rows.append(row)
        return rows
    finally:
        workbook.close()


def split_values(value: Any) -> list[str]:
    text = clean(value)
    if not text:
        return []
    # Workbook list-valued fields are stored as semicolon-separated text.
    return list(dict.fromkeys(part.strip() for part in text.split(";") if part.strip()))


def safe_id(value: Any, fallback: str = "item") -> str:
    slug = re.sub(r"[^a-zA-Z0-9_.-]+", "-", clean(value)).strip("-.").lower()
    return slug or fallback


def category_name(value: Any) -> str:
    return clean(value) or "Uncategorized"


def category_id(category: str) -> str:
    # Collection IDs should represent the spreadsheet category directly;
    # the catalog's organization identity belongs on the root Catalog.
    return safe_id(category, "uncategorized")


def parse_number(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None




def ciesin_temporal_value(value: Any) -> str | None:
    """Preserve temporal values, including BCE years, as strings."""
    if value is None or value == "":
        return None

    if isinstance(value, datetime):
        return value.isoformat(timespec="seconds")

    if isinstance(value, date):
        return value.isoformat()

    if isinstance(value, (int, float)):
        if not math.isfinite(float(value)):
            return None

        # Integer years, including BCE years, are not Excel date serials.
        if float(value).is_integer() and abs(value) <= 9999:
            return str(int(value))

    text = clean(value)
    return text or None

def is_before_stac_minimum(value: Any) -> bool:
    """
    Identify dates earlier than 0001-01-01, including BCE dates.
    """
    if value is None or value == "":
        return False

    if isinstance(value, (datetime, date)):
        return value.year < 1

    if isinstance(value, (int, float)):
        if math.isfinite(float(value)) and float(value).is_integer():
            return int(value) <= 0
        return False

    text = clean(value).upper()

    # Negative ISO-style years, e.g. -0044 or -10000-01-01.
    if re.match(r"^-\d+", text):
        return True

    # Explicit BCE / BC notation, e.g. 10000 BCE, 44 BC, BCE 10000.
    if re.search(r"\b\d[\d,]*\s*B\.?\s*C\.?\s*E?\.?\b", text):
        return True

    if re.search(r"\bB\.?\s*C\.?\s*E?\.?\s*\d", text):
        return True

    # Year zero is also before 0001-01-01.
    if re.match(r"^0000(?:[-/]|$)", text):
        return True

    return False



def as_datetime(value: Any) -> str | None:
    """Convert representable CE dates into UTC RFC3339 timestamps."""

    if value is None or value == "":
        return None

    if isinstance(value, datetime):
        parsed = value

    elif isinstance(value, date):
        parsed = datetime.combine(value, time.min)

    elif isinstance(value, (int, float)):
        if not math.isfinite(float(value)):
            return None

        # Numeric years must not be interpreted as Excel serial dates.
        if float(value).is_integer() and abs(value) <= 9999:
            year = int(value)
            if year < 1:
                return None
            parsed = datetime(year, 1, 1)
        else:
            try:
                parsed = from_excel(value)
            except (OverflowError, TypeError, ValueError):
                return None

    else:
        text = clean(value)

        # RFC3339 cannot represent BCE or negative years.
        if re.match(r"^-", text):
            return None

        # Year only: 2020
        if re.fullmatch(r"\d{4}", text):
            text += "-01-01"

        # Year and month: 2020-05
        elif re.fullmatch(r"\d{4}-\d{2}", text):
            text += "-01"

        # Compact date: 20200515
        elif re.fullmatch(r"\d{8}", text):
            text = f"{text[:4]}-{text[4:6]}-{text[6:8]}"

        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
            text += "T00:00:00Z"

        if " " in text and "T" not in text:
            text = text.replace(" ", "T", 1)

        if (
            "T" in text
            and not text.endswith("Z")
            and not re.search(r"[+-]\d{2}:\d{2}$", text)
        ):
            text += "Z"

        try:
            parsed = datetime.fromisoformat(
                text.replace("Z", "+00:00")
            )
        except (ValueError, OverflowError):
            return None

    if not (1 <= parsed.year <= 9999):
        return None

    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    else:
        parsed = parsed.astimezone(timezone.utc)

    return parsed.isoformat(timespec="seconds").replace("+00:00", "Z")




def doi_list(*values: Any) -> list[str]:
    found: list[str] = []
    for value in values:
        text = clean(value)
        for candidate in re.findall(r"10\.\d{4,9}/[^\s,;<>]+", text, flags=re.I):
            doi = candidate.rstrip(".,;:)]}").strip()
            if doi and doi.casefold() not in {item.casefold() for item in found}:
                found.append(doi)
    return found

def build_asset_indexes(asset_rows):
    by_dataset = defaultdict(list)
    by_doi = defaultdict(list)

    for row in asset_rows:
        dataset_id = clean(row.get("dataset_id")).casefold()

        if dataset_id:
            by_dataset[dataset_id].append(row)

        for doi in doi_list(row.get("dataset_doi")):
            by_doi[doi.casefold()].append(row)

    return by_dataset, by_doi

def item_dois(row: dict[str, Any]) -> list[str]:
    return doi_list(row.get("doi"), row.get("stac_nasa_doi"),
                    row.get("stac_harvard_doi"), row.get("columbia_ac_doi"))


def item_identifier(row: dict[str, Any], used: set[str]) -> str:
    base = safe_id(row.get("dataset_id") or row.get("doi") or row.get("item_title"), "item")
    item_id = base
    suffix = 2
    while item_id in used:
        item_id = f"{base}-{suffix}"
        suffix += 1
    used.add(item_id)
    return item_id


def geometry_from_bbox(row: dict[str, Any]) -> tuple[list[float] | None, dict[str, Any] | None]:
    west, south, east, north = (parse_number(row.get(field)) for field in
                                ("bbox_west", "bbox_south", "bbox_east", "bbox_north"))
    if None in (west, south, east, north):
        return None, None
    if not (-180 <= west <= 180 and -180 <= east <= 180 and -90 <= south <= 90 and -90 <= north <= 90):
        return None, None
    if south > north:
        return None, None
    bbox = [west, south, east, north]
    if west <= east:
        coordinates = [[[west, south], [east, south], [east, north], [west, north], [west, south]]]
        geometry = {"type": "Polygon", "coordinates": coordinates}
    else:
        # Split antimeridian-crossing rectangles into two polygons.
        coordinates = [
            [[[west, south], [180, south], [180, north], [west, north], [west, south]]],
            [[[-180, south], [east, south], [east, north], [-180, north], [-180, south]]],
        ]
        geometry = {"type": "MultiPolygon", "coordinates": coordinates}
    return bbox, geometry


def providers_list(value: Any) -> list[dict[str, Any]]:
    return [{"name": name, "roles": ["host"]} for name in split_values(value)]



def item_properties(
    row: dict[str, Any],
    category: str,
    dois: list[str],
) -> dict[str, Any]:

    # Parse dates that can be represented as standard RFC3339 timestamps.
    # Conventional STAC temporal values.
    start = as_datetime(row.get("temporal_start"))
    end = as_datetime(row.get("temporal_end"))
    publication_date = as_datetime(row.get("publication_date"))

    # Preserve CIESIN temporal values independently of STAC conversion.
    raw_start = ciesin_temporal_value(row.get("temporal_start"))
    raw_end = ciesin_temporal_value(row.get("temporal_end"))

    # Version status.
    latest_value = clean(row.get("is_latest_version")).casefold()

    is_latest = latest_value in {
        "yes", "y", "true", "1", "latest"
    }

    not_latest = latest_value in {
        "no", "n", "false", "0", "not latest", "not current"
    }

    version_status = (
        "current"
        if is_latest
        else "not_current"
        if not_latest
        else "unknown"
    )

    properties: dict[str, Any] = {
        "title": clean(row.get("item_title")),
        "description": clean(row.get("description")),
        "license": clean(row.get("license")) or "other",
        "keywords": split_values(row.get("tags")),
        "providers": providers_list(row.get("providers")),
        "sci:doi": dois[0] if dois else "",
        "sci:citation": clean(row.get("citation")),
        "version": clean(row.get("version")),

        "ciesin:resource_type": clean(row.get("resource_type")),
        "ciesin:category": category,
        "ciesin:is_latest_version": is_latest,
        "ciesin:version_status": version_status,
        "ciesin:dois": dois,
        "ciesin:authors": split_values(row.get("authors")),
        "ciesin:publishers": split_values(row.get("publishers")),
        "ciesin:place": split_values(row.get("place")),
        "ciesin:purpose": clean(row.get("purpose")),
        "ciesin:publication_date": publication_date,
        "ciesin:additional_resources": split_values(
            row.get("additional_resources")
        ),
        "ciesin:metadata_sources": split_values(
            row.get("metadata_sources")
        ),
        "ciesin:inventory_xml_location": clean(
            row.get("inventory_xml_location")
        ),

        # Preserve the original temporal extent, including dates outside
        # the range supported by standard RFC3339 timestamps.
        
        "ciesin:temporal_start": raw_start,
        "ciesin:temporal_end": raw_end,

    }

    # Add "latest" to the keywords for current versions.
    if is_latest:
        if "latest" not in {
            keyword.casefold()
            for keyword in properties["keywords"]
        }:
            properties["keywords"].append("latest") 

        # ---------------------------------------------------------
    # STAC temporal properties
    # ---------------------------------------------------------

    has_start = raw_start is not None
    has_end = raw_end is not None

    if has_start or has_end:

        # Convert the supplied temporal boundaries.
        start = as_datetime(row.get("temporal_start"))
        end = as_datetime(row.get("temporal_end"))

        # Clamp BCE start dates to the earliest representable date.
        if start is None and is_before_stac_minimum(
            row.get("temporal_start")
        ):
            start = "0001-01-01T00:00:00Z"

        # CASE 1: One date only (start and end are not both supplied).
        if has_start != has_end:

            single_date = end if has_end else start

            if single_date is None:
                raise ValueError(
                    f"Temporal date cannot be represented as RFC3339: "
                    f"{row.get('item_title')}"
                )

            properties["datetime"] = single_date

            properties.pop("start_datetime", None)
            properties.pop("end_datetime", None)

        # CASE 2: Both start and end dates supplied.
        else:

            if start is None or end is None:
                raise ValueError(
                    f"Cannot represent both temporal boundaries: "
                    f"{row.get('item_title')} "
                    f"({raw_start} to {raw_end})"
                )

            if start > end:
                raise ValueError(
                    f"Temporal start is after temporal end: "
                    f"{row.get('item_title')} "
                    f"({raw_start} to {raw_end})"
                )

            properties["datetime"] = None
            properties["start_datetime"] = start
            properties["end_datetime"] = end

        properties["ciesin:stac_datetime_basis"] = (
            "temporal_extent"
        )

    else: 

        # No temporal extent: use publication date.
        if publication_date is None:
            raise ValueError(
                f"Cannot create a conventional STAC datetime for: "
                f"{row.get('item_title')}. "
                f"No temporal extent or valid publication date."
            )

        properties["datetime"] = publication_date

        properties.pop("start_datetime", None)
        properties.pop("end_datetime", None)

        properties["ciesin:stac_datetime_basis"] = (
            "publication_date"
        )

    # Remove empty optional properties while retaining required
    # STAC fields and meaningful false/zero values.
    required_fields = {
        "datetime",
        "license",
    }

    return {
        key: value
        for key, value in properties.items()
        if value not in ("", [], {})
        or key in required_fields
    }

def asset_object(row: dict[str, Any]) -> dict[str, Any]:
    href = clean(row.get("href"))
    title = clean(row.get("title"))
    path = urllib.parse.urlparse(href).path
    if not title:
        title = Path(urllib.parse.unquote(path)).name or clean(row.get("asset_key"))
    media_type = clean(row.get("media_type")) or mimetypes.guess_type(path)[0] or "application/octet-stream"
    # Repository downloads often use a generic /download URL. In that case,
    # the filename supplied in the Assets workbook still identifies a PDF.
    if media_type == "application/octet-stream" and title.casefold().endswith(".pdf"):
        media_type = "application/pdf"

    file_size = parse_number(row.get("file_size_bytes"))
    
    asset: dict[str, Any] = {
        "href": href,
        "type": media_type,
        "title": title,
        "description": clean(row.get("description")),
        "roles": split_values(row.get("roles")),
        "file:size": int(file_size) if file_size is not None else None,
        "ciesin:checksum_md5": clean(row.get("checksum_md5")),
        "created": as_datetime(row.get("created")),
        "updated": as_datetime(row.get("updated")),
        "ciesin:source": clean(row.get("source")),
        "ciesin:asset_subtype": clean(row.get("asset_subtype")) or asset_subtypes.classify_asset(row),
        "ciesin:source_identifier": clean(row.get("source_identifier")),
        "ciesin:asset_match_basis": clean(row.get("asset_match_basis")),
    }
    return {key: value for key, value in asset.items() if value not in (None, "", [])}

def assets_for_item(item_row, asset_indexes):
    by_dataset, by_doi = asset_indexes

    dataset_id = clean(item_row.get("dataset_id")).casefold()
    dois = item_dois(item_row)

    # Prefer dataset_id; use DOI only when no dataset matches.
    candidates = by_dataset.get(dataset_id, []) if dataset_id else []

    if not candidates:
        candidates = [
            row
            for doi in dois
            for row in by_doi.get(doi.casefold(), [])
        ]

    matches = []
    seen = set()

    for row in candidates:
        href = clean(row.get("href"))

        if href and href not in seen:
            matches.append(row)
            seen.add(href)

    return matches


def group_assets(rows):
    assets = {}
    seen_hrefs = set()

    for row in rows:
        href = clean(row.get("href"))

        if not href or href in seen_hrefs:
            continue

        seen_hrefs.add(href)

        path_name = Path(
            urllib.parse.unquote(
                urllib.parse.urlparse(href).path
            )
        ).name

        base = safe_id(
            row.get("asset_key") or row.get("title") or path_name,
            "asset"
        )

        key = base
        suffix = 2

        while key in assets:
            key = f"{base}-{suffix}"
            suffix += 1

        assets[key] = asset_object(row)

    return assets
def collection_extent(items: list[dict[str, Any]]) -> dict[str, Any]:

    boxes = [
        item["bbox"]
        for item in items
        if item.get("bbox")
    ]

    if boxes:
        spatial_bbox = [
            min(b[0] for b in boxes),
            min(b[1] for b in boxes),
            max(b[2] for b in boxes),
            max(b[3] for b in boxes),
        ]
    else:
        spatial_bbox = [-180, -90, 180, 90]

    # Initialize dates regardless of whether bounding boxes exist.
    dates: list[str] = []

    for item in items:
        props = item["properties"]

        if props.get("ciesin:stac_datetime_basis") != "temporal_extent":
            continue

        dates.extend(
            value
            for value in (
                props.get("datetime"),
                props.get("start_datetime"),
                props.get("end_datetime"),
            )
            if isinstance(value, str) and value
        )

    interval = (
        [min(dates), max(dates)]
        if dates
        else [None, None]
    )

    return {
        "spatial": {
            "bbox": [spatial_bbox]
        },
        "temporal": {
            "interval": [interval]
        },
    }

def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def build(items_path: Path, assets_path: Path, output_dir: Path, clean_output: bool = False) -> dict[str, Any]:
    items_path = items_path.resolve()
    assets_path = assets_path.resolve()
    output_dir = output_dir.resolve()
    if not items_path.is_file():
        raise FileNotFoundError(f"Items workbook not found: {items_path}")
    if not assets_path.is_file():
        raise FileNotFoundError(f"Assets workbook not found: {assets_path}")

    stac_dir = output_dir / "stac"
    if clean_output and stac_dir.exists():
        shutil.rmtree(stac_dir)
    item_rows = read_sheet(items_path, "Items")
    asset_rows = read_sheet(assets_path, "Assets")

    asset_indexes = build_asset_indexes(asset_rows)

    category_items: dict[str, list[dict[str, Any]]] = defaultdict(list)
    items_index: list[dict[str, Any]] = []
    used_item_ids: set[str] = set()
    root_links = [{"rel": "self", "href": "./catalog.json", "type": "application/json"}]
    asset_attach_count = 0
    for row in item_rows:
        category = category_name(row.get("inventory_stac_category"))
        collection = category_id(category)
        item_id = item_identifier(row, used_item_ids)
        dois = item_dois(row)
        bbox, geometry = geometry_from_bbox(row)
        linked_asset_rows = assets_for_item(row, asset_indexes)
        assets = group_assets(linked_asset_rows)
        asset_attach_count += len(assets)
        properties = item_properties(row, category, dois)
        has_file_metadata = any(
            "file:size" in asset
            for asset in assets.values()
        )

        extensions = []

        if has_file_metadata:
            extensions.append(FILE_EXTENSION)

        if "sci:doi" in properties or "sci:citation" in properties:
            extensions.append(SCIENTIFIC_EXTENSION)
        links = [
            {"rel": "self", "href": f"./{item_id}.json", "type": "application/geo+json"},
            {"rel": "root", "href": "../../../catalog.json", "type": "application/json"},
            {"rel": "collection", "href": "../collection.json", "type": "application/json"},
        ]
        links.extend({"rel": "cite-as", "href": f"https://doi.org/{doi}", "type": "text/html"} for doi in dois)
        item = {
            "type": "Feature", "stac_version": STAC_VERSION,
            "stac_extensions": extensions, "id": item_id, "collection": collection,
            "geometry": geometry, "properties": properties, "links": links, "assets": assets,
        }
        if bbox is not None:
            item["bbox"] = bbox
        item_path = stac_dir / "collections" / collection / "items" / f"{item_id}.json"
        write_json(item_path, item)
        category_items[category].append(item)
        items_index.append({
            "id": item_id, "collection": collection, "category": category,
            "title": properties.get("title", item_id), "description": properties.get("description", ""),
            "resource_type": properties.get("ciesin:resource_type", ""),
            "datetime": properties.get("datetime"),
            "start_datetime": properties.get("start_datetime"),
            "end_datetime": properties.get("end_datetime"),

            "ciesin_temporal_start": properties.get("ciesin:temporal_start"),
            "ciesin_temporal_end": properties.get("ciesin:temporal_end"),

            # Backward compatibility with the existing browser.
            "temporal_start": properties.get("ciesin:temporal_start"),
            "temporal_end": properties.get("ciesin:temporal_end"),
            "bbox": item.get("bbox"), "asset_count": len(assets), "keywords": properties.get("keywords", []),
            "is_latest_version": properties.get("ciesin:is_latest_version", False),
            "dois": dois, "license": properties.get("license"),
            "citation": properties.get("sci:citation", ""),
            "version": properties.get("version", ""),
            "version_status": properties.get("ciesin:version_status", "unknown"),
            "authors": properties.get("ciesin:authors", []),
            "publishers": properties.get("ciesin:publishers", []),
            "providers": [provider.get("name", "") for provider in properties.get("providers", [])],
            "place": properties.get("ciesin:place", []),
            "purpose": properties.get("ciesin:purpose", ""),
            "publication_date": properties.get("ciesin:publication_date", ""),
            "additional_resources": properties.get("ciesin:additional_resources", []),
            "tags": properties.get("keywords", []),
            "item_href": item_path.relative_to(stac_dir).as_posix(),
        })

    catalog_links = list(root_links)
    category_counts: dict[str, Any] = {}
    for category, items in sorted(category_items.items(), key=lambda pair: pair[0].casefold()):
        collection = category_id(category)
        collection_items = sorted(items, key=lambda item: item["properties"].get("title", "").casefold())
        item_links = [{"rel": "item", "href": f"./items/{item['id']}.json",
                       "type": "application/geo+json", "title": item["properties"].get("title", item["id"])}
                      for item in collection_items]
        collection_doc = {
            "type": "Collection", "stac_version": STAC_VERSION, "stac_extensions": [],
            "id": collection, "title": category,
            "description": f"CIESIN datasets categorized as {category}.", "license": "other",
            "extent": collection_extent(items),
            "links": [
                {"rel": "root", "href": "../../catalog.json", "type": "application/json"},
                {"rel": "self", "href": "./collection.json", "type": "application/json"},
            ] + item_links,
        }
        write_json(stac_dir / "collections" / collection / "collection.json", collection_doc)
        catalog_links.append({"rel": "child", "href": f"./collections/{collection}/collection.json",
                              "type": "application/json", "title": collection_doc["title"]})
        category_counts[category] = {
            "collection_id": collection, "item_count": len(items),
            "asset_count": sum(len(item["assets"]) for item in items),
        }

    catalog = {
        "type": "Catalog", "stac_version": STAC_VERSION, "id": ROOT_ID,
        "title": "CIESIN SEDAC STAC Catalog",
        "description": "STAC catalog generated from the CIESIN Items and Assets workbooks.",
        "links": catalog_links,
    }
    write_json(stac_dir / "catalog.json", catalog)
    write_json(stac_dir / "items-index.json", sorted(items_index,
               key=lambda entry: (entry["category"].casefold(), entry["title"].casefold())))
    summary = {
        "items_workbook": str(items_path), "assets_workbook": str(assets_path),
        "stac_dir": str(stac_dir), "item_count": len(items_index),
        "input_asset_rows": len(asset_rows), "attached_asset_count": asset_attach_count,
        "collection_count": len(category_items), "category_counts": category_counts,
        "asset_rows_without_href": sum(not clean(row.get("href")) for row in asset_rows),
    }
    write_json(stac_dir / "summary.json", summary)
    return summary


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--items", type=Path, default=DEFAULT_ITEMS, help="Items workbook")
    parser.add_argument("--assets", type=Path, default=DEFAULT_ASSETS, help="Assets workbook")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT,
                        help="Output directory; catalog is written in its stac subfolder")
    parser.add_argument("--clean", action="store_true", help="Remove only output-dir/stac before rebuilding")
    args = parser.parse_args()
    try:
        summary = build(args.items, args.assets, args.output_dir, args.clean)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
