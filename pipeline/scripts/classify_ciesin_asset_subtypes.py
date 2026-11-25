#!/usr/bin/env python3
"""Enrich a CIESIN Assets workbook with clearer asset labels and shared-link flags.

The input workbook is preserved. By default, writes <input>_classified.xlsx.
Keeps the legacy asset_subtype column (Data/Metadata) for compatibility and adds:
  asset_type, asset_source_label, asset_format, asset_relationship,
  shared_across_datasets, related_dataset_ids, related_dataset_titles

Example (from D:\\codex):
  python .\\scripts\\classify_ciesin_asset_subtypes_improved.py --input .\\outputs\\stac_workbooks\\ciesin_assets.xlsx
"""
from __future__ import annotations
import argparse
import re
import sys
from collections import Counter, defaultdict
from copy import copy
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse, urlunparse
from openpyxl import load_workbook

GENERIC_TITLES = {"", "download", "file", "data", "document", "pdf", "xml", "record", "record page", "landing page", "doi", "link", "resource"}
METADATA_MARKERS = ("metadata", "iso 19139", "record page", "catalog record", "metadata record", "citation record")
DOC_MARKERS = ("readme", "user guide", "user manual", "documentation", "release notes", "methodology", "data dictionary", "codebook", "technical report", "technical documentation")
PUBLICATION_MARKERS = ("publication", "related publication", "journal article", "research article", "paper", "article", "working paper", "book chapter", "report")
DATA_MARKERS = ("dataset", "data file", "data download", "granule", "shapefile", "geotiff", "netcdf", "geojson", "csv", "spreadsheet")


def clean(value: Any) -> str:
    return "" if value is None else re.sub(r"\s+", " ", str(value)).strip()


def norm_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "_", clean(value).casefold()).strip("_")


def get(row: dict[str, Any], *names: str) -> str:
    normalized = {norm_header(k): clean(v) for k, v in row.items()}
    for name in names:
        if normalized.get(norm_header(name)):
            return normalized[norm_header(name)]
    return ""


def normalize_href(href: str) -> str:
    """Normalize host/scheme and remove fragments; retain path/query case and values."""
    try:
        p = urlparse(href.strip())
        if not p.scheme or not p.netloc:
            return href.strip().split("#", 1)[0]
        return urlunparse((p.scheme.casefold(), p.netloc.casefold(), p.path, p.params, p.query, ""))
    except Exception:
        return href.strip().split("#", 1)[0]


def source_label(row: dict[str, Any]) -> str:
    source = get(row, "source", "asset_source", "provider", "publisher").casefold()
    href = get(row, "href", "url")
    host = (urlparse(href).hostname or "").casefold()
    text = f"{source} {host}"
    if "cmr" in text or "earthdata" in text or "nasa" in text:
        return "NASA CMR / Earthdata"
    if "dataverse.harvard.edu" in text or "harvard dataverse" in text:
        return "Harvard Dataverse"
    if "academiccommons" in text or "columbia.edu" in text or "columbia academic commons" in text:
        return "Columbia Academic Commons"
    if "ciesin" in text or "sedac" in text:
        return "CIESIN / SEDAC"
    return clean(get(row, "source", "asset_source", "provider", "publisher")) or host or "Source not specified"


def asset_format(row: dict[str, Any]) -> str:
    href = get(row, "href", "url")
    path = unquote(urlparse(href).path).casefold()
    media = get(row, "media_type", "type", "mime_type").casefold().split(";", 1)[0]
    filename = path.rsplit("/", 1)[-1]
    ext = filename.rsplit(".", 1)[-1] if "." in filename else ""
    by_ext = {
        "pdf": "PDF", "csv": "CSV", "tsv": "TSV", "xlsx": "Excel workbook", "xls": "Excel workbook",
        "zip": "ZIP archive", "shp": "Shapefile", "tif": "GeoTIFF", "tiff": "GeoTIFF",
        "gpkg": "GeoPackage", "geojson": "GeoJSON", "json": "JSON", "xml": "XML",
        "nc": "NetCDF", "nc4": "NetCDF", "h5": "HDF5", "hdf": "HDF", "txt": "Text file",
        "doc": "Word document", "docx": "Word document", "ppt": "PowerPoint", "pptx": "PowerPoint",
        "jpg": "JPEG image", "jpeg": "JPEG image", "png": "PNG image", "html": "HTML page", "htm": "HTML page",
    }
    if ext in by_ext:
        return by_ext[ext]
    if media == "application/pdf": return "PDF"
    if "html" in media: return "HTML page"
    if "xml" in media: return "XML"
    if "json" in media: return "JSON"
    if "csv" in media: return "CSV"
    if "geotiff" in media or "tiff" in media: return "GeoTIFF"
    if "netcdf" in media: return "NetCDF"
    if media and media != "application/octet-stream": return media
    return "Format not specified"


def descriptive_asset_title(row: dict[str, Any], asset_type: str = "") -> str:
    """Return a dataset-specific display title, correcting misleading labels."""
    href = get(row, "href", "url")
    p = urlparse(href)
    host = (p.hostname or "").casefold()
    path = unquote(p.path)
    path_lower = path.casefold()
    filename = path.rsplit("/", 1)[-1]
    dataset = get(row, "item_title", "dataset_title", "item_name", "title_dataset") or get(row, "dataset_id", "item_id") or "Dataset"
    supplied = get(row, "title", "asset_title")
    media = get(row, "media_type", "type", "mime_type").casefold().split(";", 1)[0]

    if host == "search.earthdata.nasa.gov" and "/search/granules" in path_lower:
        return f"NASA Earthdata granule search — {dataset}"
    if filename.casefold().endswith("-metadata.zip") or filename.casefold() == "metadata.zip":
        return f"Metadata package — {dataset}"

    image_exts = (".jpg", ".jpeg", ".png", ".gif", ".tif", ".tiff", ".webp")
    is_image = path_lower.endswith(image_exts) or media.startswith("image/")
    if is_image:
        kind = "Map image" if any(x in path_lower for x in ("map", "thumbnail", "thumb")) else "Image"
        return f"{kind} — {dataset} ({filename})" if filename else f"{kind} — {dataset}"

    if any(x in supplied.casefold() for x in ("data set documentation", "dataset documentation", "documentation", "user guide", "readme", "release notes", "data dictionary", "codebook", "methodology")):
        return f"Dataset documentation — {dataset}"

    if "academiccommons" in host:
        if supplied and supplied.casefold() not in GENERIC_TITLES:
            return supplied
        if filename:
            return f"{dataset} — {filename}"
        return f"Columbia Academic Commons resource — {dataset}"

    return supplied or f"{dataset} — {asset_type or 'Asset'}"

def classify_type(row: dict[str, Any]) -> tuple[str, str]:
    href = get(row, "href", "url")
    p = urlparse(href)
    host = (p.hostname or "").casefold()
    path = unquote(p.path).casefold()
    title = get(row, "title", "asset_title").casefold()
    desc = get(row, "description", "asset_description").casefold()
    source = get(row, "source", "asset_source").casefold()
    roles = {
        x.strip().casefold()
        for x in re.split(r"[;,|]", get(row, "roles", "asset_roles"))
        if x.strip()
    }
    media = get(row, "media_type", "type", "mime_type").casefold().split(";", 1)[0]
    context = " ".join((title, desc, source, path))
    fmt = asset_format(row)

    filename = unquote(path.rsplit("/", 1)[-1])

    # NASA Earthdata granule search.
    if host == "search.earthdata.nasa.gov" and "/search/granules" in path:
        return "NASA Earthdata granule search", "NASA Earthdata granule search"

    # Metadata ZIP packages.
    if filename.endswith("-metadata.zip") or filename == "metadata.zip":
        return "Metadata package", "dataset metadata package"

    # Images must be identified before generic documentation/report markers.
    if media.startswith("image/") or filename.endswith(
        (".jpg", ".jpeg", ".png", ".gif", ".tif", ".tiff", ".webp")
    ):
        return "Image", "dataset image / map image"

    # Explicit appendix files.
    if "appendix" in filename or "appendix" in title:
        return "Appendix", "dataset appendix"

    # Documentation.
    if any(x in context for x in DOC_MARKERS):
        return "Documentation", "documentation"

    # Publications.
    if any(x in context for x in PUBLICATION_MARKERS):
        return "Related publication", "related publication"

    # Metadata.
    if "metadata" in roles or any(x in context for x in METADATA_MARKERS):
        return "Metadata record", "metadata"

    # DOI landing pages.
    if host in {"doi.org", "dx.doi.org"}:
        return "DOI landing page", "repository/DOI landing page"

    # Harvard Dataverse.
    if "dataverse.harvard.edu" in host and (
        "dataset.xhtml" in path or "file.xhtml" in path
    ):
        if "file.xhtml" in path or "download" in path:
            return "Dataset file", "dataset data"
        return "Repository record", "repository landing page"

    # Columbia Academic Commons.
    if "academiccommons" in host:
        is_download = (
            any(ext in path for ext in (
                ".pdf", ".doc", ".docx", ".ppt", ".pptx",
                ".csv", ".xlsx", ".xls", ".zip", ".tif", ".tiff",
                ".shp", ".gpkg", ".geojson", ".nc"
            ))
            or any(x in path for x in ("/download", "/bitstream/", "/content/"))
            or media not in {"", "text/html", "application/xhtml+xml"}
        )

        if is_download:
            if any(x in context for x in PUBLICATION_MARKERS):
                return "Related publication", "related publication"
            if any(x in context for x in DOC_MARKERS):
                return "Documentation", "documentation"
            if fmt in {"PDF", "Word document", "PowerPoint"}:
                return "Related publication or document", "related publication/document"
            return "Dataset file", "dataset data"

        return "Repository record", "repository landing page"

    return "Other / review", "unclear; review"


def legacy_subtype(asset_type: str) -> str:
    # Preserve the two-value convention used by existing browser code.
    return "Metadata" if asset_type in {"Metadata record", "Repository record", "DOI landing page", "Landing page", "Documentation", "Related publication", "Related publication or document", "Document (review type)", "NASA Earthdata granule search", "Metadata package"} else "Data"


def classify_asset(row: dict[str, Any]) -> str:
    """Backward-compatible API used by the item/asset workbook builder.

    The builder expects the legacy Data/Metadata value. The richer
    classification is available from classify_type(row).
    """
    asset_type, _relationship = classify_type(row)
    return legacy_subtype(asset_type)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, type=Path, help="Existing Assets workbook")
    parser.add_argument("--output", type=Path, help="Output path (defaults to <input>_classified.xlsx)")
    parser.add_argument("--sheet", default="Assets", help="Assets worksheet name")
    parser.add_argument("--reclassify", action="store_true", help="Recalculate existing classification fields")
    args = parser.parse_args()
    input_path = args.input.resolve()
    if not input_path.is_file():
        print(f"ERROR: input workbook not found: {input_path}", file=sys.stderr); return 1
    output_path = args.output.resolve() if args.output else input_path.with_name(f"{input_path.stem}_classified{input_path.suffix}")
    if output_path == input_path:
        print("ERROR: output path must differ from input path", file=sys.stderr); return 1
    try:
        wb = load_workbook(input_path)
        if args.sheet not in wb.sheetnames:
            raise ValueError(f"Worksheet {args.sheet!r} not found; available: {', '.join(wb.sheetnames)}")
        ws = wb[args.sheet]
        headers = [clean(c.value) for c in ws[1]]
        positions = {norm_header(h): i for i, h in enumerate(headers, 1) if h}
        if "href" not in positions and "url" not in positions:
            raise ValueError("Assets sheet must contain an href or url column")
        href_col = positions.get("href", positions.get("url"))
        # Dataset identifier/title columns are detected flexibly; if absent, shared-link checks are skipped.
        id_candidates = ("dataset_id", "item_id", "stac_item_id", "item_identifier", "id")
        title_candidates = ("dataset_title", "item_title", "stac_item_title", "item_name")
        id_col = next((positions[norm_header(n)] for n in id_candidates if norm_header(n) in positions), None)
        ds_title_col = next((positions[norm_header(n)] for n in title_candidates if norm_header(n) in positions), None)
        if "title" not in positions and "asset_title" not in positions:
            col = ws.max_column + 1
            ws.cell(1, col, "title")
            positions["title"] = col
        new_columns = ["asset_type", "asset_source_label", "asset_format", "asset_relationship", "shared_across_datasets", "related_dataset_ids", "related_dataset_titles"]
        for name in new_columns:
            if norm_header(name) not in positions:
                col = ws.max_column + 1
                ws.cell(1, col, name)
                if col > 1 and ws.cell(1, col - 1).has_style:
                    ws.cell(1, col)._style = copy(ws.cell(1, col - 1)._style)
                positions[norm_header(name)] = col
        subtype_col = positions.get("asset_subtype")
        if subtype_col is None:
            subtype_col = ws.max_column + 1
            ws.cell(1, subtype_col, "asset_subtype")
            positions["asset_subtype"] = subtype_col
        # Read rows, classify and collect cross-dataset URL usage.
        records = []
        url_datasets: dict[str, dict[str, str]] = defaultdict(dict)
        for r in range(2, ws.max_row + 1):
            row = {headers[c-1]: ws.cell(r, c).value for c in range(1, len(headers)+1) if c <= len(headers) and headers[c-1]}
            href = clean(ws.cell(r, href_col).value)
            if not href: continue
            dataset_id = clean(ws.cell(r, id_col).value) if id_col else ""
            dataset_title = clean(ws.cell(r, ds_title_col).value) if ds_title_col else ""
            if not dataset_title:
                dataset_title = dataset_id
            key = normalize_href(href)
            if dataset_id:
                url_datasets[key][dataset_id] = dataset_title or dataset_id
            records.append((r, row, key, dataset_id, dataset_title))
        counts: Counter[str] = Counter()
        for r, row, key, dataset_id, dataset_title in records:
            asset_type, relationship = classify_type(row)
            source = source_label(row)
            fmt = asset_format(row)
            other = {i: t for i, t in url_datasets[key].items() if i != dataset_id}
            shared = "Yes" if other else "No"
            values = {
                "title": descriptive_asset_title(row, asset_type),
                "asset_type": asset_type,
                "asset_source_label": source,
                "asset_format": fmt,
                "asset_relationship": relationship,
                "shared_across_datasets": shared,
                "related_dataset_ids": "; ".join(sorted(other)),
                "related_dataset_titles": "; ".join(sorted(set(other.values()))),
            }
            for name, value in values.items():
                col = positions.get(norm_header(name))
                if col is None and name == "title":
                    col = positions.get("asset_title")
                if col is None:
                    continue
                cell = ws.cell(r, col)
                if name == "title":
                    if args.reclassify or not clean(cell.value): cell.value = value
                elif args.reclassify or not clean(cell.value):
                    cell.value = value
            subtype_cell = ws.cell(r, subtype_col)
            existing = clean(subtype_cell.value).casefold()
            if args.reclassify or existing not in {"data", "metadata"}:
                subtype_cell.value = legacy_subtype(asset_type)
            else:
                subtype_cell.value = existing.title()
            counts[asset_type] += 1
        ws.freeze_panes = ws.freeze_panes or "A2"
        ws.auto_filter.ref = ws.dimensions
        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path); wb.close()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr); return 1
    print(f"Output workbook: {output_path}")
    print("Asset types:")
    for label, count in counts.most_common(): print(f"  {label}: {count}")
    if not id_col:
        print("NOTE: No dataset ID column found; cross-dataset shared-link detection was skipped.")
    if not ds_title_col:
        print("NOTE: No dataset title column found; related_dataset_titles uses IDs where available.")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
