#!/usr/bin/env python3
"""Build separate CIESIN Items and Assets Excel workbooks from the inventory.

Reads the MAIN inventory, resolves each supplied DOI through doi.org, queries
the repository identified by each DOI's redirect destination, and uses only
XML files listed in the inventory. It writes two reviewable Excel workbooks
without changing the input spreadsheet. NASA Earthdata tile assets are
excluded; documentation and metadata links may be retained.

Requires openpyxl. Run from the project root or pass explicit input paths.
"""

from __future__ import annotations

import argparse
from email.message import Message
from html.parser import HTMLParser
import json
import mimetypes
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

import extract_xml_stac_item_fields as xml_fields
import classify_ciesin_asset_subtypes as asset_subtypes
from datacite_lookup import datacite_lookup


import dedupe_harvard_doi_record_assets as harvard_dedupe



from urllib.parse import quote
 
from urllib.request import Request, urlopen
from urllib.parse import quote
from urllib.error import HTTPError, URLError


ROOT = Path(__file__).resolve().parents[1]
HARVARD_URL = "https://dataverse.harvard.edu"
CMR_URL = "https://cmr.earthdata.nasa.gov/search/collections.umm_json"
CMR_GRANULES_URL = "https://cmr.earthdata.nasa.gov/search/granules.umm_json"
CMR_GRANULE_LIMIT = 500
COLUMBIA_URL = "https://academiccommons.columbia.edu"
COLUMBIA_ASSET_CACHE: dict[str, dict[str, str]] = {}
DEFAULT_INPUT = ROOT / "inputs" / "CIESIN_doi_list_INPUT1.xlsx"
DEFAULT_XML_DIR = ROOT / "inputs" / "xmls"
DEFAULT_OUTPUT_DIR = ROOT / "outputs" / "stac_workbooks"
INPUT_COLUMNS = (
    "is_latest_version", "item_title", "inventory_stac_category", "stac_nasa_doi",
    "stac_harvard_doi", "columbia_ac_doi", "inventory_xml_location",
)
DOI_COLUMNS = ("stac_nasa_doi", "stac_harvard_doi", "columbia_ac_doi")
DOI_RE = re.compile(r"10\.\d{4,9}/[^\s\"<>]+", re.I)
DOI_TRAILING = ".,;:)]}"
DOCUMENT_TERMS = (
    "documentation", "document", "user guide", "user-guide", "user_guide", "guide",
    "readme", "codebook", "manual", "methodology", "methods", "metadata", "citation",
    "publication", "paper", "article", "report", "technical", ".pdf", ".doc", ".docx",
)

ITEM_COLUMNS = [
    "dataset_id", "is_latest_version", "item_title", "resource_type", "inventory_stac_category", "tags",
    "bbox_west", "bbox_south", "bbox_east", "bbox_north", "temporal_start", "temporal_end",
    "license", "description", "providers", "citation", "version", "additional_resources",
    "purpose", "authors", "publishers", "place", "doi", "publication_date",
    "stac_nasa_doi", "stac_harvard_doi", "columbia_ac_doi", "inventory_xml_location",
    "xml_metadata_found", "metadata_sources",
]
ASSET_COLUMNS = [
    "dataset_id", "dataset_doi", "item_title", "asset_key", "href", "media_type",
    "title", "description", "roles", "source", "source_doi", "source_identifier",
    "file_size_bytes", "checksum_md5", "created", "updated", "asset_match_basis", "asset_subtype",
]


def clean(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)
    return re.sub(r"\s+", " ", str(value)).strip()


def doi_keys(value: Any) -> set[str]:
    text = clean(value)
    text = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", text, flags=re.I)
    text = re.sub(r"^doi:\s*", "", text, flags=re.I)
    return {match.rstrip(DOI_TRAILING).casefold() for match in DOI_RE.findall(text)}


def canonical_doi(value: Any) -> str:
    keys = doi_keys(value)
    return sorted(keys)[0] if keys else ""


def slug(value: str, max_len: int = 90) -> str:
    return re.sub(r"[^a-zA-Z0-9_.-]+", "-", value or "").strip("-").lower()[:max_len].strip("-") or "dataset"


def unique_values(values: list[Any]) -> list[str]:
    found: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = clean(value)
        key = text.casefold()
        if text and key not in seen:
            seen.add(key)
            found.append(text)
    return found


def normalize_resource_type(value: Any) -> str:
    """Map repository type labels to a browser-friendly controlled value."""
    text = clean(value).casefold()
    if not text:
        return ""
    mappings = {
        "article": {"article", "articles", "journal article", "scholarly article"},
        "dataset": {"dataset", "datasets", "data set", "data sets", "collection", "data collection"},
        "report": {"report", "reports", "technical report", "working paper"},
        "book": {"book", "books", "chapter", "book chapter"},
        "software": {"software", "computer program", "code"},
        "map": {"map", "maps", "cartographic material"},
    }
    for canonical, labels in mappings.items():
        if text in labels:
            return canonical
    return "other"








def crossref_lookup(doi):
    if not doi:
        return ""

    doi = str(doi).strip()
    doi = doi.removeprefix("https://doi.org/")
    doi = doi.removeprefix("http://doi.org/")
    doi = doi.removeprefix("doi:")

    url = f"https://api.crossref.org/works/{quote(doi, safe='/')}"
    request = Request(
        url,
        headers={"User-Agent": "CIESINCitationLookup/1.0"},
    )

    try:
        with urlopen(request, timeout=15) as response:
            metadata = json.loads(response.read().decode("utf-8"))["message"]
    except (HTTPError, URLError, TimeoutError, ValueError, KeyError):
        return ""

    titles = metadata.get("title", [])
    title = titles[0].strip() if titles else ""

    if not title:
        return ""

    authors = []
    for author in metadata.get("author", []):
        family = author.get("family", "")
        given = author.get("given", "")

        if family and given:
            initials = " ".join(
                f"{part[0]}."
                for part in given.replace("-", " ").split()
                if part
            )
            authors.append(f"{family}, {initials}")
        elif family:
            authors.append(family)
        elif author.get("name"):
            authors.append(author["name"])

    if len(authors) > 7:
        author_text = ", ".join(authors[:6]) + ", … " + authors[-1]
    elif len(authors) > 1:
        author_text = ", ".join(authors[:-1]) + ", & " + authors[-1]
    else:
        author_text = authors[0] if authors else ""

    date_info = (
        metadata.get("published-print")
        or metadata.get("published-online")
        or metadata.get("issued")
        or {}
    )
    date_parts = date_info.get("date-parts", [[]])
    year = str(date_parts[0][0]) if date_parts and date_parts[0] else "n.d."

    container = (metadata.get("container-title") or [""])[0]
    publisher = metadata.get("publisher", "")
    volume = metadata.get("volume", "")
    issue = metadata.get("issue", "")
    pages = metadata.get("page", "")

    parts = [
        f"{author_text} ({year})." if author_text else f"({year}).",
        f"{title}.",
    ]

    if container:
        details = container
        if volume:
            details += f", {volume}"
        if issue:
            details += f"({issue})"
        if pages:
            details += f", {pages}"
        parts.append(f"{details}.")
    elif publisher:
        parts.append(f"{publisher}.")

    parts.append(f"https://doi.org/{doi}")
    return " ".join(parts)


class _PlainTextHTMLParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        if clean(data):
            self.parts.append(clean(data))


class _ColumbiaLandingParser(HTMLParser):
    """Collect citation metadata and file links from an Academic Commons page."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.meta: dict[str, list[str]] = defaultdict(list)
        self.links: list[str] = []
        self.link_parts: list[str] = []
        self.title_parts: list[str] = []
        self.in_title = False
        self.in_link = False

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        values = {key.casefold(): value or "" for key, value in attrs}
        if tag.casefold() == "meta":
            key = values.get("name") or values.get("property") or values.get("itemprop")
            content = clean(values.get("content"))
            if key and content:
                self.meta[key.casefold()].append(content)
        elif tag.casefold() == "a" and values.get("href"):
            self.links.append(values["href"])
            self.link_parts.append("")
            self.in_link = True
        elif tag.casefold() == "title":
            self.in_title = True

    def handle_endtag(self, tag: str) -> None:
        if tag.casefold() == "title":
            self.in_title = False
        elif tag.casefold() == "a":
            self.in_link = False

    def handle_data(self, data: str) -> None:
        if self.in_title and clean(data):
            self.title_parts.append(clean(data))
        if self.in_link and clean(data) and self.link_parts:
            self.link_parts[-1] = clean(" ".join([self.link_parts[-1], data]))


def columbia_landing_page(doi: str) -> tuple[str, _ColumbiaLandingParser]:
    """Fetch the public DOI landing page HTML (metadata only, no file body)."""
    url = f"{COLUMBIA_URL}/doi/{urllib.parse.quote(doi, safe='/')}"
    request = urllib.request.Request(url, headers={
        "User-Agent": "CIESIN-STAC-inventory-builder/1.0",
        "Accept": "text/html,application/xhtml+xml;q=0.9,*/*;q=0.8",
    })
    with urllib.request.urlopen(request, timeout=45) as response:
        page_url = response.geturl()
        html = response.read().decode("utf-8", errors="replace")
    parser = _ColumbiaLandingParser()
    parser.feed(html)
    return page_url, parser


def html_to_text(value: Any) -> str:
    parser = _PlainTextHTMLParser()
    parser.feed(clean(value))
    text = " ".join(parser.parts)
    return re.sub(r"\s+([,.;:!?])", r"\1", text).strip()


def read_inventory(path: Path, sheet_name: str) -> tuple[list[dict[str, str]], str]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Worksheet {sheet_name!r} not found. Available: {', '.join(workbook.sheetnames)}")
        sheet = workbook[sheet_name]
        values = sheet.iter_rows(values_only=True)
        headers = [clean(value) for value in next(values, ())]
        missing = [column for column in INPUT_COLUMNS if column not in headers]
        if missing:
            raise ValueError(f"Worksheet {sheet.title!r} is missing columns: {', '.join(missing)}")
        rows: list[dict[str, str]] = []
        for excel_row, row_values in enumerate(values, start=2):
            row = {header: clean(row_values[index]) if index < len(row_values) else ""
                   for index, header in enumerate(headers) if header}
            if any(row.get(column) for column in INPUT_COLUMNS):
                row["_excel_row"] = str(excel_row)
                rows.append(row)
        return rows, sheet.title
    finally:
        workbook.close()


def doi_list(value: Any) -> list[str]:
    """Return every distinct DOI in a cell, preserving first occurrence."""
    text = clean(value)
    matches = [match.rstrip(DOI_TRAILING) for match in DOI_RE.findall(text)]
    if not matches and re.match(r"^10\.\d{4,9}/", text, re.I):
        matches = [text.rstrip(DOI_TRAILING)]
    deduplicated: dict[str, str] = {}
    for match in matches:
        deduplicated.setdefault(match.casefold(), match)
    return list(deduplicated.values())


def xml_file_for_row(row: dict[str, str], xml_dir: Path) -> Path | None:
    location = clean(row.get("inventory_xml_location"))
    if not location:
        return None
    candidate = Path(location)
    checks = [candidate] if candidate.is_absolute() else [xml_dir / candidate, ROOT / candidate, xml_dir / candidate.name]
    for path in checks:
        if path.exists() and path.is_file():
            return path.resolve()
    return None


def local_tag(element: ET.Element) -> str:
    return element.tag.rsplit("}", 1)[-1].casefold()


def first_text(root: ET.Element, names: set[str]) -> str:
    for element in root.iter():
        if local_tag(element) in names:
            value = clean(element.text)
            if value:
                return value
    return ""


def content_text(root: ET.Element, names: set[str]) -> str:
    """Return full descendant text for nested ArcGIS/ISO metadata values."""
    wanted = {name.casefold() for name in names}
    for element in root.iter():
        if local_tag(element) in wanted:
            value = html_to_text(" ".join(part for part in element.itertext() if clean(part)))
            if value:
                return value
    return ""


def values(root: ET.Element, names: set[str]) -> list[str]:
    result: list[str] = []
    for element in root.iter():
        if local_tag(element) in names:
            value = clean(element.text)
            if value:
                result.append(value)
    return unique_values(result)


def first_under(root: ET.Element, parent_name: str, child_name: str) -> str:
    if root is None:
        return ""
    for parent in root.iter():
        if local_tag(parent) == parent_name.casefold():
            for child in parent.iter():
                if local_tag(child) == child_name.casefold() and clean(child.text):
                    return clean(child.text)
    return ""


def normalize_date(value: str) -> str:
    value = clean(value)
    if re.fullmatch(r"\d{8}", value):
        return f"{value[:4]}-{value[4:6]}-{value[6:8]}"
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z?", value):
        return value.rstrip("Z")
    return value


def normalize_temporal_date(value: Any) -> str:
    """Normalize CMR temporal dates, shortening midnight timestamps to dates."""
    text = normalize_date(clean(value))
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}T00:00:00(?:\.0+)?Z?", text):
        return text[:10]
    return text


def cmr_temporal_bounds(umm: dict[str, Any]) -> tuple[str, str]:
    """Return the earliest and latest dates across CMR ranges and discrete dates."""
    starts: list[str] = []
    ends: list[str] = []
    extents = umm.get("TemporalExtents", [])
    for extent in extents if isinstance(extents, list) else []:
        if not isinstance(extent, dict):
            continue
        ranges = extent.get("RangeDateTimes", [])
        for date_range in ranges if isinstance(ranges, list) else []:
            if not isinstance(date_range, dict):
                continue
            start = normalize_temporal_date(date_range.get("BeginningDateTime"))
            end = normalize_temporal_date(date_range.get("EndingDateTime"))
            if start:
                starts.append(start)
            if end:
                ends.append(end)
        singles = extent.get("SingleDateTimes", [])
        for single in singles if isinstance(singles, list) else []:
            if isinstance(single, dict):
                date = normalize_temporal_date(single.get("DateTime"))
                if date:
                    starts.append(date)
                    ends.append(date)
    return min(starts) if starts else "", max(ends) if ends else ""


def cmr_citation_text(umm: dict[str, Any], doi: str = "") -> str:
    """Read free-text or structured UMM-C collection citations."""
    citations = umm.get("CollectionCitations", [])
    formatted: list[str] = []
    for citation in citations if isinstance(citations, list) else []:
        if not isinstance(citation, dict):
            continue
        free_text = unique_values(
            list_text(citation.get("OtherCitationDetails")) +
            list_text(citation.get("Citation"))
        )
        if free_text:
            formatted.extend(free_text)
            continue

        creators = list_text(citation.get("Creators") or citation.get("Creator"))
        title = clean(citation.get("Title")) or clean(umm.get("EntryTitle"))
        release_date = normalize_temporal_date(citation.get("ReleaseDate"))
        publisher = clean(citation.get("Publisher"))
        release_place = clean(citation.get("ReleasePlace"))
        version = clean(citation.get("Version")) or clean(umm.get("Version"))
        online_resource = citation.get("OnlineResource", {})
        linkage = clean(online_resource.get("Linkage")) if isinstance(online_resource, dict) else ""
        parts = []
        if creators:
            parts.append(", ".join(creators))
        if release_date:
            parts.append(release_date)
        if title:
            parts.append(title)
        if version:
            parts.append(f"Version {version}")
        if publisher:
            parts.append(publisher)
        if release_place:
            parts.append(release_place)
        if linkage or doi:
            identifier = linkage or f"https://doi.org/{doi}"
            if identifier not in " ".join(parts):
                parts.append(identifier)
        if parts:
            formatted.append(". ".join(parts))
    return "; ".join(unique_values(formatted))


def xml_temporal(root: ET.Element) -> tuple[str, str]:
    for temporal in root.iter():
        if local_tag(temporal) != "timeperd":
            continue
        for container in temporal.iter():
            if local_tag(container) == "rngdates":
                start = first_text(container, {"begdate", "beginningdatetime"})
                end = first_text(container, {"enddate", "endingdatetime"})
                if start or end:
                    return normalize_date(start), normalize_date(end)
            if local_tag(container) == "sngdate":
                date = first_text(container, {"caldate", "datetime"})
                if date:
                    return normalize_date(date), normalize_date(date)
    return "", ""


def primary_citation(root: ET.Element) -> ET.Element | None:
    for idinfo in root.iter():
        if local_tag(idinfo) not in {"idinfo", "dataidinfo"}:
            continue
        for child in idinfo:
            if local_tag(child) in {"citation", "idcitation"}:
                for nested in child.iter():
                    if local_tag(nested) in {"citeinfo", "idcitation"}:
                        return nested
    for element in root.iter():
        if local_tag(element) == "idcitation":
            return element
    return None


def direct_text(parent: ET.Element | None, name: str) -> str:
    if parent is None:
        return ""
    for child in parent.iter():
        if local_tag(child) == name.casefold():
            value = html_to_text(" ".join(part for part in child.itertext() if clean(part)))
            if value:
                return value
    return ""


def xml_details(path: Path) -> dict[str, Any]:
    root = ET.parse(path).getroot()
    meta = xml_fields.xml_metadata(path)
    citeinfo = primary_citation(root)
    origin = direct_text(citeinfo, "origin") or "; ".join(meta.get("authors", []))
    pubdate = direct_text(citeinfo, "pubdate")
    edition = direct_text(citeinfo, "edition") or direct_text(citeinfo, "resEd")
    cite_title = direct_text(citeinfo, "title") or direct_text(citeinfo, "resTitle")
    publisher = first_under(citeinfo, "pubinfo", "publish")
    place_published = first_under(citeinfo, "pubinfo", "pubplace")
    primary_doi = direct_text(citeinfo, "onlink")
    citation_detail = direct_text(citeinfo, "otherCitDet") or direct_text(citeinfo, "othercit")
    citation = citation_detail or ". ".join(part for part in [
        origin, normalize_date(pubdate), cite_title,
        f"Edition {edition}" if edition else "", publisher, place_published,
    ] if clean(part))
    abstract_markup = content_text(root, {"abstract", "idabs"})
    abstract = html_to_text(abstract_markup)
    citation_from_description = ""
    if abstract:
        citation_match = re.search(
            r"\b(?:Recommended\s+)?(?:Dataset\s+)?Citation\b\s*:?[ \t]*(.*?)(?=\b(?:Terms of use|Data license|Contacts and data queries|Copyright)\b|$)",
            abstract, flags=re.I | re.S,
        )
        if citation_match:
            citation_from_description = clean(citation_match.group(1))
            abstract = clean(abstract[:citation_match.start()])
        abstract = re.split(r"\b(?:Terms of use|Data license|Contacts and data queries)\b", abstract,
                            maxsplit=1, flags=re.I)[0].strip()
    citation = citation or citation_from_description
    citation = re.sub(
        r"Dataset citation:\s*",
        "",
        citation,
        flags=re.IGNORECASE,
    ).strip()
    use_constraints = [content_text(root, {"useLimit", "useconst", "accconst"})]
    license_value = " | ".join(use_constraints)
    if not license_value:
        license_value = ""
    places = values(root, {"placekey", "descgeog"})
    # Keep these scoped to the primary citation. Whole-document searches also
    # pick up people and publishers cited by related publications in the XML.
    publishers = [publisher] if publisher else []
    authors = [origin] if origin else []
    purpose = first_text(root, {"purpose", "idpurp"})
    version = edition
    temporal_start, temporal_end = xml_temporal(root)
    bbox = meta.get("bbox") or []
    bbox = bbox if isinstance(bbox, list) else []
    return {
        "title": clean(meta.get("title")) or cite_title,
        "resource_type": normalize_resource_type(first_text(root, {"mdHrLvName", "resourcetype", "hierarchylevel"})),
        "description": abstract or clean(meta.get("description")),
        "purpose": html_to_text(purpose) or html_to_text(meta.get("purpose")),
        "citation": citation,
        "license": clean(meta.get("license")) or license_value,
        "authors": unique_values(authors + (meta.get("authors") or [])),
        "publishers": unique_values(publishers + (meta.get("publishers") or [])),
        "places": places,
        "tags": meta.get("keywords") or [],
        "publication_date": normalize_date(pubdate) or clean(meta.get("publication_date")),
        "version": version,
        "primary_doi": canonical_doi(primary_doi or citation or abstract),
        "bbox": bbox,
        "temporal_start": temporal_start,
        "temporal_end": temporal_end,
        "data_links": meta.get("data_links") or [],
        "metadata_links": meta.get("metadata_links") or [],
        "all_links": meta.get("all_links") or [],
        "xml_path": str(path),
    }


def list_text(value: Any, preferred_keys: tuple[str, ...] = ()) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [clean(value)] if clean(value) else []
    if isinstance(value, dict):
        for key in preferred_keys:
            if key in value:
                return list_text(value[key], preferred_keys)
        return [clean(v) for v in value.values() if clean(v)]
    if isinstance(value, list):
        out = []
        for item in value:
            if isinstance(item, dict):
                for key in preferred_keys:
                    if item.get(key):
                        out.extend(list_text(item[key], preferred_keys))
                        break
                else:
                    out.extend(clean(v) for v in item.values() if isinstance(v, str) and clean(v))
            elif clean(item):
                out.append(clean(item))
        return unique_values(out)
    return [clean(value)]


def fetch_json(url: str, headers: dict[str, str] | None = None) -> dict[str, Any]:
    request_headers = {"User-Agent": "CIESIN-STAC-inventory-builder/1.0", "Accept": "application/json"}
    request_headers.update(headers or {})
    request = urllib.request.Request(url, headers=request_headers)
    with urllib.request.urlopen(request, timeout=60) as response:
        return json.loads(response.read().decode("utf-8", errors="replace"))


def resolve_doi_url(doi: str) -> str:
    """Follow doi.org redirects and return the final repository landing URL."""
    resolver_url = "https://doi.org/" + urllib.parse.quote(doi, safe="/")
    headers = {"User-Agent": "CIESIN-STAC-inventory-builder/1.0", "Accept": "text/html,*/*;q=0.8"}
    request = urllib.request.Request(resolver_url, method="HEAD", headers=headers)
    try:
        with urllib.request.urlopen(request, timeout=45) as response:
            return response.geturl()
    except urllib.error.HTTPError as exc:
        if exc.code not in {403, 405, 501}:
            raise
    # Some DOI targets reject HEAD. A GET fallback follows the same redirects;
    # closing without reading avoids downloading the landing page body.
    request = urllib.request.Request(resolver_url, headers={**headers, "Range": "bytes=0-0"})
    with urllib.request.urlopen(request, timeout=45) as response:
        return response.geturl()


def nested_field(fields: Any, type_name: str) -> Any:
    if isinstance(fields, dict):
        if fields.get("typeName") == type_name:
            return fields.get("value", "")
        for value in fields.values():
            found = nested_field(value, type_name)
            if found not in (None, "", []):
                return found
    elif isinstance(fields, list):
        for value in fields:
            found = nested_field(value, type_name)
            if found not in (None, "", []):
                return found
    return ""


def dataverse_lookup(doi: str) -> dict[str, Any]:
    persistent_id = urllib.parse.quote("doi:" + doi, safe="")
    url = f"{HARVARD_URL}/api/datasets/:persistentId/?persistentId={persistent_id}"
    payload = fetch_json(url)
    if payload.get("status") != "OK":
        raise RuntimeError(f"Harvard Dataverse API returned status {payload.get('status')!r}")
    data = payload.get("data", {})
    version = data.get("latestVersion", {}) if isinstance(data, dict) else {}
    fields = version.get("metadataBlocks", {}).get("citation", {}).get("fields", [])
    title = nested_field(fields, "title")
    description_values = list_text(nested_field(fields, "dsDescription"), ("dsDescriptionValue", "value"))
    description = html_to_text("\n\n".join(description_values))
    authors = nested_field(fields, "author")
    author_names = list_text(authors, ("authorName", "value"))
    publisher = nested_field(fields, "producer") or nested_field(fields, "publisher")
    publisher_names = list_text(publisher, ("producerName", "publisherName", "name", "value"))
    citation_date = nested_field(fields, "publicationDate")
    subject = nested_field(fields, "subject")
    keywords = nested_field(fields, "keyword")
    subjects = unique_values(list_text(subject, ("value", "subjectValue")) +
                             list_text(keywords, ("keywordValue", "value")))
    license_info = version.get("license", "")
    if isinstance(license_info, dict):
        license_value = clean(license_info.get("name"))
        if license_value.casefold() in {"custom", "other"}:
            custom_terms = license_info.get("customTerms", "")
            license_value = clean(custom_terms.get("termsOfUse") or custom_terms.get("termsOfAccess")
                                  if isinstance(custom_terms, dict) else custom_terms) or license_value
        if not license_value:
            license_value = clean(license_info.get("uri"))
    else:
        license_value = clean(license_info)
    release_date = normalize_date(clean(version.get("releaseTime"))) or normalize_date(clean(citation_date))
    if re.match(r"^\d{4}-\d{2}-\d{2}", release_date):
        release_date = release_date[:10]
    persistent_url = clean(data.get("persistentUrl")) or f"https://doi.org/{doi}"
    citation = ""
    try:
        citation_path = f"{HARVARD_URL}/api/datasets/:persistentId/versions/:latest/citation?persistentId={persistent_id}"
        citation_payload = fetch_json(citation_path)
        citation = html_to_text(citation_payload.get("data", {}).get("message", ""))
    except Exception:
        pass
    if not citation:
        citation = ", ".join(part for part in (
            "; ".join(author_names), release_date[:4] if release_date else "",
            f'"{clean(title)}"' if clean(title) else "", persistent_url,
            "Harvard Dataverse", f"V{clean(version.get('versionNumber'))}" if clean(version.get("versionNumber")) else "",
        ) if part)
    if not citation and doi: 
        citation = crossref_lookup(doi)
    files: list[dict[str, Any]] = []
    for file_record in version.get("files", []) if isinstance(version, dict) else []:
        data_file = file_record.get("dataFile", {}) if isinstance(file_record, dict) else {}
        file_id = clean(data_file.get("id"))
        if file_id:
            files.append({
                "href": f"{HARVARD_URL}/file.xhtml?fileId={file_id}",
                "title": clean(data_file.get("filename") or file_record.get("label")),
                "description": clean(file_record.get("description")),
                "media_type": clean(data_file.get("contentType")),
                "size": clean(data_file.get("filesize")),
                "source_identifier": file_id,
                "source": "Harvard Dataverse file",
            })
    return {
        "source": "Harvard Dataverse", "doi": doi,
        "resource_type": "dataset",
        "title": clean(title), "description": clean(description),
        "authors": "; ".join(author_names), "publishers": "; ".join(publisher_names),
        "publication_date": release_date,
        "version": clean(version.get("versionNumber")),
        "license": license_value,
        "citation": citation,
        "provider": "Harvard Dataverse",
        "url": f"{HARVARD_URL}/dataset.xhtml?persistentId={persistent_id}",
        "tags": subjects, "assets": files,
    }


def cmr_lookup(doi: str) -> dict[str, Any]:
    query = urllib.parse.urlencode([("doi[]", doi), ("page_size", "100")])
    payload = fetch_json(f"{CMR_URL}?{query}", {
        "Accept": "application/vnd.nasa.cmr.umm_results+json",
        "Client-Id": "CIESIN-STAC-inventory-builder",
    })
    items = payload.get("items", [])
    matching = []
    for record in items:
        umm = record.get("umm", {})
        doi_obj = umm.get("DOI", {})
        result_doi = clean(doi_obj.get("DOI")) if isinstance(doi_obj, dict) else ""
        if doi.casefold() in doi_keys(result_doi):
            matching.append(record)
    if not matching:
        raise LookupError("CMR returned no collection with this exact DOI")
    # DOI is exact, but migrated SEDAC records can coexist under legacy and
    # current providers. Prefer the active ESDIS collection when available.
    matching.sort(key=lambda record: (
        clean(record.get("meta", {}).get("provider-id")).casefold() == "esdis",
        clean(record.get("meta", {}).get("revision-date")),
    ), reverse=True)
    record = matching[0]
    umm, meta = record.get("umm", {}), record.get("meta", {})
    concept_id = clean(meta.get("concept-id"))
    abstracts = list_text(umm.get("Abstract"), ("Description", "Abstract"))
    centers = umm.get("DataCenters", [])
    providers: list[str] = []
    for center in centers if isinstance(centers, list) else []:
        if isinstance(center, dict):
            providers.extend(list_text(center.get("Organizations"), ("LongName", "ShortName", "Name")))
    citations = umm.get("CollectionCitations", [])
    authors, publishers = [], []
    for cite in citations if isinstance(citations, list) else []:
        if isinstance(cite, dict):
            authors.extend(list_text(cite.get("Creators") or cite.get("Creator")))
            publishers.extend(list_text(cite.get("Publisher")))
    # UMM-C stores publication dates in PublicationReferences. Some legacy
    # CIESIN collections instead provide a citation release date or a CREATE
    # date, which is what Earthdata's landing page labels as Published.
    publication_date = ""
    publication_references = umm.get("PublicationReferences", [])
    for reference in publication_references if isinstance(publication_references, list) else []:
        if isinstance(reference, dict):
            publication_date = publication_date or normalize_date(clean(reference.get("PublicationDate")))
    if not publication_date:
        for cite in citations if isinstance(citations, list) else []:
            if isinstance(cite, dict):
                publication_date = normalize_date(clean(cite.get("ReleaseDate")))
                if publication_date:
                    break
    for date_group in (umm.get("DataDates", []), umm.get("MetadataDates", [])):
        if publication_date:
            break
        for date_record in date_group if isinstance(date_group, list) else []:
            if not isinstance(date_record, dict):
                continue
            if clean(date_record.get("Type")).casefold() == "create":
                publication_date = normalize_date(clean(date_record.get("Date")))
                if publication_date:
                    break
    if not publication_date:
        publication_date = normalize_date(clean(umm.get("PublicationDate")))
    start, end = cmr_temporal_bounds(umm)
    bbox: list[float] = []
    try:
        rectangles = umm["SpatialExtent"]["HorizontalSpatialDomain"]["Geometry"]["BoundingRectangles"]
        if rectangles:
            bbox = [float(rectangles[0][key]) for key in (
                "WestBoundingCoordinate", "SouthBoundingCoordinate",
                "EastBoundingCoordinate", "NorthBoundingCoordinate")]
    except (KeyError, TypeError, ValueError):
        pass


    ## build assets from NASA CMR collection and granule related URLs, excluding Earthdata tile/granule links
    
    assets: list[dict[str, Any]] = []

    # --------------------------------------------------
    # 1. Link to all granules for this collection
    # --------------------------------------------------
    if concept_id:
        granule_search_url = (
            "https://search.earthdata.nasa.gov/search/granules?"
            + urllib.parse.urlencode({"p": concept_id})
        )

        assets.append({
            "href": granule_search_url,
            "title": "NASA Earthdata Search: View all granules",
            "description": (
                "Search, discover, and access all granules "
                "associated with this dataset in NASA Earthdata Search."
            ),
            "media_type": "text/html",
            "source": "NASA CMR collection",
            "source_identifier": concept_id,
            "roles": "metadata",
        })

    # --------------------------------------------------
    # 2. Preserve collection RelatedUrls
    # --------------------------------------------------
    related_urls = umm.get("RelatedUrls", [])

    for related in related_urls if isinstance(related_urls, list) else []:
        if not isinstance(related, dict):
            continue

        href = clean(related.get("URL"))
        if not href:
            continue

        relation = clean(related.get("Type"))
        subtype = clean(related.get("Subtype"))
        description = clean(related.get("Description"))
        media_type = clean(related.get("MimeType"))

        related_description = (
            f"Related URL: {href}\n"
            f"Type: {relation}"
        )

        if subtype:
            related_description += f"\nSubtype: {subtype}"

        if description:
            related_description += f"\nDescription: {description}"

        asset = {
            "href": href,
            "title": f"Related URL: {href}",
            "description": related_description,
            "media_type": media_type,
            "source": "NASA CMR collection related URL",
            "roles": "metadata",
            "source_identifier": concept_id,
            "source_relation_type": relation,
            "source_relation_subtype": subtype,
        }

        if ( 
            not is_earthdata_tile_or_granule(asset)
            and "www.earthdata.nasa.gov/about/esdis" not in href.casefold() ):
            assets.append(asset)

    return {
        "source": "NASA CMR", "doi": doi, "title": clean(umm.get("EntryTitle")),
        "resource_type": "dataset",
        "description": "; ".join(abstracts), "authors": "; ".join(unique_values(authors)),
        "publishers": "; ".join(unique_values(publishers)), "provider": "; ".join(unique_values(providers)) or "NASA CMR / ESDIS",
        "publication_date": publication_date, "version": clean(umm.get("Version")),
        "citation": cmr_citation_text(umm, doi), "url": f"https://cmr.earthdata.nasa.gov/search/concepts/{concept_id}" if concept_id else "",
        "temporal_start": start, "temporal_end": end, "bbox": bbox,
        "tags": list_text(umm.get("ScienceKeywords"), ("Category", "Topic", "Term", "VariableLevel1", "VariableLevel2", "VariableLevel3", "DetailedVariable")),
        "assets": assets,
        # "asset_lookup_warning": granule_lookup_warning,
    }


def recursive_values(value: Any, names: set[str]) -> list[str]:
    found: list[str] = []
    if isinstance(value, dict):
        for key, child in value.items():
            if key.casefold() in names:
                found.extend(list_text(child, ("name", "value", "text", "display")))
            else:
                found.extend(recursive_values(child, names))
    elif isinstance(value, list):
        for child in value:
            found.extend(recursive_values(child, names))
    return unique_values(found)


def remove_duplicate_harvard_record_assets(assets: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], int]:
    """Drop Harvard landing-page assets when the same dataset DOI asset exists."""
    doi_assets: dict[tuple[str, str], str] = {}
    for asset in assets:
        doi = harvard_dedupe.doi_asset_doi(asset)
        if doi:
            for dataset_key in harvard_dedupe.dataset_keys(asset):
                doi_assets[(dataset_key, doi)] = asset.get("href", "")

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



def columbia_lookup(doi: str, item_title: str = "") -> dict[str, Any]:
    # Columbia's route expects the slash within a DOI to remain a path
    # separator; percent-encoding it returns a misleading "Record not found".
    route_doi = urllib.parse.quote(doi.casefold(), safe="/")
    url = f"{COLUMBIA_URL}/api/v1/record/doi/{route_doi}?format=json"
    record: dict[str, Any] = {}
    record_error = ""
    try:
        record = fetch_json(url)
        if record.get("error"):
            record_error = clean(record["error"])
    except Exception as exc:
        record_error = str(exc)
    landing_url = ""
    landing: _ColumbiaLandingParser | None = None
    if record_error or not record.get("title"):
        search_error = ""
        exact_matches = []
        if item_title:
            search_url = f"{COLUMBIA_URL}/api/v1/search?" + urllib.parse.urlencode({
                "search_type": "title", "q": item_title, "per_page": "100",
            })
            try:
                search_payload = fetch_json(search_url)
                for candidate in search_payload.get("records", []):
                    candidate_dois = doi_keys(clean(candidate.get("id")) + " " + clean(candidate.get("persistent_url")))
                    if doi.casefold() in candidate_dois:
                        exact_matches.append(candidate)
            except Exception as exc:
                search_error = str(exc)
        if len(exact_matches) == 1:
            record = exact_matches[0]
        else:
            # The public record API may not index older DOI records. Their
            # landing pages still expose citation metadata, including
            # citation_pdf_url, which is the authoritative link to the PDF.
            try:
                landing_url, landing = columbia_landing_page(doi)
            except Exception as exc:
                landing_error = str(exc)
                raise LookupError(
                    f"Academic Commons API failed ({record_error or 'empty response'}); "
                    f"title search produced {len(exact_matches)} exact DOI matches"
                    f"{': ' + search_error if search_error else ''}; "
                    f"landing page failed: {landing_error}"
                ) from exc
            page_title = " ".join(landing.title_parts).strip()
            page_title = re.sub(r"\s*\|\s*Academic Commons\s*$", "", page_title, flags=re.I)
            meta = landing.meta
            page_record_title = next(iter(meta.get("citation_title", [])), "") or page_title or item_title
            if not page_record_title:
                raise LookupError(
                    f"Academic Commons API failed ({record_error or 'empty response'}); "
                    f"title search produced {len(exact_matches)} exact DOI matches; "
                    "landing page did not expose a title"
                )
            record = {
                "title": page_record_title,
                "description": next(iter(meta.get("citation_abstract", []) + meta.get("description", [])), ""),
                "author": meta.get("citation_author", []),
                "publisher": meta.get("citation_publisher", []),
                "publication_date": next(iter(meta.get("citation_publication_date", []) + meta.get("citation_date", [])), ""),
                "persistent_url": landing_url,
                "id": doi,
            }
    titles = recursive_values(record, {"title", "item_title", "title_display"})
    descriptions = recursive_values(record, {"description", "abstract", "summary"})
    authors = recursive_values(record, {"author", "authors", "creator", "creators"})
    publishers = recursive_values(record, {"publisher", "publishers"})
    departments = recursive_values(record, {"department"})
    dates = recursive_values(record, {"publication_date", "date_issued", "issued", "published", "date"})
    versions = recursive_values(record, {"version", "version_number"})
    resource_types = recursive_values(record, {"type", "genre", "resource_type", "resource_type_ssim"})
 
    resource_paths = recursive_values(record, {"resource_paths"})
    subjects = recursive_values(record, {"subject", "columbia_series", "keywords"})
    notes = record.get("notes", "")
    notes = notes if isinstance(notes, str) else clean(notes)
    citation_match = re.search(r"Dataset citation:\s*(.*?)(?:\n\s*\n|$)", notes, flags=re.I | re.S)
    license_match = re.search(r"Data license:\s*(.*?)(?:\n\s*\n|$)", notes, flags=re.I | re.S)
    citation = clean(citation_match.group(1)) if citation_match else ""
    license_value = clean(license_match.group(1)) if license_match else ""
    record_url = clean(record.get("persistent_url"))
    if not record_url:
        record_url = f"{COLUMBIA_URL}/detail/{clean(record.get('legacy_id'))}" if clean(record.get("legacy_id")) else f"https://doi.org/{doi}"
    assets = []
    for index, path in enumerate(resource_paths, start=1):
        href = urllib.parse.urljoin(COLUMBIA_URL, path)
        file_metadata = columbia_asset_metadata(href)
        path_title = Path(urllib.parse.urlparse(href).path).name
        asset_title = file_metadata.get("title") or path_title
        if asset_title.casefold() in {"download", "file"}:
            asset_title = f"{clean(record.get('title'))} — file {index}"
        assets.append({"href": href, "title": asset_title,
                       "media_type": file_metadata.get("media_type", ""),
                       "size": file_metadata.get("size", ""),
                       "source_identifier": path,
                       "source": "Columbia Academic Commons file", "roles": "data"})
    # Some records (including older articles) omit resource_paths from the
    # API response but expose files through citation metadata or direct file
    # links on the DOI page. Probe only those declared/file-like URLs; do not
    # crawl the rest of the site.
    if landing is None:
        try:
            landing_url, landing = columbia_landing_page(doi)
        except Exception:
            landing = None
    if landing is not None:
        resource_urls: list[str] = []
        for key, values in landing.meta.items():
            if key.startswith("citation_") and any(term in key for term in (
                    "pdf_url", "data_url", "xml_url", "supplementary_material",
                    "fulltext_world_readable", "fulltext_html_url", "fulltext_pdf")):
                resource_urls.extend(values)
        file_extensions = {
            ".pdf", ".zip", ".gz", ".tar", ".csv", ".tsv", ".txt", ".json", ".xml",
            ".kml", ".kmz", ".geojson", ".gpkg", ".shp", ".shx", ".dbf", ".prj",
            ".tif", ".tiff", ".nc", ".nc4", ".grd", ".asc",
            ".hdf", ".h5", ".xlsx", ".xls", ".doc", ".docx", ".ppt", ".pptx",
            ".jpg", ".jpeg", ".png", ".jp2", ".mp4", ".mp3", ".zipx",
        }
        for index, link in enumerate(landing.links):
            parsed_link = urllib.parse.urlparse(link)
            path = parsed_link.path.casefold()
            label = landing.link_parts[index].casefold() if index < len(landing.link_parts) else ""
            filename = Path(path).name
            if (Path(filename).suffix in file_extensions
                    or re.search(r"/(?:download|files?|attachments?)(?:/|$)", path)
                    or re.search(r"\b(download|supplementary|dataset file|resource file)\b", label)):
                resource_urls.append(link)
        declared_urls = {urllib.parse.urljoin(landing_url or record_url, value).casefold()
                         for value in resource_urls}
        existing_hrefs = {clean(asset.get("href")).casefold() for asset in assets}
        for raw_href in unique_values(resource_urls):
            href = urllib.parse.urljoin(landing_url or record_url, raw_href)
            if not href or href.casefold() in existing_hrefs:
                continue
            file_metadata = columbia_asset_metadata(href)
            parsed_path = urllib.parse.urlparse(href).path
            media_type = clean(file_metadata.get("media_type"))
            disposition = clean(file_metadata.get("disposition")).casefold()
            extension = Path(urllib.parse.unquote(parsed_path)).suffix.casefold()
            is_file = (href.casefold() in declared_urls or extension in file_extensions or "attachment" in disposition
                       or (media_type and media_type not in {"text/html", "application/xhtml+xml"}))
            if not is_file:
                continue
            filename = file_metadata.get("title") or Path(urllib.parse.unquote(parsed_path)).name
            if re.search(r"(?:^|[._-])(?:logo|icon|favicon|social)(?:[._-]|$)", filename, re.I):
                continue
            if not media_type or media_type in {"text/html", "application/xhtml+xml"}:
                media_type = mimetypes.guess_type(filename or parsed_path)[0] or "application/octet-stream"
            assets.append({
                "href": href, "title": filename or f"{clean(record.get('title'))} — file",
                "media_type": media_type, "size": file_metadata.get("size", ""),
                "source_identifier": urllib.parse.urlparse(href).path,
                "source": "Columbia Academic Commons file", "roles": "data",
            })
            existing_hrefs.add(href.casefold())
    return {
        "source": "Columbia Academic Commons", "doi": doi,
        "resource_type": next((normalized for value in resource_types
                                if (normalized := value)), ""),
        "title": titles[0] if titles else "", "description": "; ".join(descriptions),
        "authors": "; ".join(authors), "publishers": "; ".join(publishers),
        "publication_date": dates[0] if dates else "",
        "version": "; ".join(versions) or (re.search(r"\bv(\d+(?:\.\d+)+)\b", titles[0], re.I).group(1)
                                                 if titles and re.search(r"\bv(\d+(?:\.\d+)+)\b", titles[0], re.I) else ""),
        "license": license_value, "citation": citation,
        "provider": "; ".join(departments) or "Columbia Academic Commons",
        "url": record_url, "tags": subjects, "assets": assets,
    }


def columbia_asset_metadata(href: str) -> dict[str, str]:
    """Read the served filename/type/size without downloading the asset body."""
    if href in COLUMBIA_ASSET_CACHE:
        return COLUMBIA_ASSET_CACHE[href]
    metadata = {"title": "", "media_type": "", "size": "", "disposition": ""}
    request = urllib.request.Request(href, method="HEAD", headers={
        "User-Agent": "CIESIN-STAC-inventory-builder/1.0",
        "Accept": "*/*",
    })
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            disposition = response.headers.get("Content-Disposition", "")
            metadata["disposition"] = disposition
            if disposition:
                parsed = Message()
                parsed["content-disposition"] = disposition
                filename = parsed.get_filename()
                if filename:
                    metadata["title"] = Path(urllib.parse.unquote(filename)).name.strip()
            metadata["media_type"] = response.headers.get_content_type() or ""
            metadata["size"] = response.headers.get("Content-Length", "")
    except Exception:
        # Preserve the useful resource URL fallback when a server disallows HEAD.
        pass
    COLUMBIA_ASSET_CACHE[href] = metadata
    time.sleep(0.2)
    return metadata


SOURCE_LOOKUPS = {
    "stac_harvard_doi": dataverse_lookup,
    "stac_nasa_doi": cmr_lookup,
    "columbia_ac_doi": columbia_lookup,
}


def source_for_resolved_url(url: str, fallback_column: str) -> tuple[str, Any]:
    host = urllib.parse.urlparse(url).hostname or ""
    host = host.casefold()
    if host == "dataverse.harvard.edu" or host.endswith(".dataverse.harvard.edu"):
        return "Harvard Dataverse", dataverse_lookup
    if "academiccommons.columbia.edu" == host or host.endswith(".academiccommons.columbia.edu"):
        return "Columbia Academic Commons", columbia_lookup
    if host.endswith("nasa.gov") or host.endswith("earthdata.nasa.gov"):
        return "NASA CMR", cmr_lookup
    labels = {
        "stac_harvard_doi": ("Harvard Dataverse", dataverse_lookup),
        "stac_nasa_doi": ("NASA CMR", cmr_lookup),
        "columbia_ac_doi": ("Columbia Academic Commons", columbia_lookup),
    }
    return labels[fallback_column]


def lookup_inventory_dois(inventory: list[dict[str, str]], pause_seconds: float) -> tuple[dict[str, list[dict[str, Any]]], list[tuple[str, str]]]:
    """Resolve and look up every DOI in every row against its destination."""
    by_row: dict[str, list[dict[str, Any]]] = defaultdict(list)
    log: list[tuple[str, str]] = []
    lookup_count = 0
    for row in inventory:
        row_num = row.get("_excel_row", "")
        for column, lookup in SOURCE_LOOKUPS.items():
            for doi in doi_list(row.get(column, "")):
                lookup_count += 1
                resolved_url = ""
                try:
                    resolved_url = resolve_doi_url(doi)
                    log.append((f"Row {row_num} DOI resolver {doi}", resolved_url))
                except Exception as exc:
                    log.append((f"Row {row_num} DOI resolver {doi}", f"redirect lookup failed: {exc}"))
                source, lookup = source_for_resolved_url(resolved_url, column)
                try:
                    record = lookup(doi, row.get("item_title", "")) if source == "Columbia Academic Commons" else lookup(doi)
                    record["input_column"] = column
                    record["resolved_url"] = resolved_url
                    if not record.get("url"):
                        record["url"] = resolved_url
                    by_row[row_num].append(record)
                    log.append((f"Row {row_num} {source} {doi}", "matched"))
                except Exception as exc:
                    log.append((f"Row {row_num} {source} {doi}", f"lookup failed: {exc}"))
                if pause_seconds:
                    time.sleep(pause_seconds)
    log.insert(0, ("Per-row DOI lookups attempted", str(lookup_count)))
    log.insert(1, ("Per-row DOI lookups matched", str(sum(len(records) for records in by_row.values()))))
    return by_row, log


def match_xml(item: dict[str, str], xml_records: list[dict[str, Any]], xml_by_path: dict[str, dict[str, Any]]) -> dict[str, Any] | None:
    location = clean(item.get("inventory_xml_location"))
    if location:
        normalized = location.replace("\\", "/").casefold().lstrip("./")
        if normalized in xml_by_path:
            return xml_by_path[normalized]
        key = Path(location).name.casefold()
        basename_matches = [record for record in xml_records
                            if Path(record["xml_location"]).name.casefold() == key]
        if len(basename_matches) == 1:
            return basename_matches[0]
    # Do not borrow metadata from another row's XML just because its title or
    # DOI happens to match. The inventory's XML location is the provenance.
    return None


def make_items(
        inventory: list[dict[str, str]],
        xml_dir: Path,
        remote_by_row: dict[str, list[dict[str, Any]]],
        log: list[tuple[str, str]],
        allow_online: bool = True,
    ) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, str]]:
    xml_records: list[dict[str, Any]] = []
    xml_errors = []
    listed_xml_paths = sorted({
        path for row in inventory
        if (path := xml_file_for_row(row, xml_dir)) is not None
    })
    for path in listed_xml_paths:
        try:
            extra = xml_details(path)
            try:
                extra["xml_location"] = str(path.relative_to(xml_dir))
            except ValueError:
                extra["xml_location"] = str(path)
            xml_records.append(extra)
        except Exception as exc:
            xml_errors.append(f"{path.name}: {exc}")
    xml_by_path: dict[str, dict[str, Any]] = {}
    for record in xml_records:
        relative_key = record["xml_location"].replace("\\", "/").casefold().lstrip("./")
        absolute_key = record["xml_path"].replace("\\", "/").casefold()
        xml_by_path[relative_key] = record
        xml_by_path[absolute_key] = record
    log.extend([("XML files listed in inventory and parsed", str(len(xml_records))),
                ("XML locations not found", str(sum(bool(clean(row.get("inventory_xml_location"))) and
                                                     xml_file_for_row(row, xml_dir) is None for row in inventory))),
                ("XML parse errors", str(len(xml_errors)))])
    if xml_errors:
        log.append(("XML error samples", " | ".join(xml_errors[:20])))

    # Keep identifiers unique for separate spreadsheet rows even if an old
    # inventory contains a DOI repeated across product records.
    doi_counts: dict[str, int] = defaultdict(int)
    for row in inventory:
        primary = next((doi for column in DOI_COLUMNS for doi in doi_list(row.get(column, ""))), "")
        if primary:
            doi_counts[primary.casefold()] += 1
    items: list[dict[str, Any]] = []
    item_context: list[dict[str, Any]] = []
    for row in inventory:
        xml_record = match_xml(row, xml_records, xml_by_path)
        remote_matches = remote_by_row.get(row.get("_excel_row", ""), [])
        title = clean(row.get("item_title")) or clean(xml_record.get("title") if xml_record else "") or next(
            (clean(match.get("title")) for match in remote_matches if clean(match.get("title"))), ""
        )
        source_dois = {column: "; ".join(doi_list(row.get(column, ""))) for column in DOI_COLUMNS}
        all_dois = unique_values([doi for column in DOI_COLUMNS for doi in doi_list(row.get(column, ""))])
        xml_doi = clean(xml_record.get("primary_doi")) if xml_record else ""
        all_dois = unique_values(all_dois + doi_list(xml_doi))
        primary_doi = next((doi for column in DOI_COLUMNS for doi in doi_list(row.get(column, ""))), "") or xml_doi
        dataset_id = primary_doi if primary_doi and doi_counts[primary_doi.casefold()] == 1 else (
            (primary_doi + "--" + slug(title or f"row-{row['_excel_row']}")) if primary_doi else slug(title or f"row-{row['_excel_row']}")
        )

        descriptions = [xml_record.get("description", "")] if xml_record else []
        descriptions += [match.get("description", "") for match in remote_matches]
        description = next((clean(value) for value in descriptions if clean(value)), "")
        authors = unique_values((xml_record.get("authors", []) if xml_record else []) +
                                [value for match in remote_matches for value in list_text(match.get("authors"))])
        publishers = unique_values((xml_record.get("publishers", []) if xml_record else []) +
                                   [value for match in remote_matches for value in list_text(match.get("publishers"))])
        providers = unique_values((xml_record.get("publishers", []) if xml_record else []) +
                                  [match.get("provider", "") for match in remote_matches] +
                                  [match["source"] for match in remote_matches])
        versions = unique_values(([xml_record.get("version", "")] if xml_record else []) +
                                 [match.get("version", "") for match in remote_matches])
        temporal_start = clean(xml_record.get("temporal_start")) if xml_record else ""
        temporal_end = clean(xml_record.get("temporal_end")) if xml_record else ""
        bbox = xml_record.get("bbox", []) if xml_record else []
        for match in remote_matches:
            if not temporal_start and match.get("temporal_start"):
                temporal_start = normalize_date(clean(match["temporal_start"]))
            if not temporal_end and match.get("temporal_end"):
                temporal_end = normalize_date(clean(match["temporal_end"]))
            if not bbox and match.get("bbox"):
                bbox = match["bbox"]
        bbox = bbox if isinstance(bbox, list) else []
        while len(bbox) < 4:
            bbox.append("")
        latest = clean(row.get("is_latest_version"))
        is_latest = latest.casefold() in {"yes", "y", "true", "1", "latest"}
        source_tags = ((xml_record.get("tags", []) if xml_record else []) +
                       [tag for match in remote_matches for tag in match.get("tags", [])])
        if not is_latest:
            # The spreadsheet's explicit version flag is authoritative; an
            # XML keyword named "latest" must not contradict a No value.
            source_tags = [tag for tag in source_tags if clean(tag).casefold() != "latest"]
        tags = unique_values(source_tags +
                             ([row.get("inventory_stac_category", "")] if row.get("inventory_stac_category") else []) +
                             (["latest"] if is_latest else []))
        xml_links = xml_record.get("all_links", []) if xml_record else []
        resources = unique_values([
            link if isinstance(link, str) else link.get("href", "")
            for link in xml_links
            if isinstance(link, (str, dict))
        ])
        resources += [clean(match.get("url")) for match in remote_matches if clean(match.get("url"))]
        resources += [clean(asset.get("href")) for match in remote_matches
                      for asset in match.get("assets", []) if isinstance(asset, dict)]
        resources = unique_values(resources)
        citation = clean(xml_record.get("citation")) if xml_record else ""
        if not citation:
            citation = next((clean(match.get("citation")) for match in remote_matches if clean(match.get("citation"))), "")
        license_value = clean(xml_record.get("license")) if xml_record else ""
        if not license_value:
            license_value = next((clean(match.get("license")) for match in remote_matches if clean(match.get("license"))), "")
        publication_date = clean(xml_record.get("publication_date")) if xml_record else ""
        if not publication_date:
            publication_date = next((normalize_date(clean(match.get("publication_date"))) for match in remote_matches
                                     if clean(match.get("publication_date"))), "")

        # Use DataCite only when the existing metadata lacks a citation
        # or publisher. Never overwrite values already found in XML or
        # the repository lookups.
        
        datacite_used = False

        if allow_online and (not citation.strip() or not publishers):
            doi_candidates = unique_values(
                [primary_doi] + all_dois
            )

            for candidate_doi in doi_candidates:
                datacite = datacite_lookup(candidate_doi)

                if not citation.strip() and datacite["citation"]:
                    citation = datacite["citation"]
                    datacite_used = True

                if not publishers and datacite["publisher"]:
                    publishers = [datacite["publisher"]]
                    datacite_used = True

                if citation.strip() and publishers:
                    break
        
        xml_file = clean(row.get("inventory_xml_location")) or (clean(xml_record.get("xml_location")) if xml_record else "")
        item = {
            "dataset_id": dataset_id, "is_latest_version": latest,
            "item_title": title,
            "resource_type": next(
                    (
                        clean(value)
                        for value in (
                            [match.get("resource_type", "") for match in remote_matches]
                            + ([xml_record.get("resource_type", "")] if xml_record else [])
                        )
                        if clean(value)
                    ),
                    ""
                ),
            "inventory_stac_category": clean(row.get("inventory_stac_category")),
            "tags": "; ".join(tags),
            "bbox_west": bbox[0] if len(bbox) > 0 else "", "bbox_south": bbox[1] if len(bbox) > 1 else "",
            "bbox_east": bbox[2] if len(bbox) > 2 else "", "bbox_north": bbox[3] if len(bbox) > 3 else "",
            "temporal_start": temporal_start, "temporal_end": temporal_end,
            "license": license_value, "description": description,
            "providers": "; ".join(providers), "citation": citation,
            "version": "; ".join(versions), "additional_resources": "; ".join(resources),
            "purpose": clean(xml_record.get("purpose")) if xml_record else "",
            "authors": "; ".join(authors), "publishers": "; ".join(publishers),
            "place": "; ".join(xml_record.get("places", [])) if xml_record else "",
            "doi": "; ".join(all_dois) or primary_doi, "publication_date": publication_date,
            "stac_nasa_doi": source_dois["stac_nasa_doi"],
            "stac_harvard_doi": source_dois["stac_harvard_doi"],
            "columbia_ac_doi": source_dois["columbia_ac_doi"],
            "inventory_xml_location": xml_file,
            "xml_metadata_found": "yes" if xml_record else "no",
            "metadata_sources": "; ".join(unique_values(
                    (["Local XML"] if xml_record else [])
                    + [match["source"] for match in remote_matches]
                    + (["DataCite"] if datacite_used else [])
                )),
        }
        items.append(item)
        item_context.append({"input": row, "item": item, "xml": xml_record, "remote": remote_matches})
    return items, item_context, {"xml_parse_errors": str(len(xml_errors))}


def doc_like(*values: Any) -> bool:
    text = " ".join(clean(value).casefold() for value in values)
    return any(term in text for term in DOCUMENT_TERMS)


def is_earthdata_tile_or_granule(row: dict[str, str]) -> bool:
    source = clean(row.get("source")).casefold()
    href = clean(row.get("href") or row.get("Download URL"))
    host = urllib.parse.urlparse(href).netloc.casefold()
    text = " ".join([source, href, clean(row.get("title")), clean(row.get("Data Description"))]).casefold()
    if "tiles.earthdata" in host or "/tiles/" in text or "/tile/" in text:
        return not doc_like(row.get("title"), row.get("description"), href)
    return "tile" in source and "nasa" in source and not doc_like(row.get("title"), row.get("description"), href)


def is_excluded_xml_asset(href: str, title: str = "") -> bool:
    """Exclude XML links that are licensing notices or branding images."""
    parsed = urllib.parse.urlparse(href)
    if (parsed.hostname and parsed.hostname.casefold().endswith("creativecommons.org")
            and parsed.path.casefold().startswith("/licenses/")):
        return True
    label = f"{Path(parsed.path).name} {title}".casefold()
    return bool(re.search(r"(?:^|[\\s._-])logos?(?:$|[\\s._-])", label))


def asset_format_label(href: str, media_type: str = "") -> str:
    """Return a human-readable file format label from MIME type or extension."""
    path = urllib.parse.unquote(urllib.parse.urlparse(href).path)
    extension = Path(path).suffix.casefold()
    mime = clean(media_type).casefold()
    labels = {
        ".pdf": "PDF document", ".csv": "CSV data", ".tsv": "tabular data",
        ".xlsx": "Excel spreadsheet", ".xls": "Excel spreadsheet",
        ".json": "JSON data", ".xml": "XML metadata", ".zip": "ZIP archive",
        ".gz": "compressed file", ".tif": "GeoTIFF raster", ".tiff": "GeoTIFF raster",
        ".shp": "Shapefile", ".gpkg": "GeoPackage", ".geojson": "GeoJSON data",
        ".kml": "KML data", ".kmz": "KMZ archive", ".nc": "NetCDF data",
        ".nc4": "NetCDF data", ".hdf": "HDF data", ".h5": "HDF5 data",
        ".doc": "Word document", ".docx": "Word document",
        ".ppt": "PowerPoint presentation", ".pptx": "PowerPoint presentation",
        ".txt": "text file", ".jpg": "image", ".jpeg": "image", ".png": "image",
    }
    if extension in labels:
        return labels[extension]
    if mime == "application/pdf":
        return "PDF document"
    if "spreadsheet" in mime or "excel" in mime:
        return "spreadsheet"
    if "json" in mime:
        return "JSON data"
    if "xml" in mime:
        return "XML metadata"
    if mime.startswith("image/"):
        return "image"
    if mime.startswith("text/") and mime not in {"text/html", "text/xml"}:
        return "text file"
    if mime in {"text/html", "application/xhtml+xml"} or not mime:
        return "web page"
    return mime.split("/", 1)[-1].replace("-", " ").replace(".", " ")


def informative_asset_title(item_title: str, href: str, title: str = "",
                            source: str = "", media_type: str = "",
                            description: str = "") -> str:
    """Create a dataset-specific title based on the actual linked resource."""
    supplied = clean(title)
    dataset = clean(item_title) or "Dataset"
    parsed_href = urllib.parse.urlparse(href)
    path_name = Path(urllib.parse.unquote(parsed_href.path)).name
    path_lower = parsed_href.path.casefold()
    host = (parsed_href.hostname or "").casefold()
    media_lower = clean(media_type).casefold().split(";", 1)[0]

    # Recognize special resource types before using generic landing-page text.
    if host == "search.earthdata.nasa.gov" and "/search/granules" in path_lower:
        return f"NASA Earthdata granule search — {dataset}"
    if path_name.casefold().endswith("-metadata.zip") or path_name.casefold() == "metadata.zip":
        return f"Metadata package — {dataset}"

    # The actual MIME type/extension takes precedence over a misleading supplied title.
    image_exts = (".jpg", ".jpeg", ".png", ".gif", ".tif", ".tiff", ".webp")
    is_image = path_lower.endswith(image_exts) or media_lower.startswith("image/")
    if is_image:
        image_kind = "Map image" if any(x in path_lower for x in ("map", "thumbnail", "thumb")) else "Image"
        if path_name and path_name.casefold() not in {"download", "file"}:
            return f"{image_kind} — {dataset} ({path_name})"
        return f"{image_kind} — {dataset}"

    doc_markers = ("data set documentation", "dataset documentation", "documentation",
                   "user guide", "user manual", "readme", "release notes", "data dictionary",
                   "codebook", "methodology")
    if any(marker in supplied.casefold() for marker in doc_markers):
        return f"Dataset documentation — {dataset}"

    generic = {"", "download", "file", "resource", "attachment", "view", "record"}
    if supplied.casefold() not in generic:
        return supplied
    if path_name and path_name.casefold() not in generic:
        stem = Path(path_name).stem
        if not re.fullmatch(r"[0-9a-f-]{24,}", stem, flags=re.I):
            return path_name
    description_text = clean(description)
    if description_text and description_text.casefold() not in {
            "download", "download file", "click here", "file download"}:
        first_sentence = re.split(r"(?<=[.!?])\\s+", description_text, maxsplit=1)[0]
        if len(first_sentence) > 120:
            first_sentence = first_sentence[:117].rstrip() + "..."
        if len(first_sentence) >= 8:
            return first_sentence
    fmt = asset_format_label(href, media_type)
    source_label = clean(source)
    if "doi" in source_label.casefold():
        kind = source_label.replace(" DOI", "").strip()
        return f"{clean(item_title) or 'Dataset'} — {kind} DOI"
    if "record" in source_label.casefold() or fmt == "web page":
        return f"{clean(item_title) or 'Dataset'} — {source_label or 'repository'} record"
    return f"{clean(item_title) or 'Dataset'} — {fmt}"


def create_asset(context: dict[str, Any], href: str, source: str, title: str = "", description: str = "",
                 media_type: str = "", roles: str = "data", source_identifier: str = "",
                 source_doi: str = "", size: str = "", checksum: str = "",
                 created: str = "", updated: str = "",
                 match_basis: str = "inventory row") -> dict[str, Any]:
    item = context["item"]
    path_name = Path(urllib.parse.unquote(urllib.parse.urlparse(href).path)).name
    resolved_media_type = media_type or mimetypes.guess_type(path_name)[0]
    if not resolved_media_type and clean(title).casefold().endswith(".pdf"):
        resolved_media_type = "application/pdf"
    resolved_media_type = resolved_media_type or mimetypes.guess_type(
        urllib.parse.urlparse(href).path)[0] or "text/html"
    display_title = informative_asset_title(
        item.get("item_title", ""), href, title, source, resolved_media_type, description
    )

    # Detect generic metadata ZIP files from either the URL or supplied title.
    metadata_zip = (
        path_name.casefold() == "-metadata.zip"
        or clean(title).casefold() == "-metadata.zip"
        or clean(description).casefold().endswith("-metadata.zip")
    )

    if metadata_zip:
        dataset_title = clean(item.get("item_title")) or "Unknown dataset"
        display_title = f"Metadata package for {dataset_title}"

    # Improve titles for individual files with descriptive filenames.
    if not metadata_zip:
        filename_stem = Path(path_name).stem
        filename_clean = re.sub(r"-\d+$", "", filename_stem)
        filename_words = re.sub(r"[-_]+", " ", filename_clean).strip()

        if filename_words and filename_words.casefold() not in {
            "download", "file", "document", "data", "file xhtml"
        }:
            if resolved_media_type.startswith("image/"):
                display_title = (
                    f"Map image — {item['item_title']} "
                    f"({filename_words})"
                )
            elif "appendix" in filename_words.casefold():
                display_title = (
                    f"{filename_words.title()} — {item['item_title']}"
                )
            elif "documentation" in filename_words.casefold():
                display_title = (
                    f"Documentation — {filename_words.replace('-', ' ')}"
                )

    # Use the descriptive title for generic or opaque filenames.
    key_name = (
        display_title
        if metadata_zip or path_name.casefold() in {"", "download", "file", "file.xhtml"}
        else path_name
    )
    row = {
        "dataset_id": item["dataset_id"], "dataset_doi": item["doi"], "item_title": item["item_title"],
        "asset_key": slug(key_name or source), "href": href,
        "media_type": resolved_media_type,
        "title": display_title, "description": clean(description),
        "roles": roles, "source": source, "source_doi": clean(source_doi),
        "source_identifier": source_identifier,
        "file_size_bytes": size, "checksum_md5": checksum, "created": created, "updated": updated,
        "asset_match_basis": match_basis,
    }
    row["asset_subtype"] = asset_subtypes.classify_asset(row)
    return row


def normalized_asset_href(href: str) -> str:
    """Normalize DOI resolver URLs for asset deduplication."""
    parsed = urllib.parse.urlparse(href.strip())
    host = (parsed.hostname or "").casefold()

    if host in {"doi.org", "dx.doi.org"}:
        doi = urllib.parse.unquote(parsed.path).strip("/").casefold()
        return f"doi:{doi}"

    return href.casefold()


def merge_asset_rows(existing: dict[str, Any], incoming: dict[str, Any]) -> None:
    """Merge provenance and useful metadata when the same URL is discovered twice."""
    for field in ("source", "source_doi", "source_identifier", "roles", "asset_match_basis"):
        existing[field] = "; ".join(unique_values(
            clean(existing.get(field)).split(";") + clean(incoming.get(field)).split(";")
        ))
    existing["description"] = "; ".join(unique_values(
        [existing.get("description", ""), incoming.get("description", "")]
    ))
    for field in ("file_size_bytes", "checksum_md5", "created", "updated"):
        if not clean(existing.get(field)) and clean(incoming.get(field)):
            existing[field] = incoming[field]
    old_type = clean(existing.get("media_type")).casefold()
    new_type = clean(incoming.get("media_type")).casefold()
    if (not old_type or old_type in {"text/html", "application/octet-stream"}) and new_type:
        existing["media_type"] = incoming["media_type"]
    old_title = clean(existing.get("title"))
    new_title = clean(incoming.get("title"))
    if (not old_title or old_title.casefold() in {"download", "file", "resource"}
            or len(new_title) > len(old_title) + 8):
        existing["title"] = new_title
    existing["asset_subtype"] = asset_subtypes.classify_asset(existing)


def is_columbia_doi_assets_panel(row: dict[str, Any]) -> bool:
    """Exclude the generic Academic Commons assets-panel link."""
    href = clean(row.get("href")).casefold()
    source = clean(row.get("source")).casefold()

    return (
        "academiccommons.columbia.edu" in href
        and "#assets-panel" in href
        and source == "columbia academic commons file"
    )


def assemble_assets(contexts: list[dict[str, Any]], log: list[tuple[str, str]]) -> list[dict[str, Any]]:
    assets: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()

    seen: dict[tuple[str, str], dict[str, Any]] = {}

    def asset_identity_href(href: str) -> str:
        """Normalize DOI resolver URLs for asset deduplication."""
        parsed = urllib.parse.urlparse(href.strip())
        host = (parsed.hostname or "").casefold()

        if host in {"doi.org", "dx.doi.org"}:
            doi = urllib.parse.unquote(parsed.path).strip("/").casefold()
            return f"doi:{doi}"

        return href.casefold()

    def add(row: dict[str, Any]) -> None:
        href = clean(row.get("href"))
        dataset_id = clean(row.get("dataset_id"))
        source = clean(row.get("source")).casefold()
        parsed_href = urllib.parse.urlparse(href)

        # Exclude individual NASA CMR granule assets.
        if clean(row.get("source")) == "NASA CMR Granule":
            return

        
        if is_columbia_doi_assets_panel(row):
            return
        
        if (parsed_href.hostname and parsed_href.hostname.casefold().endswith("creativecommons.org")
                and parsed_href.path.casefold().startswith("/licenses/")):
            return
        if source in {"xml data link", "xml metadata link"} and is_excluded_xml_asset(
                href, clean(row.get("title"))):
            return
        identity = (dataset_id, normalized_asset_href(href))
        if not href or not dataset_id:
            return
        if identity in seen:
            merge_asset_rows(seen[identity], row)
            return

        row["asset_key"] = slug(row.get("asset_key") or row.get("title") or href)
        keys = {asset["asset_key"] for asset in assets if asset["dataset_id"] == dataset_id}
        base = row["asset_key"]
        suffix = 2
        while row["asset_key"] in keys:
            row["asset_key"] = f"{base}-{suffix}"
            suffix += 1
        seen[identity] = row
        assets.append(row)

    # Add record and DOI landing pages as metadata assets.
    for context in contexts:
        item = context["item"]
        for column, source in (("stac_nasa_doi", "NASA CMR DOI"),
                               ("stac_harvard_doi", "Harvard Dataverse DOI"),
                               ("columbia_ac_doi", "Columbia Academic Commons DOI")):
            for doi in doi_list(context["input"].get(column, "")):
                add(create_asset(
                    context, f"https://doi.org/{doi}", source,
                    title=f"{clean(item.get('item_title')) or 'Dataset'} — {source.replace(' DOI', '')} DOI",
                    roles="metadata", source_identifier=doi, source_doi=doi,
                    match_basis=f"inventory column {column}"
                ))
        # Preserve each exact source record returned by the DOI lookup and all
        # file/resource URLs exposed by that record.
        for match in context.get("remote", []):
            if clean(match.get("url")):
                add(create_asset(
                    context, match["url"], match["source"] + " record",
                    title=f"{clean(item.get('item_title')) or 'Dataset'} — {match['source']} record",
                    roles="metadata", source_identifier=clean(match.get("doi")),
                    source_doi=clean(match.get("doi")), match_basis="exact DOI lookup"
                ))
            for remote_asset in match.get("assets", []):
                if not isinstance(remote_asset, dict):
                    continue
                href = clean(remote_asset.get("href"))
                if not href or is_earthdata_tile_or_granule({**remote_asset, "source": match["source"]}):
                    continue
                add(create_asset(
                    context, href, clean(remote_asset.get("source")) or match["source"] + " asset",
                    title=clean(remote_asset.get("title")), description=clean(remote_asset.get("description")),
                    media_type=clean(remote_asset.get("media_type")),
                    roles=clean(remote_asset.get("roles")) or "data",
                    source_identifier=clean(remote_asset.get("source_identifier") or match.get("doi")),
                    source_doi=clean(match.get("doi")),
                    size=clean(remote_asset.get("size")), match_basis="exact DOI lookup",
                ))

    # XML data and metadata links become individual assets. NASA Earthdata
    # tile services are excluded; downloadable granule files are retained.
    for context in contexts:
        xml_record = context.get("xml") or {}
        for link_type, collection in (("XML data link", xml_record.get("data_links", [])),
                                      ("XML metadata link", xml_record.get("metadata_links", []))):
            for link in collection:
                href = clean(link.get("href"))
                candidate = {"source": link_type, "href": href,
                             "title": clean(link.get("title")), "description": "", "source_relation_type": "metadata" if "metadata" in link_type.lower() else "data"}
                if (not href or is_excluded_xml_asset(href, clean(link.get("title")))
                        or is_earthdata_tile_or_granule(candidate)):
                    continue
                add(create_asset(context, href, link_type, title=clean(link.get("title")),
                                 media_type=clean(link.get("type")), roles="metadata" if "metadata" in link_type.lower() else "data",
                                 match_basis="inventory XML location"))

    log.append(("Assets written", str(len(assets))))
    log.append(("Asset rows with merged provenance", str(sum(
        1 for row in assets if ";" in clean(row.get("source"))
        or ";" in clean(row.get("source_doi"))
    ))))
    log.append(("Asset rows with source DOI provenance", str(sum(bool(clean(row.get("source_doi"))) for row in assets))))
    log.append(("NASA Earthdata tile assets excluded", "yes; CMR granule files are retained"))
    return assets


def style_sheet(sheet: Any) -> None:
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for cells in sheet.columns:
        lengths = [len(str(cell.value or "")) for cell in list(cells)[:200]]
        sheet.column_dimensions[get_column_letter(cells[0].column)].width = min(max(max(lengths, default=10) + 2, 12), 55)


def save_workbook(path: Path, sheet_name: str, columns: list[str], rows: list[dict[str, Any]],
                  run_log: list[tuple[str, str]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(column, "") for column in columns])
    style_sheet(sheet)
    log_sheet = workbook.create_sheet("Run log")
    log_sheet.append(["Metric", "Value"])
    for record in run_log:
        log_sheet.append(list(record))
    style_sheet(log_sheet)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)



def remove_duplicate_harvard_record_assets(assets: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], int]:
    """Drop Harvard landing-page assets when the same dataset DOI asset exists."""
    doi_assets: dict[tuple[str, str], str] = {}
    for asset in assets:
        doi = harvard_dedupe.doi_asset_doi(asset)
        if doi:
            for dataset_key in harvard_dedupe.dataset_keys(asset):
                doi_assets[(dataset_key, doi)] = asset.get("href", "")

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


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--sheet", default="Items")
    parser.add_argument("--xml-dir", type=Path, default=DEFAULT_XML_DIR)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--pause-seconds", type=float, default=1.5)
    parser.add_argument("--offline", action="store_true", help="Skip Harvard, NASA CMR, and Columbia API queries")
    args = parser.parse_args()
    if args.pause_seconds < 0:
        parser.error("--pause-seconds cannot be negative")
    input_path, xml_dir, output_dir = args.input.resolve(), args.xml_dir.resolve(), args.output_dir.resolve()
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    items_path = output_dir / f"ciesin_items_{stamp}.xlsx"
    assets_path = output_dir / f"ciesin_assets_{stamp}.xlsx"
    try:
        inventory, sheet_name = read_inventory(input_path, args.sheet)
    except Exception as exc:
        print(f"ERROR reading inventory: {exc}", file=sys.stderr)
        return 1
    logs: list[tuple[str, str]] = [
        ("Input workbook", str(input_path)), ("Input worksheet", sheet_name),
        ("Inventory rows", str(len(inventory))), ("XML directory", str(xml_dir)),
        ("Run timestamp UTC", datetime.now(timezone.utc).isoformat(timespec="seconds")),
    ]
    if args.offline:
        remote_by_row: dict[str, list[dict[str, Any]]] = {}
        logs.append(("Per-row DOI lookups", "skipped by --offline"))
    else:
        remote_by_row, source_log = lookup_inventory_dois(inventory, args.pause_seconds)
        logs.extend(source_log)
    items, contexts, _xml_summary = make_items(
            inventory,
            xml_dir,
            remote_by_row,
            logs,
            allow_online=not args.offline,
        )
    assets = assemble_assets(contexts, logs)
    # harvard dedupe
    assets, duplicate_hdv = remove_duplicate_harvard_record_assets(assets) 


    logs.extend([("Items workbook", str(items_path)), ("Assets workbook", str(assets_path)),
                 ("Latest rows tagged", str(sum("latest" in [tag.strip().casefold() for tag in item.get("tags", "").split(";")] for item in items))),
                 ("Duplicate Harvard record assets removed", str(duplicate_hdv)),
                 ])
    try:
        save_workbook(items_path, "Items", ITEM_COLUMNS, items, logs)
        save_workbook(assets_path, "Assets", ASSET_COLUMNS, assets, logs)
    except Exception as exc:
        print(f"ERROR writing output workbooks: {exc}", file=sys.stderr)
        return 1
    print(json.dumps({
        "items_workbook": str(items_path), "assets_workbook": str(assets_path),
        "inventory_items": len(items), "assets": len(assets),
        "doi_lookups_matched": sum(len(records) for records in remote_by_row.values()),
        "latest_items_tagged": sum("latest" in [tag.strip().casefold() for tag in item.get("tags", "").split(";")] for item in items),
    }, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
