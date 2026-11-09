#!/usr/bin/env python3
"""Extract STAC-ready item fields from SEDAC XML metadata and combined item CSV."""

from __future__ import annotations

import argparse
import csv
import html
from html.parser import HTMLParser
import json
from logging import root
import mimetypes
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any
from urllib.parse import urlparse


OUTPUT_COLUMNS = [
    "row_type",
    "source_row_type",
    "xml_found",
    "xml_path",
    "stac_item_id",
    "stac_collection",
    "stac_version",
    "stac_title",
    "stac_description",
    "stac_datetime",
    "stac_start_datetime",
    "stac_end_datetime",
    "stac_license",
    "sedac_category",
    "sedac_tags",
    "nasa_doi",
    "harvard_doi",
    "sedac_doi",
    "doi_href",
    "bbox",
    "geometry",
    "west",
    "south",
    "east",
    "north",
    "epsg",
    "data_format",
    "spatial_representation",
    "feature_count",
    "keywords",
    "contacts",
    "citation",
    "purpose",
    "supplemental_info",
    "data_links_json",
    "metadata_links_json",
    "all_links_json",
    "field_schema_json",
    "field_count",
    "raster_schema_json",
    "raster_band_count",
    "raster_resolution",
    "raster_resolution_unit",
    "raster_pixel_depth",
    "raster_pixel_type",
    "raster_no_data_value",
    "raster_min_value",
    "raster_max_value",
    "xml_publication_date",
    "xml_modified_date",
    "source_match_status",
    "existing_stac_id",
    "existing_assets_summary",
]


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", newline="", encoding="utf-8-sig") as handle:
        return list(csv.DictReader(handle))


def text(value: Any) -> str:
    if value is None:
        return ""
    value = html.unescape(str(value))
    value = re.sub(r"<br\s*/?>", " ", value, flags=re.IGNORECASE)
    value = re.sub(r"<[^>]+>", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def slug(value: str, max_len: int = 120) -> str:
    value = re.sub(r"[^a-zA-Z0-9_.-]+", "-", value or "").strip("-").lower()
    return value[:max_len].strip("-") or "item"


def normalize_doi(value: str | None) -> str:
    value = (value or "").strip()
    value = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", value, flags=re.IGNORECASE)
    value = re.sub(r"^doi:", "", value, flags=re.IGNORECASE)
    return value


def compact_json(value: Any) -> str:
    if value in (None, "", [], {}):
        return ""
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def local_name(tag: str) -> str:
    return tag.split("}")[-1]


def first(root: ET.Element, *tags: str) -> str:
    wanted = set(tags)
    for element in root.iter():
        if local_name(element.tag) in wanted and element.text:
            value = text(element.text)
            if value:
                return value
    return ""


def content(root: ET.Element, *tags: str) -> str:
    """Read an element's complete text, including nested ArcGIS HTML fragments."""
    wanted = {tag.casefold() for tag in tags}
    for element in root.iter():
        if local_name(element.tag).casefold() not in wanted:
            continue
        # itertext handles real nested XML nodes; text() also removes HTML when
        # the metadata stores an HTML fragment as escaped character data.
        value = text(" ".join(element.itertext()))
        if value:
            return value
    return ""


def all_values(root: ET.Element, *tags: str) -> list[str]:
    wanted = set(tags)
    values: list[str] = []
    seen: set[str] = set()
    for element in root.iter():
        if local_name(element.tag) in wanted and element.text:
            value = text(element.text)
            if value and value not in seen:
                seen.add(value)
                values.append(value)
    return values


def first_attribute(root: ET.Element, tags: tuple[str, ...], attribute: str = "value") -> str:
    wanted = {tag.casefold() for tag in tags}
    for element in root.iter():
        if local_name(element.tag).casefold() in wanted:
            value = text(element.attrib.get(attribute, ""))
            if value:
                return value
    return ""


def epsg_code(root: ET.Element) -> str:
    value = first(root, "identCode", "WKID", "LatestWKID")
    if value:
        return value
    # ArcGIS may store the coordinate-system XML as text inside <peXml>.
    projection = content(root, "peXml")
    match = re.search(r'(?:AUTHORITY\s*\[\s*"EPSG"\s*,\s*|<WKID>)(\d+)', projection, re.I)
    return match.group(1) if match else ""


def date_value(value: str) -> str:
    value = (value or "").strip()
    if not value:
        return ""
    if re.fullmatch(r"\d{8}", value):
        return f"{value[:4]}-{value[4:6]}-{value[6:8]}T00:00:00Z"
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}", value):
        return f"{value}Z"
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
        return f"{value}T00:00:00Z"
    return value


def bbox(root: ET.Element) -> tuple[list[float] | None, list[str]]:
    west = first(root, "westBL", "westbc", "WestBoundingCoordinate")
    east = first(root, "eastBL", "eastbc", "EastBoundingCoordinate")
    south = first(root, "southBL", "southbc", "SouthBoundingCoordinate")
    north = first(root, "northBL", "northbc", "NorthBoundingCoordinate")
    raw = [west, south, east, north]
    try:
        return [float(west), float(south), float(east), float(north)], raw
    except ValueError:
        return None, raw


def geometry_from_bbox(box: list[float] | None) -> dict[str, Any] | None:
    if not box:
        return None
    west, south, east, north = box
    return {
        "type": "Polygon",
        "coordinates": [[[west, south], [east, south], [east, north], [west, north], [west, south]]],
    }


def urls(root: ET.Element) -> list[str]:
    pattern = re.compile(r"https?://[^\s<>'\"\\]+")
    seen: set[str] = set()
    out: list[str] = []
    for element in root.iter():
        values = list(element.attrib.values())
        if element.text:
            values.append(element.text)
        if element.tail:
            values.append(element.tail)
        for value in values:
            for match in pattern.findall(html.unescape(value or "")):
                url = match.rstrip(".,);]")
                if excluded_asset_link(url):
                    continue
                if url not in seen:
                    seen.add(url)
                    out.append(url)
    return out


class _EmbeddedAnchorParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.current_href = ""
        self.current_text: list[str] = []
        self.titles: dict[str, str] = {}

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.casefold() == "a":
            self.current_href = dict(attrs).get("href") or ""
            self.current_text = []

    def handle_data(self, data: str) -> None:
        if self.current_href and text(data):
            self.current_text.append(text(data))

    def handle_endtag(self, tag: str) -> None:
        if tag.casefold() == "a" and self.current_href:
            label = text(" ".join(self.current_text))
            if label:
                self.titles[html.unescape(self.current_href)] = label
            self.current_href = ""
            self.current_text = []


def embedded_link_titles(root: ET.Element) -> dict[str, str]:
    parser = _EmbeddedAnchorParser()
    for element in root.iter():
        for value in (element.text, element.tail):
            if value and "<a" in value.casefold():
                parser.feed(value)
        # ArcGIS metadata often stores description HTML as real XML child
        # elements, so links are represented as <a href="...">label</a>.
        if local_name(element.tag).casefold() == "a":
            href = element.attrib.get("href") or element.attrib.get("xlink:href") or ""
            label = text(" ".join(element.itertext()))
            if href and label:
                parser.titles[html.unescape(href)] = label
    return parser.titles



def xml_format(root: ET.Element) -> str:
    """Identify the two XML metadata formats supported by this extractor."""
    names = {local_name(element.tag).casefold() for element in root.iter()}

    if "arcgisprofile" in names or "dataidinfo" in names:
        return "esri"

    if "idinfo" in names and "distinfo" in names:
        return "fgdc"

    return "unknown"


def is_doi_resolver(href: str) -> bool:
    """A DOI resolver is a citation/record link, not a downloadable asset."""
    host = (urlparse(href).hostname or "").casefold()
    return host in {"doi.org", "dx.doi.org"}


def unique_asset_links(
    links: list[dict[str, str]],
) -> list[dict[str, str]]:
    """Deduplicate extracted links while preserving their first occurrence."""
    result: list[dict[str, str]] = []
    seen: set[str] = set()

    for link in links:
        href = link.get("href", "").strip()
        key = href.rstrip("/").casefold()

        if not href or key in seen:
            continue

        seen.add(key)
        result.append(link)

    return result


def classify_xml_link(
    href: str,
    title: str = "",
    context: str = "",
) -> tuple[str, str]:
    """
    Return (collection, role) for an XML-derived link.

    The collection is either 'data' or 'metadata'.
    Roles retain more specific meaning for downstream asset processing.
    """
    parsed = urlparse(href)
    filename = Path(parsed.path).name.casefold()
    label = text(title).casefold()
    context = text(context).casefold()

    # Citation resolvers are not downloadable resources.
    if is_doi_resolver(href):
        return "exclude", ""

    # Preview images are useful assets, but not dataset data files.
    if filename.endswith((".jpg", ".jpeg", ".png", ".gif", ".webp")):
        return "metadata", "thumbnail"

    if label.endswith((".jpg", ".jpeg", ".png", ".gif", ".webp")):
        return "metadata", "thumbnail"

    # Explicit documentation labels/extensions take precedence over
    # generic URL tokens such as '/download'.
    documentation_terms = (
        "release notes",
        "documentation",
        "readme",
        "codebook",
        "methodology",
        "user guide",
        "technical report",
    )
    if any(term in label for term in documentation_terms):
        return "metadata", "documentation"

    if filename.endswith((".pdf", ".doc", ".docx", ".txt")):
        return "metadata", "documentation"

    if label.endswith((".pdf", ".doc", ".docx", ".txt")):
        return "metadata", "documentation"

    # Actual data formats.
    data_extensions = (
        ".gpkg", ".zip", ".tif", ".tiff", ".csv", ".nc",
        ".shp", ".shx", ".dbf", ".prj", ".geojson",
        ".json", ".parquet", ".xlsx", ".xls",
    )
    if filename.endswith(data_extensions) or label.endswith(data_extensions):
        return "data", "data"

    # Service endpoints and explicit data access paths.
    lower_url = href.casefold()
    if any(token in lower_url for token in (
        "/featureserver/",
        "/mapserver/",
        "/rest/services/",
        "/api/access/datafile/",
        "/datafile/",
        "/export",
    )):
        return "data", "data"

    # A generic landing/download page is not automatically a data file.
    return "metadata", "landing_page"



# if xml is FGDC, extract related publications from crossref/citeinfo elements    
def extract_fgdc_related_publications(
        root: ET.Element,
    ) -> list[dict[str, str]]:
        """Extract related publications from FGDC crossref elements."""
        publications: list[dict[str, str]] = []

        for crossref in root.iter():
            if local_name(crossref.tag).casefold() != "crossref":
                continue

            citation = next(
                (
                    child for child in crossref
                    if local_name(child.tag).casefold() == "citeinfo"
                ),
                None,
            )

            if citation is None:
                continue

            title = ""
            publication_date = ""
            publisher = ""
            authors = []
            href = ""

            for element in citation.iter():
                tag = local_name(element.tag).casefold()
                value = text(element.text)

                if tag == "title" and value and not title:
                    title = value
                elif tag == "pubdate" and value and not publication_date:
                    publication_date = value
                elif tag == "publish" and value and not publisher:
                    publisher = value
                elif tag == "origin" and value:
                    authors.append(value)
                elif tag == "onlink" and value and not href:
                    href = value

            if not href.startswith(("http://", "https://")):
                continue

            if is_doi_resolver(href):
                # Retain the DOI as the publication's asset link.
                pass

            citation_parts = [
                value for value in (
                    title,
                    "; ".join(dict.fromkeys(authors)),
                    publisher,
                ) if value
            ]

            citation = ". ".join(citation_parts)

            if publication_date:
                citation += f". {publication_date[:4]}"

            description_parts = []

            if publication_date:
                description_parts.append(
                    f"Published {date_value(publication_date)}"
                )

            if publisher:
                description_parts.append(f"Publisher: {publisher}")

            if authors:
                description_parts.append(
                    f"Authors: {'; '.join(dict.fromkeys(authors))}"
                )

            description = ". ".join(description_parts)
            if description:
                description += "."

            publications.append({
                "href": href,
                "title": f"Related Publication: {citation}",
                "type": "text/html",
                "roles": "related publication",
                "description": " | ".join(
                    value for value in description_parts if value
                ),
                "source_element": "crossref/citeinfo",
            })

        return unique_asset_links(publications)


def extract_fgdc_asset_links(
    root: ET.Element,
) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    """Extract distribution URLs and browse images from FGDC metadata."""
    data_links: list[dict[str, str]] = []
    metadata_links: list[dict[str, str]] = []

    related_publications = extract_fgdc_related_publications(root)
    metadata_links.extend(related_publications)

    # Find the FGDC distribution information section first.
    for distinfo in root.iter():
        if local_name(distinfo.tag).casefold() != "distinfo":
            continue

        # Extract distribution links from each digital format separately.
        for digform in distinfo.iter():
            if local_name(digform.tag).casefold() != "digform":
                continue

            formname = ""
            accinstr = ""
            links = []

            for element in digform.iter():
                tag = local_name(element.tag).casefold()
                value = text(element.text)

                if tag == "formname" and value and not formname:
                    formname = value

                elif tag == "accinstr" and value and not accinstr:
                    accinstr = value

                elif tag == "networkr":
                    if value.startswith(("http://", "https://")):
                        links.append(value)

            for href in links:
                collection, role = classify_xml_link(
                    href,
                    context="FGDC distribution",
                )

                if collection == "exclude":
                    continue

                link = {
                    "href": href,
                    "title": (
                            f"Related link: {accinstr}"
                            if accinstr
                            else formname or Path(urlparse(href).path).name or "Related distribution link"
                        ),
                    "description": accinstr,
                    "type": mimetypes.guess_type(urlparse(href).path)[0] or "text/html",
                    "roles": role,
                    "source_element": "distinfo/stdorder/digform/digtopt/onlinopt/computer/networka/networkr",
                }

                if collection == "data":
                    data_links.append(link)
                else:
                    metadata_links.append(link)

    # Extract browse images separately.
    for browse in root.iter():
        if local_name(browse.tag).casefold() != "browse":
            continue

        for element in browse.iter():
            if local_name(element.tag).casefold() != "browsen":
                continue

            href = text(element.text)
            if not href.startswith(("http://", "https://")):
                continue

            metadata_links.append({
                "href": href,
                "title": "Browse image",
                "type": mimetypes.guess_type(urlparse(href).path)[0] or "image/jpeg",
                "roles": "thumbnail",
                "source_element": "idinfo/browse/browsen",
            })

    return (
        unique_asset_links(data_links),
        unique_asset_links(metadata_links),
    )

def extract_esri_asset_links(
    root: ET.Element,
) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    """Extract linked resources from the Esri idAbs description."""
    data_links: list[dict[str, str]] = []
    metadata_links: list[dict[str, str]] = []

    for element in root.iter():
        if local_name(element.tag).casefold() != "idabs":
            continue

        parser = _EmbeddedAnchorParser()

        # Parse embedded HTML wherever it occurs inside idAbs.
        for child in element.iter():
            for value in (child.text, child.tail):
                if value and "<a" in value.casefold():
                    parser.feed(value)

        # Handles anchors represented as actual XML child elements.
        for child in element.iter():
            if local_name(child.tag).casefold() != "a":
                continue

            href = (
                child.attrib.get("href")
                or child.attrib.get("xlink:href")
                or ""
            )
            label = text(" ".join(child.itertext()))
            if href:
                parser.titles[href] = label or Path(
                    urlparse(href).path
                ).name

        for href, label in parser.titles.items():
            href = html.unescape(href).strip()
            if not href.startswith(("http://", "https://")):
                continue

            collection, role = classify_xml_link(
                href,
                title=label,
                context="Esri idAbs",
            )
            if collection == "exclude":
                continue

            link = {
                "href": href,
                "title": label or Path(urlparse(href).path).name,
                "type": media_type_from_title(
                    label,
                    mimetypes.guess_type(urlparse(href).path)[0] or "text/html",
                ),
                "roles": role,
                "source_element": "dataIdInfo/idAbs",
            }

            if collection == "data":
                data_links.append(link)
            else:
                metadata_links.append(link)

    return (
        unique_asset_links(data_links),
        unique_asset_links(metadata_links),
    )


def extract_xml_asset_links(
    root: ET.Element,
) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    """Dispatch asset extraction to the matching XML format parser."""
    fmt = xml_format(root)

    if fmt == "fgdc":
        return extract_fgdc_asset_links(root)

    if fmt == "esri":
        return extract_esri_asset_links(root)

    return [], []



def excluded_asset_link(href: str, title: str = "") -> bool:
    """Links that describe licensing or branding are not dataset resources."""
    parsed = urlparse(href)
    if (parsed.hostname and parsed.hostname.casefold().endswith("creativecommons.org")
            and parsed.path.casefold().startswith("/licenses/")):
        return True
    filename = Path(parsed.path).name.casefold()
    label = f"{filename} {title}".casefold()
    return bool(re.search(r"(?:^|[\\s._-])logos?(?:$|[\\s._-])", label))


def split_links(links: list[str]) -> tuple[list[dict[str, str]], list[dict[str, str]]]:
    data_links: list[dict[str, str]] = []
    metadata_links: list[dict[str, str]] = []
    for url in links:
        if excluded_asset_link(url):
            continue
        guessed, _ = mimetypes.guess_type(url)
        record = {
            "href": url,
            "type": guessed or "text/html",
            "title": Path(re.sub(r"[?#].*$", "", url)).name or url,
        }
        lower = url.lower()
        if any(token in lower for token in [
            "/download", "/export", "/data/", "/datafile/", "/api/access/datafile/",
            "/featureserver/", "/mapserver/", "/rest/services/",
            ".gpkg", ".zip", ".tif", ".tiff", ".csv", ".nc", ".shp",
        ]):
            data_links.append(record)
        else:
            metadata_links.append(record)
    return data_links, metadata_links


def media_type_from_title(title: str, fallback: str) -> str:
    suffix = Path(title.split("?", 1)[0]).suffix.casefold()
    known = {
        ".gpkg": "application/geopackage+sqlite3",
        ".shp": "application/zip",
        ".gdb": "application/zip",
    }
    return known.get(suffix) or mimetypes.guess_type(title)[0] or fallback


def field_schema(root: ET.Element) -> list[dict[str, str]]:
    fields = []
    for attr in root.iter():
        if local_name(attr.tag) != "attr":
            continue
        record: dict[str, str] = {}
        for child in attr:
            name = local_name(child.tag)
            if name in {"attrlabl", "attalias", "attrtype", "attwidth", "attrdef", "attrdefs"}:
                record[name] = text(child.text)
        if record:
            fields.append(record)
    return fields


def raster_schema(root: ET.Element) -> dict[str, Any]:
    general: dict[str, str] = {}
    for element in root.iter():
        name = local_name(element.tag)
        if name in {
            "PixelDepth",
            "HasColormap",
            "CompressionType",
            "NumBands",
            "Format",
            "HasPyramids",
            "SourceType",
            "PixelType",
            "NoDataValue",
        }:
            general[name] = text(element.text)

    bands: list[dict[str, str]] = []
    for band in root.iter():
        if local_name(band.tag) != "Band":
            continue
        record: dict[str, str] = {"schema_type": "raster_band"}
        for child in band.iter():
            name = local_name(child.tag)
            if name in {"dimDescrp", "maxVal", "minVal", "bitsPerVal"}:
                record[name] = text(child.text)
            elif name == "UOM":
                record["unit"] = child.attrib.get("type", "") or text(child.text)
        bands.append(record)

    resolutions: list[dict[str, str]] = []
    for dim in root.iter():
        if local_name(dim.tag) != "axisDimension":
            continue
        record: dict[str, str] = {"axis_type": dim.attrib.get("type", "")}
        for child in dim.iter():
            name = local_name(child.tag)
            if name == "dimSize":
                record["dim_size"] = text(child.text)
            elif name == "value":
                record["resolution"] = text(child.text)
                record["resolution_unit"] = child.attrib.get("uom", "")
        if len(record) > 1:
            resolutions.append(record)

    schema = {
        "schema_type": "raster",
        "general": general,
        "bands": bands,
        "axis_dimensions": resolutions,
    }
    if not general and not bands and not resolutions:
        return {}
    return {key: value for key, value in schema.items() if value not in ({}, [])}


def license_id(*values: str) -> str:
    lower = " ".join(values).lower()
    if "creativecommons.org/licenses/by-sa/4.0" in lower or "cc by-sa 4.0" in lower:
        return "CC-BY-SA-4.0"
    if "creativecommons.org/licenses/by/4.0" in lower or "cc by 4.0" in lower:
        return "CC-BY-4.0"
    return "other"


def contact_names(root: ET.Element, *party_tags: str) -> list[str]:
    """Extract organization/person labels from FGDC and ISO/ArcGIS contacts."""
    wanted = {name.casefold() for name in party_tags}
    result: list[str] = []
    for element in root.iter():
        if local_name(element.tag).casefold() not in wanted:
            continue
        value = text(" ".join(element.itertext()))
        if value:
            result.append(value)
    return list(dict.fromkeys(result))

## extracting citations from different formats of XMLs 
def extract_citation(root: ET.Element, abstract: str = "") -> str:
    """Extract a dataset citation from different XML metadata structures."""

    def clean(value: str) -> str:
        value = text(value)
        value = re.sub(
            r"^(?:Dataset|Recommended)\s+citation\s*:\s*",
            "",
            value,
            flags=re.IGNORECASE,
        )
        return value.strip()

    # 1. Prefer an explicitly labeled citation in the abstract.
    if abstract:
        match = re.search(
            r"\b(?:Recommended\s+)?(?:Dataset\s+)?Citation\s*:\s*"
            r"(.*?)"
            r"(?=\b(?:Terms of use|Data license|Contacts and data queries|"
            r"Copyright)\b|$)",
            abstract,
            flags=re.IGNORECASE | re.DOTALL,
        )

        if match:
            citation = clean(match.group(1))
            

            if citation:
                return citation

    # 2. Look for common structured citation fields.
    citation_tags = {
        "othercitdet",
        "othercit",
        "datasetcitation",
        "citationtext",
    }

    for element in root.iter():
        if local_name(element.tag).casefold() not in citation_tags:
            continue

        citation = clean(" ".join(element.itertext()))
        if citation and len(citation) > 20:
            return citation

    # 3. Fall back to the broader citation element.
    for element in root.iter():
        if local_name(element.tag).casefold() in {
            "idcitation",
            "citeinfo",
        }:
            citation = clean(" ".join(element.itertext()))
            if citation and len(citation) > 20:
                return citation

    return ""


def xml_metadata(path: Path) -> dict[str, Any]:
    root = ET.parse(path).getroot()
    box, raw_box = bbox(root)
    all_urls = urls(root)

    # Extract actual resources according to the XML metadata format.
    data_links, metadata_links = extract_xml_asset_links(root)
    raster = raster_schema(root)
    fields = field_schema(root)
    abstract_markup = content(root, "idAbs", "abstract")
    abstract = text(abstract_markup)
    embedded_citation = ""
    if abstract:
        citation_match = re.search(
            r"\b(?:Recommended\s+)?(?:Dataset\s+)?Citation\b\s*:?[ \t]*(.*?)(?=\b(?:Terms of use|Data license|Contacts and data queries|Copyright)\b|$)",
            abstract, flags=re.I | re.S,
        )
        if citation_match:
            embedded_citation = text(citation_match.group(1))
            abstract = text(abstract[:citation_match.start()]) 

            abstract = re.split(r"\b(?:Terms of use|Data license|Contacts and data queries)\b", abstract,
                                        maxsplit=1, flags=re.I)[0].strip()

    if not embedded_citation:
            embedded_citation = extract_citation(root, abstract) 

    embedded_citation = re.sub(r"^Dataset citation:\s*", "", embedded_citation, flags=re.IGNORECASE)

        

        
        
    use_limit = " ".join(
        value for value in (content(root, "useLimit"), first(root, "accconst", "useconst")) if value
    )
    # License URLs in ArcGIS descriptions may exist only in anchor attributes.
    license_context = " ".join([use_limit, *all_urls])
    pubdate = date_value(first(root, "pubDate", "pubdate"))
    moddate = date_value(first(root, "mdDateSt", "ModDate"))
    citation_value = content(root, "otherCitDet", "othercit")
    if not citation_value:
        citation_value = content(root, "idCitation")

    # Use either citation source, then remove the prefix from the final value.
    citation = citation_value or embedded_citation
    citation = re.sub(
        r"^(?:Dataset|Recommended)\s+citation\s*:?\s*",
        "",
        citation,
        flags=re.IGNORECASE,
    )
    return {
        "title": first(root, "resTitle", "title"),
        "description": abstract or content(root, "idPurp", "purpose", "suppInfo"),
        "publication_date": pubdate,
        "modified_date": moddate,
        "license": license_id(license_context),
        "bbox": box,
        "raw_bbox": raw_box,
        "geometry": geometry_from_bbox(box),
        "epsg": epsg_code(root),
        "format": "; ".join(all_values(root, "formatName", "geoform")),
        "spatial_representation": first(root, "SpatRepTypCd", "spref") or first_attribute(root, ("SpatRepTypCd",)),
        "feature_count": first(root, "enttypc", "geoObjCnt", "efeacnt"),
        "keywords": sorted(set(all_values(root, "keyword", "themekey", "placekey"))),
        "contacts": sorted(set(all_values(root, "rpOrgName", "cntorg", "cntper", "eMailAdd", "cntemail"))),
        "authors": contact_names(root, "origin", "rpOrgName", "displayName"),
        "publishers": contact_names(root, "publish", "rpOrgName", "displayName"),
        "citation": citation,
        "purpose": content(root, "idPurp", "purpose"),
        "supplemental_info": content(root, "suppInfo", "supplinf"),
        "version": first(root, "edition", "resEd"),
        "resource_type": first(root, "mdHrLvName", "resourcetype", "hierarchylevel"),
        "data_links": data_links,
        "metadata_links": metadata_links,
        "all_links": all_urls,
        "fields": fields,
        "raster": raster,
    }


def build_row(input_row: dict[str, str], xml_dir: Path) -> dict[str, str]:
    xml_location = input_row.get("sedac_xml_location") or input_row.get("XML LOCATION") or ""
    xml_path = xml_dir / xml_location if xml_location else None
    meta: dict[str, Any] = {}
    if xml_path and xml_path.exists():
        meta = xml_metadata(xml_path)

    title = meta.get("title") or input_row.get("properties.title") or input_row.get("sedac_data_product_name") or input_row.get("input_dataset_title")
    doi = normalize_doi(input_row.get("sedac_doi") or input_row.get("nasa_doi") or input_row.get("properties.sci:doi"))
    box = meta.get("bbox")
    keywords = sorted(set((meta.get("keywords") or []) + [v.strip() for v in re.split(r"[;,]", input_row.get("sedac_stac_tags", "")) if v.strip()] + ([input_row.get("sedac_stac_category", "")] if input_row.get("sedac_stac_category") else [])))
    raster = meta.get("raster") or {}
    raster_general = raster.get("general", {}) if isinstance(raster, dict) else {}
    raster_bands = raster.get("bands", []) if isinstance(raster, dict) else []
    axis_dimensions = raster.get("axis_dimensions", []) if isinstance(raster, dict) else []
    first_resolution = axis_dimensions[0] if axis_dimensions else {}
    first_band = raster_bands[0] if raster_bands else {}

    return {
        "row_type": "stac_item_candidate",
        "source_row_type": input_row.get("row_type", ""),
        "xml_found": "yes" if meta else "no",
        "xml_path": str(xml_path) if xml_path else "",
        "stac_item_id": input_row.get("id") or slug(title or doi),
        "stac_collection": input_row.get("collection") or "ciesin-sedac",
        "stac_version": input_row.get("stac_version") or "1.1.0",
        "stac_title": title or "",
        "stac_description": meta.get("description") or input_row.get("properties.description", ""),
        "stac_datetime": meta.get("publication_date") or input_row.get("properties.datetime", ""),
        "stac_start_datetime": input_row.get("properties.start_datetime", ""),
        "stac_end_datetime": input_row.get("properties.end_datetime", ""),
        "stac_license": meta.get("license") or input_row.get("properties.license", "") or "other",
        "sedac_category": input_row.get("sedac_stac_category", ""),
        "sedac_tags": input_row.get("sedac_stac_tags", ""),
        "nasa_doi": input_row.get("nasa_doi", ""),
        "harvard_doi": input_row.get("harvard_doi", ""),
        "sedac_doi": doi,
        "doi_href": f"https://doi.org/{doi}" if doi else "",
        "bbox": compact_json(box),
        "geometry": compact_json(meta.get("geometry")),
        "west": str(box[0]) if box else "",
        "south": str(box[1]) if box else "",
        "east": str(box[2]) if box else "",
        "north": str(box[3]) if box else "",
        "epsg": meta.get("epsg", ""),
        "data_format": meta.get("format", ""),
        "spatial_representation": meta.get("spatial_representation", ""),
        "feature_count": meta.get("feature_count", ""),
        "keywords": compact_json(keywords),
        "contacts": compact_json(meta.get("contacts", [])),
        "citation": meta.get("citation", ""),
        "purpose": meta.get("purpose", ""),
        "supplemental_info": meta.get("supplemental_info", ""),
        "data_links_json": compact_json(meta.get("data_links", [])),
        "metadata_links_json": compact_json(meta.get("metadata_links", [])),
        "all_links_json": compact_json(meta.get("all_links", [])),
        "field_schema_json": compact_json(meta.get("fields", [])),
        "field_count": str(len(meta.get("fields", []))) if meta else "0",
        "raster_schema_json": compact_json(raster),
        "raster_band_count": str(len(raster_bands)) if raster else "0",
        "raster_resolution": first_resolution.get("resolution", ""),
        "raster_resolution_unit": first_resolution.get("resolution_unit", ""),
        "raster_pixel_depth": raster_general.get("PixelDepth", ""),
        "raster_pixel_type": raster_general.get("PixelType", ""),
        "raster_no_data_value": raster_general.get("NoDataValue", ""),
        "raster_min_value": first_band.get("minVal", ""),
        "raster_max_value": first_band.get("maxVal", ""),
        "xml_publication_date": meta.get("publication_date", ""),
        "xml_modified_date": meta.get("modified_date", ""),
        "source_match_status": input_row.get("source_match_status", ""),
        "existing_stac_id": input_row.get("id", ""),
        "existing_assets_summary": input_row.get("assets_summary", ""),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-csv", type=Path, default=Path("inputs/ciesin_stac_items_with_sedac_category_and_unmatched_sedac.csv"))
    parser.add_argument("--xml-dir", type=Path, default=Path("inputs/xmls"))
    parser.add_argument("--output-dir", type=Path, default=Path("outputs/xml_stac_item_fields"))
    args = parser.parse_args()

    rows = read_csv(args.input_csv)
    output_rows = [build_row(row, args.xml_dir) for row in rows]
    args.output_dir.mkdir(parents=True, exist_ok=True)
    output_csv = args.output_dir / "ciesin_stac_item_candidates_from_xml.csv"
    with output_csv.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=OUTPUT_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(output_rows)

    summary = {
        "input_rows": len(rows),
        "output_csv": str(output_csv),
        "xml_found": sum(1 for row in output_rows if row["xml_found"] == "yes"),
        "xml_missing": sum(1 for row in output_rows if row["xml_found"] != "yes"),
        "rows_with_data_links": sum(1 for row in output_rows if row["data_links_json"]),
        "rows_with_bbox": sum(1 for row in output_rows if row["bbox"]),
    }
    summary_path = args.output_dir / "ciesin_stac_item_candidates_from_xml_summary.json"
    summary_path.write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(summary, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
