#!/usr/bin/env python3
"""Discover possible CIESIN datasets via public metadata APIs and write a review workbook.

The script is read-only against the input spreadsheet and remote repositories. It
does not edit the inventory or publish STAC records.
"""
from __future__ import annotations

import argparse
import difflib
import json
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
from html.parser import HTMLParser

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError as exc:
    raise SystemExit("Install dependencies with: python -m pip install -r requirements-discovery.txt") from exc

ORG_TERMS = ("ciesin", "sedac", "center for integrated earth system information")
NASA = "https://cmr.earthdata.nasa.gov/search"
HARVARD = "https://dataverse.harvard.edu"
# COLUMBIA_OAI = "https://academiccommons.columbia.edu/catalog/oai"
# COLUMBIA_SEARCH = "https://academiccommons.columbia.edu/search.json"

REQUIRED = ("item_title", "inventory_stac_category", "stac_nasa_doi", "stac_harvard_doi",
            "columbia_ac_doi",   "inventory_xml_location")
DOI_RE = re.compile(r"10\.\d{4,9}/[^\s<>\"']+", re.I)
DOI_TRAILING = ".,;:)]}>'\""
HEADERS = {"User-Agent": "CIESIN-STAC-discovery/1.0"}



COLUMBIA_URL = "https://academiccommons.columbia.edu"


class ColumbiaLandingParser(HTMLParser):
    """Extract metadata and links from an Academic Commons record."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.meta = defaultdict(list)
        self.links = []
        self.title_parts = []
        self._in_title = False
        self._current_link = None

    def handle_starttag(self, tag, attrs):
        values = {
            key.casefold(): value or ""
            for key, value in attrs
        }

        tag = tag.casefold()

        if tag == "meta":
            key = (
                values.get("name")
                or values.get("property")
                or values.get("itemprop")
            )

            content = values.get("content", "").strip()

            if key and content:
                self.meta[key.casefold()].append(content)

        elif tag == "a" and values.get("href"):
            self._current_link = {
                "href": values["href"],
                "text": "",
                "title": values.get("title", ""),
                "aria_label": values.get("aria-label", ""),
            }
            self.links.append(self._current_link)

        elif tag == "title":
            self._in_title = True

    def handle_data(self, data):
        data = re.sub(r"\s+", " ", data).strip()

        if not data:
            return

        if self._in_title:
            self.title_parts.append(data)

        if self._current_link is not None:
            current = self._current_link["text"]
            self._current_link["text"] = (
                f"{current} {data}".strip()
            )

    def handle_endtag(self, tag):
        tag = tag.casefold()

        if tag == "title":
            self._in_title = False

        elif tag == "a":
            self._current_link = None


def columbia_landing_page(doi):
    """Retrieve and parse an Academic Commons DOI landing page."""

    url = (
        f"{COLUMBIA_URL}/doi/"
        f"{urllib.parse.quote(doi, safe='/')}"
    )

    request = urllib.request.Request(
        url,
        headers={
            **HEADERS,
            "Accept": "text/html,application/xhtml+xml,*/*",
        },
    )

    with urllib.request.urlopen(request, timeout=45) as response:
        page_url = response.geturl()
        html = response.read().decode(
            "utf-8", errors="replace"
        )

    parser = ColumbiaLandingParser()
    parser.feed(html)

    return page_url, parser


def columbia_meta(parser, *keys):
    """Return the first nonempty value for the supplied metadata keys."""

    for key in keys:
        values = parser.meta.get(key.casefold(), [])

        for value in values:
            if value.strip():
                return value.strip()

    return ""


def columbia_record_metadata(doi):
    """Extract useful metadata from an Academic Commons landing page."""

    page_url, parser = columbia_landing_page(doi)

    title = columbia_meta(
        parser,
        "citation_title",
        "dc.title",
        "dcterms.title",
        "og:title",
    )

    if not title:
        title = " ".join(parser.title_parts)

    description = columbia_meta(
        parser,
        "citation_abstract",
        "description",
        "og:description",
        "dc.description",
        "dcterms.abstract",
    )

    date = columbia_meta(
        parser,
        "citation_publication_date",
        "citation_date",
        "dc.date",
        "dcterms.date",
        "article:published_time",
    )

    record_doi = columbia_meta(
        parser,
        "citation_doi",
        "dc.identifier",
        "dcterms.identifier",
        "doi",
    )

    # Prefer a DOI explicitly identified in the landing page.
    extracted = extract_dois(record_doi)

    final_doi = (
        sorted(extracted)[0]
        if extracted
        else norm_doi(doi)
    )

    return {
        "title": title,
        "doi": final_doi,
        "url": page_url,
        "description": description,
        "date": date,
        "parser": parser,
    }


def norm(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip().casefold()


def norm_doi(value: Any) -> str:
    text = str(value or "").strip()
    text = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", text, flags=re.I)
    text = re.sub(r"^doi:\s*", "", text, flags=re.I)
    match = DOI_RE.search(text)
    return (match.group(0).rstrip(DOI_TRAILING) if match else text).casefold()


def extract_dois(*values: Any) -> set[str]:
    out: set[str] = set()
    for value in values:
        matches = DOI_RE.findall(str(value or ""))
        if matches:
            out.update(m.rstrip(DOI_TRAILING).casefold() for m in matches)
        elif value and re.match(r"^\s*(?:doi:\s*)?10\.\d{4,9}/", str(value), re.I):
            out.add(norm_doi(value))
    return out


def get_json(url: str, attempts: int = 3) -> Any:
    request = urllib.request.Request(url, headers={**HEADERS, "Accept": "application/json"})
    last_error: Exception | None = None
    for attempt in range(attempts):
        try:
            with urllib.request.urlopen(request, timeout=45) as response:
                return json.loads(response.read().decode("utf-8", errors="replace"))
        except (urllib.error.URLError, TimeoutError, json.JSONDecodeError) as exc:
            last_error = exc
            if attempt + 1 < attempts:
                time.sleep(attempt + 1)
    raise RuntimeError(f"Request failed after {attempts} attempts: {last_error}")


def read_input(path: Path, sheet_name: str | None) -> tuple[list[dict[str, str]], str]:
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    iterator = ws.iter_rows(values_only=True)
    headers = [str(v or "").strip() for v in next(iterator, ())]
    missing = [name for name in REQUIRED if name not in headers]
    if missing:
        title = ws.title
        wb.close()
        raise ValueError(f"Worksheet {title!r} is missing required columns: {', '.join(missing)}")
    rows = []
    for values in iterator:
        row = {h: str(values[i] or "").strip() if i < len(values) else ""
               for i, h in enumerate(headers) if h}
        if any(row.get(name) for name in REQUIRED):
            rows.append(row)
    title = ws.title
    wb.close()
    return rows, title


def known_keys(rows: list[dict[str, str]]) -> tuple[set[str], set[str], list[str]]:
    ids, titles, title_values = set(), set(), []
    for row in rows:
        title = row.get("item_title", "")
        if norm(title):
            titles.add(norm(title))
            title_values.append(title)
        ids.update(extract_dois(*(row.get(col, "") for col in
            ("stac_nasa_doi", "stac_harvard_doi", "columbia_ac_doi"))))
    return ids, titles, title_values


def mentions_ciesin(value: Any) -> bool:
    text = norm(value)
    return any(term in text for term in ORG_TERMS)


def hit(source: str, title: str, identifier: str = "", url: str = "",
        description: str = "", date: str = "", query: str = "",
        record_type: str = "") -> dict[str, str]:
    identifier = norm_doi(identifier)
    return {
        "item_title": title.strip(), "inventory_stac_category": "", "candidate_doi": identifier,
        "stac_nasa_doi": identifier if source == "NASA CMR" else "",
        "stac_harvard_doi": identifier if source == "Harvard Dataverse" else "",
        "columbia_ac_doi": identifier if source == "Columbia Academic Commons" else "", 
        "source": source, "source_url": url, "description": description,
        "source_date": date, "search_query": query, "record_type": record_type,
    }


def search_nasa(query: str, page_cap: int) -> list[dict[str, str]]:
    found, scanned = [], 0
    for page in range(1, page_cap + 1):
        search_field = "doi[]" if DOI_RE.search(query) else "keyword"
        params = urllib.parse.urlencode({search_field: query, "page_size": 2000, "page_num": page})
        payload = get_json(f"{NASA}/collections.umm_json?{params}")
        items = payload.get("items", [])
        scanned += len(items)
        for item in items:
            umm = item.get("umm", {})
            if not mentions_ciesin(json.dumps(umm, ensure_ascii=False)):
                continue
            meta = item.get("meta", {})
            concept = str(meta.get("concept-id", ""))
            citations = umm.get("CollectionCitations") or [{}]
            found.append(hit("NASA CMR", str(umm.get("EntryTitle") or umm.get("ShortName") or ""),
                str(umm.get("DOI") or ""),
                f"https://cmr.earthdata.nasa.gov/search/concepts/{concept}.umm_json" if concept else "",
                str(umm.get("Abstract") or ""), str(citations[0].get("ReleaseDate") or ""), query,
                "NASA CMR collection"))
        if not items or len(items) < 2000 or scanned >= int(payload.get("hits", 0)):
            break
    return found


def dataverse_fields(dataset: dict[str, Any]) -> dict[str, Any]:
    result = {}
    for block in (dataset.get("latestVersion") or {}).get("metadataBlocks", {}).values():
        for field in block.get("fields", []):
            result[field.get("typeName", "")] = field.get("value")
    return result


def search_harvard(query: str, page_cap: int) -> list[dict[str, str]]:
    found, per_page = [], 100
    for page in range(page_cap):
        params = urllib.parse.urlencode({"q": query, "type": "dataset", "per_page": per_page, "start": page * per_page})
        data = get_json(f"{HARVARD}/api/search?{params}").get("data", {})
        items = data.get("items", [])
        for item in items:
            if not mentions_ciesin(json.dumps(item, ensure_ascii=False)):
                continue
            persistent_id = str(item.get("global_id") or "").removeprefix("doi:")
            detail = {}
            if persistent_id:
                detail_params = urllib.parse.urlencode({"persistentId": persistent_id})
                detail = get_json(f"{HARVARD}/api/datasets/:persistentId?{detail_params}").get("data", {})
            fields = dataverse_fields(detail) if detail else {}
            if not mentions_ciesin(json.dumps(fields, ensure_ascii=False) + json.dumps(item, ensure_ascii=False)):
                continue
            url = str(detail.get("persistentUrl") or
                       (f"https://doi.org/{persistent_id}" if persistent_id else item.get("url", "")))
            found.append(hit("Harvard Dataverse", str(item.get("name") or fields.get("title") or ""),
                persistent_id, url, str(item.get("description") or ""),
                str(item.get("published_at") or ""), query, "Dataverse dataset"))
        total = int(data.get("total_count", 0))
        if not items or len(items) < per_page or (page + 1) * per_page >= total:
            break
    return found


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def xml_values(node: ET.Element, name: str) -> list[str]:
    return [" ".join((el.text or "").split()) for el in node.iter()
            if local_name(el.tag) == name and (el.text or "").strip()]

import socket

def get_html(url: str, attempts: int = 4) -> str:
    """Retrieve HTML with longer timeouts and exponential backoff."""

    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/130.0.0.0 Safari/537.36"
            ),
            "Accept": (
                "text/html,application/xhtml+xml,"
                "application/xml;q=0.9,*/*;q=0.8"
            ),
            "Accept-Language": "en-US,en;q=0.9",
            "Connection": "close",
        },
    )

    last_error = None

    for attempt in range(attempts):
        try:
            with urllib.request.urlopen(
                request,
                timeout=90,
            ) as response:

                html = response.read().decode(
                    "utf-8", errors="replace"
                )

                if not html.strip():
                    raise RuntimeError(
                        "Academic Commons returned an empty response."
                    )

                return html

        except (
            urllib.error.URLError,
            urllib.error.HTTPError,
            socket.timeout,
            TimeoutError,
            ConnectionError,
        ) as exc:

            last_error = exc

            print(
                f"Academic Commons request attempt "
                f"{attempt + 1}/{attempts} failed: {exc}"
            )

            if attempt + 1 < attempts:
                wait = 5 * (2 ** attempt)
                print(f"Retrying in {wait} seconds...")
                time.sleep(wait)

    raise RuntimeError(
        f"HTML request failed after {attempts} attempts: "
        f"{last_error}"
    )



def search_columbia(
    page_cap: int,
    from_date: str = "",
    known_dois: set[str] | None = None,
    known_titles: set[str] | None = None,
) -> list[dict[str, str]]:
    """
    Discover CIESIN records using the Academic Commons REST API.

    Searches by CIESIN and SEDAC keywords and by current and legacy
    CIESIN department names.

    Retrieves paginated JSON results and extracts DOI, title,
    abstract, date, and persistent URL.

    Skips records already present in the inventory and deduplicates
    results across searches.
    """

    known_dois = known_dois or set()
    known_titles = known_titles or set()

    found: list[dict[str, str]] = []
    seen: set[str] = set()

    api_url = f"{COLUMBIA_URL}/api/v1/search"

    headers = {
        "User-Agent": "CIESIN-STAC-discovery/1.0",
        "Accept": "application/json",
    }

    # Search terms and department filters.
    searches = [
        {
            "name": "CIESIN keyword",
            "params": {
                "q": "ciesin",
                "search_type": "keyword",
            },
        },
        {
            "name": "SEDAC keyword",
            "params": {
                "q": "SEDAC",
                "search_type": "keyword",
            },
        },
        {
            "name": "Current CIESIN department",
            "params": {
                "department[]":
                    "Center for Integrated Earth System Information",
            },
        },
        {
            "name": "Legacy CIESIN department",
            "params": {
                "department[]":
                    "Center for International Earth Science Information Network",
            },
        },
    ]

    per_page = 25

    for search in searches:

        search_name = search["name"]
        base_params = search["params"]
 

        for page in range(1, page_cap + 1):

            params = {
                **base_params,
                "page": page,
                "per_page": per_page,
            }

            search_url = (
                f"{api_url}?"
                + urllib.parse.urlencode(params)
            )

            try:
                request = urllib.request.Request(
                    search_url,
                    headers=headers,
                )

                with urllib.request.urlopen(
                    request,
                    timeout=90,
                ) as response:

                    response_text = response.read().decode(
                        "utf-8",
                        errors="replace",
                    )

                data = json.loads(response_text)

            except Exception as exc:
                raise RuntimeError(
                    f"Academic Commons API failed during "
                    f"'{search_name}', page {page}: {exc}"
                ) from exc

            records = data.get("records", [])
            total = data.get("total_number_of_results", 0)

             

            if not records:
                break

            page_new = 0

            for api_record in records:

                # DOI is provided directly in the API's id field.
                doi = norm_doi(
                    api_record.get("id", "")
                )

                if not doi:
                    continue

                # Avoid duplicates across all four searches.
                if doi in seen:
                    continue

                seen.add(doi)

                title = re.sub(
                    r"\s+",
                    " ",
                    str(api_record.get("title") or ""),
                ).strip()

                if not title:
                    continue

                # Skip records already in the inventory.
                if doi in known_dois:
                    continue

                if norm(title) in known_titles:
                    continue

                # Extract additional metadata from the API.
                description = str(
                    api_record.get("abstract") or ""
                ).strip()

                date = str(
                    api_record.get("date") or ""
                ).strip()

                # Prefer the repository's persistent URL.
                record_url = (
                    api_record.get("persistent_url")
                    or f"{COLUMBIA_URL}/doi/{doi}"
                )

                record = hit(
                    source="Columbia Academic Commons",
                    title=title,
                    identifier=doi,
                    url=record_url,
                    description=description,
                    date=date,
                    query=search_name,
                    record_type="Academic Commons record",
                )

                found.append(record)
                page_new += 1

                 

            

            # Stop when all API results have been retrieved.
            if page * per_page >= total:
                break

            if page < page_cap:
                time.sleep(1)

    print(
        f"\nColumbia | Total unique new records retrieved: "
        f"{len(found)}"
    )

    return found



def filter_merge(hits: list[dict[str, str]], known_dois: set[str],
                 known_titles: set[str], title_values: list[str]) -> list[dict[str, Any]]:
    grouped: dict[str, dict[str, Any]] = {}
    for row in hits:
        title, identifier = row.get("item_title", ""), norm_doi(row.get("candidate_doi", ""))
        if (identifier and identifier in known_dois) or (norm(title) and norm(title) in known_titles):
            continue
        key = f"doi:{identifier}" if identifier else f"title:{norm(title)}"
        if key in ("doi:", "title:"):
            continue
        item = grouped.setdefault(key, {**row, "_sources": set(), "_urls": set(), "_queries": set()})
        item["_sources"].add(row.get("source", ""))
        item["_urls"].add(row.get("source_url", "")) if row.get("source_url") else None
        item["_queries"].add(row.get("search_query", "")) if row.get("search_query") else None
        if not item.get("description") and row.get("description"):
            item["description"] = row["description"]
        best, score = "", 0.0
        for known in title_values:
            current = difflib.SequenceMatcher(None, norm(title), norm(known)).ratio()
            if current > score:
                best, score = known, current
        if score >= 0.84 and score > item.get("title_match_score", 0):
            item["possible_existing_title_match"] = best
            item["title_match_score"] = round(score, 3)
    results = []
    for row in grouped.values():
        row["sources"] = "; ".join(sorted(filter(None, row.pop("_sources"))))
        row["source_urls"] = "; ".join(sorted(row.pop("_urls")))
        row["search_queries"] = "; ".join(sorted(filter(None, row.pop("_queries"))))
        if row.get("possible_existing_title_match"):
            row["review_flag"] = "Possible similar existing title"
        elif "confirmation needed" in row.get("record_type", "").casefold():
            row["review_flag"] = "Confirm record is a dataset"
        else:
            row["review_flag"] = "Review as potential new item"
        results.append(row)
    return sorted(results, key=lambda row: (row.get("item_title", "").casefold(), row.get("candidate_doi", "")))


def add_table(ws: Any, headers: list[str], rows: list[dict[str, Any]], name: str) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([row.get(header, "") for header in headers])
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="28536B")
        cell.alignment = Alignment(vertical="top", wrap_text=True)
    for col in ws.columns:
        width = max(14, min(55, max((len(str(c.value or "")) for c in col[:200]), default=10) + 2))
        ws.column_dimensions[col[0].column_letter].width = width
    if rows:
        table = Table(displayName=name, ref=ws.dimensions)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(table)


def write_workbook(path: Path, candidates: list[dict[str, Any]], log: list[dict[str, str]],
                   input_path: Path, sheet_name: str, run_time: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    summary = wb.active
    summary.title = "Run Summary"
    for row in [
        ["CIESIN potential item discovery", ""], ["Organization", "Center for Integrated Earth System Information (CIESIN)"],
        ["Input workbook", str(input_path)], ["Input worksheet", sheet_name], ["Run started (UTC)", run_time],
        ["Potential records after exact DOI/title filtering", len(candidates)],
        ["Review note", "Discovery leads only. Confirm CIESIN relationship, scope, identifiers, category, and XML metadata before adding records to the inventory."],
    ]:
        summary.append(row)
    summary.column_dimensions["A"].width = 45
    summary.column_dimensions["B"].width = 115
    summary["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    summary["A1"].fill = PatternFill("solid", fgColor="28536B")
    summary["B1"].fill = PatternFill("solid", fgColor="28536B")
    summary.sheet_view.showGridLines = False
    for row in summary.iter_rows(min_row=2):
        row[0].font = Font(bold=True)
        row[1].alignment = Alignment(vertical="top", wrap_text=True)

    items = wb.create_sheet("Potential Items")
    headers = ["review_flag", "item_title", "inventory_stac_category", "candidate_doi",
        "stac_nasa_doi", "stac_harvard_doi", "columbia_ac_doi",  
        "sources", "record_type", "source_url", "source_urls", "description", "source_date", "search_queries",
        "possible_existing_title_match", "title_match_score"]
    add_table(items, headers, candidates, "PotentialItems")
    items.column_dimensions["L"].width = 65
    items.column_dimensions["K"].width = 55
    items.column_dimensions["N"].width = 40
    for row in items.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        if row[0].value == "Possible similar existing title":
            row[0].fill = PatternFill("solid", fgColor="FFF2CC")

    log_ws = wb.create_sheet("Source Log")
    add_table(log_ws, ["source", "status", "records_fetched", "candidates_after_filter", "details"], log, "SourceLog")
    log_ws.column_dimensions["E"].width = 100
    for row in log_ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    wb.save(path)


def discover(input_path: Path, output_path: Path, sheet: str | None,
             page_cap: int, columbia_from: str, pause: float,
             last_n: int | None = None) -> dict[str, Any]:
    started = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    rows, sheet_name = read_input(input_path, sheet)
    known_dois, known_titles, title_values = known_keys(rows)
    selected_rows = rows[-last_n:] if last_n else rows
    hits: list[dict[str, str]] = []
    log: list[dict[str, str]] = []
    if last_n:
        nasa_terms = list(dict.fromkeys(
            norm_doi(row.get("stac_nasa_doi")) or row.get("item_title", "")
            for row in selected_rows if row.get("stac_nasa_doi") or row.get("item_title")))
        harvard_terms = list(dict.fromkeys(
            norm_doi(row.get("stac_harvard_doi")) or row.get("item_title", "")
            for row in selected_rows if row.get("stac_harvard_doi") or row.get("item_title")))
        selected_dois = set().union(*(extract_dois(row.get("columbia_ac_doi", ""))
                                      for row in selected_rows))
        selected_titles = {norm(row.get("item_title", "")) for row in selected_rows if row.get("item_title")}
        def columbia_sample() -> list[dict[str, str]]:
            found = search_columbia(page_cap, columbia_from)
            return [r for r in found if (r.get("candidate_doi") and r["candidate_doi"] in selected_dois)
                    or norm(r.get("item_title", "")) in selected_titles]
        tasks = [
            ("NASA CMR", lambda: [r for q in nasa_terms for r in search_nasa(q, 1)]),
            ("Harvard Dataverse", lambda: [r for q in harvard_terms for r in search_harvard(q, 1)]),
            ("Columbia Academic Commons", lambda: search_columbia(page_cap, columbia_from, known_dois, known_titles)),
        ]
    else:
        tasks = [
            ("NASA CMR", lambda: [r for q in ("CIESIN", "SEDAC") for r in search_nasa(q, page_cap)]),
            ("Harvard Dataverse", lambda: [r for q in ("CIESIN", "SEDAC") for r in search_harvard(q, page_cap)]),
            ("Columbia Academic Commons", lambda: search_columbia(page_cap, columbia_from, known_dois, known_titles)),
        ]
    for source, task in tasks:
        try:
            found = task()
            hits.extend(found)
            log.append({"source": source, "status": "completed", "records_fetched": str(len(found)), "candidates_after_filter": "", "details": f"Configured page cap: {page_cap}. Confirm coverage if source results may exceed this cap."})
        except Exception as exc:
            log.append({"source": source, "status": "failed", "records_fetched": "0", "candidates_after_filter": "", "details": str(exc)})
        if pause:
            time.sleep(pause)
    candidates = filter_merge(hits, known_dois, known_titles, title_values)
    source_counts: dict[str, int] = defaultdict(int)
    for row in candidates:
        for source in row.get("sources", "").split("; "):
            if source:
                source_counts[source] += 1
    for row in log:
        row["candidates_after_filter"] = str(source_counts.get(row["source"], 0))
    write_workbook(output_path, candidates, log, input_path, sheet_name, started)
    return {"input_workbook": str(input_path), "input_sheet": sheet_name, "input_rows": len(rows),
            "rows_used_as_lookup_seeds": len(selected_rows),
            "known_doi_count": len(known_dois), "candidate_count": len(candidates),
            "output_workbook": str(output_path), "source_log": log}


def main() -> int:
    repo = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-xlsx", type=Path, default=repo / "inputs" / "CIESIN_doi_list_INPUT1.xlsx")
    parser.add_argument("--output-xlsx", type=Path, help="Defaults to a timestamped file under outputs/discovery.")
    parser.add_argument("--sheet", help="Worksheet name; default is the active worksheet.")
    parser.add_argument("--max-pages-per-query", type=int, default=20)
    parser.add_argument("--columbia-from", default="", help="Optional OAI-PMH lower bound YYYY-MM-DD; default harvests from the start.")
    parser.add_argument("--pause-seconds", type=float, default=0.5)
    parser.add_argument("--last-n", type=int, help="Use only the last N input rows as source lookup seeds; duplicate checks still use the full workbook.")
    args = parser.parse_args()
    if args.max_pages_per_query < 1:
        parser.error("--max-pages-per-query must be at least 1")
    if args.last_n is not None and args.last_n < 1:
        parser.error("--last-n must be at least 1")
    if args.columbia_from and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", args.columbia_from):
        parser.error("--columbia-from must be YYYY-MM-DD")
    output = args.output_xlsx
    if output is None:
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        output = repo / "outputs" / "discovery" / f"ciesin_potential_items_{stamp}.xlsx"
    try:
        result = discover(args.input_xlsx, output, args.sheet, args.max_pages_per_query, args.columbia_from, args.pause_seconds, args.last_n)
    except Exception as exc:
        print(f"Discovery failed: {exc}", file=sys.stderr)
        return 1
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
