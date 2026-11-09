
import json
from urllib.error import HTTPError, URLError
from urllib.parse import quote
from urllib.request import Request, urlopen


def clean(value):
    """Convert a value to a trimmed string."""
    if value is None:
        return ""

    return " ".join(str(value).split())


def canonical_doi(doi):
    """Normalize a DOI or DOI URL to its DOI identifier."""
    doi = clean(doi)

    for prefix in (
        "https://doi.org/",
        "http://doi.org/",
        "https://dx.doi.org/",
        "http://dx.doi.org/",
        "doi:",
    ):
        if doi.lower().startswith(prefix):
            doi = doi[len(prefix):]
            break

    return doi.strip().rstrip(".,; ")


def datacite_lookup(doi):
    """
    Retrieve citation and publisher metadata from DataCite.

    Parameters
    ----------
    doi : str
        DOI identifier or DOI URL.

    Returns
    -------
    dict
        Dictionary with 'citation' and 'publisher' keys.
        Missing metadata is returned as an empty string.
    """
    doi = canonical_doi(doi)

    result = {
        "citation": "",
        "publisher": "",
    }

    if not doi:
        return result

    url = (
        "https://api.datacite.org/dois/"
        f"{quote(doi, safe='/')}"
    )

    request = Request(
        url,
        headers={
            "Accept": "application/json",
            "User-Agent": "CIESINMetadataLookup/1.0",
        },
    )

    try:
        with urlopen(request, timeout=20) as response:
            payload = json.loads(
                response.read().decode("utf-8")
            )

    except (
        HTTPError,
        URLError,
        TimeoutError,
        ValueError,
        OSError,
    ):
        return result

    metadata = (
        payload.get("data", {}).get("attributes", {})
    )

    # Extract title and publisher.
    titles = metadata.get("titles") or []
    title = ""

    for entry in titles:
        if isinstance(entry, dict) and entry.get("title"):
            title = clean(entry["title"])
            break

    publisher = clean(metadata.get("publisher"))
    year = clean(metadata.get("publicationYear"))

    # Extract creator names.
    authors = [
        clean(creator.get("name"))
        for creator in metadata.get("creators", [])
        if isinstance(creator, dict)
        and clean(creator.get("name"))
    ]

    # Format author names for a readable citation.
    if len(authors) > 7:
        author_text = ", ".join(authors[:6])
        author_text += ", … " + authors[-1]
    elif len(authors) > 1:
        author_text = (
            ", ".join(authors[:-1])
            + ", & "
            + authors[-1]
        )
    elif authors:
        author_text = authors[0]
    else:
        author_text = ""

    # Build a basic citation.
    if title:
        year_text = year or "n.d."

        if author_text:
            citation = f"{author_text} ({year_text})."
        else:
            citation = f"({year_text})."

        citation += f" {title}."

        if publisher:
            citation += f" {publisher}."

        citation += f" https://doi.org/{doi}"

        result["citation"] = citation

    result["publisher"] = publisher

    return result
