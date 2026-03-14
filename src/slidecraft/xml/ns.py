"""XML namespace constants and helpers for Open XML documents."""

from __future__ import annotations

# Office Open XML namespaces
NS: dict[str, str] = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "ep": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}

# Relationship type URIs
RT: dict[str, str] = {
    "CORE_PROPERTIES": (
        "http://schemas.openxmlformats.org/package/2006/relationships"
        "/metadata/core-properties"
    ),
    "EXTENDED_PROPERTIES": (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        "/extended-properties"
    ),
    "OFFICE_DOCUMENT": (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        "/officeDocument"
    ),
    "SLIDE": (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        "/slide"
    ),
    "SLIDE_LAYOUT": (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        "/slideLayout"
    ),
    "SLIDE_MASTER": (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        "/slideMaster"
    ),
    "THEME": (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        "/theme"
    ),
    "IMAGE": (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        "/image"
    ),
    "CHART": (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        "/chart"
    ),
    "HYPERLINK": (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        "/hyperlink"
    ),
    "TABLE_STYLES": (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        "/tableStyles"
    ),
}

# Content type constants
CT: dict[str, str] = {
    "PRESENTATION": (
        "application/vnd.openxmlformats-officedocument.presentationml"
        ".presentation.main+xml"
    ),
    "SLIDE": (
        "application/vnd.openxmlformats-officedocument.presentationml"
        ".slide+xml"
    ),
    "SLIDE_LAYOUT": (
        "application/vnd.openxmlformats-officedocument.presentationml"
        ".slideLayout+xml"
    ),
    "SLIDE_MASTER": (
        "application/vnd.openxmlformats-officedocument.presentationml"
        ".slideMaster+xml"
    ),
    "THEME": "application/vnd.openxmlformats-officedocument.theme+xml",
    "CORE_PROPERTIES": (
        "application/vnd.openxmlformats-package.core-properties+xml"
    ),
    "EXTENDED_PROPERTIES": (
        "application/vnd.openxmlformats-officedocument.extended-properties+xml"
    ),
    "RELS": (
        "application/vnd.openxmlformats-package.relationships+xml"
    ),
    "PNG": "image/png",
    "JPEG": "image/jpeg",
    "GIF": "image/gif",
    "SVG": "image/svg+xml",
    "EMF": "image/x-emf",
    "WMF": "image/x-wmf",
    "CHART": (
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    ),
}


def qn(tag: str) -> str:
    """Expand a prefixed tag name to Clark notation.

    Example: qn("p:sld") -> "{http://...presentationml/2006/main}sld"
    """
    prefix, local = tag.split(":", 1)
    uri = NS.get(prefix)
    if uri is None:
        raise KeyError(f"Unknown namespace prefix: {prefix!r}")
    return f"{{{uri}}}{local}"


def nsmap(*prefixes: str) -> dict[str, str]:
    """Build a namespace map for the given prefixes."""
    return {p: NS[p] for p in prefixes}
