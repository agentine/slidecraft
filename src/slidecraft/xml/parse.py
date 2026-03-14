"""XML parsing and serialization utilities."""

from __future__ import annotations

import xml.etree.ElementTree as ET

from slidecraft.xml.ns import NS


def parse_xml(xml_bytes: bytes) -> ET.Element:
    """Parse XML bytes into an Element, with namespace prefixes registered."""
    for prefix, uri in NS.items():
        ET.register_namespace(prefix, uri)
    return ET.fromstring(xml_bytes)


def serialize_xml(
    element: ET.Element, *, xml_declaration: bool = True
) -> bytes:
    """Serialize an Element to XML bytes."""
    for prefix, uri in NS.items():
        ET.register_namespace(prefix, uri)
    if xml_declaration:
        result: bytes = ET.tostring(
            element, encoding="UTF-8", xml_declaration=True
        )
        return result
    return ET.tostring(element, encoding="unicode").encode("utf-8")
