"""Parse and manage [Content_Types].xml."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

from slidecraft.xml.ns import NS

_CT_NS = NS["ct"]


@dataclass
class ContentTypeMap:
    """Maps part names and extensions to content types."""

    _defaults: dict[str, str] = field(default_factory=dict)
    _overrides: dict[str, str] = field(default_factory=dict)

    @classmethod
    def from_xml(cls, xml_bytes: bytes) -> ContentTypeMap:
        """Parse [Content_Types].xml bytes."""
        root = ET.fromstring(xml_bytes)
        ct = cls()
        for child in root:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "Default":
                ext = child.get("Extension", "")
                ctype = child.get("ContentType", "")
                ct._defaults[ext.lower()] = ctype
            elif tag == "Override":
                part = child.get("PartName", "")
                ctype = child.get("ContentType", "")
                ct._overrides[part] = ctype
        return ct

    def content_type_for(self, part_name: str) -> str | None:
        """Look up the content type for a part name."""
        if part_name in self._overrides:
            return self._overrides[part_name]
        ext = part_name.rsplit(".", 1)[-1].lower() if "." in part_name else ""
        return self._defaults.get(ext)

    def add_override(self, part_name: str, content_type: str) -> None:
        """Add or update a content type override for a part."""
        self._overrides[part_name] = content_type

    def add_default(self, extension: str, content_type: str) -> None:
        """Add or update a default content type for an extension."""
        self._defaults[extension.lower()] = content_type

    def to_xml(self) -> bytes:
        """Serialize back to [Content_Types].xml bytes."""
        root = ET.Element(f"{{{_CT_NS}}}Types")
        for ext, ctype in sorted(self._defaults.items()):
            ET.SubElement(
                root,
                f"{{{_CT_NS}}}Default",
                Extension=ext,
                ContentType=ctype,
            )
        for part_name, ctype in sorted(self._overrides.items()):
            ET.SubElement(
                root,
                f"{{{_CT_NS}}}Override",
                PartName=part_name,
                ContentType=ctype,
            )
        ET.register_namespace("", _CT_NS)
        result: bytes = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
        return result
