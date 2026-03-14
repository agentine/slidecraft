"""Parse and manage .rels relationship files."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from collections.abc import Iterator
from dataclasses import dataclass, field

from slidecraft.xml.ns import NS

_REL_NS = NS["pr"]


@dataclass
class Relationship:
    """A single relationship entry."""

    r_id: str
    rel_type: str
    target: str
    is_external: bool = False


@dataclass
class RelationshipCollection:
    """A collection of relationships from a .rels file."""

    _rels: list[Relationship] = field(default_factory=list)
    _next_id: int = 1

    @classmethod
    def from_xml(cls, xml_bytes: bytes) -> RelationshipCollection:
        """Parse .rels XML bytes."""
        root = ET.fromstring(xml_bytes)
        coll = cls()
        max_id = 0
        for child in root:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "Relationship":
                r_id = child.get("Id", "")
                rel_type = child.get("Type", "")
                target = child.get("Target", "")
                target_mode = child.get("TargetMode", "")
                rel = Relationship(
                    r_id=r_id,
                    rel_type=rel_type,
                    target=target,
                    is_external=(target_mode == "External"),
                )
                coll._rels.append(rel)
                num = int(r_id.replace("rId", "")) if r_id.startswith("rId") else 0
                if num > max_id:
                    max_id = num
        coll._next_id = max_id + 1
        return coll

    @classmethod
    def new(cls) -> RelationshipCollection:
        """Create an empty relationship collection."""
        return cls()

    def __iter__(self) -> Iterator[Relationship]:
        return iter(self._rels)

    def __len__(self) -> int:
        return len(self._rels)

    def get_by_id(self, r_id: str) -> Relationship | None:
        """Find a relationship by its rId."""
        for rel in self._rels:
            if rel.r_id == r_id:
                return rel
        return None

    def get_by_type(self, rel_type: str) -> list[Relationship]:
        """Find all relationships of a given type."""
        return [r for r in self._rels if r.rel_type == rel_type]

    def add(
        self,
        rel_type: str,
        target: str,
        *,
        is_external: bool = False,
        r_id: str | None = None,
    ) -> Relationship:
        """Add a new relationship and return it."""
        if r_id is None:
            r_id = f"rId{self._next_id}"
            self._next_id += 1
        rel = Relationship(
            r_id=r_id,
            rel_type=rel_type,
            target=target,
            is_external=is_external,
        )
        self._rels.append(rel)
        return rel

    def to_xml(self) -> bytes:
        """Serialize to .rels XML bytes."""
        root = ET.Element(f"{{{_REL_NS}}}Relationships")
        for rel in self._rels:
            attrib: dict[str, str] = {
                "Id": rel.r_id,
                "Type": rel.rel_type,
                "Target": rel.target,
            }
            if rel.is_external:
                attrib["TargetMode"] = "External"
            ET.SubElement(root, f"{{{_REL_NS}}}Relationship", attrib=attrib)
        ET.register_namespace("", _REL_NS)
        result: bytes = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
        return result
