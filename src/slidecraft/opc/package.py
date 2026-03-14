"""OPC package — read/write .pptx ZIP containers."""

from __future__ import annotations

import io
import os
import zipfile
from collections.abc import Iterator
from pathlib import Path

from slidecraft.opc.content_types import ContentTypeMap
from slidecraft.opc.part import Part
from slidecraft.opc.relationships import RelationshipCollection


class OpcPackage:
    """An Open Packaging Convention (OPC) package.

    Represents a .pptx file as a collection of parts with content types
    and relationships.
    """

    def __init__(self) -> None:
        self._content_types = ContentTypeMap()
        self._rels = RelationshipCollection.new()
        self._parts: dict[str, Part] = {}

    @classmethod
    def open(cls, path: str | os.PathLike[str] | io.BytesIO) -> OpcPackage:
        """Open an existing .pptx file."""
        pkg = cls()

        if isinstance(path, io.BytesIO):
            zf = zipfile.ZipFile(path, "r")
        else:
            zf = zipfile.ZipFile(Path(path), "r")

        with zf:
            # Parse [Content_Types].xml
            ct_bytes = zf.read("[Content_Types].xml")
            pkg._content_types = ContentTypeMap.from_xml(ct_bytes)

            # Parse package-level relationships
            if "_rels/.rels" in zf.namelist():
                rels_bytes = zf.read("_rels/.rels")
                pkg._rels = RelationshipCollection.from_xml(rels_bytes)

            # Load all parts
            for name in zf.namelist():
                if name == "[Content_Types].xml":
                    continue
                if name.endswith(".rels"):
                    continue
                part_name = f"/{name}" if not name.startswith("/") else name
                content_type = pkg._content_types.content_type_for(part_name) or ""
                blob = zf.read(name)

                # Check for part-level rels
                rels_path = _rels_path_for(name)
                part_rels: RelationshipCollection | None = None
                if rels_path in zf.namelist():
                    part_rels = RelationshipCollection.from_xml(zf.read(rels_path))

                part = Part(part_name, content_type, blob, part_rels)
                pkg._parts[part_name] = part

        return pkg

    def save(self, path: str | os.PathLike[str] | io.BytesIO) -> None:
        """Save the package to a .pptx file."""
        if isinstance(path, io.BytesIO):
            buf = path
        else:
            buf = io.BytesIO()

        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            # Write [Content_Types].xml
            zf.writestr("[Content_Types].xml", self._content_types.to_xml())

            # Write package-level rels
            if len(self._rels) > 0:
                zf.writestr("_rels/.rels", self._rels.to_xml())

            # Write parts
            for part_name, part in sorted(self._parts.items()):
                zip_name = part_name.lstrip("/")
                zf.writestr(zip_name, part.blob)

                # Write part-level rels
                if len(part.rels) > 0:
                    rels_path = _rels_path_for(zip_name)
                    zf.writestr(rels_path, part.rels.to_xml())

        if not isinstance(path, io.BytesIO):
            Path(path).write_bytes(buf.getvalue())

    @property
    def content_types(self) -> ContentTypeMap:
        return self._content_types

    @property
    def rels(self) -> RelationshipCollection:
        return self._rels

    @property
    def parts(self) -> dict[str, Part]:
        return self._parts

    def get_part(self, part_name: str) -> Part | None:
        """Get a part by its name."""
        return self._parts.get(part_name)

    def add_part(self, part: Part) -> None:
        """Add or replace a part in the package."""
        self._parts[part.part_name] = part
        if part.content_type:
            self._content_types.add_override(part.part_name, part.content_type)

    def remove_part(self, part_name: str) -> None:
        """Remove a part from the package."""
        self._parts.pop(part_name, None)

    def iter_parts(self) -> Iterator[Part]:
        """Iterate over all parts."""
        yield from self._parts.values()


def _rels_path_for(part_path: str) -> str:
    """Compute the .rels path for a given part path.

    E.g. 'ppt/presentation.xml' -> 'ppt/_rels/presentation.xml.rels'
    """
    parts = part_path.rsplit("/", 1)
    if len(parts) == 1:
        return f"_rels/{parts[0]}.rels"
    return f"{parts[0]}/_rels/{parts[1]}.rels"
