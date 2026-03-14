"""Base Part class for OPC package parts."""

from __future__ import annotations

from slidecraft.opc.relationships import RelationshipCollection


class Part:
    """A part within an OPC package.

    Each part has a name (path within the ZIP), content type, and binary blob.
    Parts can also have their own relationships.
    """

    def __init__(
        self,
        part_name: str,
        content_type: str,
        blob: bytes,
        rels: RelationshipCollection | None = None,
    ) -> None:
        self._part_name = part_name
        self._content_type = content_type
        self._blob = blob
        self._rels = rels or RelationshipCollection.new()

    @property
    def part_name(self) -> str:
        """The part name (path within the package, e.g. '/ppt/presentation.xml')."""
        return self._part_name

    @property
    def content_type(self) -> str:
        return self._content_type

    @property
    def blob(self) -> bytes:
        return self._blob

    @blob.setter
    def blob(self, value: bytes) -> None:
        self._blob = value

    @property
    def rels(self) -> RelationshipCollection:
        return self._rels

    def __repr__(self) -> str:
        return f"Part({self._part_name!r}, {self._content_type!r})"
