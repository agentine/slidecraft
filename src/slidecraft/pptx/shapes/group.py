"""Group shape."""

from __future__ import annotations

from slidecraft.pptx.shapes.shape import BaseShape, ShapeCollection


class GroupShape(BaseShape):
    """A group of shapes."""

    @property
    def shapes(self) -> ShapeCollection:
        """Get the nested shapes in this group."""
        return ShapeCollection(self._element, self._slide)
