"""Slide, SlideLayout, and SlideMaster classes."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from typing import TYPE_CHECKING

from slidecraft.opc.part import Part
from slidecraft.pptx.shapes.shape import ShapeCollection
from slidecraft.xml.ns import qn
from slidecraft.xml.parse import parse_xml, serialize_xml

if TYPE_CHECKING:
    from slidecraft.pptx.presentation import Presentation


class SlideMaster:
    """A slide master defining base formatting."""

    def __init__(self, part: Part) -> None:
        self._part = part
        self._element: ET.Element | None = None
        self._layouts: list[SlideLayout] = []

    @property
    def element(self) -> ET.Element:
        if self._element is None:
            self._element = parse_xml(self._part.blob)
        return self._element

    @property
    def name(self) -> str:
        csld = self.element.find(qn("p:cSld"))
        if csld is not None:
            return csld.get("name", "")
        return ""

    @property
    def slide_layouts(self) -> list[SlideLayout]:
        return list(self._layouts)

    def _add_layout(self, layout: SlideLayout) -> None:
        self._layouts.append(layout)


class SlideLayout:
    """A slide layout defining placeholder positions."""

    def __init__(self, part: Part, slide_master: SlideMaster | None = None) -> None:
        self._part = part
        self._slide_master = slide_master
        self._element: ET.Element | None = None

    @property
    def element(self) -> ET.Element:
        if self._element is None:
            self._element = parse_xml(self._part.blob)
        return self._element

    @property
    def name(self) -> str:
        csld = self.element.find(qn("p:cSld"))
        if csld is not None:
            return csld.get("name", "")
        return ""

    @property
    def slide_master(self) -> SlideMaster | None:
        return self._slide_master

    @property
    def part(self) -> Part:
        return self._part


class Slide:
    """A single slide in a presentation."""

    def __init__(
        self,
        part: Part,
        slide_layout: SlideLayout | None = None,
        slide_id: int = 0,
        presentation: Presentation | None = None,
    ) -> None:
        self._part = part
        self._slide_layout = slide_layout
        self._slide_id = slide_id
        self._presentation = presentation
        self._element: ET.Element | None = None

    @property
    def element(self) -> ET.Element:
        if self._element is None:
            self._element = parse_xml(self._part.blob)
        return self._element

    @property
    def slide_layout(self) -> SlideLayout | None:
        return self._slide_layout

    @property
    def slide_id(self) -> int:
        return self._slide_id

    @property
    def name(self) -> str:
        csld = self.element.find(qn("p:cSld"))
        if csld is not None:
            return csld.get("name", "")
        return ""

    @name.setter
    def name(self, value: str) -> None:
        csld = self.element.find(qn("p:cSld"))
        if csld is not None:
            csld.set("name", value)
            self._sync_blob()

    @property
    def part(self) -> Part:
        return self._part

    @property
    def shapes(self) -> ShapeCollection:
        """Get shapes on this slide."""
        sp_tree = self.sp_tree
        if sp_tree is None:
            raise ValueError("Slide has no shape tree")
        return ShapeCollection(sp_tree, self)

    @property
    def sp_tree(self) -> ET.Element | None:
        """Get the shape tree element."""
        csld = self.element.find(qn("p:cSld"))
        if csld is not None:
            return csld.find(qn("p:spTree"))
        return None

    def _sync_blob(self) -> None:
        """Write current element back to part blob."""
        if self._element is not None:
            self._part.blob = serialize_xml(self._element)
