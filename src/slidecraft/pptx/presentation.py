"""Presentation class — top-level entry point."""

from __future__ import annotations

import io
import os
import xml.etree.ElementTree as ET
from collections.abc import Iterator

from slidecraft.opc.package import OpcPackage
from slidecraft.opc.part import Part
from slidecraft.pptx.defaults import (
    PRESENTATION_XML,
    SLIDE_LAYOUT_XML,
    SLIDE_MASTER_XML,
    SLIDE_XML,
    THEME_XML,
)
from slidecraft.pptx.slide import Slide, SlideLayout, SlideMaster
from slidecraft.util.units import Emu
from slidecraft.xml.ns import CT, RT, qn
from slidecraft.xml.parse import parse_xml, serialize_xml


class Presentation:
    """A PowerPoint presentation.

    Create a new blank presentation::

        prs = Presentation()

    Open an existing .pptx file::

        prs = Presentation.open("my.pptx")
    """

    def __init__(self) -> None:
        self._pkg = _new_blank_package()
        pres_part = self._pkg.get_part("/ppt/presentation.xml")
        assert pres_part is not None
        self._element = parse_xml(pres_part.blob)
        self._slides = SlideCollection(self)
        self._slide_masters: list[SlideMaster] = []
        self._slide_layouts: list[SlideLayout] = []
        self._load_masters_and_layouts()

    @classmethod
    def open(cls, path: str | os.PathLike[str] | io.BytesIO) -> Presentation:
        """Open an existing .pptx file."""
        prs = object.__new__(cls)
        prs._pkg = OpcPackage.open(path)
        pres_part = prs._pkg.get_part("/ppt/presentation.xml")
        if pres_part is None:
            raise ValueError("Not a valid .pptx file: missing presentation.xml")
        prs._element = parse_xml(pres_part.blob)
        prs._slides = SlideCollection(prs)
        prs._slide_masters = []
        prs._slide_layouts = []
        prs._load_masters_and_layouts()
        prs._load_existing_slides()
        return prs

    def save(self, path: str | os.PathLike[str] | io.BytesIO) -> None:
        """Save the presentation to a .pptx file."""
        # Sync slide element trees back to their part blobs
        for slide in self._slides:
            slide._sync_blob()
        # Sync presentation XML back to part
        pres_part = self._pkg.get_part("/ppt/presentation.xml")
        if pres_part is not None:
            pres_part.blob = serialize_xml(self._element)
        self._pkg.save(path)

    @property
    def slides(self) -> SlideCollection:
        return self._slides

    @property
    def slide_width(self) -> Emu:
        sz = self._element.find(qn("p:sldSz"))
        if sz is not None:
            return Emu(int(sz.get("cx", "12192000")))
        return Emu(12192000)

    @slide_width.setter
    def slide_width(self, value: int) -> None:
        sz = self._element.find(qn("p:sldSz"))
        if sz is not None:
            sz.set("cx", str(int(value)))

    @property
    def slide_height(self) -> Emu:
        sz = self._element.find(qn("p:sldSz"))
        if sz is not None:
            return Emu(int(sz.get("cy", "6858000")))
        return Emu(6858000)

    @slide_height.setter
    def slide_height(self, value: int) -> None:
        sz = self._element.find(qn("p:sldSz"))
        if sz is not None:
            sz.set("cy", str(int(value)))

    @property
    def slide_layouts(self) -> list[SlideLayout]:
        return list(self._slide_layouts)

    @property
    def slide_masters(self) -> list[SlideMaster]:
        return list(self._slide_masters)

    def _load_masters_and_layouts(self) -> None:
        """Load slide masters and layouts from the package."""
        pres_part = self._pkg.get_part("/ppt/presentation.xml")
        if pres_part is None:
            return

        # Find slide masters from presentation rels
        for rel in pres_part.rels.get_by_type(RT["SLIDE_MASTER"]):
            master_path = _resolve_rel_target("/ppt/presentation.xml", rel.target)
            master_part = self._pkg.get_part(master_path)
            if master_part is None:
                continue
            master = SlideMaster(master_part)
            self._slide_masters.append(master)

            # Find layouts from master rels
            for layout_rel in master_part.rels.get_by_type(RT["SLIDE_LAYOUT"]):
                layout_path = _resolve_rel_target(master_path, layout_rel.target)
                layout_part = self._pkg.get_part(layout_path)
                if layout_part is None:
                    continue
                layout = SlideLayout(layout_part, slide_master=master)
                master._add_layout(layout)
                self._slide_layouts.append(layout)

    def _load_existing_slides(self) -> None:
        """Load existing slides from the package."""
        pres_part = self._pkg.get_part("/ppt/presentation.xml")
        if pres_part is None:
            return

        sld_id_lst = self._element.find(qn("p:sldIdLst"))
        if sld_id_lst is None:
            return

        for sld_id_elem in sld_id_lst.findall(qn("p:sldId")):
            r_id = sld_id_elem.get(qn("r:id"), "")
            slide_id = int(sld_id_elem.get("id", "0"))
            rel = pres_part.rels.get_by_id(r_id)
            if rel is None:
                continue
            slide_path = _resolve_rel_target("/ppt/presentation.xml", rel.target)
            slide_part = self._pkg.get_part(slide_path)
            if slide_part is None:
                continue

            # Find the layout for this slide
            layout: SlideLayout | None = None
            for layout_rel in slide_part.rels.get_by_type(RT["SLIDE_LAYOUT"]):
                layout_path = _resolve_rel_target(slide_path, layout_rel.target)
                for sl in self._slide_layouts:
                    if sl.part.part_name == layout_path:
                        layout = sl
                        break
                break

            slide = Slide(slide_part, layout, slide_id, self)
            self._slides._slides.append(slide)


class SlideCollection:
    """Collection of slides in a presentation."""

    def __init__(self, presentation: Presentation) -> None:
        self._presentation = presentation
        self._slides: list[Slide] = []

    def add(self, layout: SlideLayout | None = None) -> Slide:
        """Add a new slide using the given layout."""
        if layout is None and self._presentation._slide_layouts:
            layout = self._presentation._slide_layouts[0]

        prs = self._presentation
        slide_num = len(self._slides) + 1
        slide_path = f"/ppt/slides/slide{slide_num}.xml"

        # Find a unique slide ID
        used_ids = {s.slide_id for s in self._slides}
        slide_id = 256
        while slide_id in used_ids:
            slide_id += 1

        # Create slide part
        slide_part = Part(slide_path, CT["SLIDE"], SLIDE_XML)
        prs._pkg.add_part(slide_part)

        # Add relationship from slide to layout
        if layout is not None:
            # Compute relative path from slide to layout
            layout_rel_target = _relative_path(slide_path, layout.part.part_name)
            slide_part.rels.add(RT["SLIDE_LAYOUT"], layout_rel_target)

        # Add relationship from presentation to slide
        pres_part = prs._pkg.get_part("/ppt/presentation.xml")
        if pres_part is not None:
            rel = pres_part.rels.add(RT["SLIDE"], f"slides/slide{slide_num}.xml")

            # Update sldIdLst in presentation XML
            sld_id_lst = prs._element.find(qn("p:sldIdLst"))
            if sld_id_lst is None:
                sld_id_lst = _make_element("p:sldIdLst")
                prs._element.append(sld_id_lst)
            sld_id_elem = _make_element("p:sldId")
            sld_id_elem.set("id", str(slide_id))
            sld_id_elem.set(qn("r:id"), rel.r_id)
            sld_id_lst.append(sld_id_elem)

        slide = Slide(slide_part, layout, slide_id, prs)
        self._slides.append(slide)
        return slide

    def remove(self, index: int) -> None:
        """Remove a slide by index."""
        if index < 0 or index >= len(self._slides):
            raise IndexError(f"Slide index {index} out of range")
        slide = self._slides.pop(index)

        # Remove from package
        self._presentation._pkg.remove_part(slide.part.part_name)

        # Remove from sldIdLst
        sld_id_lst = self._presentation._element.find(qn("p:sldIdLst"))
        if sld_id_lst is not None:
            for sld_id_elem in list(sld_id_lst):
                if sld_id_elem.get("id") == str(slide.slide_id):
                    sld_id_lst.remove(sld_id_elem)
                    break

    def __iter__(self) -> Iterator[Slide]:
        return iter(self._slides)

    def __len__(self) -> int:
        return len(self._slides)

    def __getitem__(self, index: int) -> Slide:
        return self._slides[index]


def _new_blank_package() -> OpcPackage:
    """Create a minimal blank .pptx package."""
    pkg = OpcPackage()

    # Content types
    ct = pkg.content_types
    ct.add_default("rels", CT["RELS"])
    ct.add_default("xml", "application/xml")
    ct.add_default("png", CT["PNG"])
    ct.add_default("jpeg", CT["JPEG"])

    # Presentation part
    pres_part = Part("/ppt/presentation.xml", CT["PRESENTATION"], PRESENTATION_XML)
    pkg.add_part(pres_part)
    pkg.rels.add(RT["OFFICE_DOCUMENT"], "ppt/presentation.xml")

    # Slide master
    master_part = Part("/ppt/slideMasters/slideMaster1.xml", CT["SLIDE_MASTER"], SLIDE_MASTER_XML)
    pkg.add_part(master_part)
    pres_part.rels.add(RT["SLIDE_MASTER"], "slideMasters/slideMaster1.xml")

    # Slide layout
    layout_part = Part("/ppt/slideLayouts/slideLayout1.xml", CT["SLIDE_LAYOUT"], SLIDE_LAYOUT_XML)
    pkg.add_part(layout_part)
    master_part.rels.add(RT["SLIDE_LAYOUT"], "../slideLayouts/slideLayout1.xml")
    layout_part.rels.add(RT["SLIDE_MASTER"], "../slideMasters/slideMaster1.xml")

    # Theme
    theme_part = Part("/ppt/theme/theme1.xml", CT["THEME"], THEME_XML)
    pkg.add_part(theme_part)
    master_part.rels.add(RT["THEME"], "../theme/theme1.xml")

    return pkg


def _resolve_rel_target(source_path: str, target: str) -> str:
    """Resolve a relative relationship target to an absolute part name."""
    if target.startswith("/"):
        return target
    source_dir = source_path.rsplit("/", 1)[0]
    parts = (source_dir + "/" + target).split("/")
    resolved: list[str] = []
    for p in parts:
        if p == "..":
            if resolved:
                resolved.pop()
        elif p and p != ".":
            resolved.append(p)
    return "/" + "/".join(resolved)


def _relative_path(from_path: str, to_path: str) -> str:
    """Compute a relative path from one part to another."""
    from_parts = from_path.strip("/").split("/")[:-1]  # directory of from
    to_parts = to_path.strip("/").split("/")

    # Find common prefix
    common = 0
    for a, b in zip(from_parts, to_parts):
        if a != b:
            break
        common += 1

    ups = len(from_parts) - common
    rel_parts = [".."] * ups + to_parts[common:]
    return "/".join(rel_parts)


def _make_element(tag: str) -> ET.Element:
    """Create an element from a prefixed tag name like 'p:sldIdLst'."""
    return ET.Element(qn(tag) if ":" in tag else tag)
