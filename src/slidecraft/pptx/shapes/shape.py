"""Base shape and shape collection."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from collections.abc import Iterator
from typing import TYPE_CHECKING

from slidecraft.pptx.enum import MSO_SHAPE_TYPE
from slidecraft.pptx.text import TextFrame
from slidecraft.util.units import Emu
from slidecraft.xml.ns import CT, RT, qn

if TYPE_CHECKING:
    from slidecraft.pptx.shapes.picture import Image
    from slidecraft.pptx.shapes.table import Table
    from slidecraft.pptx.slide import Slide


class BaseShape:
    """Base class for all shape types."""

    def __init__(self, element: ET.Element, slide: Slide | None = None) -> None:
        self._element = element
        self._slide = slide

    @property
    def shape_id(self) -> int:
        nvpr = self._get_nv_pr()
        if nvpr is not None:
            cnvpr = nvpr.find(qn("p:cNvPr"))
            if cnvpr is not None:
                return int(cnvpr.get("id", "0"))
        return 0

    @property
    def name(self) -> str:
        nvpr = self._get_nv_pr()
        if nvpr is not None:
            cnvpr = nvpr.find(qn("p:cNvPr"))
            if cnvpr is not None:
                return cnvpr.get("name", "")
        return ""

    @name.setter
    def name(self, value: str) -> None:
        nvpr = self._get_nv_pr()
        if nvpr is not None:
            cnvpr = nvpr.find(qn("p:cNvPr"))
            if cnvpr is not None:
                cnvpr.set("name", value)

    @property
    def shape_type(self) -> MSO_SHAPE_TYPE:
        tag = self._element.tag
        local = tag.split("}")[-1] if "}" in tag else tag
        if local == "sp":
            # Check if it has a placeholder
            nvpr = self._get_nv_pr()
            if nvpr is not None:
                nv_pr = nvpr.find(qn("p:nvPr"))
                if nv_pr is not None and nv_pr.find(qn("p:ph")) is not None:
                    return MSO_SHAPE_TYPE.PLACEHOLDER
            return MSO_SHAPE_TYPE.AUTO_SHAPE
        if local == "pic":
            return MSO_SHAPE_TYPE.PICTURE
        if local == "graphicFrame":
            return MSO_SHAPE_TYPE.TABLE
        if local == "grpSp":
            return MSO_SHAPE_TYPE.GROUP
        return MSO_SHAPE_TYPE.AUTO_SHAPE

    @property
    def left(self) -> Emu:
        off = self._get_offset()
        return Emu(int(off.get("x", "0"))) if off is not None else Emu(0)

    @left.setter
    def left(self, value: int) -> None:
        off = self._get_offset()
        if off is not None:
            off.set("x", str(int(value)))

    @property
    def top(self) -> Emu:
        off = self._get_offset()
        return Emu(int(off.get("y", "0"))) if off is not None else Emu(0)

    @top.setter
    def top(self, value: int) -> None:
        off = self._get_offset()
        if off is not None:
            off.set("y", str(int(value)))

    @property
    def width(self) -> Emu:
        ext = self._get_extent()
        return Emu(int(ext.get("cx", "0"))) if ext is not None else Emu(0)

    @width.setter
    def width(self, value: int) -> None:
        ext = self._get_extent()
        if ext is not None:
            ext.set("cx", str(int(value)))

    @property
    def height(self) -> Emu:
        ext = self._get_extent()
        return Emu(int(ext.get("cy", "0"))) if ext is not None else Emu(0)

    @height.setter
    def height(self, value: int) -> None:
        ext = self._get_extent()
        if ext is not None:
            ext.set("cy", str(int(value)))

    @property
    def has_text_frame(self) -> bool:
        return self._element.find(qn("p:txBody")) is not None

    @property
    def text_frame(self) -> TextFrame:
        txbody = self._element.find(qn("p:txBody"))
        if txbody is None:
            raise ValueError("Shape does not have a text frame")
        return TextFrame(txbody)

    @property
    def element(self) -> ET.Element:
        return self._element

    def _get_nv_pr(self) -> ET.Element | None:
        """Get the non-visual properties container (nvSpPr, nvPicPr, etc.)."""
        for tag in ("p:nvSpPr", "p:nvPicPr", "p:nvGrpSpPr", "p:nvGraphicFramePr"):
            el = self._element.find(qn(tag))
            if el is not None:
                return el
        return None

    def _get_xfrm(self) -> ET.Element | None:
        sppr = self._element.find(qn("p:spPr"))
        if sppr is not None:
            return sppr.find(qn("a:xfrm"))
        return None

    def _get_offset(self) -> ET.Element | None:
        xfrm = self._get_xfrm()
        if xfrm is not None:
            return xfrm.find(qn("a:off"))
        return None

    def _get_extent(self) -> ET.Element | None:
        xfrm = self._get_xfrm()
        if xfrm is not None:
            return xfrm.find(qn("a:ext"))
        return None


class ShapeCollection:
    """Collection of shapes on a slide."""

    def __init__(self, sp_tree: ET.Element, slide: Slide | None = None) -> None:
        self._sp_tree = sp_tree
        self._slide = slide

    def _next_id(self) -> int:
        """Get the next available shape ID."""
        max_id = 1
        for shape in self:
            if shape.shape_id > max_id:
                max_id = shape.shape_id
        return max_id + 1

    def add_textbox(
        self, left: int, top: int, width: int, height: int
    ) -> BaseShape:
        """Add a text box shape."""
        shape_id = self._next_id()
        sp = _make_textbox_element(shape_id, left, top, width, height)
        self._sp_tree.append(sp)
        return BaseShape(sp, self._slide)

    def add_shape(
        self,
        auto_shape_type: int,
        left: int,
        top: int,
        width: int,
        height: int,
    ) -> BaseShape:
        """Add an auto shape."""
        shape_id = self._next_id()
        sp = _make_auto_shape_element(
            shape_id, auto_shape_type, left, top, width, height
        )
        self._sp_tree.append(sp)
        return BaseShape(sp, self._slide)

    def add_table(
        self, rows: int, cols: int, left: int, top: int, width: int, height: int
    ) -> Table:
        """Add a table shape."""
        from slidecraft.pptx.shapes.table import Table, make_table_element

        shape_id = self._next_id()
        gf = make_table_element(shape_id, rows, cols, left, top, width, height)
        self._sp_tree.append(gf)
        return Table(gf, self._slide)

    def add_picture(
        self,
        image: Image,
        left: int,
        top: int,
        width: int | None = None,
        height: int | None = None,
    ) -> BaseShape:
        """Add a picture shape. Image must be an Image object."""
        from slidecraft.opc.part import Part
        from slidecraft.pptx.shapes.picture import make_picture_element

        if self._slide is None:
            raise ValueError("Cannot add picture without a slide context")

        # Determine dimensions (default to image native size at 96 DPI)
        if width is None:
            width = int(image.width * 914400 / 96) if image.width else 914400
        if height is None:
            height = int(image.height * 914400 / 96) if image.height else 914400

        # Add image as a part in the package
        prs = self._slide._presentation
        if prs is None:
            raise ValueError("Slide not attached to presentation")

        img_num = len([
            p for p in prs._pkg.parts
            if p.startswith("/ppt/media/")
        ]) + 1
        img_path = f"/ppt/media/image{img_num}.{image.ext}"
        img_part = Part(img_path, image.content_type, image.blob)
        prs._pkg.add_part(img_part)

        # Add content type default
        prs._pkg.content_types.add_default(image.ext, image.content_type)

        # Add relationship from slide to image
        rel_target = f"../media/image{img_num}.{image.ext}"
        rel = self._slide._part.rels.add(RT["IMAGE"], rel_target)

        shape_id = self._next_id()
        pic = make_picture_element(
            shape_id, rel.r_id, f"Picture {shape_id}", left, top, width, height
        )
        self._sp_tree.append(pic)
        return BaseShape(pic, self._slide)

    def __iter__(self) -> Iterator[BaseShape]:
        for child in self._sp_tree:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag in ("sp", "pic", "graphicFrame", "grpSp"):
                yield BaseShape(child, self._slide)

    def __len__(self) -> int:
        return sum(1 for _ in self)

    def __getitem__(self, index: int) -> BaseShape:
        shapes = list(self)
        return shapes[index]


def _make_textbox_element(
    shape_id: int, left: int, top: int, width: int, height: int
) -> ET.Element:
    """Create XML element for a text box shape."""
    sp = ET.Element(qn("p:sp"))

    # nvSpPr
    nv = ET.SubElement(sp, qn("p:nvSpPr"))
    cnvpr = ET.SubElement(nv, qn("p:cNvPr"))
    cnvpr.set("id", str(shape_id))
    cnvpr.set("name", f"TextBox {shape_id}")
    cnvsppr = ET.SubElement(nv, qn("p:cNvSpPr"))
    cnvsppr.set("txBox", "1")
    ET.SubElement(nv, qn("p:nvPr"))

    # spPr
    sppr = ET.SubElement(sp, qn("p:spPr"))
    xfrm = ET.SubElement(sppr, qn("a:xfrm"))
    off = ET.SubElement(xfrm, qn("a:off"))
    off.set("x", str(int(left)))
    off.set("y", str(int(top)))
    ext = ET.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(int(width)))
    ext.set("cy", str(int(height)))
    prstgeom = ET.SubElement(sppr, qn("a:prstGeom"))
    prstgeom.set("prst", "rect")

    # txBody
    txbody = ET.SubElement(sp, qn("p:txBody"))
    ET.SubElement(txbody, qn("a:bodyPr"))
    ET.SubElement(txbody, qn("a:lstStyle"))
    ET.SubElement(txbody, qn("a:p"))

    return sp


def _make_auto_shape_element(
    shape_id: int,
    preset: int,
    left: int,
    top: int,
    width: int,
    height: int,
) -> ET.Element:
    """Create XML element for an auto shape."""
    # Map common preset values to names
    preset_names: dict[int, str] = {
        1: "rect",
        2: "roundRect",
        3: "ellipse",
        4: "diamond",
        5: "triangle",
        6: "rtTriangle",
        7: "parallelogram",
        8: "trapezoid",
        9: "pentagon",
        10: "hexagon",
    }
    prst = preset_names.get(preset, "rect")

    sp = ET.Element(qn("p:sp"))

    nv = ET.SubElement(sp, qn("p:nvSpPr"))
    cnvpr = ET.SubElement(nv, qn("p:cNvPr"))
    cnvpr.set("id", str(shape_id))
    cnvpr.set("name", f"Shape {shape_id}")
    ET.SubElement(nv, qn("p:cNvSpPr"))
    ET.SubElement(nv, qn("p:nvPr"))

    sppr = ET.SubElement(sp, qn("p:spPr"))
    xfrm = ET.SubElement(sppr, qn("a:xfrm"))
    off = ET.SubElement(xfrm, qn("a:off"))
    off.set("x", str(int(left)))
    off.set("y", str(int(top)))
    ext = ET.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(int(width)))
    ext.set("cy", str(int(height)))
    prstgeom = ET.SubElement(sppr, qn("a:prstGeom"))
    prstgeom.set("prst", prst)

    # Auto shapes also get a text body
    txbody = ET.SubElement(sp, qn("p:txBody"))
    ET.SubElement(txbody, qn("a:bodyPr"))
    ET.SubElement(txbody, qn("a:lstStyle"))
    ET.SubElement(txbody, qn("a:p"))

    return sp
