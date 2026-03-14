"""Tests for GroupShape and Placeholder."""

from __future__ import annotations

import xml.etree.ElementTree as ET

from slidecraft.pptx.enum import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from slidecraft.pptx.shapes.group import GroupShape
from slidecraft.pptx.shapes.placeholder import Placeholder, PlaceholderFormat
from slidecraft.xml.ns import qn


def _make_group_element(
    shape_id: int = 1,
    name: str = "Group 1",
) -> ET.Element:
    """Build a minimal p:grpSp XML element."""
    grp = ET.Element(qn("p:grpSp"))

    nv = ET.SubElement(grp, qn("p:nvGrpSpPr"))
    cnvpr = ET.SubElement(nv, qn("p:cNvPr"))
    cnvpr.set("id", str(shape_id))
    cnvpr.set("name", name)
    ET.SubElement(nv, qn("p:cNvGrpSpPr"))
    ET.SubElement(nv, qn("p:nvPr"))

    grppr = ET.SubElement(grp, qn("p:grpSpPr"))
    xfrm = ET.SubElement(grppr, qn("a:xfrm"))
    off = ET.SubElement(xfrm, qn("a:off"))
    off.set("x", "0")
    off.set("y", "0")
    ext = ET.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", "1000000")
    ext.set("cy", "1000000")
    choff = ET.SubElement(xfrm, qn("a:chOff"))
    choff.set("x", "0")
    choff.set("y", "0")
    chext = ET.SubElement(xfrm, qn("a:chExt"))
    chext.set("cx", "1000000")
    chext.set("cy", "1000000")

    return grp


def _add_child_sp(parent: ET.Element, shape_id: int, name: str) -> ET.Element:
    """Add a child p:sp to a group element."""
    sp = ET.SubElement(parent, qn("p:sp"))

    nv = ET.SubElement(sp, qn("p:nvSpPr"))
    cnvpr = ET.SubElement(nv, qn("p:cNvPr"))
    cnvpr.set("id", str(shape_id))
    cnvpr.set("name", name)
    ET.SubElement(nv, qn("p:cNvSpPr"))
    ET.SubElement(nv, qn("p:nvPr"))

    sppr = ET.SubElement(sp, qn("p:spPr"))
    xfrm = ET.SubElement(sppr, qn("a:xfrm"))
    off = ET.SubElement(xfrm, qn("a:off"))
    off.set("x", "100")
    off.set("y", "200")
    ext = ET.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", "300")
    ext.set("cy", "400")
    prstgeom = ET.SubElement(sppr, qn("a:prstGeom"))
    prstgeom.set("prst", "rect")

    return sp


def _make_placeholder_sp(
    shape_id: int = 1,
    name: str = "Title 1",
    ph_type: str = "title",
    ph_idx: str = "0",
) -> ET.Element:
    """Build a minimal p:sp with placeholder properties."""
    sp = ET.Element(qn("p:sp"))

    nv = ET.SubElement(sp, qn("p:nvSpPr"))
    cnvpr = ET.SubElement(nv, qn("p:cNvPr"))
    cnvpr.set("id", str(shape_id))
    cnvpr.set("name", name)
    ET.SubElement(nv, qn("p:cNvSpPr"))
    nvpr = ET.SubElement(nv, qn("p:nvPr"))
    ph = ET.SubElement(nvpr, qn("p:ph"))
    ph.set("type", ph_type)
    ph.set("idx", ph_idx)

    sppr = ET.SubElement(sp, qn("p:spPr"))
    xfrm = ET.SubElement(sppr, qn("a:xfrm"))
    off = ET.SubElement(xfrm, qn("a:off"))
    off.set("x", "0")
    off.set("y", "0")
    ext = ET.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", "500000")
    ext.set("cy", "300000")

    return sp


class TestGroupShape:
    def test_empty_group(self) -> None:
        grp_el = _make_group_element()
        group = GroupShape(grp_el)
        assert len(group.shapes) == 0

    def test_group_with_children(self) -> None:
        grp_el = _make_group_element()
        _add_child_sp(grp_el, 2, "Rect 1")
        _add_child_sp(grp_el, 3, "Rect 2")
        group = GroupShape(grp_el)
        assert len(group.shapes) == 2

    def test_group_child_names(self) -> None:
        grp_el = _make_group_element()
        _add_child_sp(grp_el, 2, "Shape A")
        _add_child_sp(grp_el, 3, "Shape B")
        group = GroupShape(grp_el)
        names = [s.name for s in group.shapes]
        assert names == ["Shape A", "Shape B"]

    def test_group_shape_id(self) -> None:
        grp_el = _make_group_element(shape_id=10, name="My Group")
        group = GroupShape(grp_el)
        assert group.shape_id == 10
        assert group.name == "My Group"

    def test_group_shape_type(self) -> None:
        grp_el = _make_group_element()
        group = GroupShape(grp_el)
        assert group.shape_type == MSO_SHAPE_TYPE.GROUP

    def test_group_iterate_children(self) -> None:
        grp_el = _make_group_element()
        _add_child_sp(grp_el, 2, "Child 1")
        _add_child_sp(grp_el, 3, "Child 2")
        _add_child_sp(grp_el, 4, "Child 3")
        group = GroupShape(grp_el)
        ids = [s.shape_id for s in group.shapes]
        assert ids == [2, 3, 4]

    def test_group_getitem(self) -> None:
        grp_el = _make_group_element()
        _add_child_sp(grp_el, 2, "First")
        _add_child_sp(grp_el, 3, "Second")
        group = GroupShape(grp_el)
        assert group.shapes[0].name == "First"
        assert group.shapes[1].name == "Second"


class TestPlaceholder:
    def test_title_placeholder(self) -> None:
        sp = _make_placeholder_sp(ph_type="title", ph_idx="0")
        ph = Placeholder(sp)
        fmt = ph.placeholder_format
        assert fmt is not None
        assert fmt.type == PP_PLACEHOLDER.TITLE
        assert fmt.idx == 0

    def test_body_placeholder(self) -> None:
        sp = _make_placeholder_sp(ph_type="body", ph_idx="1")
        ph = Placeholder(sp)
        fmt = ph.placeholder_format
        assert fmt is not None
        assert fmt.type == PP_PLACEHOLDER.BODY
        assert fmt.idx == 1

    def test_subtitle_placeholder(self) -> None:
        sp = _make_placeholder_sp(ph_type="subTitle", ph_idx="1")
        ph = Placeholder(sp)
        fmt = ph.placeholder_format
        assert fmt is not None
        assert fmt.type == PP_PLACEHOLDER.SUBTITLE

    def test_center_title_placeholder(self) -> None:
        sp = _make_placeholder_sp(ph_type="ctrTitle", ph_idx="0")
        ph = Placeholder(sp)
        fmt = ph.placeholder_format
        assert fmt is not None
        assert fmt.type == PP_PLACEHOLDER.CENTER_TITLE

    def test_slide_number_placeholder(self) -> None:
        sp = _make_placeholder_sp(ph_type="sldNum", ph_idx="12")
        ph = Placeholder(sp)
        fmt = ph.placeholder_format
        assert fmt is not None
        assert fmt.type == PP_PLACEHOLDER.SLIDE_NUMBER
        assert fmt.idx == 12

    def test_footer_placeholder(self) -> None:
        sp = _make_placeholder_sp(ph_type="ftr", ph_idx="11")
        ph = Placeholder(sp)
        fmt = ph.placeholder_format
        assert fmt is not None
        assert fmt.type == PP_PLACEHOLDER.FOOTER

    def test_date_placeholder(self) -> None:
        sp = _make_placeholder_sp(ph_type="dt", ph_idx="10")
        ph = Placeholder(sp)
        fmt = ph.placeholder_format
        assert fmt is not None
        assert fmt.type == PP_PLACEHOLDER.DATE

    def test_object_placeholder(self) -> None:
        sp = _make_placeholder_sp(ph_type="obj", ph_idx="7")
        ph = Placeholder(sp)
        fmt = ph.placeholder_format
        assert fmt is not None
        assert fmt.type == PP_PLACEHOLDER.OBJECT

    def test_unknown_placeholder_type(self) -> None:
        sp = _make_placeholder_sp(ph_type="unknown", ph_idx="99")
        ph = Placeholder(sp)
        fmt = ph.placeholder_format
        assert fmt is not None
        assert fmt.type is None
        assert fmt.idx == 99

    def test_no_placeholder(self) -> None:
        """Shape without placeholder properties returns None."""
        sp = ET.Element(qn("p:sp"))
        nv = ET.SubElement(sp, qn("p:nvSpPr"))
        cnvpr = ET.SubElement(nv, qn("p:cNvPr"))
        cnvpr.set("id", "1")
        cnvpr.set("name", "Shape 1")
        ET.SubElement(nv, qn("p:cNvSpPr"))
        ET.SubElement(nv, qn("p:nvPr"))  # no p:ph child
        ET.SubElement(sp, qn("p:spPr"))
        ph = Placeholder(sp)
        assert ph.placeholder_format is None

    def test_placeholder_shape_type(self) -> None:
        sp = _make_placeholder_sp(ph_type="title", ph_idx="0")
        ph = Placeholder(sp)
        assert ph.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER

    def test_placeholder_shape_properties(self) -> None:
        sp = _make_placeholder_sp(shape_id=5, name="Title Placeholder")
        ph = Placeholder(sp)
        assert ph.shape_id == 5
        assert ph.name == "Title Placeholder"


class TestPlaceholderFormat:
    def test_idx_default(self) -> None:
        """idx defaults to 0 when not specified."""
        ph_el = ET.Element(qn("p:ph"))
        ph_el.set("type", "title")
        fmt = PlaceholderFormat(ph_el)
        assert fmt.idx == 0

    def test_type_none_when_missing(self) -> None:
        """type returns None when type attribute is absent."""
        ph_el = ET.Element(qn("p:ph"))
        fmt = PlaceholderFormat(ph_el)
        assert fmt.type is None
