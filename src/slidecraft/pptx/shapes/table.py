"""Table shape support."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from typing import TYPE_CHECKING

from slidecraft.pptx.text import TextFrame
from slidecraft.util.units import Emu
from slidecraft.xml.ns import qn

if TYPE_CHECKING:
    from slidecraft.pptx.slide import Slide


class _Cell:
    """A single cell in a table."""

    def __init__(self, tc: ET.Element) -> None:
        self._element = tc

    @property
    def text_frame(self) -> TextFrame:
        txbody = self._element.find(qn("a:txBody"))
        if txbody is None:
            txbody = ET.SubElement(self._element, qn("a:txBody"))
            ET.SubElement(txbody, qn("a:bodyPr"))
            ET.SubElement(txbody, qn("a:p"))
        return TextFrame(txbody)

    @property
    def text(self) -> str:
        return self.text_frame.text

    @text.setter
    def text(self, value: str) -> None:
        self.text_frame.text = value

    @property
    def is_merge_origin(self) -> bool:
        return (
            self._element.get("gridSpan") is not None
            or self._element.get("rowSpan") is not None
        )

    @property
    def is_spanned(self) -> bool:
        return (
            self._element.get("hMerge") == "1"
            or self._element.get("vMerge") == "1"
        )

    def merge(self, other: _Cell) -> None:
        """Merge this cell with another (same row for horizontal merge)."""
        # Simple horizontal merge support
        self._element.set("gridSpan", "2")

    @property
    def element(self) -> ET.Element:
        return self._element


class _Row:
    """A row in a table."""

    def __init__(self, tr: ET.Element) -> None:
        self._element = tr

    @property
    def cells(self) -> list[_Cell]:
        return [_Cell(tc) for tc in self._element.findall(qn("a:tc"))]

    @property
    def height(self) -> Emu:
        return Emu(int(self._element.get("h", "370840")))

    @height.setter
    def height(self, value: int) -> None:
        self._element.set("h", str(int(value)))

    @property
    def element(self) -> ET.Element:
        return self._element


class _Column:
    """A column in a table."""

    def __init__(self, grid_col: ET.Element) -> None:
        self._element = grid_col

    @property
    def width(self) -> Emu:
        return Emu(int(self._element.get("w", "0")))

    @width.setter
    def width(self, value: int) -> None:
        self._element.set("w", str(int(value)))


class Table:
    """A table shape."""

    def __init__(self, graphic_frame: ET.Element, slide: Slide | None = None) -> None:
        self._element = graphic_frame
        self._slide = slide

    @property
    def _tbl(self) -> ET.Element:
        graphic = self._element.find(qn("a:graphic"))
        if graphic is None:
            raise ValueError("No graphic element")
        gd = graphic.find(qn("a:graphicData"))
        if gd is None:
            raise ValueError("No graphicData element")
        tbl = gd.find(qn("a:tbl"))
        if tbl is None:
            raise ValueError("No table element")
        return tbl

    @property
    def rows(self) -> list[_Row]:
        return [_Row(tr) for tr in self._tbl.findall(qn("a:tr"))]

    @property
    def columns(self) -> list[_Column]:
        grid = self._tbl.find(qn("a:tblGrid"))
        if grid is None:
            return []
        return [_Column(gc) for gc in grid.findall(qn("a:gridCol"))]

    def cell(self, row: int, col: int) -> _Cell:
        """Get a cell by row and column index."""
        rows = self.rows
        if row < 0 or row >= len(rows):
            raise IndexError(f"Row {row} out of range")
        cells = rows[row].cells
        if col < 0 or col >= len(cells):
            raise IndexError(f"Column {col} out of range")
        return cells[col]

    @property
    def element(self) -> ET.Element:
        return self._element


def make_table_element(
    shape_id: int,
    rows: int,
    cols: int,
    left: int,
    top: int,
    width: int,
    height: int,
) -> ET.Element:
    """Create a graphicFrame element containing a table."""
    gf = ET.Element(qn("p:graphicFrame"))

    # nvGraphicFramePr
    nvgfpr = ET.SubElement(gf, qn("p:nvGraphicFramePr"))
    cnvpr = ET.SubElement(nvgfpr, qn("p:cNvPr"))
    cnvpr.set("id", str(shape_id))
    cnvpr.set("name", f"Table {shape_id}")
    cnvgfpr = ET.SubElement(nvgfpr, qn("p:cNvGraphicFramePr"))
    locks = ET.SubElement(cnvgfpr, qn("a:graphicFrameLocks"))
    locks.set("noGrp", "1")
    ET.SubElement(nvgfpr, qn("p:nvPr"))

    # xfrm
    xfrm = ET.SubElement(gf, qn("p:xfrm"))
    off = ET.SubElement(xfrm, qn("a:off"))
    off.set("x", str(int(left)))
    off.set("y", str(int(top)))
    ext = ET.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(int(width)))
    ext.set("cy", str(int(height)))

    # graphic
    graphic = ET.SubElement(gf, qn("a:graphic"))
    gd = ET.SubElement(graphic, qn("a:graphicData"))
    gd.set("uri", "http://schemas.openxmlformats.org/drawingml/2006/table")

    tbl = ET.SubElement(gd, qn("a:tbl"))
    tbl_pr = ET.SubElement(tbl, qn("a:tblPr"))
    tbl_pr.set("firstRow", "1")
    tbl_pr.set("bandRow", "1")

    # Grid
    grid = ET.SubElement(tbl, qn("a:tblGrid"))
    col_width = width // cols
    for _ in range(cols):
        gc = ET.SubElement(grid, qn("a:gridCol"))
        gc.set("w", str(col_width))

    # Rows
    row_height = height // rows
    for _ in range(rows):
        tr = ET.SubElement(tbl, qn("a:tr"))
        tr.set("h", str(row_height))
        for _ in range(cols):
            tc = ET.SubElement(tr, qn("a:tc"))
            txbody = ET.SubElement(tc, qn("a:txBody"))
            ET.SubElement(txbody, qn("a:bodyPr"))
            ET.SubElement(txbody, qn("a:p"))
            tc_pr = ET.SubElement(tc, qn("a:tcPr"))
            tc_pr.set("marL", "91440")
            tc_pr.set("marR", "91440")
            tc_pr.set("marT", "45720")
            tc_pr.set("marB", "45720")

    return gf
