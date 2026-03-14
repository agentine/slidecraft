"""Chart shape support."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from enum import IntEnum
from typing import TYPE_CHECKING

from slidecraft.pptx.enum import ChartType
from slidecraft.xml.ns import NS, qn

if TYPE_CHECKING:
    from slidecraft.pptx.slide import Slide

_C_NS = NS["c"]


class LegendPosition(IntEnum):
    """Legend position."""

    BOTTOM = 0
    TOP = 1
    LEFT = 2
    RIGHT = 3
    CORNER = 4


_LEGEND_POS_XML: dict[LegendPosition, str] = {
    LegendPosition.BOTTOM: "b",
    LegendPosition.TOP: "t",
    LegendPosition.LEFT: "l",
    LegendPosition.RIGHT: "r",
    LegendPosition.CORNER: "tr",
}


class ChartSeries:
    """A data series for a chart."""

    def __init__(self, name: str, values: list[float]) -> None:
        self.name = name
        self.values = values


class ChartData:
    """Chart data container with categories and series."""

    def __init__(self) -> None:
        self.categories: list[str] = []
        self._series: list[ChartSeries] = []

    def add_series(self, name: str, values: list[float]) -> ChartSeries:
        """Add a data series."""
        s = ChartSeries(name, values)
        self._series.append(s)
        return s

    @property
    def series(self) -> list[ChartSeries]:
        return list(self._series)


class Chart:
    """A chart shape."""

    def __init__(self, graphic_frame: ET.Element, slide: Slide | None = None) -> None:
        self._element = graphic_frame
        self._slide = slide

    @property
    def element(self) -> ET.Element:
        return self._element


def generate_chart_xml(
    chart_type: ChartType,
    chart_data: ChartData,
    *,
    has_legend: bool = True,
    legend_position: LegendPosition = LegendPosition.BOTTOM,
) -> bytes:
    """Generate chart XML for a c:chartSpace element."""
    ET.register_namespace("c", _C_NS)
    ET.register_namespace("a", NS["a"])
    ET.register_namespace("r", NS["r"])

    chart_space = ET.Element(f"{{{_C_NS}}}chartSpace")
    chart = ET.SubElement(chart_space, f"{{{_C_NS}}}chart")

    # Plot area
    plot_area = ET.SubElement(chart, f"{{{_C_NS}}}plotArea")

    # Chart-type specific element
    _make_chart_type_element(chart_type, chart_data, plot_area)

    # Category axis + value axis for non-pie charts
    if chart_type != ChartType.PIE:
        cat_ax = ET.SubElement(plot_area, f"{{{_C_NS}}}catAx")
        ax_id = ET.SubElement(cat_ax, f"{{{_C_NS}}}axId")
        ax_id.set("val", "1")
        scaling = ET.SubElement(cat_ax, f"{{{_C_NS}}}scaling")
        orient = ET.SubElement(scaling, f"{{{_C_NS}}}orientation")
        orient.set("val", "minMax")
        delete = ET.SubElement(cat_ax, f"{{{_C_NS}}}delete")
        delete.set("val", "0")
        axpos = ET.SubElement(cat_ax, f"{{{_C_NS}}}axPos")
        axpos.set("val", "b")
        cross = ET.SubElement(cat_ax, f"{{{_C_NS}}}crossAx")
        cross.set("val", "2")

        val_ax = ET.SubElement(plot_area, f"{{{_C_NS}}}valAx")
        ax_id2 = ET.SubElement(val_ax, f"{{{_C_NS}}}axId")
        ax_id2.set("val", "2")
        scaling2 = ET.SubElement(val_ax, f"{{{_C_NS}}}scaling")
        orient2 = ET.SubElement(scaling2, f"{{{_C_NS}}}orientation")
        orient2.set("val", "minMax")
        delete2 = ET.SubElement(val_ax, f"{{{_C_NS}}}delete")
        delete2.set("val", "0")
        axpos2 = ET.SubElement(val_ax, f"{{{_C_NS}}}axPos")
        axpos2.set("val", "l")
        cross2 = ET.SubElement(val_ax, f"{{{_C_NS}}}crossAx")
        cross2.set("val", "1")

    # Legend
    if has_legend:
        legend = ET.SubElement(chart, f"{{{_C_NS}}}legend")
        pos = ET.SubElement(legend, f"{{{_C_NS}}}legendPos")
        pos.set("val", _LEGEND_POS_XML.get(legend_position, "b"))

    result: bytes = ET.tostring(chart_space, encoding="UTF-8", xml_declaration=True)
    return result


def _make_chart_type_element(
    chart_type: ChartType, chart_data: ChartData, plot_area: ET.Element
) -> ET.Element:
    """Create the chart-type specific element."""
    tag_map: dict[ChartType, str] = {
        ChartType.BAR: "barChart",
        ChartType.BAR_CLUSTERED: "barChart",
        ChartType.COLUMN: "barChart",
        ChartType.COLUMN_CLUSTERED: "barChart",
        ChartType.LINE: "lineChart",
        ChartType.PIE: "pieChart",
    }
    tag = tag_map.get(chart_type, "barChart")
    elem = ET.SubElement(plot_area, f"{{{_C_NS}}}{tag}")

    # Bar/column direction and grouping
    if tag == "barChart":
        bar_dir = ET.SubElement(elem, f"{{{_C_NS}}}barDir")
        if chart_type in (ChartType.BAR, ChartType.BAR_CLUSTERED):
            bar_dir.set("val", "bar")
        else:
            bar_dir.set("val", "col")
        grouping = ET.SubElement(elem, f"{{{_C_NS}}}grouping")
        grouping.set("val", "clustered")
    elif tag == "lineChart":
        grouping = ET.SubElement(elem, f"{{{_C_NS}}}grouping")
        grouping.set("val", "standard")

    # Series
    for idx, series in enumerate(chart_data.series):
        ser = ET.SubElement(elem, f"{{{_C_NS}}}ser")
        ser_idx = ET.SubElement(ser, f"{{{_C_NS}}}idx")
        ser_idx.set("val", str(idx))
        ser_order = ET.SubElement(ser, f"{{{_C_NS}}}order")
        ser_order.set("val", str(idx))

        # Series name
        tx = ET.SubElement(ser, f"{{{_C_NS}}}tx")
        str_ref = ET.SubElement(tx, f"{{{_C_NS}}}strRef")
        str_cache = ET.SubElement(str_ref, f"{{{_C_NS}}}strCache")
        pt_count = ET.SubElement(str_cache, f"{{{_C_NS}}}ptCount")
        pt_count.set("val", "1")
        pt = ET.SubElement(str_cache, f"{{{_C_NS}}}pt")
        pt.set("idx", "0")
        v = ET.SubElement(pt, f"{{{_C_NS}}}v")
        v.text = series.name

        # Category reference
        if chart_data.categories:
            cat = ET.SubElement(ser, f"{{{_C_NS}}}cat")
            cat_str_ref = ET.SubElement(cat, f"{{{_C_NS}}}strRef")
            cat_cache = ET.SubElement(cat_str_ref, f"{{{_C_NS}}}strCache")
            cat_count = ET.SubElement(cat_cache, f"{{{_C_NS}}}ptCount")
            cat_count.set("val", str(len(chart_data.categories)))
            for ci, cat_name in enumerate(chart_data.categories):
                cat_pt = ET.SubElement(cat_cache, f"{{{_C_NS}}}pt")
                cat_pt.set("idx", str(ci))
                cat_v = ET.SubElement(cat_pt, f"{{{_C_NS}}}v")
                cat_v.text = cat_name

        # Value reference
        val = ET.SubElement(ser, f"{{{_C_NS}}}val")
        num_ref = ET.SubElement(val, f"{{{_C_NS}}}numRef")
        num_cache = ET.SubElement(num_ref, f"{{{_C_NS}}}numCache")
        val_count = ET.SubElement(num_cache, f"{{{_C_NS}}}ptCount")
        val_count.set("val", str(len(series.values)))
        for vi, value in enumerate(series.values):
            val_pt = ET.SubElement(num_cache, f"{{{_C_NS}}}pt")
            val_pt.set("idx", str(vi))
            val_v = ET.SubElement(val_pt, f"{{{_C_NS}}}v")
            val_v.text = str(value)

    # Axis IDs for non-pie
    if tag != "pieChart":
        ax1 = ET.SubElement(elem, f"{{{_C_NS}}}axId")
        ax1.set("val", "1")
        ax2 = ET.SubElement(elem, f"{{{_C_NS}}}axId")
        ax2.set("val", "2")

    return elem


def make_chart_graphic_frame(
    shape_id: int,
    r_id: str,
    left: int,
    top: int,
    width: int,
    height: int,
) -> ET.Element:
    """Create a graphicFrame element for a chart."""
    gf = ET.Element(qn("p:graphicFrame"))

    # nvGraphicFramePr
    nvgfpr = ET.SubElement(gf, qn("p:nvGraphicFramePr"))
    cnvpr = ET.SubElement(nvgfpr, qn("p:cNvPr"))
    cnvpr.set("id", str(shape_id))
    cnvpr.set("name", f"Chart {shape_id}")
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

    # graphic with chart reference
    graphic = ET.SubElement(gf, qn("a:graphic"))
    gd = ET.SubElement(graphic, qn("a:graphicData"))
    gd.set("uri", "http://schemas.openxmlformats.org/drawingml/2006/chart")
    chart_ref = ET.SubElement(gd, qn("c:chart"))
    chart_ref.set(qn("r:id"), r_id)

    return gf
