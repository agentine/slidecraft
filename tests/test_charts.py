"""Tests for chart shapes."""

from __future__ import annotations

import io

from slidecraft.pptx.enum import ChartType
from slidecraft.pptx.presentation import Presentation
from slidecraft.pptx.shapes.chart import (
    ChartData,
    LegendPosition,
    generate_chart_xml,
)
from slidecraft.util.units import Inches


class TestChartData:
    def test_create_chart_data(self) -> None:
        cd = ChartData()
        cd.categories = ["Q1", "Q2", "Q3", "Q4"]
        cd.add_series("Sales", [100, 200, 150, 300])
        cd.add_series("Expenses", [80, 150, 120, 200])
        assert len(cd.series) == 2
        assert cd.categories == ["Q1", "Q2", "Q3", "Q4"]
        assert cd.series[0].name == "Sales"
        assert cd.series[0].values == [100, 200, 150, 300]


class TestChartXML:
    def test_bar_chart_xml(self) -> None:
        cd = ChartData()
        cd.categories = ["A", "B", "C"]
        cd.add_series("S1", [1.0, 2.0, 3.0])
        xml = generate_chart_xml(ChartType.BAR, cd)
        assert b"chartSpace" in xml
        assert b"barChart" in xml
        assert b"bar" in xml  # barDir=bar

    def test_column_chart_xml(self) -> None:
        cd = ChartData()
        cd.categories = ["X", "Y"]
        cd.add_series("D", [10.0, 20.0])
        xml = generate_chart_xml(ChartType.COLUMN, cd)
        assert b"barChart" in xml
        assert b"col" in xml  # barDir=col

    def test_line_chart_xml(self) -> None:
        cd = ChartData()
        cd.categories = ["Jan", "Feb"]
        cd.add_series("Revenue", [1000.0, 1200.0])
        xml = generate_chart_xml(ChartType.LINE, cd)
        assert b"lineChart" in xml

    def test_pie_chart_xml(self) -> None:
        cd = ChartData()
        cd.categories = ["A", "B", "C"]
        cd.add_series("Share", [30.0, 50.0, 20.0])
        xml = generate_chart_xml(ChartType.PIE, cd)
        assert b"pieChart" in xml
        # Pie charts should not have axis elements
        assert b"catAx" not in xml

    def test_legend_position(self) -> None:
        cd = ChartData()
        cd.categories = ["A"]
        cd.add_series("S", [1.0])
        xml = generate_chart_xml(
            ChartType.BAR, cd, legend_position=LegendPosition.TOP
        )
        assert b"legendPos" in xml

    def test_no_legend(self) -> None:
        cd = ChartData()
        cd.categories = ["A"]
        cd.add_series("S", [1.0])
        xml = generate_chart_xml(ChartType.BAR, cd, has_legend=False)
        assert b"legend" not in xml


class TestChartOnSlide:
    def test_add_bar_chart(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        cd = ChartData()
        cd.categories = ["Q1", "Q2", "Q3"]
        cd.add_series("Sales", [100.0, 200.0, 150.0])
        cd.add_series("Costs", [80.0, 160.0, 120.0])
        shape = slide.shapes.add_chart(
            ChartType.BAR, cd, Inches(1), Inches(1), Inches(6), Inches(4)
        )
        assert len(slide.shapes) == 1

    def test_add_pie_chart(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        cd = ChartData()
        cd.categories = ["Red", "Blue", "Green"]
        cd.add_series("Colors", [40.0, 35.0, 25.0])
        slide.shapes.add_chart(
            ChartType.PIE, cd, Inches(2), Inches(2), Inches(4), Inches(4)
        )
        assert len(slide.shapes) == 1

    def test_chart_roundtrip(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        cd = ChartData()
        cd.categories = ["A", "B"]
        cd.add_series("Data", [10.0, 20.0])
        slide.shapes.add_chart(
            ChartType.COLUMN, cd, Inches(1), Inches(1), Inches(5), Inches(3)
        )

        slide._sync_blob()

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation.open(buf)
        assert len(prs2.slides[0].shapes) == 1
        # Verify chart part exists
        chart_parts = [p for p in prs2._pkg.parts if "charts" in p]
        assert len(chart_parts) == 1
