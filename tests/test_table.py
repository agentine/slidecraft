"""Tests for table shapes."""

from __future__ import annotations

import io

from slidecraft.pptx.presentation import Presentation
from slidecraft.util.units import Inches


class TestTable:
    def test_add_table(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tbl = slide.shapes.add_table(3, 4, Inches(1), Inches(2), Inches(6), Inches(3))
        assert len(tbl.rows) == 3
        assert len(tbl.columns) == 4

    def test_cell_text(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tbl = slide.shapes.add_table(2, 2, 0, 0, 1000000, 500000)
        tbl.cell(0, 0).text = "Header 1"
        tbl.cell(0, 1).text = "Header 2"
        tbl.cell(1, 0).text = "Data 1"
        tbl.cell(1, 1).text = "Data 2"
        assert tbl.cell(0, 0).text == "Header 1"
        assert tbl.cell(1, 1).text == "Data 2"

    def test_row_height(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tbl = slide.shapes.add_table(2, 2, 0, 0, 1000000, 500000)
        row = tbl.rows[0]
        row.height = Inches(1)
        assert abs(row.height.inches - 1.0) < 0.01

    def test_column_width(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tbl = slide.shapes.add_table(2, 2, 0, 0, 1000000, 500000)
        col = tbl.columns[0]
        col.width = Inches(3)
        assert abs(col.width.inches - 3.0) < 0.01

    def test_cell_index_error(self) -> None:
        import pytest

        prs = Presentation()
        slide = prs.slides.add()
        tbl = slide.shapes.add_table(2, 2, 0, 0, 1000000, 500000)
        with pytest.raises(IndexError):
            tbl.cell(5, 0)
        with pytest.raises(IndexError):
            tbl.cell(0, 5)

    def test_table_roundtrip(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tbl = slide.shapes.add_table(2, 3, Inches(1), Inches(1), Inches(6), Inches(2))
        tbl.cell(0, 0).text = "A1"
        tbl.cell(0, 1).text = "B1"
        tbl.cell(1, 2).text = "C2"

        slide._sync_blob()

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation.open(buf)
        assert len(prs2.slides) == 1
        # The table is a shape on the slide
        shapes = prs2.slides[0].shapes
        assert len(shapes) >= 1
