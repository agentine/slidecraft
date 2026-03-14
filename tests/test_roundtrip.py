"""Comprehensive round-trip tests."""

from __future__ import annotations

import io

from slidecraft import (
    PP_ALIGN,
    ChartData,
    ChartType,
    Inches,
    Presentation,
    RGBColor,
)
from slidecraft.pptx.shapes.picture import Image


def _make_png_bytes() -> bytes:
    """Create a minimal PNG."""
    import struct
    import zlib

    sig = b"\x89PNG\r\n\x1a\n"
    w, h = 10, 10
    ihdr_data = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
    ihdr_crc = struct.pack(">I", zlib.crc32(b"IHDR" + ihdr_data) & 0xFFFFFFFF)
    ihdr = struct.pack(">I", 13) + b"IHDR" + ihdr_data + ihdr_crc
    idat_data = zlib.compress(b"\x00" * (w * 3 + 1) * h)
    idat_crc = struct.pack(">I", zlib.crc32(b"IDAT" + idat_data) & 0xFFFFFFFF)
    idat = struct.pack(">I", len(idat_data)) + b"IDAT" + idat_data + idat_crc
    iend_crc = struct.pack(">I", zlib.crc32(b"IEND") & 0xFFFFFFFF)
    iend = struct.pack(">I", 0) + b"IEND" + iend_crc
    return sig + ihdr + idat + iend


class TestFullRoundTrip:
    def test_complex_presentation(self) -> None:
        """Create a presentation with multiple content types and round-trip it."""
        prs = Presentation()

        # Slide 1: Title text
        s1 = prs.slides.add()
        tb = s1.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
        tf = tb.text_frame
        tf.text = "Welcome to SlideCraft"
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.runs[0]
        run.font.bold = True
        run.font.color = RGBColor(0, 0, 128)
        run.font.name = "Arial"

        # Slide 2: Table
        s2 = prs.slides.add()
        tbl = s2.shapes.add_table(3, 3, Inches(1), Inches(1), Inches(6), Inches(3))
        for r in range(3):
            for c in range(3):
                tbl.cell(r, c).text = f"R{r}C{c}"

        # Slide 3: Image
        s3 = prs.slides.add()
        img = Image.from_blob(_make_png_bytes())
        s3.shapes.add_picture(img, Inches(2), Inches(2), Inches(4), Inches(3))

        # Slide 4: Chart
        s4 = prs.slides.add()
        cd = ChartData()
        cd.categories = ["Q1", "Q2", "Q3", "Q4"]
        cd.add_series("Revenue", [100.0, 150.0, 130.0, 200.0])
        s4.shapes.add_chart(ChartType.COLUMN, cd, Inches(1), Inches(1), Inches(6), Inches(4))

        # Sync all slides
        for s in prs.slides:
            s._sync_blob()

        # Save
        buf = io.BytesIO()
        prs.save(buf)
        assert buf.tell() > 0

        # Reload
        buf.seek(0)
        prs2 = Presentation.open(buf)
        assert len(prs2.slides) == 4

        # Verify slide 1 text
        s1_shapes = prs2.slides[0].shapes
        assert len(s1_shapes) == 1
        assert s1_shapes[0].text_frame.text == "Welcome to SlideCraft"

        # Verify slide 2 has a shape (table)
        assert len(prs2.slides[1].shapes) >= 1

        # Verify slide 3 has a picture
        assert len(prs2.slides[2].shapes) >= 1

        # Verify slide 4 has a chart
        assert len(prs2.slides[3].shapes) >= 1

    def test_empty_presentation_roundtrip(self) -> None:
        """Empty presentation should round-trip."""
        prs = Presentation()
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation.open(buf)
        assert len(prs2.slides) == 0
        assert len(prs2.slide_layouts) >= 1

    def test_many_slides(self) -> None:
        """Presentation with many slides."""
        prs = Presentation()
        for i in range(20):
            slide = prs.slides.add()
            tb = slide.shapes.add_textbox(0, 0, Inches(5), Inches(1))
            tb.text_frame.text = f"Slide {i + 1}"
            slide._sync_blob()

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation.open(buf)
        assert len(prs2.slides) == 20
