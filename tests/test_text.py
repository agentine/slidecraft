"""Tests for shapes and text."""

from __future__ import annotations

import io

from slidecraft.pptx.enum import PP_ALIGN
from slidecraft.pptx.presentation import Presentation
from slidecraft.util.color import RGBColor
from slidecraft.util.units import Inches


class TestShapeCollection:
    def test_add_textbox(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        shapes = slide.shapes
        tb = shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
        assert tb.shape_id > 0
        assert len(shapes) == 1

    def test_add_multiple_shapes(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        slide.shapes.add_textbox(0, 0, 100, 100)
        slide.shapes.add_textbox(0, 0, 100, 100)
        slide.shapes.add_shape(1, 0, 0, 100, 100)
        assert len(slide.shapes) == 3

    def test_shape_position(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(3), Inches(4))
        assert abs(tb.left.inches - 1.0) < 0.01
        assert abs(tb.top.inches - 2.0) < 0.01
        assert abs(tb.width.inches - 3.0) < 0.01
        assert abs(tb.height.inches - 4.0) < 0.01

    def test_shape_name(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        assert "TextBox" in tb.name
        tb.name = "My TextBox"
        assert tb.name == "My TextBox"

    def test_set_position(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        tb.left = Inches(2)
        tb.top = Inches(3)
        assert abs(tb.left.inches - 2.0) < 0.01
        assert abs(tb.top.inches - 3.0) < 0.01

    def test_getitem(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        slide.shapes.add_textbox(0, 0, 100, 100)
        slide.shapes.add_textbox(0, 0, 200, 200)
        assert slide.shapes[0].width == 100
        assert slide.shapes[1].width == 200


class TestTextFrame:
    def test_textbox_has_text_frame(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        assert tb.has_text_frame

    def test_set_text(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        tf = tb.text_frame
        tf.text = "Hello"
        assert tf.text == "Hello"

    def test_multiline_text(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        tf = tb.text_frame
        tf.text = "Line 1\nLine 2"
        assert tf.text == "Line 1\nLine 2"
        assert len(tf.paragraphs) == 2

    def test_add_paragraph(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        tf = tb.text_frame
        tf.clear()
        p1 = tf.paragraphs[0]
        p1.text = "First"
        p2 = tf.add_paragraph()
        p2.text = "Second"
        assert len(tf.paragraphs) == 2
        assert tf.text == "First\nSecond"

    def test_clear(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        tf = tb.text_frame
        tf.text = "Hello"
        tf.clear()
        assert tf.text == ""
        assert len(tf.paragraphs) == 1

    def test_word_wrap(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        tf = tb.text_frame
        tf.word_wrap = True
        assert tf.word_wrap is True
        tf.word_wrap = False
        assert tf.word_wrap is False


class TestRun:
    def test_add_run(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = "Hello"
        assert p.text == "Hello"
        assert len(p.runs) == 1

    def test_multiple_runs(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        r1 = p.add_run()
        r1.text = "Hello "
        r2 = p.add_run()
        r2.text = "World"
        assert p.text == "Hello World"


class TestFont:
    def test_bold(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = "Bold"
        run.font.bold = True
        assert run.font.bold is True

    def test_italic(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.font.italic = True
        assert run.font.italic is True

    def test_font_name(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.font.name = "Arial"
        assert run.font.name == "Arial"

    def test_font_color(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.font.color = RGBColor(255, 0, 0)
        assert run.font.color == RGBColor(255, 0, 0)

    def test_underline(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.font.underline = True
        assert run.font.underline is True


class TestParagraphFormatting:
    def test_alignment(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        assert p.alignment == PP_ALIGN.CENTER

    def test_level(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(0, 0, 100, 100)
        tf = tb.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.level = 2
        assert p.level == 2


class TestShapesRoundTrip:
    def test_textbox_roundtrip(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        tb = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(5), Inches(1))
        tb.text_frame.text = "Hello World"

        # Need to sync slide XML
        slide._sync_blob()

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation.open(buf)
        slide2 = prs2.slides[0]
        assert len(slide2.shapes) == 1
        shape = slide2.shapes[0]
        assert shape.text_frame.text == "Hello World"
        assert abs(shape.left.inches - 1.0) < 0.01
