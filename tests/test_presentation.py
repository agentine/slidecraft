"""Tests for the presentation model."""

from __future__ import annotations

import io

from slidecraft.pptx.presentation import Presentation
from slidecraft.util.units import Inches


class TestPresentation:
    def test_new_presentation(self) -> None:
        prs = Presentation()
        assert len(prs.slides) == 0
        assert len(prs.slide_layouts) >= 1
        assert len(prs.slide_masters) >= 1

    def test_add_slide(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        assert len(prs.slides) == 1
        assert slide.slide_id == 256

    def test_add_multiple_slides(self) -> None:
        prs = Presentation()
        prs.slides.add()
        prs.slides.add()
        prs.slides.add()
        assert len(prs.slides) == 3

    def test_slide_width_height(self) -> None:
        prs = Presentation()
        # Default widescreen: 13.333 x 7.5 inches
        assert int(prs.slide_width) == 12192000
        assert int(prs.slide_height) == 6858000

    def test_set_slide_dimensions(self) -> None:
        prs = Presentation()
        prs.slide_width = Inches(10)
        prs.slide_height = Inches(7.5)
        assert abs(prs.slide_width.inches - 10.0) < 0.01

    def test_save_and_open_roundtrip(self) -> None:
        prs = Presentation()
        prs.slides.add()
        prs.slides.add()

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation.open(buf)
        assert len(prs2.slides) == 2
        assert len(prs2.slide_layouts) >= 1

    def test_remove_slide(self) -> None:
        prs = Presentation()
        prs.slides.add()
        prs.slides.add()
        assert len(prs.slides) == 2
        prs.slides.remove(0)
        assert len(prs.slides) == 1

    def test_slide_layout_name(self) -> None:
        prs = Presentation()
        layout = prs.slide_layouts[0]
        assert layout.name == "Blank"

    def test_slide_getitem(self) -> None:
        prs = Presentation()
        s1 = prs.slides.add()
        s2 = prs.slides.add()
        assert prs.slides[0].slide_id == s1.slide_id
        assert prs.slides[1].slide_id == s2.slide_id

    def test_iterate_slides(self) -> None:
        prs = Presentation()
        prs.slides.add()
        prs.slides.add()
        ids = [s.slide_id for s in prs.slides]
        assert len(ids) == 2
