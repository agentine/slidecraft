"""Tests for picture shapes."""

from __future__ import annotations

import io
import struct

from slidecraft.pptx.presentation import Presentation
from slidecraft.pptx.shapes.picture import Image, _detect_image_type
from slidecraft.util.units import Inches


def _make_png_bytes(width: int = 100, height: int = 50) -> bytes:
    """Create a minimal valid PNG file."""
    # PNG signature
    sig = b"\x89PNG\r\n\x1a\n"
    # IHDR chunk
    ihdr_data = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    import zlib

    ihdr_crc = struct.pack(">I", zlib.crc32(b"IHDR" + ihdr_data) & 0xFFFFFFFF)
    ihdr = struct.pack(">I", 13) + b"IHDR" + ihdr_data + ihdr_crc
    # IEND chunk
    iend_crc = struct.pack(">I", zlib.crc32(b"IEND") & 0xFFFFFFFF)
    iend = struct.pack(">I", 0) + b"IEND" + iend_crc
    # Empty IDAT (minimal)
    idat_data = zlib.compress(b"\x00" * (width * 3 + 1) * height)
    idat_crc = struct.pack(">I", zlib.crc32(b"IDAT" + idat_data) & 0xFFFFFFFF)
    idat = struct.pack(">I", len(idat_data)) + b"IDAT" + idat_data + idat_crc
    return sig + ihdr + idat + iend


class TestImageDetection:
    def test_detect_png(self) -> None:
        blob = _make_png_bytes()
        ct, ext = _detect_image_type(blob)
        assert ext == "png"
        assert "png" in ct

    def test_detect_jpeg(self) -> None:
        blob = b"\xff\xd8\xff\xe0" + b"\x00" * 100
        ct, ext = _detect_image_type(blob)
        assert ext == "jpeg"

    def test_image_from_blob(self) -> None:
        blob = _make_png_bytes(200, 150)
        img = Image.from_blob(blob)
        assert img.ext == "png"
        assert img.width == 200
        assert img.height == 150


class TestPicture:
    def test_add_picture(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        blob = _make_png_bytes(100, 50)
        img = Image.from_blob(blob)
        pic = slide.shapes.add_picture(img, Inches(1), Inches(1))
        assert len(slide.shapes) == 1

    def test_add_picture_with_dimensions(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        blob = _make_png_bytes()
        img = Image.from_blob(blob)
        pic = slide.shapes.add_picture(
            img, Inches(1), Inches(1), width=Inches(3), height=Inches(2)
        )
        assert abs(pic.width.inches - 3.0) < 0.01
        assert abs(pic.height.inches - 2.0) < 0.01

    def test_picture_roundtrip(self) -> None:
        prs = Presentation()
        slide = prs.slides.add()
        blob = _make_png_bytes(100, 50)
        img = Image.from_blob(blob)
        slide.shapes.add_picture(img, Inches(1), Inches(1), Inches(3), Inches(2))

        slide._sync_blob()

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation.open(buf)
        shapes = prs2.slides[0].shapes
        assert len(shapes) >= 1
        # Verify the image part exists
        img_parts = [p for p in prs2._pkg.parts if "media" in p]
        assert len(img_parts) == 1
