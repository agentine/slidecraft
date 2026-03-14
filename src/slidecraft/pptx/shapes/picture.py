"""Picture shape support."""

from __future__ import annotations

import io
import struct
import xml.etree.ElementTree as ET
from typing import TYPE_CHECKING

from slidecraft.xml.ns import CT, qn

if TYPE_CHECKING:
    from slidecraft.pptx.slide import Slide


class Image:
    """An image embedded in a presentation."""

    def __init__(
        self,
        blob: bytes,
        content_type: str,
        ext: str,
        width: int = 0,
        height: int = 0,
    ) -> None:
        self.blob = blob
        self.content_type = content_type
        self.ext = ext
        self.width = width
        self.height = height

    @classmethod
    def from_blob(cls, blob: bytes) -> Image:
        """Create an Image from raw bytes, auto-detecting type."""
        content_type, ext = _detect_image_type(blob)
        w, h = _detect_dimensions(blob, ext)
        return cls(blob, content_type, ext, w, h)

    @classmethod
    def from_file(cls, path: str) -> Image:
        """Create an Image from a file path."""
        with open(path, "rb") as f:
            blob = f.read()
        return cls.from_blob(blob)


class Picture:
    """A picture shape on a slide."""

    def __init__(self, element: ET.Element, slide: Slide | None = None) -> None:
        self._element = element
        self._slide = slide

    @property
    def crop_left(self) -> float:
        fill = self._get_blip_fill()
        if fill is not None:
            src_rect = fill.find(qn("a:srcRect"))
            if src_rect is not None:
                val = src_rect.get("l", "0")
                return int(val) / 100000.0
        return 0.0

    @crop_left.setter
    def crop_left(self, value: float) -> None:
        fill = self._ensure_blip_fill()
        src_rect = fill.find(qn("a:srcRect"))
        if src_rect is None:
            src_rect = ET.SubElement(fill, qn("a:srcRect"))
        src_rect.set("l", str(int(value * 100000)))

    @property
    def crop_top(self) -> float:
        fill = self._get_blip_fill()
        if fill is not None:
            src_rect = fill.find(qn("a:srcRect"))
            if src_rect is not None:
                val = src_rect.get("t", "0")
                return int(val) / 100000.0
        return 0.0

    @crop_top.setter
    def crop_top(self, value: float) -> None:
        fill = self._ensure_blip_fill()
        src_rect = fill.find(qn("a:srcRect"))
        if src_rect is None:
            src_rect = ET.SubElement(fill, qn("a:srcRect"))
        src_rect.set("t", str(int(value * 100000)))

    @property
    def crop_right(self) -> float:
        fill = self._get_blip_fill()
        if fill is not None:
            src_rect = fill.find(qn("a:srcRect"))
            if src_rect is not None:
                val = src_rect.get("r", "0")
                return int(val) / 100000.0
        return 0.0

    @crop_right.setter
    def crop_right(self, value: float) -> None:
        fill = self._ensure_blip_fill()
        src_rect = fill.find(qn("a:srcRect"))
        if src_rect is None:
            src_rect = ET.SubElement(fill, qn("a:srcRect"))
        src_rect.set("r", str(int(value * 100000)))

    @property
    def crop_bottom(self) -> float:
        fill = self._get_blip_fill()
        if fill is not None:
            src_rect = fill.find(qn("a:srcRect"))
            if src_rect is not None:
                val = src_rect.get("b", "0")
                return int(val) / 100000.0
        return 0.0

    @crop_bottom.setter
    def crop_bottom(self, value: float) -> None:
        fill = self._ensure_blip_fill()
        src_rect = fill.find(qn("a:srcRect"))
        if src_rect is None:
            src_rect = ET.SubElement(fill, qn("a:srcRect"))
        src_rect.set("b", str(int(value * 100000)))

    @property
    def element(self) -> ET.Element:
        return self._element

    def _get_blip_fill(self) -> ET.Element | None:
        return self._element.find(qn("p:blipFill"))

    def _ensure_blip_fill(self) -> ET.Element:
        fill = self._get_blip_fill()
        if fill is None:
            fill = ET.SubElement(self._element, qn("p:blipFill"))
        return fill


def make_picture_element(
    shape_id: int,
    r_id: str,
    name: str,
    left: int,
    top: int,
    width: int,
    height: int,
) -> ET.Element:
    """Create a p:pic element for an image."""
    pic = ET.Element(qn("p:pic"))

    # nvPicPr
    nv = ET.SubElement(pic, qn("p:nvPicPr"))
    cnvpr = ET.SubElement(nv, qn("p:cNvPr"))
    cnvpr.set("id", str(shape_id))
    cnvpr.set("name", name)
    cnvpicpr = ET.SubElement(nv, qn("p:cNvPicPr"))
    locks = ET.SubElement(cnvpicpr, qn("a:picLocks"))
    locks.set("noChangeAspect", "1")
    ET.SubElement(nv, qn("p:nvPr"))

    # blipFill
    blip_fill = ET.SubElement(pic, qn("p:blipFill"))
    blip = ET.SubElement(blip_fill, qn("a:blip"))
    blip.set(qn("r:embed"), r_id)
    stretch = ET.SubElement(blip_fill, qn("a:stretch"))
    ET.SubElement(stretch, qn("a:fillRect"))

    # spPr
    sppr = ET.SubElement(pic, qn("p:spPr"))
    xfrm = ET.SubElement(sppr, qn("a:xfrm"))
    off = ET.SubElement(xfrm, qn("a:off"))
    off.set("x", str(int(left)))
    off.set("y", str(int(top)))
    ext = ET.SubElement(xfrm, qn("a:ext"))
    ext.set("cx", str(int(width)))
    ext.set("cy", str(int(height)))
    prstgeom = ET.SubElement(sppr, qn("a:prstGeom"))
    prstgeom.set("prst", "rect")

    return pic


def _detect_image_type(blob: bytes) -> tuple[str, str]:
    """Detect image content type from header bytes."""
    if blob[:8] == b"\x89PNG\r\n\x1a\n":
        return CT["PNG"], "png"
    if blob[:2] == b"\xff\xd8":
        return CT["JPEG"], "jpeg"
    if blob[:4] == b"GIF8":
        return CT["GIF"], "gif"
    if blob[:4] == b"\x01\x00\x00\x00" or blob[:4] == b" EMF":
        return CT["EMF"], "emf"
    if blob[:6] in (b"<?xml ", b"<svg "):
        return CT["SVG"], "svg"
    # Default to PNG
    return CT["PNG"], "png"


def _detect_dimensions(blob: bytes, ext: str) -> tuple[int, int]:
    """Detect image width and height in pixels."""
    if ext == "png" and len(blob) >= 24:
        w = struct.unpack(">I", blob[16:20])[0]
        h = struct.unpack(">I", blob[20:24])[0]
        return w, h
    if ext == "jpeg" and len(blob) > 2:
        return _jpeg_dimensions(blob)
    return 0, 0


def _jpeg_dimensions(blob: bytes) -> tuple[int, int]:
    """Extract dimensions from JPEG."""
    stream = io.BytesIO(blob)
    stream.read(2)  # SOI marker
    while True:
        marker = stream.read(2)
        if len(marker) < 2:
            break
        if marker[0] != 0xFF:
            break
        m = marker[1]
        if m == 0xC0 or m == 0xC2:  # SOF0 or SOF2
            stream.read(3)  # length + precision
            h_bytes = stream.read(2)
            w_bytes = stream.read(2)
            if len(h_bytes) == 2 and len(w_bytes) == 2:
                h = struct.unpack(">H", h_bytes)[0]
                w = struct.unpack(">H", w_bytes)[0]
                return w, h
            break
        else:
            length_bytes = stream.read(2)
            if len(length_bytes) < 2:
                break
            length = struct.unpack(">H", length_bytes)[0]
            stream.read(length - 2)
    return 0, 0
