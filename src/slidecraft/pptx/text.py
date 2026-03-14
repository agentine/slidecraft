"""Text support: TextFrame, Paragraph, Run, Font."""

from __future__ import annotations

import xml.etree.ElementTree as ET

from slidecraft.pptx.enum import PP_ALIGN, pp_align_from_xml
from slidecraft.util.color import RGBColor
from slidecraft.util.units import Emu
from slidecraft.xml.ns import qn


class Font:
    """Character-level formatting."""

    def __init__(
        self, rpr: ET.Element | None = None, parent: ET.Element | None = None
    ) -> None:
        self._rpr = rpr
        self._parent = parent  # The a:r element that owns this font

    def _ensure_rpr(self) -> ET.Element:
        if self._rpr is None:
            self._rpr = ET.Element(qn("a:rPr"))
            if self._parent is not None:
                self._parent.insert(0, self._rpr)
        return self._rpr

    @property
    def name(self) -> str | None:
        if self._rpr is None:
            return None
        latin = self._rpr.find(qn("a:latin"))
        if latin is not None:
            return latin.get("typeface")
        return None

    @name.setter
    def name(self, value: str | None) -> None:
        rpr = self._ensure_rpr()
        latin = rpr.find(qn("a:latin"))
        if value is None:
            if latin is not None:
                rpr.remove(latin)
            return
        if latin is None:
            latin = ET.SubElement(rpr, qn("a:latin"))
        latin.set("typeface", value)

    @property
    def size(self) -> Emu | None:
        if self._rpr is None:
            return None
        sz = self._rpr.get("sz")
        if sz is not None:
            # sz is in hundredths of a point, convert to EMU
            return Emu(int(int(sz) * 12700 / 100))
        return None

    @size.setter
    def size(self, value: int | None) -> None:
        rpr = self._ensure_rpr()
        if value is None:
            if "sz" in rpr.attrib:
                del rpr.attrib["sz"]
            return
        # Convert EMU to hundredths of a point
        hundredths = int(int(value) * 100 / 12700)
        rpr.set("sz", str(hundredths))

    @property
    def bold(self) -> bool | None:
        if self._rpr is None:
            return None
        b = self._rpr.get("b")
        if b is None:
            return None
        return b == "1" or b.lower() == "true"

    @bold.setter
    def bold(self, value: bool | None) -> None:
        rpr = self._ensure_rpr()
        if value is None:
            if "b" in rpr.attrib:
                del rpr.attrib["b"]
            return
        rpr.set("b", "1" if value else "0")

    @property
    def italic(self) -> bool | None:
        if self._rpr is None:
            return None
        i = self._rpr.get("i")
        if i is None:
            return None
        return i == "1" or i.lower() == "true"

    @italic.setter
    def italic(self, value: bool | None) -> None:
        rpr = self._ensure_rpr()
        if value is None:
            if "i" in rpr.attrib:
                del rpr.attrib["i"]
            return
        rpr.set("i", "1" if value else "0")

    @property
    def underline(self) -> bool | None:
        if self._rpr is None:
            return None
        u = self._rpr.get("u")
        if u is None:
            return None
        return u != "none"

    @underline.setter
    def underline(self, value: bool | None) -> None:
        rpr = self._ensure_rpr()
        if value is None:
            if "u" in rpr.attrib:
                del rpr.attrib["u"]
            return
        rpr.set("u", "sng" if value else "none")

    @property
    def color(self) -> RGBColor | None:
        if self._rpr is None:
            return None
        solid = self._rpr.find(qn("a:solidFill"))
        if solid is not None:
            srgb = solid.find(qn("a:srgbClr"))
            if srgb is not None:
                val = srgb.get("val", "")
                if len(val) == 6:
                    return RGBColor.from_string(val)
        return None

    @color.setter
    def color(self, value: RGBColor | None) -> None:
        rpr = self._ensure_rpr()
        # Remove existing fill
        for fill in list(rpr.findall(qn("a:solidFill"))):
            rpr.remove(fill)
        if value is None:
            return
        solid = ET.SubElement(rpr, qn("a:solidFill"))
        srgb = ET.SubElement(solid, qn("a:srgbClr"))
        srgb.set("val", str(value))

    @property
    def element(self) -> ET.Element | None:
        return self._rpr


class Run:
    """A run of text with consistent formatting."""

    def __init__(self, element: ET.Element) -> None:
        self._element = element

    @property
    def text(self) -> str:
        t = self._element.find(qn("a:t"))
        if t is not None and t.text is not None:
            return t.text
        return ""

    @text.setter
    def text(self, value: str) -> None:
        t = self._element.find(qn("a:t"))
        if t is None:
            t = ET.SubElement(self._element, qn("a:t"))
        t.text = value

    @property
    def font(self) -> Font:
        rpr = self._element.find(qn("a:rPr"))
        return Font(rpr, parent=self._element)

    @property
    def element(self) -> ET.Element:
        return self._element

    def _sync_font(self) -> None:
        """Ensure font element is attached to run."""
        rpr = self._element.find(qn("a:rPr"))
        font_el = self.font.element
        if font_el is not None and rpr is None:
            self._element.insert(0, font_el)


class Paragraph:
    """A paragraph within a text frame."""

    def __init__(self, element: ET.Element) -> None:
        self._element = element

    @property
    def runs(self) -> list[Run]:
        return [Run(r) for r in self._element.findall(qn("a:r"))]

    def add_run(self) -> Run:
        """Add a new run to this paragraph."""
        r_elem = ET.SubElement(self._element, qn("a:r"))
        ET.SubElement(r_elem, qn("a:t"))
        return Run(r_elem)

    @property
    def text(self) -> str:
        parts: list[str] = []
        for child in self._element:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "r":
                t = child.find(qn("a:t"))
                if t is not None and t.text is not None:
                    parts.append(t.text)
            elif tag == "br":
                parts.append("\n")
        return "".join(parts)

    @text.setter
    def text(self, value: str) -> None:
        # Remove existing runs
        for r in list(self._element.findall(qn("a:r"))):
            self._element.remove(r)
        for br in list(self._element.findall(qn("a:br"))):
            self._element.remove(br)
        # Add a single run
        if value:
            run = self.add_run()
            run.text = value

    @property
    def alignment(self) -> PP_ALIGN | None:
        ppr = self._element.find(qn("a:pPr"))
        if ppr is not None:
            algn = ppr.get("algn")
            if algn is not None:
                return pp_align_from_xml(algn)
        return None

    @alignment.setter
    def alignment(self, value: PP_ALIGN | None) -> None:
        ppr = self._element.find(qn("a:pPr"))
        if value is None:
            if ppr is not None and "algn" in ppr.attrib:
                del ppr.attrib["algn"]
            return
        if ppr is None:
            ppr = ET.SubElement(self._element, qn("a:pPr"))
            # Insert pPr at beginning
            self._element.remove(ppr)
            self._element.insert(0, ppr)
        ppr.set("algn", value.to_xml())

    @property
    def level(self) -> int:
        ppr = self._element.find(qn("a:pPr"))
        if ppr is not None:
            lvl = ppr.get("lvl")
            if lvl is not None:
                return int(lvl)
        return 0

    @level.setter
    def level(self, value: int) -> None:
        ppr = self._element.find(qn("a:pPr"))
        if ppr is None:
            ppr = ET.SubElement(self._element, qn("a:pPr"))
            self._element.remove(ppr)
            self._element.insert(0, ppr)
        ppr.set("lvl", str(value))

    @property
    def space_before(self) -> Emu | None:
        ppr = self._element.find(qn("a:pPr"))
        if ppr is not None:
            spc = ppr.find(qn("a:spcBef"))
            if spc is not None:
                pts = spc.find(qn("a:spcPts"))
                if pts is not None:
                    val = pts.get("val")
                    if val is not None:
                        # val is hundredths of a point
                        return Emu(int(int(val) * 12700 / 100))
        return None

    @space_before.setter
    def space_before(self, value: int | None) -> None:
        ppr = self._element.find(qn("a:pPr"))
        if ppr is None:
            ppr = ET.SubElement(self._element, qn("a:pPr"))
            self._element.remove(ppr)
            self._element.insert(0, ppr)
        spc = ppr.find(qn("a:spcBef"))
        if value is None:
            if spc is not None:
                ppr.remove(spc)
            return
        if spc is None:
            spc = ET.SubElement(ppr, qn("a:spcBef"))
        pts = spc.find(qn("a:spcPts"))
        if pts is None:
            pts = ET.SubElement(spc, qn("a:spcPts"))
        pts.set("val", str(int(int(value) * 100 / 12700)))

    @property
    def space_after(self) -> Emu | None:
        ppr = self._element.find(qn("a:pPr"))
        if ppr is not None:
            spc = ppr.find(qn("a:spcAft"))
            if spc is not None:
                pts = spc.find(qn("a:spcPts"))
                if pts is not None:
                    val = pts.get("val")
                    if val is not None:
                        return Emu(int(int(val) * 12700 / 100))
        return None

    @space_after.setter
    def space_after(self, value: int | None) -> None:
        ppr = self._element.find(qn("a:pPr"))
        if ppr is None:
            ppr = ET.SubElement(self._element, qn("a:pPr"))
            self._element.remove(ppr)
            self._element.insert(0, ppr)
        spc = ppr.find(qn("a:spcAft"))
        if value is None:
            if spc is not None:
                ppr.remove(spc)
            return
        if spc is None:
            spc = ET.SubElement(ppr, qn("a:spcAft"))
        pts = spc.find(qn("a:spcPts"))
        if pts is None:
            pts = ET.SubElement(spc, qn("a:spcPts"))
        pts.set("val", str(int(int(value) * 100 / 12700)))

    @property
    def element(self) -> ET.Element:
        return self._element


class TextFrame:
    """A text frame containing paragraphs."""

    def __init__(self, txbody: ET.Element) -> None:
        self._element = txbody

    @property
    def paragraphs(self) -> list[Paragraph]:
        return [Paragraph(p) for p in self._element.findall(qn("a:p"))]

    def add_paragraph(self) -> Paragraph:
        """Add a new paragraph."""
        p = ET.SubElement(self._element, qn("a:p"))
        return Paragraph(p)

    def clear(self) -> None:
        """Remove all paragraphs and add one empty paragraph."""
        for p in list(self._element.findall(qn("a:p"))):
            self._element.remove(p)
        ET.SubElement(self._element, qn("a:p"))

    @property
    def text(self) -> str:
        return "\n".join(p.text for p in self.paragraphs)

    @text.setter
    def text(self, value: str) -> None:
        self.clear()
        lines = value.split("\n")
        paras = self.paragraphs
        if paras:
            paras[0].text = lines[0]
        for line in lines[1:]:
            p = self.add_paragraph()
            p.text = line

    @property
    def word_wrap(self) -> bool | None:
        bp = self._element.find(qn("a:bodyPr"))
        if bp is not None:
            wrap = bp.get("wrap")
            if wrap is not None:
                return wrap == "square"
        return None

    @word_wrap.setter
    def word_wrap(self, value: bool | None) -> None:
        bp = self._element.find(qn("a:bodyPr"))
        if bp is None:
            bp = ET.SubElement(self._element, qn("a:bodyPr"))
            self._element.remove(bp)
            self._element.insert(0, bp)
        if value is None:
            if "wrap" in bp.attrib:
                del bp.attrib["wrap"]
        elif value:
            bp.set("wrap", "square")
        else:
            bp.set("wrap", "none")

    @property
    def element(self) -> ET.Element:
        return self._element
