"""Fluent XML element builder."""

from __future__ import annotations

import xml.etree.ElementTree as ET

from slidecraft.xml.ns import qn


class XmlBuilder:
    """Fluent builder for creating XML elements.

    Usage:
        el = (XmlBuilder("p:sld")
              .set("xmlns:a", NS["a"])
              .child("p:cSld")
                .child("p:spTree")
                .up()
              .up()
              .build())
    """

    def __init__(self, tag: str, text: str | None = None) -> None:
        resolved = qn(tag) if ":" in tag else tag
        self._element = ET.Element(resolved)
        if text is not None:
            self._element.text = text
        self._stack: list[ET.Element] = []

    def set(self, name: str, value: str) -> XmlBuilder:
        """Set an attribute on the current element."""
        self._element.set(name, value)
        return self

    def child(
        self, tag: str, text: str | None = None, **attribs: str
    ) -> XmlBuilder:
        """Add a child element and descend into it."""
        resolved = qn(tag) if ":" in tag else tag
        child = ET.SubElement(self._current, resolved, attrib=dict(attribs))
        if text is not None:
            child.text = text
        self._stack.append(self._element)
        self._element = child
        return self

    def up(self) -> XmlBuilder:
        """Return to the parent element."""
        if not self._stack:
            raise ValueError("Already at root element")
        self._element = self._stack.pop()
        return self

    @property
    def _current(self) -> ET.Element:
        return self._element

    def build(self) -> ET.Element:
        """Return the root element."""
        while self._stack:
            self._element = self._stack.pop()
        return self._element
