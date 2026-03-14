"""Placeholder shapes."""

from __future__ import annotations

import xml.etree.ElementTree as ET

from slidecraft.pptx.enum import PP_PLACEHOLDER
from slidecraft.pptx.shapes.shape import BaseShape
from slidecraft.xml.ns import qn


class PlaceholderFormat:
    """Format information for a placeholder shape."""

    def __init__(self, ph_element: ET.Element) -> None:
        self._ph = ph_element

    @property
    def type(self) -> PP_PLACEHOLDER | None:
        idx = self._ph.get("type")
        type_map: dict[str, PP_PLACEHOLDER] = {
            "title": PP_PLACEHOLDER.TITLE,
            "body": PP_PLACEHOLDER.BODY,
            "ctrTitle": PP_PLACEHOLDER.CENTER_TITLE,
            "subTitle": PP_PLACEHOLDER.SUBTITLE,
            "dt": PP_PLACEHOLDER.DATE,
            "sldNum": PP_PLACEHOLDER.SLIDE_NUMBER,
            "ftr": PP_PLACEHOLDER.FOOTER,
            "obj": PP_PLACEHOLDER.OBJECT,
        }
        if idx is not None:
            return type_map.get(idx)
        return None

    @property
    def idx(self) -> int:
        return int(self._ph.get("idx", "0"))


class Placeholder(BaseShape):
    """A placeholder shape on a slide."""

    @property
    def placeholder_format(self) -> PlaceholderFormat | None:
        nvpr = self._get_nv_pr()
        if nvpr is not None:
            nv_pr = nvpr.find(qn("p:nvPr"))
            if nv_pr is not None:
                ph = nv_pr.find(qn("p:ph"))
                if ph is not None:
                    return PlaceholderFormat(ph)
        return None
