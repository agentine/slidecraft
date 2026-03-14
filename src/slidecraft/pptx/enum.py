"""PowerPoint enumerations."""

from __future__ import annotations

from enum import IntEnum


class PP_ALIGN(IntEnum):
    """Paragraph alignment."""

    LEFT = 0
    CENTER = 1
    RIGHT = 2
    JUSTIFY = 3
    DISTRIBUTE = 4

    def to_xml(self) -> str:
        return _PP_ALIGN_XML[self]


_PP_ALIGN_XML: dict[PP_ALIGN, str] = {
    PP_ALIGN.LEFT: "l",
    PP_ALIGN.CENTER: "ctr",
    PP_ALIGN.RIGHT: "r",
    PP_ALIGN.JUSTIFY: "just",
    PP_ALIGN.DISTRIBUTE: "dist",
}

_PP_ALIGN_FROM_XML: dict[str, PP_ALIGN] = {v: k for k, v in _PP_ALIGN_XML.items()}


def pp_align_from_xml(val: str) -> PP_ALIGN | None:
    """Convert XML alignment value to PP_ALIGN."""
    return _PP_ALIGN_FROM_XML.get(val)


class MSO_SHAPE_TYPE(IntEnum):
    """Shape types."""

    AUTO_SHAPE = 1
    PICTURE = 13
    TABLE = 19
    CHART = 3
    TEXT_BOX = 17
    GROUP = 6
    PLACEHOLDER = 14
    FREEFORM = 5
    LINE = 9


class PP_PLACEHOLDER(IntEnum):
    """Placeholder types."""

    TITLE = 0
    BODY = 1
    CENTER_TITLE = 3
    SUBTITLE = 4
    DATE = 10
    SLIDE_NUMBER = 12
    FOOTER = 11
    OBJECT = 7
    VERTICAL_BODY = 14
    VERTICAL_TITLE = 15


class ChartType(IntEnum):
    """Chart types."""

    BAR = 1
    BAR_CLUSTERED = 2
    COLUMN = 3
    COLUMN_CLUSTERED = 4
    LINE = 5
    PIE = 6
