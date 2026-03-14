"""slidecraft — Modern Python library for creating and modifying PowerPoint (.pptx) files."""

from slidecraft.pptx.enum import (
    PP_ALIGN,
    PP_PLACEHOLDER,
    ChartType,
    MSO_SHAPE_TYPE,
)
from slidecraft.pptx.presentation import Presentation
from slidecraft.pptx.shapes.chart import ChartData
from slidecraft.pptx.shapes.picture import Image
from slidecraft.pptx.slide import Slide, SlideLayout
from slidecraft.pptx.text import Font, Paragraph, Run, TextFrame
from slidecraft.util.color import RGBColor
from slidecraft.util.units import Cm, Emu, Inches, Pt

__all__ = [
    "ChartData",
    "ChartType",
    "Cm",
    "Emu",
    "Font",
    "Image",
    "Inches",
    "MSO_SHAPE_TYPE",
    "PP_ALIGN",
    "PP_PLACEHOLDER",
    "Paragraph",
    "Presentation",
    "Pt",
    "RGBColor",
    "Run",
    "Slide",
    "SlideLayout",
    "TextFrame",
]
