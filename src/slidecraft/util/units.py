"""Unit conversion utilities for PowerPoint measurements.

PowerPoint uses English Metric Units (EMU) internally.
1 inch = 914400 EMU
1 cm = 360000 EMU
1 pt = 12700 EMU
"""

from __future__ import annotations


class Emu(int):
    """English Metric Unit — the base measurement in Open XML."""

    def __new__(cls, value: int) -> Emu:
        return super().__new__(cls, value)

    @property
    def inches(self) -> float:
        return self / 914400.0

    @property
    def cm(self) -> float:
        return self / 360000.0

    @property
    def pt(self) -> float:
        return self / 12700.0

    def __repr__(self) -> str:
        return f"Emu({int(self)})"


def Inches(value: float) -> Emu:  # noqa: N802
    """Convert inches to EMU."""
    return Emu(int(round(value * 914400)))


def Cm(value: float) -> Emu:  # noqa: N802
    """Convert centimeters to EMU."""
    return Emu(int(round(value * 360000)))


def Pt(value: float) -> Emu:  # noqa: N802
    """Convert points to EMU."""
    return Emu(int(round(value * 12700)))
