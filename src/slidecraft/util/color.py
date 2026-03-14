"""Color utilities for PowerPoint."""

from __future__ import annotations


class RGBColor:
    """An RGB color value."""

    __slots__ = ("_r", "_g", "_b")

    def __init__(self, r: int, g: int, b: int) -> None:
        for name, val in (("r", r), ("g", g), ("b", b)):
            if not 0 <= val <= 255:
                raise ValueError(
                    f"{name} must be 0-255, got {val}"
                )
        self._r = r
        self._g = g
        self._b = b

    @classmethod
    def from_string(cls, hex_str: str) -> RGBColor:
        """Create from a 6-character hex string like 'FF0000'."""
        s = hex_str.lstrip("#")
        if len(s) != 6:
            raise ValueError(f"Expected 6-char hex string, got {hex_str!r}")
        return cls(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))

    @property
    def r(self) -> int:
        return self._r

    @property
    def g(self) -> int:
        return self._g

    @property
    def b(self) -> int:
        return self._b

    def __str__(self) -> str:
        return f"{self._r:02X}{self._g:02X}{self._b:02X}"

    def __repr__(self) -> str:
        return f"RGBColor(0x{self._r:02X}, 0x{self._g:02X}, 0x{self._b:02X})"

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, RGBColor):
            return NotImplemented
        return (self._r, self._g, self._b) == (other._r, other._g, other._b)

    def __hash__(self) -> int:
        return hash((self._r, self._g, self._b))


# Common colors
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)
RED = RGBColor(255, 0, 0)
GREEN = RGBColor(0, 255, 0)
BLUE = RGBColor(0, 0, 255)
