"""Text box shape."""

from __future__ import annotations

from slidecraft.pptx.shapes.shape import BaseShape


class TextBox(BaseShape):
    """A text box shape — a shape whose primary purpose is text."""

    @property
    def has_text_frame(self) -> bool:
        return True
