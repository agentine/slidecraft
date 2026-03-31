"""Images: generate a small PNG programmatically and insert it into a slide."""

import struct
import zlib

from slidecraft import Image, Inches, PP_ALIGN, Presentation, Pt


def make_png(width: int, height: int, color: tuple[int, int, int]) -> bytes:
    """Generate a minimal solid-color PNG in pure Python."""
    raw = b""
    for _ in range(height):
        raw += b"\x00" + bytes(color) * width
    compressed = zlib.compress(raw)

    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", compressed) + chunk(b"IEND", b"")


prs = Presentation()
slide = prs.slides.add()

# Add a title
tb = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
p = tb.text_frame.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
run = p.add_run()
run.text = "Image Demo"
run.font.size = Pt(28)
run.font.bold = True

# Generate and insert a colored rectangle
png_bytes = make_png(200, 150, (0, 120, 200))
img = Image.from_blob(png_bytes)
slide.shapes.add_picture(img, Inches(3), Inches(2), Inches(4), Inches(3))

prs.save("images.pptx")
print("Saved images.pptx")
