# slidecraft

Modern Python library for creating and modifying PowerPoint (.pptx) files.

A modern, type-annotated replacement for python-pptx. Pure Python, zero dependencies, Python 3.10+.

## Installation

```bash
pip install slidecraft
```

## Quick Start

```python
from slidecraft import Presentation, Inches, RGBColor, PP_ALIGN

# Create a new presentation
prs = Presentation()

# Add a slide
slide = prs.slides.add()

# Add a text box
tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
tf = tb.text_frame
tf.text = "Hello, SlideCraft!"
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
run = p.runs[0]
run.font.bold = True
run.font.color = RGBColor(0, 0, 128)

# Save
prs.save("output.pptx")
```

## Features

- **Presentations** — Create new or open existing .pptx files
- **Slides** — Add, remove, reorder slides with layout support
- **Text** — Rich text with fonts, colors, alignment, paragraphs, runs
- **Shapes** — Text boxes, auto shapes, placeholders
- **Tables** — Create tables with rows, columns, cell text
- **Images** — Insert PNG, JPEG images with auto-detection
- **Charts** — Bar, column, line, pie charts with data binding

## API Reference

### Presentation

```python
from slidecraft import Presentation

prs = Presentation()              # New blank presentation
prs = Presentation.open("in.pptx")  # Open existing
prs.save("out.pptx")             # Save

prs.slides                        # SlideCollection
prs.slide_layouts                 # List of SlideLayout
prs.slide_width                   # Emu
prs.slide_height                  # Emu
```

### Slides

```python
slide = prs.slides.add()          # Add with default layout
slide = prs.slides.add(layout)    # Add with specific layout
prs.slides.remove(0)              # Remove by index
slide.shapes                      # ShapeCollection
```

### Text

```python
tb = slide.shapes.add_textbox(left, top, width, height)
tf = tb.text_frame
tf.text = "Hello"
tf.word_wrap = True

p = tf.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
p.level = 1

run = p.add_run()
run.text = "Bold text"
run.font.bold = True
run.font.italic = True
run.font.name = "Arial"
run.font.size = Pt(24)
run.font.color = RGBColor(255, 0, 0)
run.font.underline = True
```

### Tables

```python
table = slide.shapes.add_table(rows, cols, left, top, width, height)
table.cell(0, 0).text = "Header"
table.rows[0].height = Inches(0.5)
table.columns[0].width = Inches(2)
```

### Images

```python
from slidecraft.pptx.shapes.picture import Image

img = Image.from_file("photo.png")
slide.shapes.add_picture(img, Inches(1), Inches(1), Inches(4), Inches(3))
```

### Charts

```python
from slidecraft import ChartData, ChartType

cd = ChartData()
cd.categories = ["Q1", "Q2", "Q3", "Q4"]
cd.add_series("Sales", [100, 200, 150, 300])
cd.add_series("Costs", [80, 160, 120, 240])

slide.shapes.add_chart(ChartType.COLUMN, cd, Inches(1), Inches(1), Inches(6), Inches(4))
```

### Units

```python
from slidecraft import Inches, Cm, Pt, Emu

Inches(1)    # 914400 EMU
Cm(2.54)     # ~914400 EMU
Pt(72)       # 914400 EMU
Emu(914400)  # Direct EMU value
```

## Migration from python-pptx

| python-pptx | slidecraft |
|---|---|
| `from pptx import Presentation` | `from slidecraft import Presentation` |
| `prs = Presentation()` | `prs = Presentation()` |
| `prs.slides.add_slide(layout)` | `prs.slides.add(layout)` |
| `from pptx.util import Inches` | `from slidecraft import Inches` |
| `from pptx.util import Pt` | `from slidecraft import Pt` |
| `from pptx.dml.color import RGBColor` | `from slidecraft import RGBColor` |
| `from pptx.enum.text import PP_ALIGN` | `from slidecraft import PP_ALIGN` |

## License

MIT
