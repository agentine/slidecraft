# slidecraft

[![PyPI](https://img.shields.io/pypi/v/slidecraft)](https://pypi.org/project/slidecraft/)
[![Python](https://img.shields.io/pypi/pyversions/slidecraft)](https://pypi.org/project/slidecraft/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

Modern Python library for creating and modifying PowerPoint (.pptx) files.

A modern, type-annotated replacement for python-pptx. Pure Python, zero dependencies, Python 3.10+.

## Why slidecraft?

python-pptx has 20M+ monthly downloads but has been dormant since August 2024 with a single maintainer and no funding. slidecraft is a drop-in modern replacement: fully typed, actively maintained, and zero external dependencies.

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

prs = Presentation()                  # New blank presentation
prs = Presentation.open("in.pptx")   # Open existing
prs.save("out.pptx")                 # Save

prs.slides                            # SlideCollection
prs.slide_layouts                     # list[SlideLayout]
prs.slide_masters                     # list[SlideMaster]
prs.slide_width                       # Emu (read/write)
prs.slide_height                      # Emu (read/write)
```

### Slides

```python
slide = prs.slides.add()              # Add with default layout
slide = prs.slides.add(layout)        # Add with specific layout
prs.slides.remove(0)                  # Remove by index
for slide in prs.slides:              # Iterate
    ...
slide = prs.slides[0]                 # Index access
len(prs.slides)                       # Slide count
slide.shapes                          # ShapeCollection
```

### Text

```python
tb = slide.shapes.add_textbox(left, top, width, height)
tf = tb.text_frame
tf.text = "Hello"                     # Set all text (replaces paragraphs)
tf.word_wrap = True
tf.paragraphs                         # list[Paragraph]
p = tf.paragraphs[0]
p2 = tf.add_paragraph()               # Append a new paragraph
tf.clear()                            # Remove all text

p.text = "Line of text"
p.alignment = PP_ALIGN.CENTER
p.level = 1                           # Indent level (0-based)
p.space_before = Pt(12)               # Space above paragraph
p.space_after = Pt(6)                 # Space below paragraph
p.runs                                # list[Run]
run = p.add_run()                     # Append a new run

run.text = "Bold text"
run.font.bold = True
run.font.italic = True
run.font.underline = True
run.font.name = "Arial"
run.font.size = Pt(24)
run.font.color = RGBColor(255, 0, 0)
```

### Tables

```python
from slidecraft import Inches, Presentation

prs = Presentation()
slide = prs.slides.add()
table = slide.shapes.add_table(3, 2, Inches(1), Inches(1), Inches(6), Inches(3))
table.cell(0, 0).text = "Name"
table.cell(0, 1).text = "Score"
table.cell(1, 0).text = "Alice"
table.cell(1, 1).text = "95"
table.cell(2, 0).text = "Bob"
table.cell(2, 1).text = "87"
table.rows[0].height = Inches(0.5)
table.columns[0].width = Inches(3)
prs.save("tables.pptx")
```

### Images

```python
from slidecraft import Image, Inches, Presentation

prs = Presentation()
slide = prs.slides.add()
img = Image.from_file("photo.png")
slide.shapes.add_picture(img, Inches(1), Inches(1), Inches(4), Inches(3))
prs.save("images.pptx")
```

### Charts

```python
from slidecraft import ChartData, ChartType, Inches, Presentation

prs = Presentation()
slide = prs.slides.add()
cd = ChartData()
cd.categories = ["Q1", "Q2", "Q3", "Q4"]
cd.add_series("Sales", [100, 200, 150, 300])
cd.add_series("Costs", [80, 160, 120, 240])
slide.shapes.add_chart(ChartType.COLUMN, cd, Inches(1), Inches(1), Inches(6), Inches(4))
prs.save("charts.pptx")
```

See [examples/](examples/) for more complete runnable scripts.

### Units

```python
from slidecraft import Inches, Cm, Pt, Emu

Inches(1)    # 914400 EMU
Cm(2.54)     # ~914400 EMU
Pt(72)       # 914400 EMU
Emu(914400)  # Direct EMU value
```

### Enumerations

```python
from slidecraft import PP_ALIGN, PP_PLACEHOLDER, ChartType, MSO_SHAPE_TYPE

# PP_ALIGN — paragraph alignment
PP_ALIGN.LEFT
PP_ALIGN.CENTER
PP_ALIGN.RIGHT
PP_ALIGN.JUSTIFY
PP_ALIGN.DISTRIBUTE

# ChartType
ChartType.BAR
ChartType.BAR_CLUSTERED
ChartType.COLUMN
ChartType.COLUMN_CLUSTERED
ChartType.LINE
ChartType.PIE

# MSO_SHAPE_TYPE — identify shape kinds
MSO_SHAPE_TYPE.AUTO_SHAPE
MSO_SHAPE_TYPE.PICTURE
MSO_SHAPE_TYPE.TABLE
MSO_SHAPE_TYPE.CHART
MSO_SHAPE_TYPE.TEXT_BOX
MSO_SHAPE_TYPE.GROUP
MSO_SHAPE_TYPE.PLACEHOLDER

# PP_PLACEHOLDER — placeholder types
PP_PLACEHOLDER.TITLE
PP_PLACEHOLDER.BODY
PP_PLACEHOLDER.CENTER_TITLE
PP_PLACEHOLDER.SUBTITLE
PP_PLACEHOLDER.DATE
PP_PLACEHOLDER.FOOTER
PP_PLACEHOLDER.SLIDE_NUMBER
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
