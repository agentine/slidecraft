# slidecraft — Modern Python PowerPoint Library

**Replaces:** python-pptx (20.5M downloads/month, bus factor=1, dormant since Aug 2024)
**Package name:** slidecraft (verified available on PyPI)
**Language:** Python
**License:** MIT

## Problem

python-pptx is the only open-source Python library for creating, reading, and updating PowerPoint (.pptx) files. It has 20.5M monthly downloads and 29.3K GitHub dependents. The project has been dormant since August 2024 (19 months), with 442 open issues and 78 unresolved PRs. It has a single maintainer (scanny) and no funding. The only fork attempt (python-pptx-ng) failed to gain traction (beta, last release Jan 2024).

## Scope

A modern, type-annotated Python library for creating, reading, and updating PowerPoint 2007+ (.pptx) files. The .pptx format is Open XML (ECMA-376), a ZIP container with XML parts.

### Core Deliverables

1. **OPC Package Layer** — Read/write ZIP-based Open Packaging Convention containers. Parse `[Content_Types].xml`, relationships (`.rels`), and content parts.

2. **Presentation Model** — Object model for presentations, slides, slide layouts, and slide masters. Support for adding, removing, reordering slides.

3. **Shapes & Text** — Shape types: text boxes, auto shapes, placeholders, group shapes. Rich text support: paragraphs, runs, fonts, colors, alignment, bullet lists.

4. **Tables** — Create and modify tables with rows, columns, cell merging, and cell formatting.

5. **Images** — Insert, resize, and crop images (PNG, JPEG, SVG, EMF, WMF).

6. **Charts** — Basic chart support: bar, column, line, pie charts with data binding.

7. **Slide Layouts & Masters** — Access and apply built-in slide layouts. Theme color and font scheme support.

### Non-Goals (v1)

- Animations and transitions
- SmartArt
- VBA/macros
- Audio/video embedding
- Full theme editing
- Slide comments and notes (stretch goal)

## Architecture

```
slidecraft/
  opc/           # Open Packaging Convention (ZIP + XML)
    package.py   # Read/write .pptx ZIP containers
    content_types.py
    relationships.py
    part.py
  pptx/          # PowerPoint-specific model
    presentation.py   # Top-level Presentation object
    slide.py          # Slide, SlideLayout, SlideMaster
    shapes/
      shape.py        # Base shape, auto shapes
      textbox.py      # Text frames, paragraphs, runs
      placeholder.py  # Placeholder shapes
      group.py        # Group shapes
      picture.py      # Image shapes
      table.py        # Table shape
      chart.py        # Chart shapes
    text.py           # TextFrame, Paragraph, Run, Font
    theme.py          # Theme colors, fonts
    enum.py           # PowerPoint enumerations
  xml/            # XML helpers
    ns.py         # Namespace management
    builder.py    # Fluent XML element builder
    parse.py      # XML parsing utilities
  util/
    units.py      # EMU, Inches, Cm, Pt, Emu conversions
    color.py      # RGBColor, theme colors
```

### Key Design Decisions

- **Pure Python** — Use `xml.etree.ElementTree` from stdlib (no lxml dependency). Optional lxml acceleration.
- **Type annotations** — Full typing throughout, compatible with mypy/pyright.
- **Modern Python** — Python 3.10+ only. Dataclasses, `|` unions, match statements where appropriate.
- **Immutable-first** — Prefer returning new objects over mutation where practical.
- **Lazy loading** — Parse XML parts on demand, not on file open.
- **Compatibility** — Read files created by python-pptx and Microsoft Office. Write files that open correctly in PowerPoint, LibreOffice, Google Slides.

## Migration Path

Users should be able to migrate incrementally. The API will be similar to python-pptx where it makes sense, but modernized:

```python
# python-pptx
from pptx import Presentation
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[0])

# slidecraft (similar but typed)
from slidecraft import Presentation
prs = Presentation()
slide = prs.slides.add(prs.slide_layouts[0])
```

## Testing Strategy

- Unit tests for XML generation/parsing
- Round-trip tests: create → save → reload → verify
- Compatibility tests: open files created by PowerPoint, LibreOffice, python-pptx
- Golden file tests with known-good .pptx files
