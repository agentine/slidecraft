# Changelog

## 0.1.2 (2026-03-31)

- Fix: add `c:autoTitleDeleted` to chart XML to prevent blank pages in PowerPoint
- Docs: fix README code examples to be fully copy-pastable
- New: add `examples/` directory with 6 runnable example scripts (basic presentation, charts, tables, images, text formatting, multi-slide)

## 0.1.1 (2026-03-31)

- Fix: sync slide element trees to part blobs in `Presentation.save()` — ensures modified slides are written correctly
- Security: pin all CI GitHub Actions to commit SHAs for supply chain security
- CI: bump `actions/setup-python` from 5 to 6, `astral-sh/setup-uv` from 5 to 7

## 0.1.0 (2026-03-14)

Initial release.

- OPC package layer (read/write .pptx ZIP containers)
- Presentation model (Presentation, Slide, SlideLayout, SlideMaster)
- Shapes and text (TextFrame, Paragraph, Run, Font, ShapeCollection)
- Tables (rows, columns, cell text, cell merge)
- Images (PNG, JPEG with auto-detection); `Image` re-exported from top-level package
- Charts (bar, column, line, pie with data binding)
- Unit conversions (Inches, Cm, Pt, Emu)
- RGBColor support
- Type annotations (mypy --strict clean)
- Python 3.10+ support
- Security: centralized XML parsing via `parse_xml()` helper with ZIP entry size limits
- Lint: ruff enforced in CI; all unused imports and dead code removed
