# Changelog

## [Unreleased]

### Added
- Re-export `Image` class from top-level `slidecraft` package for direct import

### Fixed
- Centralized XML parsing via `parse_xml()` helper; added ZIP entry size limits to prevent ZIP bomb attacks
- Removed unused imports and dead code; added ruff to CI for lint enforcement

## 0.1.0 (2026-03-14)

Initial release.

- OPC package layer (read/write .pptx ZIP containers)
- Presentation model (Presentation, Slide, SlideLayout, SlideMaster)
- Shapes and text (TextFrame, Paragraph, Run, Font, ShapeCollection)
- Tables (rows, columns, cell text, cell merge)
- Images (PNG, JPEG with auto-detection)
- Charts (bar, column, line, pie with data binding)
- Unit conversions (Inches, Cm, Pt, Emu)
- RGBColor support
- Type annotations (mypy --strict clean)
- Python 3.10+ support
