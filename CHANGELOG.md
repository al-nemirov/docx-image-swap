# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

---

## [1.0] - 2025-03-19

### Added

- Initial release of DOCX Image Swap
- CLI runner (`run.py`) with step-by-step pipeline execution
- GUI runner (`run.pyw`) built with ttkbootstrap (dark theme)
- 4-step pipeline:
  - **Step 1** — Extract all images (inline + floating) from DOCX
  - **Step 2** — Manual image swap pause (user replaces files)
  - **Step 3** — Reinject images back into the document
  - **Step 4** — Save output with configurable filename template
- JPEG optimization with configurable quality (1-100)
- DPI normalization to fix stretched images
- Aspect ratio preservation on reinsertion
- Configurable max dimensions (pixels and inches)
- Full document structure preservation (text, styles, formatting)
- JSON-based configuration (`config.json`)
- Per-step enable/disable support
