# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/), and this project adheres to
[Semantic Versioning](https://semver.org/).

## [1.0.0] - 2026-08-19

### Highlights
- **Cross-platform.** One codebase now runs on **Windows, macOS, and Linux** — the macOS-only `textutil` dependency is no longer required.
- **Fully self-contained.** Final `.docx` rendering uses pure `python-docx` (`scripts/render_docx.py`), with **zero dependency on WorkBuddy's built-in skill** — the whole pipeline works in any AI tool that can run Python.
- **All documentation, prompts, and code comments are in English.**

### Added
- `scripts/render_docx.py`: a more capable Markdown → `.docx` renderer — six-section heading colors, vocabulary tables with header shading, ⚠️ common-mistake cards (light-red), a metadata box (light-blue), page numbers in the footer, and cross-platform CJK font selection (PingFang SC on macOS, Microsoft YaHei on Windows, etc.).
- **Zero-manual initialization.** On first run the script:
  - auto-installs any missing Python dependencies (`requirements.txt`),
  - auto-installs `tesseract` via the platform package manager when missing,
  - auto-downloads missing OCR language packs (default `deu+eng`) to `~/.notes_ocr_tessdata`.
- Scanned-PDF rendering via pure-pip **PyMuPDF** (`pymupdf`) — **no poppler required**.
- Cross-platform `.doc` / `.ppt` parsing chain: macOS `textutil` (docx-preserving) → Windows **Word COM** (`pywin32`) → **LibreOffice** → Linux `antiword`. Tables are preserved on every platform.

### Fixed (all 10 issues from the community bug list)
- **#5** OCR language code corrected from `de+eng` → `deu+eng`.
- **#6** `requirements.txt` now includes `pdf2image`.
- **#1** Same-named files of different formats (e.g. `a.docx` / `a.pdf`) no longer overwrite each other — the extension is embedded in the output name.
- **#2** Incremental mode now re-parses a file when it changes: `processed.log` records `name ⇥ mtime ⇥ size`.
- **#3** Re-runs clean up stale `_partN.md` / old outputs before writing.
- **#4 & #7** Mixed PDFs: each page is OCR'd only when it has `< 30` chars, so scanned pages are no longer dropped and short-text PDFs are no longer mis-detected as scans.
- **#8** PPT **tables** are now extracted.
- **#9** Exit code `2` when files existed but nothing was parsed, so schedulers/automations can detect failures.
- **#10** Sub-directories are now traversed (not just the top level).

### Changed
- Removed the redundant `scripts/build_docx.py` (superseded by `render_docx.py`).
- `SKILL.md`, `README.md`, and `references/*` rewritten in English and updated to the new cross-platform, self-contained pipeline.

### Known limitations
- `.doc` / `.ppt` parsing still needs one of Word / LibreOffice / textutil / antiword on the platform; without any of them those formats are skipped (with a clear hint).
- OCR requires `tesseract` to be installed; text-based notes (e.g. German B1) never trigger OCR.

## [0.2.0-beta] - 2026-08-17
- Self-contained `.docx` generation via `build_docx.py` (`python-docx`), removing the `tencent-docx` dependency.

## [0.1.0-beta] - 2026-08-17
- Initial beta release. macOS-focused, Chinese documentation, `tencent-docx`-based rendering.
