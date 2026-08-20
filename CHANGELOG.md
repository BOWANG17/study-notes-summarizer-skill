# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/), and this project adheres to
[Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added
- `.zip` / `.rar` archives are now **auto-extracted and recursed into** (nested archives supported); contents are parsed like ordinary files with `归档名__内部相对路径` naming.
- `.xlsx` spreadsheets are now parsed (each sheet → a Markdown table) via `openpyxl`.
- **Speech-to-text for audio/video recordings** (`.mp3`/`.m4a`/`.wav`/`.flac`/`.ogg`/`.aac`/`.mp4`/`.mov`/`.webm`/`.m4v`):
  transcribed locally with **faster-whisper** (free, CPU, no cloud API). Audio decoding uses PyAV, which
  bundles its own ffmpeg — **no system ffmpeg install needed**. The Whisper model auto-downloads
  from HuggingFace on first use; configurable via `NOTES_WHISPER_MODEL` (default `small`) and `NOTES_WHISPER_LANG`.

### Fixed
- **tessdata location on Intel Macs**: `_default_tessdata_dir()` now probes `tesseract --print-tessdata-dir` and the Homebrew `share/tessdata` layouts (`/usr/local` and `/opt/homebrew`), so OCR works when brew installs the binary in `bin/` but language packs in `share/tessdata`.
- **Watermark-overlay PDFs are now detected and re-OCR'd**: PDFs whose text layer is mostly repeated
  header/watermark lines (typical of slide decks exported to PDF — the real content lives in page images)
  were previously mis-detected as "text-extractable" and the real content was silently lost. `parse_pdf`
  now checks for a watermark-overlay signature (repeated lines > 40% of all lines, or < 15 unique content
  lines) and re-OCRs the whole file instead of trusting the noise.
- **Low-memory OCR of large scanned PDFs**: `_ocr_pdf` now streams page-by-page with PyMuPDF (render one
  page → OCR → release) instead of materializing every page image in memory, so 100+ page decks no longer
  risk out-of-memory kills (poppler remains the fallback path).
- Default OCR language is now **`chi_sim+deu+eng`** (Chinese + German + English, subject-agnostic) instead of `deu+eng` — Chinese notes are recognized out of the box.
- rar install hint updated: Homebrew removed the `unrar` formula, so the message now suggests `brew install unar` (or Linux `apt install unrar`); `rarfile` auto-detects unrar / unar / bsdtar.

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
