---
name: study-notes-summarizer
description: "This skill turns a folder of mixed-format study notes (Word docx or doc, PDF text or scanned, images, PowerPoint pptx, Excel xlsx, audio/video recordings, and zip/rar archives containing nested notes) into organized, per-month summary Word documents. Audio and video (lecture recordings, etc.) are transcribed to text via local, free Whisper speech-to-text. Notes are classified into six smart sections: vocabulary, grammar, listening, speaking, reading, writing. A section is omitted when the source notes contain nothing for it. Use this skill when a user wants to consolidate many scattered notes — including recordings — into clean revision summaries, especially before an exam. Subject-agnostic (language, exam prep, coursework). All parsing uses free local tools; no paid API required."
agent_created: true
---

# Study Notes Summarizer

## Overview

Turn a pile of scattered study notes (Word / PDF / images / PPT, possibly dozens of files) into
**per-month, section-classified** exam-prep summary Word documents. The whole pipeline uses only
free local tools — no paid API or external connector required.

- **Input**: a local folder containing the user's original note files (`.docx` / `.doc` / `.pdf` / images / `.pptx` / `.xlsx` / audio+video recordings / `.zip` / `.rar`).
- **Processing**: ① parse everything into Markdown (**cross-platform, no longer macOS-dependent**) →
  ② aggregate by month (filename `x.xx`) → ③ intelligently classify into six sections →
  ④ render into a formatted `.docx` with **scripts/render_docx.py (pure python-docx, zero WorkBuddy dependency)**.
- **Output**: `final/{Month} Notes Summary.docx`, plus `final/{Subject} Core Material Compilation.docx`
  for undated general material.

## When To Use

- The user says "organize my notes into a summary / exam-prep review doc / monthly digest".
- The user has a bunch of `.docx`/`.pdf`/image/PPT notes they want merged, classified, and de-duplicated.
- The user is preparing for an exam and needs material organized by vocabulary/grammar/listening-speaking-reading-writing.
- Trigger keywords: notes summary, review materials, exam-prep organization, monthly archive, classify documents.

## Pipeline

### Step 0 — Confirm parameters
Confirm with the user (or read memory): ① path to the notes source folder; ② subject name (e.g. "German B1");
③ preference for where undated material should go. The source folder is **only read/parsed — never moved,
renamed, or deleted**.

### Step 1 — Parse (scripts/parse_notes.py)
Run the unified parser to convert all supported formats in the source folder into Markdown:
```bash
python3 scripts/parse_notes.py --source "source-folder" --out "workspace/parsed" --log "workspace/parsed/processed.log"
```
- The script dispatches by extension: `.docx`→python-docx, `.doc`→cross-platform fallback chain
  (convert to `.docx` first via macOS `textutil` / Windows Word COM / LibreOffice, then read with
  python-docx to **preserve tables**; text-only fallback when docx conversion is unavailable),
  `.pdf` text→pdfplumber,
  `.pdf` scanned/images→tesseract OCR (scanned PDFs rendered via PyMuPDF, no poppler needed),
  `.pdf` watermark-overlay→auto-detected (text layer is mostly repeated header/watermark lines while the
  real content lives in page images, e.g. slide decks exported to PDF) and the whole file is re-OCR'd,
  `.pptx`→python-pptx, `.ppt`→LibreOffice/textutil, `.xlsx`→openpyxl (each sheet → Markdown table),
  audio/video (`.mp3`/`.m4a`/`.wav`/`.flac`/`.ogg`/`.aac`/`.mp4`/`.mov`/...)→local Whisper speech-to-text
  (faster-whisper; model auto-downloaded, language auto-detected),
  `.zip`/`.rar`→auto-extract and recurse into nested notes (PDF/DOCX/.../audio), images→OCR.
- `processed.log` records processed files, so **re-runs never re-parse**; if an engine is missing the file
  is skipped with a hint and retried automatically once the engine is installed.
- Oversized files are auto-split into `_part1.md / _part2.md`.

### Step 2 — Read and classify
Read `parsed/*.md`, decide which sections each note hits per `references/section_guide.md`,
and flag original errors for the ⚠️ Common-mistake annotations.

### Step 3 — Generate summary (references/summary_prompt.md)
Follow the template in `references/summary_prompt.md`:
- Merge files of the same month (`x.xx` → corresponding month) into one Markdown summary.
- Organize as "Vocabulary → Grammar → Listening → Speaking → Reading → Writing",
  **omitting any section with no content**.
- Capture original errors as ⚠️ Common-mistake annotations.
- Undated general material becomes its own "Core Material Compilation" doc (unless the user asks to fold it into a month).

### Step 4 — Render Word (pure python-docx, no external dependency)
Feed the Step 3 Markdown to **scripts/render_docx.py** to render the `.docx`,
output `final/{Month} Notes Summary.docx`. This script uses the open-source `python-docx` library,
has no dependency on WorkBuddy's built-in skill, and behaves identically on Windows / macOS / Linux.
```bash
python3 scripts/render_docx.py "summaries/August Notes Summary.md" -o "final/August Notes Summary.docx"
```
- Formatting rules are implemented inside the script, matching `references/section_guide.md`:
  six-section heading colors, vocabulary tables (header shading), ⚠️ common-mistake cards with light-red
  shading, metadata box with light-blue shading, page numbers in the footer.
- Undated material's "Core Material Compilation" is rendered the same way:
  `python3 scripts/render_docx.py "summaries/Core Material Compilation.md" -o "final/German B1 Core Material Compilation.docx"`.

### Step 5 — Record and deliver
- Update the project's `processed.log` (done automatically by the script) and memory
  (which months are generated, which are pending).
- Present the generated `.docx` to the user via present_files.

## Parameters

| Parameter | Description | Default |
|---|---|---|
| Source folder | Directory of the user's original note files (absolute path) | Required |
| Subject name | e.g. "German B1", used for titles and compilation doc naming | Required |
| Workspace | Where parsed/ summaries/ final/ live | Sibling of source folder or specified |
| Undated material destination | Separate compilation / fold into a month / both | Separate compilation by default |

## Prerequisites (all free, cross-platform, **zero-manual**)

This skill is **zero-manual to initialize**: on first run the script auto-installs whatever environment is
missing, so the end user usually needs to install nothing themselves.

- **Python dependencies: auto-installed.** At startup the script detects missing libraries
  (`python-docx / pdfplumber / python-pptx / img2pdf / pytesseract / pillow / pymupdf / openpyxl / rarfile / [pywin32 on Windows] / pdf2image`)
  and `pip install`s them automatically (manifest in `requirements.txt`). You can also pre-install manually:
  ```bash
  pip install -r requirements.txt
  ```
  Offline environments can set `NOTES_SKIP_DEP_INSTALL=1` to disable auto-install.
- **OCR engine `tesseract`: auto-installed when missing.** If the script finds no tesseract, it calls the
  platform's native package manager to install it and re-detects: Windows→`winget install UB-Mannheim.TesseractOCR`;
  macOS→`brew install tesseract` (Linux/apt needs sudo and prints a manual prompt). You can also run
  `winget install UB-Mannheim.TesseractOCR` manually. Set `NOTES_SKIP_TESSERACT_INSTALL=1` to disable.
  Note: for a subject like German B1 where notes are all text-based, OCR is never triggered; OCR only applies to scans, photos, and watermark-overlay slide-deck PDFs.
- **`.doc` / legacy `.ppt` parsing (cross-platform, table-preserving)**: the script converts `.doc` to
  `.docx` with the best available engine and reads it via python-docx, so **tables are preserved on every
  platform** (no macOS-only text flattening). Engine order: 1) macOS built-in `textutil` (`-convert docx`,
  plain-text fallback if that fails); 2) Windows **Microsoft Word** (`pywin32`, auto-installed) via COM;
  3) **LibreOffice** (soffice / libreoffice); 4) Linux `antiword` (text-only). If none are available it
  prints a clear install hint; or you can "Save As .docx" in Word/WPS before feeding.
- **Scanned PDF rendering uses pure-pip PyMuPDF (`pymupdf`), no poppler needed**; the poppler route is kept
  as an option (`brew install poppler` / `apt install poppler-utils` + `pdf2image`).
- **OCR language packs auto-download**: default `chi_sim+deu+eng` (Chinese+German+English, subject-agnostic); a missing pack
  is **auto-downloaded from GitHub on first run** to `~/.notes_ocr_tessdata`. Override with
  `NOTES_OCR_LANG=eng` to recognize English only, etc. Install all packs offline with `brew install tesseract-lang` (macOS).
- **Speech-to-text for audio/video recordings (zero cost, fully local)**: `.mp3`/`.m4a`/`.wav`/`.flac`/`.ogg`/`.aac`
  and video containers (`.mp4`/`.mov`/`.webm`/`.m4v`) are transcribed with **faster-whisper** (open-source, runs on
  CPU, no cloud, no paid API). Audio decoding uses **PyAV**, which **bundles its own ffmpeg**, so no separate
  ffmpeg install is needed. The Whisper model is **auto-downloaded from HuggingFace on first use** and cached;
  default model is `small` (good multilingual/Chinese accuracy). Tune with `NOTES_WHISPER_MODEL=base` (faster, lighter)
  or `medium` (best accuracy), and `NOTES_WHISPER_LANG=zh` (or `en`) to force a language instead of auto-detect.
- **Final `.docx` rendering uses open-source `python-docx`** (in `requirements.txt`), done by
  `scripts/render_docx.py`, **zero built-in WorkBuddy skill dependency**, fully local and cross-platform.

If the environment is still unusable, the script prints a clear self-check list and install hints — it never fails silently.

## Notes

- **Read-only, never modify originals**: parsing and summarizing happen on copies/derived files; the source
  folder's originals stay untouched.
- **Incremental-friendly**: `processed.log` guarantees "drop one, parse one, resume on interruption", ideal for
  long-accumulated notes.
- **Subject-agnostic**: German B1 is just an example; swap in English, grad-school exams, any subject by
  changing the subject name and section weighting.
- **Scan/handwriting image quality** determines OCR quality; clean print is best, sloppy handwriting may need human proofreading.
