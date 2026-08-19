# Study Notes Summarizer

> A WorkBuddy Skill that turns a folder of mixed-format study notes (Word / PDF / images / PPT) into organized, per-month, section-classified revision-summary Word documents — using only free local tools, no paid API.

*A WorkBuddy skill that turns a folder of mixed-format study notes (Word / PDF / images / PPT) into organized, per-month, section-classified revision-summary Word documents — using only free local tools, no paid API.*

---

## Features

- **Multi-format parsing**: `.docx` / `.doc` / `.pdf` (text & scanned) / `.pptx` / images, all unified into Markdown.
- **Fully free local engine**: `python-docx`, `macOS textutil`, `pdfplumber`, `python-pptx`, `tesseract OCR`, `img2pdf` — no network, no per-use billing.
- **Month aggregation**: notes whose filename carries `x.xx` (e.g. `7.14`, `8.16`) are auto-grouped into that month's summary.
- **Six-section smart classification**: `Vocabulary / Grammar / Listening / Speaking / Reading / Writing`; a section with no source content is **omitted entirely**.
- **⚠️ Common-mistake annotations**: automatically catches errors in the source and flags them for exam avoidance.
- **Incremental-friendly**: `processed.log` records processed files; re-runs never re-parse, resume on interruption.
- **Cross-platform tolerance**: missing an engine gracefully skips with an install hint, never fails silently.
- **Subject-agnostic**: German B1 is just an example; swap to English, grad exams, any subject by changing the subject name.

## Directory structure

```
study-notes-summarizer/
├── SKILL.md              # Skill definition: triggers, full pipeline, parameters
├── scripts/
│   ├── parse_notes.py    # Unified free parser (multi-format → Markdown)
│   └── render_docx.py    # Markdown → formatted .docx (pure python-docx)
├── references/
│   ├── section_guide.md  # Six-section classification rules
│   └── summary_prompt.md # Summary-generation prompt template
├── requirements.txt      # Python dependency manifest
├── LICENSE               # MIT License
└── README.md             # This file
```

## Prerequisites

This skill is **zero-manual to initialize**: on first run the script auto-installs the needed environment, so the end user usually installs nothing themselves.

- **Python dependencies: auto-installed.** At startup the script detects and `pip install`s missing libraries (manifest in `requirements.txt`). Manual pre-install:
  ```bash
  pip install -r requirements.txt
  ```
  Offline: set `NOTES_SKIP_DEP_INSTALL=1`.
- **OCR engine `tesseract`: auto-installed when missing.** Windows→`winget install UB-Mannheim.TesseractOCR`; macOS→`brew install tesseract` (Linux/apt needs sudo and prints a manual hint). Manual: `winget install UB-Mannheim.TesseractOCR`. Disable with `NOTES_SKIP_TESSERACT_INSTALL=1`.
  Note: German B1 notes are all text-based and never trigger OCR; OCR only applies to scans/images.
- **`.doc` / legacy `.ppt` parsing (cross-platform)**: the script auto-picks: Windows **Microsoft Word** (`pywin32`, auto-installed) → **LibreOffice** → macOS `textutil` → Linux `antiword`. Otherwise it prints a clear hint; or "Save As .docx" in Word/WPS first.
- **Image / scanned-PDF OCR**:
  - Scanned-PDF rendering prefers pure-pip **PyMuPDF** (`pymupdf`), **no poppler needed**; poppler route available (`brew install poppler` / `apt install poppler-utils` + `pdf2image`).
  - Default OCR language `deu+eng`; missing packs **auto-download** to `~/.notes_ocr_tessdata` on first run, override via `NOTES_OCR_LANG` env var.
- Final `.docx` rendered by open-source **python-docx** (in `requirements.txt`) via `scripts/render_docx.py`, **zero built-in WorkBuddy skill dependency**, fully local and cross-platform.

If the environment is still unusable, the script prints a clear self-check list and install hints — never fails silently.

## Installation

Install as a WorkBuddy custom skill:

```bash
# Option A: clone into the skills directory
git clone <repo-url> ~/.workbuddy/skills/study-notes-summarizer

# Option B: copy manually
cp -r study-notes-summarizer ~/.workbuddy/skills/
```

After install, say "use study-notes-summarizer to organize my notes" in a chat to trigger it.

## Usage

**Simplest flow** (let the AI take over):

1. Drop your original note files into any local folder (e.g. `~/Desktop/my-notes`).
2. Tell WorkBuddy: *"Use study-notes-summarizer to organize `~/Desktop/my-notes`, subject is German B1"*.
3. The skill auto: parses → aggregates by month → classifies into six sections → uses `scripts/render_docx.py` (pure python-docx) to generate `final/{Month} Notes Summary.docx`.

**Run just the parser** (command line):

```bash
python3 scripts/parse_notes.py \
  --source "/path/to/notes" \
  --out   "/path/to/notes/../parsed" \
  --log   "/path/to/notes/../parsed/processed.log"
```

Arguments:

| Arg | Description | Default |
|---|---|---|
| `--source` | Notes source folder (required) | — |
| `--out` | Parsed output directory | `<source>/../parsed` |
| `--log` | `processed.log` path | `<out>/processed.log` |
| `--force` | Ignore `processed.log`, force re-parse all | off |

## Parameters / Configuration

| Parameter | Description | Default |
|---|---|---|
| Source folder | Directory of original note files | Required |
| Subject name | e.g. "German B1", used for titles and compilation naming | Required |
| Workspace | Where `parsed/` `summaries/` `final/` live | Sibling of source or specified |
| Undated material destination | Separate compilation / fold into a month / both | Separate compilation by default |

## How it works

```
feed folder → parse_notes.py parses to Markdown
           → AI aggregates by month (x.xx) + classifies into six sections
           → render_docx.py (pure python-docx) renders to .docx
           → can be re-run on a schedule daily (incremental)
```

See the Pipeline section of `SKILL.md` for details.

## Notes / Limitations

- **Read-only, never modify originals**: parsing and summarizing happen on derived files; source originals stay untouched.
- **`.doc` / `.ppt` legacy formats are cross-platform**: on Windows parsed by Microsoft Word (pywin32), or via LibreOffice / macOS textutil / Linux antiword — no manual conversion needed.
- **Default OCR language `deu+eng`** (German+English, fits German study). Override via env var: `NOTES_OCR_LANG=chi_sim+eng` for Chinese+English, etc.; missing packs auto-download on first run.
- **Scan / handwriting image quality** determines OCR quality; clean print is best, sloppy handwriting may need human proofreading.

## License

[MIT](./LICENSE) — free to use, modify, and distribute, including commercially.
