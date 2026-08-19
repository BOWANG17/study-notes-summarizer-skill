# Monthly/Batch Summary Generation Prompt Template

The following steps turn "already-parsed Markdown notes" into a structured exam-prep summary.
Read `references/section_guide.md` first to understand the six sections and smart-parsing rules.

## Input
- The contents of several `parsed/*.md` files for one month (or one batch).
- Subject name (e.g. "German B1"), month label (e.g. "July 2026").

## Steps
1. **Skim each file** and tag which sections it hits (Vocabulary / Grammar / Listening / Speaking / Reading / Writing).
2. **De-duplicate and merge**: merge repeated vocab, patterns, and templates across files, keeping the most complete wording and examples.
3. **Organize by section**: strictly follow "Vocabulary → Grammar → Listening → Speaking → Reading → Writing".
4. **Smart omission**: if a section has no content in this batch, don't write its heading or body.
5. **⚠️ Common-mistake annotations**: pull errors from the source and mark them with blockquotes (see section_guide).
6. **Give quick-recall hints**: at the end of each section, note in 1–3 sentences the points "most likely tested / most easily mistaken before the exam".

## Output format (Markdown)
```
# {Subject} {Month} Notes Summary

> Covered files: {file1}, {file2}… (N files in total)
> Generated: {YYYY-MM-DD}

## Vocabulary
(themed vocab table; omit entirely if none)

## Grammar
(grammar points + examples + ⚠️ common mistakes; omit entirely if none)

## Listening
(omit entirely if none)

## Speaking
(omit entirely if none)

## Reading
(omit entirely if none)

## Writing
(writing templates + point-deduction traps + ⚠️ common mistakes; omit entirely if none)
```

## Convert to Word document
Feed this Markdown to **scripts/render_docx.py** (pure python-docx, no built-in WorkBuddy skill needed) to render:
```bash
python3 scripts/render_docx.py this_summary.md -o "final/{Month} Notes Summary.docx"
```
The script auto-produces a clean, page-numbered `.docx`: six-section heading colors, vocab tables (header shading),
⚠️ common-mistake cards with light-red shading, metadata box with light-blue shading.
Output path convention: `final/{Month} Notes Summary.docx` (e.g. `final/July Notes Summary.docx`).

## Separate compilation for undated files
If this batch includes undated general material, produce a separate `final/{Subject} Core Material Compilation.docx`
with the same structure, but its covered-files note points to those undated files.
