# Six-Section Summary Structure

This skill unifies notes from any subject into the six sections below. When generating each
monthly/topical summary, the model **only outputs sections that actually contain content for that
month/batch; a section with no relevant content is simply omitted** — no padding.

| Section | Meaning | Example content to include |
|---|---|---|
| **Vocabulary** | Words, phrases, collocations | Themed vocab tables, synonym distinctions, easily-confused words |
| **Grammar** | Grammar points, sentence patterns, rules | Tenses, cases, conjunctions, clauses, passive voice, preposition usage, with example sentences |
| **Listening** | Listening-related | High-frequency listening words, liaison/reduction, common scenarios, intensive-listening tips |
| **Speaking** | Speaking-related | Spoken patterns, topic material, common Q&A templates, fluency tips |
| **Reading** | Reading-related | Reading strategies, long/complex sentences, question-type techniques, high-frequency topic vocab |
| **Writing** | Writing-related | Writing templates, argument structure, universal sentence patterns, common point-deduction traps |

## Smart-parsing rules
- Read each parsed Markdown and decide which sections it hits.
- One note may span multiple sections (e.g. both speaking and writing practice) — split into the corresponding sections.
- **A section with zero content this month/batch → that whole section is not generated.**
- Section order is fixed: Vocabulary → Grammar → Listening → Speaking → Reading → Writing.

## ⚠️ Common-mistake annotation convention
Errors in the original notes (spelling, conjugation, collocation, subject-verb agreement, case errors, etc.)
should be marked under the relevant section with a blockquote:
`> ⚠️ Common mistake: <wrong original> → <correct form> (<reason>)`. This is the most valuable part of an
exam-prep summary and is on by default.

## Month-aggregation rules
- Notes whose filename contains `x.xx` (e.g. `7.14`, `8.2`) are grouped by month: all `7.xx` → July summary,
  all `8.xx` → August summary, and so on.
- Undated general material (e.g. "Pronoun Summary Table", "Speaking Corpus") is by default produced as a
  separate **Core Material Compilation** document, not forced into a month; the user may instead fold it
  into a month or keep both.
- If a month has many files (>8) or one file is very large, split into topical batches first, then merge
  into one big monthly document at the end.
