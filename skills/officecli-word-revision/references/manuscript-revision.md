# Manuscript revision mode

Use `scenario=manuscript_revision` when the requested outcome is a native Word redline against an existing manuscript.

- The source document is the only format master. Audit its effective styles, direct formatting, captions, tables, fields, and page layout before editing.
- A project adapter identifies the exact paragraphs, tables, figures, citations, and immutable result keys to change.
- Iterate pagination on a clean candidate, then compile tracked insertions/deletions. Deliver tracked, accepted, and reject-all verification copies.
- If no journal-specific layout is supplied, enable and follow [generic-journal-layout.md](generic-journal-layout.md). Apply its table rules to the final data tables, not to arrow/diagram carrier tables, and verify pagination only after Microsoft Word renders the accepted DOCX.
- Preserve live Zotero fields. Citation text that must remain refreshable is inserted through the Zotero Word plugin or copied from a valid plugin-generated field with a unique citation ID.
- If an adapter invokes the Zotero Word macro, start Zotero in the background first, bind automation to the owned document by name/reference, and reject the layout if visible semantic text changes across refresh. Never use application-level `activate`, which can redirect unrelated keyboard input into Word.
- When results are supplied, verify an immutable `analysis_manifest`; do not launch the statistical program from the Word workflow.

Schema 1.0 jobs remain supported. New jobs should use schema 2.0 and declare the scenario explicitly.

Initialize a new project after preparing its content source and adapter:

```bash
wordrev init --scenario manuscript-revision \
  --source /absolute/path/original.docx \
  --content-source /absolute/path/revised-content.docx \
  --adapter /absolute/path/project_adapter.py \
  --revision-author Codex \
  --out /absolute/path/manuscript.job.json

wordrev audit --job /absolute/path/manuscript.job.json --resume
wordrev layout --job /absolute/path/manuscript.job.json --resume
wordrev build --job /absolute/path/manuscript.job.json --resume
wordrev verify --job /absolute/path/manuscript.job.json \
  --run-dir /absolute/path/run-dir --resume
```

Formatting-only thesis work does not use a project manuscript adapter. Route it to `thesis_format` instead of adding thesis-wide rules to this job.
