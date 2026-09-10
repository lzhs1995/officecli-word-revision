# Thesis formatting mode

Use `scenario=thesis_format` when institutional formatting rules, rather than the source document's appearance, define the target layout.

## Start a job

Create a job first, then edit the generated JSON to add the audited style and section mappings.

```bash
# One chapter. Repeat --style for every role that exists in the chapter.
wordrev init --scenario thesis-format --scope chapter \
  --source /absolute/path/chapter-04.docx \
  --profile ruc-doctoral-2026 \
  --style body=/styles/Normal \
  --style chapter=/styles/Heading1 \
  --style heading1=/styles/Heading2 \
  --style heading2=/styles/Heading3 \
  --out /absolute/path/chapter-04.job.json

# Several chapters remain separate outputs; no merge step exists.
wordrev init --scenario thesis-format --scope chapter-batch \
  --source /absolute/path/chapter-01.docx \
  --source /absolute/path/chapter-02.docx \
  --profile ruc-doctoral-2026 \
  --out /absolute/path/chapters.job.json

# Already merged thesis. Page-number changes require an explicit section_map.
wordrev init --scenario thesis-format --scope full-thesis \
  --source /absolute/path/thesis.docx \
  --profile ruc-doctoral-2026 \
  --thesis-title "论文题目" \
  --out /absolute/path/thesis.job.json
```

Run the stages with resume support:

```bash
wordrev audit --job /absolute/path/thesis.job.json --resume
wordrev layout --job /absolute/path/thesis.job.json --resume
wordrev build --job /absolute/path/thesis.job.json --resume
wordrev verify --job /absolute/path/thesis.job.json \
  --run-dir /absolute/path/run-dir --resume
```

On the first Microsoft Word run, grant Word access to the stable staging directory recorded in `word_pdf_staging_dir`; later jobs reuse that authorization. Set `WORDREV_STAGING_DIR` before `wordrev init`, or pass `--word-pdf-staging-dir`, when the machine already has a dedicated authorized folder. Do not use a different staging folder for every job on macOS.

## Scopes

- `chapter`: format one chapter while preserving its section orientations and content. Do not create a global TOC or restart whole-thesis page numbering.
- `chapter_batch`: run the chapter workflow independently for each listed source. Share the profile and audit cache; do not concatenate documents.
- `full_thesis`: format an already merged thesis. Validate front matter order and, when explicit anchors are supplied, configure front-matter/body page numbering and global fields.

## Required decisions

- Map semantic roles such as `body`, `chapter`, `heading1`, `heading2`, `caption`, `footnote`, and `bibliography` to existing style paths. The audit may propose mappings, but build must stop on ambiguity.
- Every mapped `/styles/...` selector must exist in the source's audited Word styles. A role name that exists in the profile is not enough; a missing source selector fails before the layout batch runs.
- A custom profile JSON filename (without `.json`) must match its `profile_id`; `profile_version` uses `X.Y.Z`, and the profile must supply page/margin, style, header/footer, numbering, table, and rule sections. Page geometry applied after OfficeCLI remains profile-driven rather than falling back to RUC constants.
- Supply explicit section anchors before changing page-number regimes or creating fields. `validate_only` never reorders major document blocks.
- Keep `field_policy=preserve` unless the job supplies caption labels, bookmark names, and insertion anchors. Do not convert hand-typed captions or references by guesswork.
- Keep `table_policy=report_only` unless the tables to format are explicitly selected or all body tables are known to be statistical tables.
- Declare intentional recto blank pages in `visual_qa.allowed_blank_pages`; undeclared blank pages fail visual QA.
- `visual_qa.touched_pages` and `allowed_blank_pages` are 1-indexed. Page 0 is invalid, and an explicitly requested page beyond the produced PDF fails verification with a clear error.

## Output and QA

New RUC jobs default to `*_RUC格式化清洁稿_YYYYMMDD.docx` plus the authoritative Microsoft Word PDF, rule-by-rule difference report, structure/field QA, and a key-page contact sheet. The initialization command freezes `output.delivery_date` so resume behavior does not change across calendar days. `output.tracked_formatting=true` additionally creates `*_RUC格式修订稿_YYYYMMDD.docx`; it does not replace semantic or structural verification. Jobs created before this field was introduced keep their established output names.

On macOS the postprocessor refreshes TOC, SEQ, REF, PAGEREF, PAGE, NUMPAGES, and STYLEREF fields selectively through Microsoft Word. It skips Zotero ADDIN fields and legacy form fields. Automatic open-time field refresh is disabled on the produced clean file to avoid modal prompts.

Use `wordrev benchmark --job JOB --cold 3 --warm 5` only after verification passes. Benchmark a `chapter_batch` through its child jobs because the batch deliberately has no aggregate document.

Verify that paragraph/table/footnote text, canonical equation XML, embedded media payloads, Zotero fields, and unrelated Word fields are preserved. Counts alone do not prove equation or image preservation. Review all sections, not only the final section. Never delete blank paragraphs or section breaks solely because `view issues` reports them.

The key-page contact sheet automatically samples pages 1–3, the last page, and the entry/exit boundaries of every landscape-page run. Use `visual_qa.touched_pages` for known TOC, caption, footnote, reference, or other high-risk pages that cannot be inferred from PDF geometry alone.
