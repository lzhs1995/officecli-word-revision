---
name: officecli-word-revision
description: Revise journal manuscripts with native Word tracked changes or format doctoral theses with institution-specific styles, sections, headers, page numbering, TOC, captions, and cross-references using OfficeCLI and Microsoft Word. Use for academic DOCX revision, Zotero-preserving edits, dissertation formatting, or repeatable Word QA workflows.
---

# OfficeCLI Academic Word

Use the `wordrev` CLI for document production. First identify the scenario; do not mix content revision and thesis-wide formatting in one write job.

## Scenario Routing

- `manuscript_revision`: journal papers, proposals, or reports that require native Word insertions/deletions, an accepted copy, Zotero fields, or immutable statistical-result traceability. Read [manuscript-revision.md](references/manuscript-revision.md).
- `thesis_format`: doctoral-thesis page layout, styles, sections, headers/footers, page numbering, TOC, captions, cross-references, footnotes, and reference-list layout. Read [thesis-format.md](references/thesis-format.md). When `profile=ruc-doctoral-2026`, also read [ruc-doctoral-2026.md](references/ruc-doctoral-2026.md).

Use separate sequential jobs when both are needed: finish and approve content revisions first, then format the accepted document.

## Common Workflow

1. Create or inspect a schema 2.0 job using [job.schema.json](references/job.schema.json). Existing schema 1.0 manuscript jobs remain valid.
2. Keep `notebooklm` set to `disabled` and `statistics.allow_execution` set to `false`.
3. Audit before any layout write:

```bash
wordrev audit --job /absolute/path/job.json --resume
```

4. Review the audit and resolve ambiguous style or section mappings. Then build the clean layout candidate:

```bash
wordrev layout --job /absolute/path/job.json --resume
```

5. Build and verify. Native tracked revisions are compiled only for `manuscript_revision`, or when thesis `output.tracked_formatting=true` is explicitly requested.

```bash
wordrev build --job /absolute/path/job.json --resume
wordrev verify --run-dir /absolute/path/run-dir --resume
```

6. Inspect the generated PDF and key-page contact sheet. The contact sheet automatically includes the title/body opening sample, the last page, and both sides of each landscape run; add other required pages with 1-indexed `visual_qa.touched_pages`. Never overwrite the source or an existing final delivery.

For manuscripts without an explicit venue style guide, read [generic-journal-layout.md](references/generic-journal-layout.md). New manuscript jobs opt in to this policy through `publication_layout`; legacy jobs remain unchanged until the field is added deliberately.

## Guardrails

- In `manuscript_revision`, treat the source DOCX as the only format master. In `thesis_format`, the selected versioned institution profile is the format authority while the source remains the content authority.
- Serialize all Word, Zotero, and same-file OfficeCLI writes. Microsoft Word automation is protected by a cross-process lock; a lock timeout is a hard failure with owner/path/wait diagnostics, never permission to start a concurrent writer.
- Word helper scripts must address the owned document explicitly and must not bring Word to the foreground with application-level `activate`. A timeout is recorded as return code 124 and terminates the launched process group before the phase fails.
- Run read-only QA in parallel only after the DOCX hashes are frozen.
- Preserve Zotero fields as live `ADDIN ZOTERO_ITEM CSL_CITATION` fields; never type an author-year citation as a substitute.
- Do not invoke NotebookLM. Local notes may document requirements, but they do not authorize a NotebookLM call.
- Do not launch R, Stata, Python analysis, or model estimation. Consume `analysis_manifest.json` and stable result keys only.
- On macOS, refresh Word fields and export the authoritative PDF through Microsoft Word automation; do not rely on `officecli refresh` for DOCX.
- Treat `view issues` as a review queue. Never delete blank paragraphs, breaks, or section boundaries solely because they are reported as issues.
- Use `--resume` so unchanged phases remain cached.
- For `publication_layout` jobs, absent an explicit width rule, `table.fit_mode` defaults to `content`: native Word AutoFit to Contents after font/indent normalization. Save its geometry and native execution evidence; OOXML/flextable autofit alone is insufficient. Inspect long intervals and descriptive-table continuation in the PDF.
- The clean layout candidate must equal the **accepted** content, not a composite containing old deleted paragraphs/tables. Tune pagination on this candidate before compiling native revisions.
- Statistical manifests must enumerate every consumed artifact with `path` and SHA-256. Missing hashes fail closed; legacy jobs may use a versioned sidecar, never an unverified skip.
- A Zotero VBA launcher returning is not Refresh completion. Keep the owned document open until an update is observed, verify all dynamic fields, and inspect any modal/timeout before resuming. Record CSL disambiguation display changes separately from non-citation content.
- QA reports must name their coverage: full row/column/key sets, units, fields, and cross-section terminology. File integrity, original-model convergence, bootstrap finite estimates, replicate-level convergence, semantic test cases, and causal assumptions are separate statuses.
- After changing the pipeline, adapter, schema, or profile, treat older manifests and benchmarks as historical evidence until current implementation fingerprints have completed `build` and `verify`. Scope acceptance claims precisely: distinguish a synthetic fixture from a real chapter, and a repeated-source batch from heterogeneous chapters.
- Treat `run_manifest.json` command ranges as global half-open indexes: `commands[command_start_index:command_end_index]`. Do not infer a phase failure from commands outside that range.
- A nonblank PDF is not visual acceptance. When `publication_layout` is present, the adapter must return a passing `publication_layout_qa` based on the final Word-rendered PDF, including object-page narrative checks.
- Resolve actual typography, not just style names. A style can inherit a different size through `basedOn`; do not clear direct formatting until the intended effective font is known. A user-approved `format_contract` binds body/table/caption typography and object count to final-file QA.
- For new or revised Chinese manuscript jobs, explicitly populate `format_contract` before writing: table Chinese uses SimSun, Latin letters/numbers use Times New Roman; use the user's or audited mother's point size, not a past project's hardcoded size. Apply `han-left-nonhan-right`: cells containing Han are left-aligned; Latin/numeric-only cells are right-aligned, including headers. Mixed-language cells retain script-specific fonts. A thesis profile or explicit user rule takes precedence.
- The current manuscript's approved body/table size is 10.5 pt; this is a job value, not a universal thesis setting. All target tables must be covered, including unchanged rows of a touched descriptive table. Record exceptions only for genuine diagram carriers, never for uninspected old numerical tables.
- Table zero indentation must override both character units and length units explicitly. Deleting `firstLineChars` can inherit a two-character indent that Word materializes on save. Test after Word accepts/saves, not only in pre-render XML.
- Parse `w:b`/`w:bCs` values: an element with `val=0`, `false` or `off` means false. Never use element existence as the bold state or automatically bold all table headers.
- Word can remove redundant `w:jc` on save. Validate resolved paragraph alignment (including style inheritance and Word's left default), not element presence. In cumulative redlines, a fully row-deleted old table must not interrupt the accepted-view caption binding to its replacement. `tblGridChange` supports an ID but not author/date attributes; keep authorship on the associated format revisions and manifest, never invent invalid XML.
- Bind every table/figure to its actual caption: table caption above, figure caption below unless the user's format authority says otherwise. Caption text being present on a page is not a placement check. Count diagram carrier tables as figures, not data tables.
- The pipeline's shared final table/format gates are mandatory when configured; project adapters may add checks but cannot bypass them. Preserve negative mutation fixtures for inherited font size, Word-roundtrip indent, bold=false, wrong cell alignment, and missing/misplaced captions.
- Confirm clean-copy revision count independently in OOXML, including `pPr/rPr/rPrChange` on the last paragraph mark. OfficeCLI's revision query and Word bulk acceptance may omit that edge case. PDF export must hide revision display as well as disable revision printing; a gray review gutter or formatting balloon fails visual acceptance.
- After native content-fit, check the final Word PDF for atomic numeric tokens, long confidence intervals and formulas. A reported split not reproduced in the current PDF is not authority to change widths. For explicitly bounded decimal-result tables, use `scripts/rendered_numeric_qa.py` and retain negative split-token fixtures; never concatenate PDF lines to conceal splitting. Report the exact covered tables and manually inspect uncovered ones.
- Wording/field QA must enumerate document, footnotes, endnotes, headers and footers, not just body paragraphs. Evidence classes remain separate: structure, visual rendering, numeric source agreement, execution reproduction and methodological validity. Passing one does not establish the others.

See [workflow.md](references/workflow.md) for cache invalidation and failure recovery.
