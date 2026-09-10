# Pipeline stages and invalidation

The engine supports `manuscript_revision` and `thesis_format`. Schema 1.0 jobs are normalized to the manuscript scenario without rewriting the job file.

| Stage | Invalidated by | Main outputs |
|---|---|---|
| toolchain | TTL or tool version policy | version record |
| audit | source hash, OfficeCLI version, audit selectors | style/structure inventory, source PDF/contact |
| layout | scenario, scope, profile, source/content/assets/analysis hashes, style/section/table/field/publication-layout policies, citation contract | clean DOCX and layout PDF |
| revision | layout/metadata/author/date/adapter hashes | tracked DOCX and OfficeCLI batch; manuscript jobs only |
| word-postprocess | tracked or clean-layout hash and field contract | accepted/rejected manuscript files, or the clean thesis file |
| verify | final DOCX hashes, publication-layout policy and QA contract | numeric, structural and PDF QA |

Every phase is recorded in `run_manifest.json`. A phase is reusable only when its fingerprint matches and every recorded output still has the same SHA-256.

The JSON printed by `wordrev build` describes its terminal `word-postprocess` phase, not total wall time for every preceding phase. Use phase timestamps/durations in `run_manifest.json`, or a fresh `wordrev benchmark`, for performance claims. A pre-upgrade `status=complete` is not current-code proof: require current implementation fingerprints plus a freshly completed `verify` artifact.

`run_manifest.json` stores one append-only command array. Every phase records a global half-open range, `commands[command_start_index:command_end_index]`, so later `build` and `verify` invocations remain attributable. When exporting a CSV summary, encode repeated run samples as a JSON array inside the CSV cell; the canonical machine-readable evidence remains `benchmark.json`.

Before creating `run_dir`, `cache_dir`, chapter-batch definitions, or benchmark directories, runtime validation checks the statistics no-execution contract, scenario policies, output types/date, style-map value types, and the selected profile's identity, semantic version, and required sections. A malformed job therefore fails without leaving a partly initialized run tree.

`layout` is deliberately separated from `revision`: paginate and style the clean candidate first. Do not repeatedly compile hundreds of tracked changes to tune a page break.

The clean candidate must contain exactly the intended accepted content. A combined old/new staging document is only a revision compiler input, not a pagination candidate. Do not solve overflow by reducing the inherited table font to an unauthorized size.

For statistical jobs, use `{"artifacts":[{"path":"/absolute/results.csv","sha256":"<64 hex>"}]}`. Every consumed CSV/image must be listed. The parser accepts fully hash-bound legacy `data/script/outputs`, but rejects unhashed `tables/scripts` contracts instead of silently skipping them. Preserve the old manifest and create a canonical sidecar. Actual artifact hashes participate in cache invalidation; no document change starts a statistical estimator.

Audit invalidation follows source, OfficeCLI, selectors and audit implementation, not the complete project adapter. A content checkpoint stores matching prepared/metadata hashes before Word PDF export. If export fails, resume from that checkpoint; failed staging files are retained for diagnosis rather than unlinked while Word might still own them.

ZoteroRefresh is an asynchronous launcher. The owned-document helper waits for a document update, and the adapter verifies citation identities and non-citation content. A changed author-year suffix can be valid CSL disambiguation and must be logged. An unexpected modal is a failure requiring inspection, not proof that Refresh succeeded. Do not close unrelated documents or terminate the user's Word application to recover one job.

Acceptance reports should separately cover (1) frozen results integrity; (2) model/replicate states actually logged; (3) scale/semantic test cases; (4) manuscript wording and complete cell mappings; (5) native fields/revisions/OpenXML; (6) Word pagination and human visual review. Check counts summarize coverage but never replace it.

`final_format_qa.json` is a pipeline-owned gate, not an optional adapter claim. It runs on the Word-saved accepted file and, for `format_contract` jobs, the accepted-view content of the native redline. A passing numeric or OpenXML report cannot override a failing font, indentation, alignment, or caption-object contract. Typography policy and shared helper hashes invalidate layout/verification caches. Retain the prior report as historical evidence when a later, broader gate finds defects.

The pipeline never installs or upgrades OfficeCLI during document production. If the cached release check detects a newer version, upgrade separately, run a smoke test, then resume; the changed version naturally invalidates the format audit.

The Word stage accepts/rejects manuscript revisions, or selectively refreshes thesis fields, serially. On macOS it uses Microsoft Word automation for authoritative pagination and PDF export. It must not update Zotero ADDIN or legacy form fields as part of a blanket F9 operation.

Subprocess timeouts are hard failures recorded with `returncode=124`, `timed_out=true`, and `timeout_seconds`; the pipeline terminates the launched process group before writing the failed phase. Word/Zotero adapters must still close or identify any application-owned helper that the GUI application launches outside that process group.

Once the final files are frozen, independent read-only checks may run concurrently. NotebookLM has no code path in this workflow. Thesis jobs preserve semantic content and never execute statistics.

Use `wordrev benchmark --job JOB --cold 3 --warm 5` to measure a verified job. Both sample counts must be positive integers. If `benchmark_dir` is omitted, results are stored beside `run_dir` in a `*_benchmarks` directory; each invocation receives a collision-proof timestamped archive/work pair even when two runs start in the same second. Timing claims are job- and machine-specific; manuscript results are not promises for a full thesis.

Label benchmark fixtures honestly. Reusing one chapter twice is useful for proving independent child jobs and shared-cache restoration, but does not demonstrate heterogeneous chapter style mappings.

For opt-in journal layout, verify coverage before counting violations: target data-table cells must equal the alignment-, indent-, spacing-, and margin-checked sets. Locate figures from `pdfimages -list`, tables and narrative from PDF text, and fail any object page without the configured amount of narrative. A page-pixel nonblank check is only a baseline corruption guard.
