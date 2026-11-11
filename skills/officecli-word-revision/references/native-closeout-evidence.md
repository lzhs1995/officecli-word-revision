# Native document closeout evidence

Apply these rules to Word/Zotero and RUC document production. This reference does not enable NotebookLM inside a `wordrev` job; external review remains the owning thesis workflow's responsibility.

- Separate content revision from institution-format jobs. Freeze each input and record OfficeCLI/Word versions, institution profile, adapter hashes and phase fingerprints before final verification.
- Serialize Word/Zotero operations through the shared OS lock and explicit task ownership. Offline review of frozen files can proceed in parallel. A previous owner's pending transaction belongs to that owner; neither lock-file deletion nor a document-count guess authorizes takeover.
- Before normally closing a saved user document, make and hash a dated recovery copy. A disk copy does not preserve unsaved memory edits. Do not discard edits, force-quit applications or answer a new modal under an unrelated one-time authorization.
- When establishing an unverified or changed Zotero route, use a small native test before chapter-scale work. Reuse still-valid evidence for an unchanged route. Start from a citation without a bibliography when testing `ZoteroAddEditBibliography`; keep positive controls immutable. A synthetic static field fixture cannot establish native refresh/save/reopen capability.
- Macro return and HTTP acceptance are not native completion. Require visible citation/bibliography consistency, save/close/reopen evidence and a real-field audit. Same-item references in different fields are valid; duplicate URIs within one field, `dontUpdate`, TEMP fields and fake bibliographies fail.
- Preserve raw `-128`, `-1712` and `-1743` evidence with the failing phase and sender. Diagnose the exact owned document and modal state. Retry only after the original operation has returned and evidence supports changed conditions, within the task's existing authorization and bounds; an uncertain operation must first be recovered.
- Use the existing Word-container PDF staging service. Bind its complete path, retain timeout staging and samples, normally close only the identified owned copy, and acknowledge idle pending through the official interface. Do not reset TCC, delete a shared lock or clear another task's pending state.
- Submission PDF evidence requires a successful Word export receipt bound to the current bytes, valid output hash/pages, unchanged input/staging DOCX and confirmed cleanup. Existing evidence can be reused for the same bytes and scope. A PDF left by a failed attempt, an unverified directory artifact or a preview exporter is not equivalent evidence.
- Resolve effective Word styles after save. Explicitly zero character and length first-line/hanging indents in tables. Use Han/mixed-left and non-Han-right alignment under the approved profile; absence of a direct alignment element does not prove incorrect alignment.
- Validate the running thesis-title header against an independent title authority, not only against the job's own title field. A chapter title mistakenly configured as the thesis title can otherwise pass self-consistency QA. Check inherited indentation and actual PDF centering; a header-only change still creates a new PDF acceptance identity.
- Count captions by their resolved caption style and actual object binding. Narrative paragraphs beginning with a table/figure reference are not captions, and multiple image representations can belong to one figure.
- A full contact sheet checks global layout; detailed checks must state inspected pages. An unchanged-page review may be inherited only with explicit page-geometry/pixel equivalence and new-hash binding; inspect every changed page. Do not claim a fresh manual inspection of pages reviewed by equivalence.
- Keep document integrity, numerical archive matching, full model reproduction, Bootstrap replicate convergence and scientific validity as separate statuses. Preserve warnings and scope limitations in the final report.
- A clean native roundtrip does not validate a method-table definition. Check categorical logic and the current chapter's terminology when importing historical tables; a confirmed new semantic defect prevents closeout despite older passing reviews.
- After a frozen artifact has been consumed elsewhere, write a new final manifest instead of overwriting its old manifest. Keep raw failures, later successful receipts and any report corrections in the audit trail.
- Bind unchanged companion files by their exact hashes when only one document changes. Refresh all downstream evidence for the changed document; never transfer a passing review to a new hash or let a prior pass conceal a newly confirmed defect.
- When the controlling workflow changes from dual agent to authorized solo execution, keep the same native/format/visual gates and label self-review accurately. The absent executor is not a reason to abandon authorized work or to claim independent consensus.


### 页眉重建后的正文缩进继承

R18实证：RUC的`_replace_story_text`重建页眉段落时仅设置center，原先显式零缩进被删除；正文默认`firstLineChars=200`继承到页眉，Word保存后出现`firstLine=420`，PDF页眉实际右偏约10.5pt。页眉标题与字体正确、`jc=center`以及常规RUC PASS都不能覆盖这一缺陷。应在布局器完成页眉重建后显式覆盖长度和字符两类缩进，保存重开后按样式继承重新审计，并在最终PDF测量正文版心中心。负控制必须覆盖“长度为0但字符缩进仍继承”的情形；任务本地修复不要改写已归档候选或共享脚本。

### Running stories and bounded visual inheritance

- Apply the zero-indent repair after the layout engine rebuilds headers and footers. Check all applicable length/character indentation attributes against the installed OOXML schema, including start/end aliases. Preserve PAGE fields and odd/even alignment. An offline rebuild proof is a prerequisite, never evidence that Word has saved and reopened the candidate.
- Re-audit resolved stories after Word saves and reopens. Derive the expected center and outer edge from the current DOCX section margins, verify those margins against the independent versioned institution profile, and compare the actual PDF positions. Text, font and paragraph alignment alone do not establish position. Require a PDF-page-to-section map for mixed or landscape sections instead of applying portrait constants.
- When the task requires full running-story verification, render the applicable headers/footers with expected and actual guide lines; bind measurement/style dependencies and record image review. Optional manifest keys must not bypass those declared requirements. This is a scoped task check, not a new gate for unrelated or already frozen documents.
- If body regions are unchanged, precisely scoped geometry/word-position/pixel equivalence may bind a prior body-only visual review. Every changed header/footer remains a current review obligation. Preserve earlier whole-document visual failures and do not call an equivalence proof a new full-page manual inspection.

### Audit resource use and content successors

- `audit.render_source=false` does not prove a whole `wordrev audit` invocation is offline: toolchain discovery may still query Word. Perform ZIP/XML/schema/field checks independently while another task owns Word, and defer the complete pipeline until a formal slot handoff.
Treat a desktop reservation as a bounded native work package, not the lifetime of the research task. Prepare immutable inputs and review scripts before admission. After the owned native operation is confirmed complete and closed, frozen ZIP/XML, numeric and PDF review can proceed independently. Return the registered resource promptly once its original release conditions are actually met; do not retain it solely for content review, NotebookLM, other-document preparation or the final narrative report. A later native phase requires a valid reservation under the original controller and current authority; never reuse a released token or change an in-flight attempt.

Inspect the actual command path before calling it offline. `wordrev audit` can query Word during toolchain discovery even with `render_source=false`; `layout` and `verify` can export through Word. Run independent frozen-file checks while waiting, rather than treating those complete pipeline commands as offline. The existing PDF service allows unrelated open documents, but this does not relax the registered broker's handoff or pending-state requirements.

- A visual contract for header/footer-only changes cannot require unchanged body pixels after an authorized content correction. Bind the exact allowed text delta to source and final DOCX, compare all actual PDF pages, review every changed page, and prove the remaining content and numerical results are preserved. Never compare a new PDF to itself as evidence that the prior body was retained.
- Bind the RUC QA's actual source hash to the repaired candidate, in addition to checking job paths and accepted/native output hashes. A self-consistent output chain from the wrong input must fail. After adding this guard, verify existing immutable receipts and content deltas offline instead of replaying an already successful native transaction.
- A same-hash companion-file reuse map must validate the referenced acceptances, their states and exact file hashes. Keep planned evidence paths distinct from accepted evidence. Preserve old maps and scope strings; derive current report scope from the approved frozen manifest. Imported copy, lock and session dependencies belong in the tool fingerprint set, with late snapshots labeled by their true capture time.

### Zotero Web API receipts and per-item versions

This is a Web API/library rule, not proof of concurrent native Word integration.
In an observed successful multi-item create response, item A had version 2 and
item B version 3, while `Last-Modified-Version` was 3. The captured official
[write documentation](https://www.zotero.org/support/dev/web_api/v3/write_requests)
describes a shared version for successful objects, so this observed difference
remains a documentation/response discrepancy with cause unresolved. Preserve the
raw HTTP response and
its per-member success/failure map; do not classify a successful create as failed
solely because those numbers differ.

For a subsequent conditional single-item write, bind the exact item key and its
own current version (for example, `If-Unmodified-Since-Version: 2` for A and `3`
for B), rather than copying the multi-create response's library header to both.
If a local assertion failed after successful creation, keep that original failure,
reconcile the exact created keys/versions from its receipt, and continue only the
remaining authorized operations under the original request budget. Do not repeat
the create POST. Preserve UNKNOWN if the original response cannot settle the write.
Successful creation or offline tests alone do not establish concurrent update
correctness, a real 412 conflict response or protection against stale overwrites;
each requires its own actual response and readback before reporting PASS.

這是 Web API／文獻庫規則，不證明原生 Word 整合可以並行。一次已觀察到的成功
多項建立回應中，項目 A 的 version 為 2、B 為 3，而 `Last-Modified-Version`
為 3。官方寫入說明描述成功項目共用版本，因此這次差異應保留為文件與實際回應不一致，
成因尚未確認；不能宣稱官方從無相等描述或服務端異常已定位。保留原 HTTP 回應及逐項成功／
失敗結果，不能只因版本數字不同就判定建立失敗。

後續對單一項目的條件寫入，必須綁定該項目的確切 key 及其目前版本，例如 A
使用 `If-Unmodified-Since-Version: 2`、B 使用 `3`；不能把多項建立回應的
文獻庫標頭一律套給兩者。成功建立後若本地斷言失敗，保留原失敗紀錄，以原回執
核實已建立的 key／版本，只接續原請求預算內尚未執行的已授權操作，不重送
建立 POST。原回應不足以確認寫入時，繼續保留 UNKNOWN。成功建立或離線測試
不等於並行更新、真實 412 衝突或防止舊版本覆寫已通過；各項必須有實際回應及
讀回證據後才能報 PASS。

A bounded test in an isolated synthetic library subsequently returned HTTP 204
for both conditional item updates: A 2→5 and B 3→4. Their client request intervals
overlapped by 3.076 seconds; this does not measure simultaneous server execution
or establish a general throughput gain. An A update with stale version 2 returned
412; the final GET still reported version 5 and the accepted updated title. The
whole sequence used eight requests and four attempted writes, including the
rejected stale write, with no create replay or production-library writes. This
validates that particular Web API conditional-write sequence, not native Word
plugin concurrency, arbitrary multiwriter loads or every conflict scenario.

後續在隔離的合成測試文獻庫中，兩個條件式項目更新皆實際回傳 HTTP 204：A 由
2→5、B 由 3→4。客戶端請求區間重疊 3.076 秒，不等於量測到服務端同時執行，
也不能外推一般吞吐效能。以舊版本 2 再更新 A，實際回傳 412；最後 GET 仍為
版本 5，且已接受的新標題未被覆寫。全流程累計 8 個請求、4 次寫入嘗試（包含
被拒的舊版寫入），未重送建立 POST，未寫入生產文獻庫。這只驗證該次 Web API
條件寫入鏈，不證明原生 Word 插件並行、任意多寫負載或所有衝突情境。
