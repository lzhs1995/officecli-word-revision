# Scripted Word PDF on macOS

Use `scripts/word_pdf_service.py INPUT.docx OUTPUT.pdf`. The thesis-refiner
`docx-to-pdf.js --backend=word` and the thesis adapter call this same service.
Requires Microsoft Word, Python 3, `pdfinfo` and `pdftotext` (Poppler). No OfficeCLI
PDF exporter, GUI clicks, VBA installation or cloud upload is used.

## Output path is a capability

Apple Events permission, Full Disk Access and sandbox access to a new output
directory are distinct. In the 2026-09-15 Chapter 6 test, Word opened the 13 MB
DOCX and resolved its full path, but PDF export timed out after 180 seconds.
A native process sample showed `MbuOpenSecurityScopedResourceSessionForWrite ->
MbuObtainSecurityBookmarksForFileURLs -> runModalForWindow`. Word was waiting
for output-path authorization; increasing the timeout did not address that wait.

The verified route:

1. Claim the exclusive output and copy the frozen DOCX byte-for-byte into a unique run directory under
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wordrev-pdf/`.
2. Acquire the persistent OS lock at `~/.cache/wordrev/word-automation.lock`.
3. Open that copy; bind Word's document by its complete POSIX path and record the actual Word PID/user.
4. Export a new PDF inside that same Word container directory; close only the owned document without saving.
5. Freeze the exact native PDF bytes while holding the lock, then release the lock.
6. Validate PDF header, parsing, nonempty text, page count and hashes outside the native lock.
7. Publish without overwriting an existing file; remove only this successful run's private staging.

Export does not rewrite the source or refresh Zotero. Unsaved Word edits are not
implicitly consumed. Freeze/save the candidate before export. Other open Word
documents are not an automatic failure. All participating writers must use the
shared lock. A lock file's existence does not mean its OS lock is held; never
unlink a persistent flock file to force entry.

## Evidence and recovery

Every run writes `OUTPUT_STEM.word-run-ID/result.json`, `stages.jsonl`, stage
AppleScripts, `pdfinfo.txt` and extracted text beside the requested PDF.
Export completion requires `success=true`, `cleanup=PASS`,
`input_unchanged=true`, the exact PDF hash and a parsed page count. This does not
establish visual, NLM, citation-display or statistical acceptance. Quartz metadata
alone is not proof that the intended Word document was exported.

`-1712` is a timeout, not a diagnosis. Inspect the failed stage and take a native
`sample WORD_PID 3 1 -file sample.txt`. A bookmark modal, unresolved Zotero
operation, permission denial (`-1743`) and active pagination require different
recovery. Do not repeat exports, close unrelated documents, kill all Word
processes, suppress alerts or reset TCC automatically. The service retains
uncertain staging and a pending recovery marker.

`native_closed=true` and `native-closed.json` establish the owned native phase's
closure and frozen output, not QA success or external ticket release. A later QA
failure returns `NATIVE_CLOSED_QA_FAILED` without creating a native pending marker.
A close failure returns failure, retains the unpublished native file and marks
pending; no QA or publication is attempted. The result fields and successful
`EXPORT_PASS` contract remain compatible, but a close failure is no longer
reported as an export success. `os_lock_wait_seconds` and
`os_lock_held_seconds` measure only the OS lock. This function does not manage
the resource broker: `broker_managed=false`, `logical_reservation_seconds=null`.
Its caller must record reservation acquisition/release and duration separately.
The optional `on_native_closed` callback runs after lock release and before QA;
it can invoke an already-authorized outer reconciliation protocol, not infer it.

New jobs may select `--worker-id mac-test|windows-test --worker-registry PATH`.
Registration binds machine, user, OS session, current-user Word instance and an
exclusive output root. Unknown workers, alternate lock paths, mismatched identities
and unverified Zotero routes are rejected. Aliases do not create another native
lane on the same instance; Word-only and legacy Word/Zotero work share the original
lock and pending marker. No shared broker resource names are changed here.
Windows uses the same API with a local byte-range lock and one bounded COM
transaction; it refuses an existing Word process in the registered session.
It binds the owned document's `Window.Hwnd` to the unique newborn PID/start time,
closes its document and requests a typed normal Quit only for its empty instance.
The child receipt records `QUIT_REQUESTED`; after that PowerShell host exits, a
separate read-only process observer must confirm disappearance before the parent
can record native closure. See [Windows worker evidence](windows-word-pdf.md). `--native-only` retains a closed frozen PDF for
external offline QA and reports `NATIVE_CLOSED_AWAITING_QA`, never full PASS.

After explicitly resolving the owned operation and verifying Word has no open
documents (or is not running), acknowledge it with:

```bash
python3 scripts/word_pdf_service.py --ack-pending
```

This serialized check refuses while Word has open/unreadable documents. Preserve
the failed receipt before recovery. Do not manually delete the pending marker.

## Separate OfficeCLI policy

Word PDF export neither calls OfficeCLI nor queries GitHub releases. For audits,
record the selected binary, version, digest and smoke test.
`WORDREV_OFFICECLI=/absolute/path/officecli` selects an isolated tested binary.
Keep explicit job version requirements and upgrade outside production; do not
classify `require_latest` mismatch as Word PDF failure or silently disable it.
A binary smoke test does not establish RUC layout acceptance.

The Chapter 6 regression produced 76 A4 pages with the EXPORT stage completing
in about 13 seconds after container staging. The same source hash previously
timed out at 180 seconds outside the container. This is a verified case, not a
guarantee that every future `-1712` has this cause.

Microsoft documents the separate external-file consent requirement for sandboxed
Office on Mac in [Request access to multiple files](https://learn.microsoft.com/en-us/office/vba/office-mac/grantaccesstomultiplefiles).
The container workaround above was verified locally; it is not a quoted Microsoft
recommendation or a claim that changing AppleScript to VBA removes sandboxing.

## Bind an unexecuted benchmark to its actual code

Before running a prepared benchmark, inspect and fingerprint the source paths and
remote bundle it will actually execute, including imported runtime/worker modules
and its QA entry. An installed release or a passing test of that release does not
upgrade a previously frozen benchmark. If that unused plan needs a fix, preserve
it and create an explicit successor under the original controller and authority.
Prove the old members have not executed; unknown execution is not an unused member.
Give the successor distinct attempt/output paths, retain the original locks and
resource protocol, and bind QA to the successor plan/version explicitly. Preparation
and negative tests are not completed serial/parallel measurements.

執行尚未開始的基準測試前，核對並固定實際將執行的來源路徑、遠端套件、匯入的
runtime／worker 模組及 QA 入口。已安裝新版或新版測試通過，不會自動更新舊的
凍結計畫。需要修正尚未執行的計畫時，保留原件，在原控制器與原授權下建立明確
後繼；先證明舊成員沒有執行，執行未知不等於未執行。後繼使用獨立嘗試與輸出
路徑，沿用原鎖及資源協議，QA 顯式綁定後繼計畫與版本。準備或負例測試通過
不代表已取得串行／並行效能結果。

## Release a closed failure without changing its outcome

Keep business outcome, native closure and logical reservation state separate.
A failed export or QA result can still have a confirmed closed native phase. Only
when the original task-bound receipt proves that exact attempt closed, reconcile
through the original controller's fresh observe/release checks and record the
release readback. Preserve the failure and do not rerun the export just to return
the reservation. If closure or dispatch is UNKNOWN, preserve the original pending
state and attempt budget; a failure code, zero windows or a new empty directory is
not closure evidence. Do not release a foreign reservation or bypass its checks.

業務結果、原生階段關閉和邏輯資源票據必須分開記錄。匯出或 QA 失敗仍可能已有
確定關閉的原生階段；僅在原任務綁定回執證明同一次操作已關閉後，才經原控制器
重新觀察、核准歸還並讀回票據狀態。業務仍為失敗，不為歸還資源而重跑匯出。
關閉或派送若為 UNKNOWN，保留原 pending、嘗試次數與額度；失敗碼、零視窗或
新的空目錄都不是關閉證據。不得代歸還他方票據或跳過原檢查。
