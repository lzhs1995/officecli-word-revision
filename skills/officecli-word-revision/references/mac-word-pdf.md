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

1. Acquire the persistent OS lock at `~/.cache/wordrev/word-automation.lock`.
2. Copy the frozen DOCX byte-for-byte into a unique run directory under
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wordrev-pdf/`.
3. Open that copy; bind Word's document by its complete POSIX path.
4. Export a new PDF inside that same Word container directory.
5. Validate PDF header, parsing, nonempty text, page count and hashes.
6. Publish to the requested destination without overwriting an existing file.
7. Close only the owned document without saving; remove successful staging.

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
