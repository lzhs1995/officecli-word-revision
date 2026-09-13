# macOS TCC recovery for Zotero -> Word automation

## Symptom

On macOS, Microsoft Word may show the Zotero ribbon while Refresh, Add/Edit Citation, or Add/Edit Bibliography has no effect. Zotero may show its native missing-permission warning. Full Disk Access for Word and Zotero does not grant the separate Apple Events permission.

## Evidence-first diagnosis

1. Test the real Zotero ribbon route on a disposable copy.
2. Inspect System Settings -> Privacy & Security -> Automation -> Zotero -> Microsoft Word.
3. Treat a Terminal-to-Word `osascript` `-1743` result as evidence about the Terminal sender only.
4. A valid route must produce genuine `ZOTERO_ITEM` and `ZOTERO_BIBL` fields; a visible ribbon or plain bibliography text is insufficient.

## Recovery when `tccutil` cannot reset the entry

If a targeted command such as `tccutil reset AppleEvents org.zotero.zotero` fails, the per-user TCC database may be stale, permission-corrupted, or inconsistent. Preserve a backup and record the operation before using the manual recovery documented by the user:

```text
~/Library/Application Support/com.apple.TCC/TCC.db
```

With Word and Zotero closed, move the user-level `TCC.db` to a dated backup/trash location, restart macOS, and re-authorize only required Automation entries. This is a last-resort local recovery, not an automatic pipeline step: it can clear unrelated privacy grants and must never be performed without a rollback copy. Re-test Zotero -> Word after restart.

A successful recovery proves that rebuilding TCC resolved the local failure; it does not by itself prove database corruption or establish a universal fix for every Mac.

## Acceptance after recovery

On a safe copy, run Zotero Refresh and Add/Edit Bibliography, save, reopen, count live fields, export through Microsoft Word, and rerun content/NLM and RUC visual gates. Do not mark the thesis complete until post-insertion gates pass.
