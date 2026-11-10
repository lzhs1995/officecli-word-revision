# Windows Word-only PDF worker

The bounded Windows backend uses installed Microsoft Word, PowerShell, and its
loaded Microsoft.Office.Interop.Word assembly. Default PDF QA also requires
Poppler (`pdfinfo` and `pdftotext`). It does not install applications, start Zotero,
refresh citations, or make a second writer safe on the same Word instance.

One synthetic document has completed the native export/close/exit lifecycle and
separate PDF checks. That document exercised tables, equations, a footnote and
tracked changes. This is bounded evidence, not acceptance of every Windows
version, arbitrary manuscripts, live Zotero fields, or a performance claim.

## Explicit registration and invocation

The registry is local configuration, not a distributable machine inventory. From
the repository root, obtain the actual identity without starting Word:

```powershell
python -c "import sys,json; sys.path.insert(0,'skills/officecli-word-revision/scripts'); from word_workers import local_identity; print(json.dumps(local_identity()))"
```

Create an existing, exclusive output directory and fill a private registry with
that exact identity. The values below are placeholders, not valid registration:

```json
{
  "schema": 1,
  "workers": {
    "windows-test": {
      "machine": "EXACT-MACHINE",
      "user": "exact-user",
      "session": "exact-session-number",
      "platform": "Windows",
      "word_instance": "Microsoft Word/current-user-session",
      "zotero_instance": null,
      "output_root": "C:\\word-test\\outputs"
    }
  }
}
```

The input must be frozen and the output must not already exist. This command
starts Word and performs a native operation; it is not a read-only preflight:

```powershell
python skills/officecli-word-revision/scripts/word_pdf_service.py C:\word-test\input.docx C:\word-test\outputs\result.pdf --worker-id windows-test --worker-registry C:\word-test\workers.json --native-only
```

`--native-only` retains the closed, frozen PDF in the evidence directory and
reports `NATIVE_CLOSED_AWAITING_QA` with `success=false`. Exit code 0 means this
limited phase completed with unchanged input and cleanup; it does not mean PDF
QA passed or the requested publication path exists. Read the returned evidence
paths and perform offline QA. Omit the flag when the service should perform its
normal PDF QA and publication after releasing the OS lock.

## Ownership and two separate exit receipts

The worker rejects mismatched machine/user/session, an existing Word process in
its session, unknown registrations, alternative lock/pending paths and an
unverified Zotero route. Word-only aliases share the original per-user lock and
pending marker with legacy Word/Zotero work. Cross-device orchestration and any
external resource reservation remain the caller's responsibility.

1. After creation, require one newly started Word process and an empty, private
   automation instance. Open only the staged read-only input, then bind its
   `Window.Hwnd` to the recorded PID, birth time and session before export.
2. Export document content without markup, close only the owned document, and
   verify that document/window/protected-view collections and background activity
   are known empty. Use the loaded PIA for a typed normal
   `Quit(ref 0, ref Type.Missing, ref Type.Missing)` call.
3. Preserve `native-com.json` with binding schema 3 and `QUIT_REQUESTED`. A returned
   Quit call is not proof that the application process has already disappeared.
4. Once the original PowerShell host exits with code 0, a fresh read-only observer
   checks the exact PID/birth/session and remaining Word processes. It never
   attaches COM or requests Quit. Only a valid empty observation permits
   `native-parent.json` to report `COMPLETE` and `native_closed=true`.

The parent preserves the child's actual return code and receipt hash. A missing
or invalid child receipt, child error, remaining/reused PID, foreign Word process,
identity mismatch, observer failure or unrecordable parent receipt leaves closure
unconfirmed. Later empty processes do not rewrite a failed child as success.
The outer service retains pending recovery evidence and does not publish an
uncertain native output. Do not replay the export, delete pending state, force
quit Word or close another document to make the result pass.

`native-closed.json` records frozen output and owned native closure; offline QA,
visual review and any external ticket release remain separate. Synthetic unit
tests exercise these failure contracts without invoking Word or COM.
