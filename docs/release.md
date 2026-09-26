# Release boundaries

Current development version: Python `0.1.0.dev2`; RUC profile `1.1.0`.
This is a source/developer preview. It is not a completed native production
acceptance, a statistical audit, or a claim that all earlier defects are fixed.

## Automated gate

- Run all synthetic unit tests in this source tree.
- Build/install the wheel in an isolated environment; verify `wordrev --version`
  and packaged skill/schema/profile resources.
- Preserve the same software version in the VERSION file and Python metadata; a release tag, when published, must refer to that exact commit.
- Publish the source archive and SHA-256 only from a clean, committed tree.
- Use local machines or GitHub Actions, never VPS/production builds.

## Before a stable native release

- Reproduce the macOS Zotero→Word route after any TCC recovery on a disposable DOCX. Verify live `ZOTERO_ITEM` and `ZOTERO_BIBL` fields and observed Word Refresh completion; do not treat Full Disk Access or a Terminal `-1743` probe as sufficient.

- Resolve/reproduce the `track revisions` / AutoFit `-10006` failure on a disposable
  document without closing unrelated Word documents.
- Validate accepted/rejected native changes and all affected document stories.
- Confirm actual Word-saved content-fit, typography, atomic numbers, captions and
  full PDF layout. Do not infer native success from mocked unit tests.
- Verify dynamic Zotero citations using the Word plugin and an observed Refresh
  completion. No private library export is required for a public test fixture.
- Supply a portable manuscript adapter example. Existing private project adapters
  are not distributed; generic thesis formatting is a separate scenario.
- Report Windows/native support only after it is implemented and tested.

## Source of truth and rollback

The repository is the development source. Installed skills and `wordrev` commands
are versioned outputs, not the place to make untracked changes. Install a tested
commit/tag in a separate environment, back up any existing skill before replacing it,
and retain the prior environment for rollback. Running tasks keep their pinned
scripts; installation alone neither migrates their bindings nor proves client reload.
