# Release boundaries

Current development version: Python `0.1.0.dev2`; RUC profile `1.1.0`.
This is a source/developer preview. It is not a completed native production
acceptance, a statistical audit, or a claim that all earlier defects are fixed.

## Native synthetic verification, 2026-09-26

Microsoft Word 16.113.2 saved, closed and reopened the final synthetic example
twice, with nine pages each time. The declared structural checks, resolved TOC
targets, per-page footnote restart, separate bibliography paragraphs, PDF export
and owned-resource cleanup passed. The author inspected the nine-page contact
sheet and detailed figure/table and bibliography pages. Synthetic citations were
preserved; native Zotero Refresh was not run. This example does not implement or
certify the complete RUC double-sided template.

Three separate, immutable fixture versions were executed once each. Earlier
evidence is retained privately:

| Version | Observed issue | Disposition |
|---|---|---|
| v1 | Document-default eachPage disappeared on save; generated TOC names changed | Set explicit ordered section properties; compare actual target locations and text while retaining raw snapshots |
| v2 | Mechanisms passed, but a literal newline placed two bibliography entries in one paragraph | Add an optional actual-paragraph check and a paragraph-spanning synthetic field |
| v3 | Two notes in the same section on different pages both display 1; entries occupy separate paragraphs | Accepted for the stated synthetic mechanisms and visual scope |

The validated implementation is included in commit
`f899c2a011011bb64ecec85f61d987f4eb20f3b5`; 153 synthetic tests passed locally
and on Ubuntu/macOS CI. Native evidence fingerprints:

- Word-exported synthetic PDF: `91f33ef495a604f9484c92f401b0bf7d0e742b6630290036cb7c4d1cda699a37`.
- Author acceptance: `d3b451829ee33c775da0c84cffbeb4f13e3f62ad426f82f6f9f7a60dac270770`.

Original application logs and native fixture artifacts remain in the private
task archive. The public generator and tests are reproducible examples; Word's
metadata and generated identifiers do not promise byte-identical future files.

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

When an existing task pins installed script bytes, install the new wheel in a
separate version-and-commit directory. Point a distinct command and the skill's
new-task entrypoint at that environment; keep the old command and scripts in
place. Record the resolved executable, package version, packaged profile and
schema hashes, and the previous entrypoint bytes. Verify the installed command
with the synthetic positive and negative examples. Rollback restores the
entrypoint after checking that it has not changed again; it does not remove an
environment that another task has started using.
