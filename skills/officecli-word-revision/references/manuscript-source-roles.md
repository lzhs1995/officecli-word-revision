# Source-bound manuscript paragraph roles

A uniform body rule must not turn an audited abstract, keyword paragraph or
bibliography into body text. New manuscript jobs may add
`format_contract.paragraph_roles`, an exact `path`, `bytes`, `sha256` pointer
to a sidecar matching [the selector schema](manuscript-source-roles.schema.json).
The source DOCX in that sidecar must be the job's exact `content_source`.
The sidecar supplies selectors only; actual source formatting remains the
authority. Existing jobs without this pointer keep their uniform contract.

Each entry declares the zero-based **accepted-view direct body paragraph**
index, role, full visible-text SHA and source style ID. The accepted-view count
includes ordinary blank paragraphs and excludes entirely deleted paragraphs.
A deleted paragraph mark with surviving content requires a new explicit mapping;
the checker refuses to guess how Word merges it.

Supported source roles are `abstract`, `keywords` and `bibliography_entry`.
Abstract and keyword paragraphs have explicit Chinese or English labels followed
by a colon. All such labels in the bound source must be selected. Bibliography
entries require the one explicit “参考文献”, “References” or “Bibliography”
heading index and an exclusive end index; every nonempty paragraph between
those anchors must be selected. Empty field terminators remain unselected.
Caption, image, equation, heading and note roles cannot be relabelled to bypass
their shared checks. Other layouts require their own reviewed contract extension.

Both shared apply and final audit use the same immutable source binding.
Special paragraphs retain source spacing, character/length indentation,
alignment, paragraph-mark typography and per-position run formatting, including
complex-script size. Missing `line`/`lineRule` values are retained as missing.
Equivalent run splitting/merging passes; moving fonts between character positions
fails. Tabs, break types and literal whitespace stay distinct. Apply can split
plain text/control runs at source font boundaries, but refuses unsafe splits
involving field instructions or existing run-format change records.

For a manuscript with already accepted minimal body/table repairs, the separate
`apply_source_roles(document, format_contract)` entry restores only the selected
paragraphs. An already correct role is a zero-change operation; this entry never
normalizes unrelated body/table properties. The full final `audit_contract`
still checks all body/table/caption/object rules. Full `apply_contract` retains
its existing uniform formatting scope.

This is a checked source contract, not an ignored-error list. All nonselected
body paragraphs, numerical tables, captions and object-count gates continue.
Sidecar/source drift, unknown keys, duplicate/out-of-range selectors, missing
roles, changed selected text and wrong formatting fail. Selected text changes
from real Zotero Refresh need a separately reviewed, freshly bound source
contract; this feature never authorizes skipping Refresh or replacing live
fields with plain text.

Pipeline preflight validates the binding before loading a project adapter or
creating run directories. Every phase rechecks its immutable input pins before
cache reuse; the role helper and both schemas participate in implementation
fingerprints. A new helper is not proof of native Word acceptance, final PDF
layout or NotebookLM review.
