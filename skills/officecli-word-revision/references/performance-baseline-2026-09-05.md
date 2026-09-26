# Performance baseline, 2026-09-05

These measurements are evidence for this workstation, Word version, and fixtures. They are not promises for arbitrary theses.

| Scenario | Fixture | Cold/new-run median | Cold/new-run P95 | No-change resume median | Result |
|---|---|---:|---:|---:|---|
| `manuscript_revision` | v7.1 gold regression, native revisions + Zotero + 162 numeric checks | 41.74 s | 42.06 s | 0.015 s | pass |
| `thesis_format` | four-page RUC chapter sample, clean output | 5.55 s | 38.05 s | 0.005 s | pass |

The first thesis run after the final implementation-cache change took 41.66 seconds; the next two independent run directories took 5.01 and 5.55 seconds with shared audit/layout cache. The final v7.1 runs took 42.09, 40.63, and 41.74 seconds. Every v7.1 run retained the 608-revision contract and 162/162 numeric checks.

Cache evidence:

- A new run directory may report `shared_cache_hit=true` for `audit` and `layout`; it still runs Word postprocessing and verification.
- A no-change `--resume` in the same run directory validates recorded output hashes and returns without reopening Word.
- A source, profile, style/section/table/field contract, adapter, or engine hash change invalidates the affected phase.

The original workstation-specific receipts are retained in the owner's private
archive. They are historical local measurements, not artifacts shipped here or
acceptance of later versions. Use the current release's tests and a separately
scoped native fixture when measuring a new implementation.
