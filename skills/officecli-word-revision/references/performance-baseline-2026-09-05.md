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

Evidence files:

- Manuscript benchmark: `/Users/lzhs/Desktop/20251202 CMAverse_4_way Decomposition/20260904_basic_regression_eight_chains_v7_1/performance/word_revision_pipeline_v1/benchmarks/20260905T091100_xf95e0j6/benchmark.json`
- Thesis benchmark: `/Users/lzhs/Desktop/wordrev_v71_bench/thesis_skill_test_20260905/benchmarks/20260905T091004_hhm17a2b/benchmark.json`
- Real Chapter 4 verification: `/Users/lzhs/Desktop/wordrev_v71_bench/thesis_ch4_test_20260905/run/verification.json`
- Full-thesis field verification: `/Users/lzhs/Desktop/wordrev_v71_bench/thesis_full_test_20260905/run/verification.json`
- Chapter-batch verification: `/Users/lzhs/Desktop/wordrev_v71_bench/thesis_batch_test_20260905/run/chapter_batch_manifest.json`
