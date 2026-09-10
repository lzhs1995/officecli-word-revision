# Initial publication validation — 2026-09-10

Local command:

```text
python -m unittest discover -s skills/officecli-word-revision/scripts -p 'test_*.py' -v
Ran 86 tests in 3.338s
OK
```

The tests cover synthetic document transformations, contracts, mocked native
commands, numeric-token checks and process locks. They do not open Microsoft Word
or establish native Zotero/Word acceptance. A python-docx style-id lookup
deprecation warning remains.

The source CLI returned `wordrev 0.1.0.dev1`; `git diff --check` passed.
The separate Codex skill-frontmatter checker could not run in the bundled Python
because PyYAML was absent. This is not recorded as a passed check.

The known native AutoFit `-10006` failure is documented in the README. No native
Word/Zotero experiment, R/Stata execution or manuscript change was performed for
this publication. GitHub Actions results are a separate, externally visible
record, not inferred from this local result.
