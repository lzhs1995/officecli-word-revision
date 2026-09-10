# Optional RStudio integration

No runtime dependency on ClaudeR is introduced. This repository is the single
source for Word automation. The RStudio workbench remains the source for live R
execution, compareGroups and paired-mval skills.

The producer freezes statistical outputs and supplies a contract such as:

```json
{"artifacts": [{"path": "/absolute/path/table.csv", "sha256": "<actual SHA-256>"}]}
```

Use actual file hashes, not the example placeholder. The Word job must enumerate
every consumed result/image, use stable keys, and declare
`statistics.allow_execution=false`. A manuscript adapter maps the intended rows,
figures and text. Editing fonts or pagination must not start R or Stata.

Updating one repository does not require reinstalling the other. Future optional
integration should pin a tested release and link here rather than copy skill
source. This initial release does not claim a newly tested end-to-end RStudio to
Word workflow or alter the existing workbench repository.
