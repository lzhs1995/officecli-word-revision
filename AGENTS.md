# Agent entry point

Read [README.md](README.md) (or [繁體中文](README.zh-TW.md)), [docs/agent-quickstart.md](docs/agent-quickstart.md), then [the canonical skill](skills/officecli-word-revision/SKILL.md) before working on documents.

- Select the manuscript or thesis scenario explicitly; preserve original inputs and existing task identities.
- Treat `audit` as potentially native: it checks tools and defaults to Word rendering.
- Serialize operations sharing a Word/Zotero integration instance; honor shared ownership, locks and unresolved operations. Do not invent a parallel lane by changing paths.
- Keep statistical execution and NotebookLM disabled in this workflow. Separate document QA from scientific acceptance.
- Use public synthetic examples; do not commit private manuscripts, credentials, personal paths or execution logs.
- A client may or may not discover this file automatically. When unsure, explicitly read the linked entry points; do not claim universal agent support.

This thin entry does not replace the skill's scenario rules or grant permission beyond the active task.
