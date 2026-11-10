# Agent quickstart / Agent快速開始

This guide is a portable reading and execution entry point. It is not a claim that every client automatically discovers `AGENTS.md`, `CLAUDE.md`, `GEMINI.md` or a skill directory. Ask the client to read these files explicitly if discovery is unavailable. / 本頁是通用閱讀與執行入口，不保證任何客戶端自動載入；必要時明確指定閱讀路徑。

## Read before acting

1. Read [AGENTS.md](../AGENTS.md), [README](../README.md) [繁體中文說明](../README.zh-TW.md) or [简体中文指南](../README.zh-CN.md), then [SKILL.md](../skills/officecli-word-revision/SKILL.md).
2. Identify `manuscript_revision` versus `thesis_format`. Follow the selected reference; finish accepted content before a separate formatting job when both are needed.
3. Identify source and accepted content, the exact tool version, adapter/profile, output directory and current execution state. Never infer an existing job's completion from a file merely being present.

閱讀入口後，先辨別稿件修訂或論文排版；明確來源、已確認內容、版本、適配器／規範、輸出及原作業狀態。兩種工作都有時，先完成內容，再分開處理格式。

## Install and inspect

Run from a clone of this repository, with Python 3.10 or newer:

```sh
git checkout --detach 47f9d43be54f9320c14d32a39d4f0ce073bb4320
python3 -m venv .venv
.venv/bin/python -m pip install .
.venv/bin/wordrev --version
.venv/bin/wordrev --help
.venv/bin/wordrev init --help
```

The [package configuration](../pyproject.toml) defines the CLI entry point; these help/version calls do not start Word. Keep CLI and skill pinned to the same version. This installation does not configure MCP, Zotero or operating-system permissions. Register the skill only through the client's supported mechanism; do not overwrite an active installation.

命令採POSIX shell。安裝、版本與說明查詢不啟動Word；客戶端技能註冊須依其支援方式，CLI與技能固定同版，不覆寫執行中安裝。

## Start with an offline fixture

```sh
.venv/bin/python skills/officecli-word-revision/scripts/make_thesis_fixture.py --out-dir ./agent-example
.venv/bin/python skills/officecli-word-revision/scripts/thesis_integrity.py --docx ./agent-example/synthetic-thesis.docx --contract ./agent-example/contract.json --output ./agent-example/check.json
```

Use a new output directory. This creates synthetic material and checks its contract without Word, Zotero or network calls. It does not validate a real manuscript, native pagination or live Zotero refresh. / 使用新目錄；合成檢查不代表真實文稿、Word分頁或Zotero更新已通過。

## Real-job boundary

Inspect a schema 2.0 job against [job.schema.json](../skills/officecli-word-revision/references/job.schema.json). Existing schema 1.0 manuscript jobs remain supported. `init` creates a job configuration; it does not resolve every mapping or approve production writes.

- `manuscript-revision` initialization needs one source, `--content-source` and `--adapter`.
- `thesis-format` initialization needs `--scope` and a matching profile/style mapping. The bundled RUC profile is an example from a supplied snapshot, not current university certification. Other institutions require their own profile and checked rule provenance.
- Keep `notebooklm` disabled and `statistics.allow_execution=false`. Consume approved statistical outputs by manifest and hash; do not estimate models.
- `audit` runs a toolchain check and defaults to native source rendering. `layout` also invokes the audit. Neither is an offline-only entry point.

真實作業須審查設定與對應關係。`audit`預設可能啟動Word；即使只想看報告，也先協調原生資源。不得因設定檔已產生就推論寫入或學術判斷已獲批准。

Once the task authorizes the operation and the native resource is available, follow the scenario's `audit → layout → build → verify` sequence. Use `--resume` to reuse unchanged phases, but do not replay a timed-out native transaction until its original result is reconciled. Inspect the final Word-rendered PDF, including touched pages, before calling it visually accepted.

## Shared instances and recovery

One Word application/Zotero integration instance has one native writer. All agents must honor its shared resource coordination and OS lock. Output-directory isolation is not application-instance isolation. Read-only work on frozen files may proceed independently.

Keep an uncertain operation's controller, identity, pending state and original evidence intact. Reconcile the original result before retrying. A stale lock file, missing stdout or timeout is not permission to force-quit a process, clear ownership or relaunch a job. macOS is the current native target; Windows has not completed full production acceptance.

同一原生實例只允許一個寫入者；逾時後保留原身份及證據，先讀回，不清除未知狀態或強制關閉他方程式。Windows仍未完成完整產製驗收。

## Report and publish

Report what changed, exact source/output versions, what checks actually ran, and unresolved limitations. Distinguish structural PASS, native completion, field refresh, visual review, statistical source agreement and scientific judgment. The collaboration and thesis-review projects do not grant extra Word concurrency or automatically install this package.

Public contributions contain generic code/rules and synthetic fixtures. Exclude real manuscripts, datasets, credentials, personal paths, private adapters and session logs. Keep dated maintenance reports marked as historical. A documentation-only change does not establish a native defect as fixed.

交付時分開說明結構、原生完成、欄位、視覺與科研判斷；公開內容不包含真實文稿、資料、憑證、私人路徑或會話日誌。
