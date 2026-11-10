# Security policy

## Reporting a vulnerability

Use GitHub's **Report a vulnerability** option in the repository's Security tab when private reporting is available: [Security advisories](https://github.com/lzhs1995/officecli-word-revision/security/advisories). This document does not assert that the repository setting is enabled.

If that option is unavailable, open a minimal issue requesting a private reporting route, without exploit details or sensitive attachments. Do not post credentials, unpublished manuscripts, real participant data, private file paths or full execution/session logs. Prefer a small synthetic DOCX or a redacted description that reproduces the problem.

Include the exact commit/package version, operating system and application versions, expected versus observed behavior, and the affected trust boundary. Preserve the original result locally; do not keep retrying an uncertain native Word operation just to reproduce it.

## Scope and updates

This repository is a development preview. No long-term support or security backport schedule is promised. Use reviewed, pinned revisions and consult release notes; do not replace an installation that owns an active native task.

Relevant reports include source/output isolation failures, unsafe application ownership or cleanup, unintended execution, citation-field corruption and disclosure of private content. A formatting defect can still be reported as an ordinary issue using synthetic material. Microsoft Word, Zotero, OfficeCLI and their plugins are separate dependencies with their own security channels.

The project keeps file/structure validation, native application completion and scholarly correctness separate. A successful static test does not establish that arbitrary documents or macros are safe to execute. Configure native applications and document trust according to your environment; application installation or permission changes are not silently performed by this package.

## 安全通報

若GitHub已啟用私人漏洞通報，請沿Security頁的「Report a vulnerability」提交；未啟用時，可先開不含敏感內容的issue詢問私人通報管道。不要公開憑證、未發表稿件、真實研究資料、私人路徑或完整會話日誌。請提供固定版本、環境、預期與實際結果，優先使用合成範例。

本專案為開發預覽，沒有承諾長期維護或回補時程。原生操作結果不明時保留原證據，先核回結果，不反覆執行或清除他方資源。
