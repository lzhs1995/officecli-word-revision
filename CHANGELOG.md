# Changelog

## Unreleased — 2026-09-13

- Document macOS Zotero→Word TCC recovery: distinguish Full Disk Access from Automation, keep Terminal `-1743` sender-specific, and describe backed-up per-user TCC database recovery only as a last resort.

## 0.1.0.dev1 — 2026-09-10

- 首次将现有本地 Word skill、两场景、schema/profile、流水线及合成测试纳入独立 Git 仓库。
- 增加 Python 包、`wordrev` 入口、版本查询、MIT 许可及合成测试 CI。
- 可执行路径通过 PATH/`WORDREV_OFFICECLI` 查找；数值 QA 默认复用当前 Python；RTK 可选。
- 保留历史说明与本机路径；不包含私人论文、数据、研究适配器或运行产物。
- 开发预览：Word 原生 AutoFit 的已知 `-10006` 失败仍待定位；未进行新一轮 Word/Zotero 实测。
- 本机现用 skill 与 RStudio workbench 不被此次发布替换或修改。
