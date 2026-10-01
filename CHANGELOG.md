## Unreleased — 2026-10-01 documentation

- Separate actual Word release from unaccepted NLM snapshots; link the existing-diagnostic reuse workflow without changing runtime admission or release rules.
- Record bounded HID admission failures separately from Word occupancy, TCC denial, process completion and unknown NLM remote outcomes; retain original release conditions and unverified concurrency status.
- Record the user-supplied Word-first/Zotero-second Automation retry, keeping its unverified recovery status distinct from historical TCC rebuild evidence.
- Document subsequent Refresh persistence with preserved strict metadata failures; distinguish profile selection and file-access prompts from Apple Events denial.
- Retain the unverified Word+NLM overlap boundary and original delivery/queue ownership. No runtime or installed skill changes.

## 0.1.0.dev3 — 2026-09-26

Integrate prior rendering/table regression coverage without reverting current Word locking, equation semantics, field-refresh or RUC profile 1.1.0. Document root/path and source-version boundaries.

# Changelog

## 0.1.0.dev2 — 2026-09-26

- 增加全文统稿指南、37条来源矩阵、电子附录发布策略和可运行合成正反例；校规、用户要求、工程建议及未决项分列。
- 增加可选只读 `thesis_integrity` 契约，检查逐节设置、域/书签、原生脚注、题注配对/顺序、主稿书目及英中排序；接入最终检查和缓存指纹。
- RUC profile 1.1.0修正三级标题为宋体，新增目录、文后标题及图/表题分离角色；保留罗马大小写和标题段距解释的来源边界。
- 整合双单位缩进、主题字体、可增长行高、表头和渲染区域检查；空规则/区域或未命中目标的渲染检查不能通过。
- 公式同时保留原XML身份与版本化语义投影，避免把Word仅格式序列化误判为数学内容变化。
- Word适配器/Runner/PDF服务共用持久锁并支持同线程嵌套；刷新失败保留暂存文档并在锁内记录pending，字段更新错误不能当成功。
- 原生PDF导出使用Word容器内副本，成功、输入不变和资源归还分别记录；不覆盖现役绑定或重审已冻结论文。
- 合成样例逐节设置脚注并提供同节跨页重编号检查；保存快照按自动目录书签的实际位置/文字比较，保留原始代码，拒绝同文错指和缓存变化。

## Unreleased — 2026-09-13

- Document macOS Zotero→Word TCC recovery: distinguish Full Disk Access from Automation, keep Terminal `-1743` sender-specific, and describe backed-up per-user TCC database recovery only as a last resort.

## 0.1.0.dev1 — 2026-09-10

- 首次将现有本地 Word skill、两场景、schema/profile、流水线及合成测试纳入独立 Git 仓库。
- 增加 Python 包、`wordrev` 入口、版本查询、MIT 许可及合成测试 CI。
- 可执行路径通过 PATH/`WORDREV_OFFICECLI` 查找；数值 QA 默认复用当前 Python；RTK 可选。
- 保留历史说明与本机路径；不包含私人论文、数据、研究适配器或运行产物。
- 开发预览：Word 原生 AutoFit 的已知 `-10006` 失败仍待定位；未进行新一轮 Word/Zotero 实测。
- 本机现用 skill 与 RStudio workbench 不被此次发布替换或修改。
# 2026-09-15

- Added the staged macOS Word PDF service and shared-lock recovery contract.
