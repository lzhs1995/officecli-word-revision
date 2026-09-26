# OfficeCLI Academic Word

Independent academic Word skill and executable pipeline. Skill name:
`officecli-word-revision`; CLI: `wordrev`.

**开发预览 `0.1.0.dev2`。** 本版加入全文统稿经验、规则来源矩阵、可选的整稿检查契约及
可生成的合成正反例。软件版本、作业 schema 2.0、RUC profile 1.1.0、整稿契约1.0分别管理。
自动检查只证明其声明的范围；原生分页、Zotero Refresh、视觉和内容审读分别验收。

## 功能与范围

- `manuscript_revision`：以原稿为格式母版，先审计，后建立接受内容的清洁候选，
  再生成原生修订、接受稿、拒绝核验稿和 QA。**需要项目适配器**；私人论文适配器不在本仓库。
- `thesis_format`：已有通用适配器，支持单章、章节批处理和合并后的整篇论文。
  规范配置与正文内容分离，不自动合并章节。
- 阶段哈希、缓存与恢复、统计结果清单核验、表格/字体/缩进/题注和字段检查。
- 可选 `thesis_integrity`：内容块、逐节页码/页眉页脚、三目录、脚注、题注配对、
  REF/SEQ 与 Zotero 域、唯一主稿书目及英中排序；契约绑定SHA，检查只读。
- Word 原生 AutoFit、Zotero 动态字段和 Word PDF 均需本机应用参与；
  不能把 OOXML 或模拟单元测试通过当作实际 Word 分页/Refresh 成功。
- 不执行统计估计、不调用 NotebookLM、不覆盖原稿。

## 已知限制：使用前必读

1. 当前应用自动化面向 **macOS + Microsoft Word**。Windows 完整生产流程未验证，
   Linux 只适合静态/合成单元测试，不提供原生 Word 后处理。
2. 最近一项本地任务在原生 AutoFit 前设置 `track revisions` 时报告 Word
   `-10006`。本次发布保留该已知问题，没有重新启动 Word 修复或宣称已解决。
3. 正文修订仍需明确的项目适配器、内容来源及样式锚点，不是任意两份 DOCX 的通用无配置比较器。
4. 历史说明中的本机路径和性能记录仅为上下文。新的经验与示例使用通用规则和合成输入，
   不发布真实论文附件或原始执行日志，历史时间也不能承诺任意文档的速度。
5. 内置 RUC 配置源于用户提供的规范快照，不代表学校最新官方规范或官方认证。
6. 只发布通用源码、规则与合成测试；论文、CFPS 数据、私有适配器、运行产物不在发布范围。
7. OfficeCLI、zotero-mcp 与 Computer Use 都不能创建 Word 活 Zotero 域。插入/刷新只走 Word+Zotero.dotm 宏；Refresh 不能从无到有创建书目。

## 安装与使用

先独立安装 Microsoft Word、OfficeCLI 和 Poppler（`pdfinfo`、`pdftoppm`、
`pdftotext`、`pdfimages`）。有引用任务时另需 Zotero 桌面端、Word 插件及所配置的引用工具。
应用安装、权限授予和插件配置不由本包静默执行。RTK 可选。

建议在新目录和独立虚拟环境试用；不要替换正在工作的生产 skill：

```bash
git clone https://github.com/lzhs1995/officecli-word-revision.git
cd officecli-word-revision
python3 -m venv .venv
.venv/bin/python -m pip install .
.venv/bin/wordrev --version
.venv/bin/wordrev --help
```

将 `skills/officecli-word-revision` 复制到客户端的 skill 目录即可启用技能。
现有目标必须先备份，不能覆盖正在执行的任务。CLI 与 skill 来自同一固定提交。
可在独立版本目录安装新版，现役任务继续读取原固定路径；在安全边界才为后继任务切换。
安装不重配 MCP，也不自动迁移既有任务。

```text
$officecli-word-revision 审计这份论文的格式；使用 manuscript_revision 场景。
先检查作业配置与实际 Word 能力，不覆盖原稿，不重新估计统计模型。
```

已具备有效作业配置时：

```bash
wordrev audit --job /absolute/path/job.json --resume
wordrev layout --job /absolute/path/job.json --resume
wordrev build --job /absolute/path/job.json --resume
wordrev verify --job /absolute/path/job.json --run-dir /absolute/path/run --resume
```

详细入口见 [SKILL.md](skills/officecli-word-revision/SKILL.md)，
以及 [期刊修订](skills/officecli-word-revision/references/manuscript-revision.md)、
[学位论文排版](skills/officecli-word-revision/references/thesis-format.md)。
格式规格不明确时先解决映射，不直接开始写入。

## 全文统稿指南与可运行示例

[全文统稿指南](skills/officecli-word-revision/references/full-thesis-formatting.md)覆盖人大要求、摘要与三目录、
原生脚注与书目、分节判断、图表排序与美观、保存重开及恢复踩坑。
[来源矩阵](skills/officecli-word-revision/references/ruc-rule-sources.json)区分学校存档、用户要求和工程建议；
[真实审议与用户裁决](skills/officecli-word-revision/references/format-review-decisions.md)记录5次NLM返回、71原引、33项裁决及新稿默认选择；
[电子附录策略](skills/officecli-word-revision/references/electronic-appendices.md)说明版本、许可和实际成员校验。

以下命令从本仓根执行，输出目录必须尚不存在：

```bash
python skills/officecli-word-revision/scripts/make_thesis_fixture.py --out-dir /absolute/run/valid
python skills/officecli-word-revision/scripts/thesis_integrity.py --docx /absolute/run/valid/synthetic-thesis.docx --contract /absolute/run/valid/contract.json --output /absolute/run/valid/check.json
python skills/officecli-word-revision/scripts/make_thesis_fixture.py --case appendix-restart --out-dir /absolute/run/bad-restart
```

反例还可选 `missing-reference`、`bibliography-order`、`clipped-image`。生成器退出0表示已生成并核实预期结果；
对反例单独运行检查器会退出2，报告具体失败。合成ADDIN只测试结构，不能用作原生Zotero创建/Refresh的证据。
接入现有job的方法、覆盖范围与退出码见[整稿契约](skills/officecli-word-revision/references/thesis-integrity.md)。

## 与 RStudio workbench 的关系

本仓库独立于 [clauder-rstudio-workbench](https://github.com/lzhs1995/clauder-rstudio-workbench)。
RStudio/compareGroups 负责计算和结果来源；Word 流水线只读取冻结的 CSV/图片和
`analysis_manifest.json`。不会把格式变更转化为统计重跑。
见 [衔接说明](docs/rstudio-integration.md)。未向 RStudio 仓库复制源码或修改安装器。

## 开发、测试与版本

```bash
python -m unittest discover -s skills/officecli-word-revision/scripts -p 'test_*.py' -v
python -m pip wheel . --no-deps --wheel-dir dist
```

CI 在 Ubuntu/macOS 运行合成单元测试与打包检查，**不提供 Microsoft Word 或 Zotero 原生验收**。
当前 CI 实际结果以 Actions 为准。手工验收检查表见 [发布与验收](docs/release.md)。
只在源码仓库修改、测试、提交；运行目录是安装副本。正式 Git 历史从本次导入开始，
不伪造早期提交记录。每次修改保留可审阅 diff，失败测试和应用限制不得隐藏。

## 许可与依赖

本仓库代码采用 MIT。OfficeCLI、Microsoft Word、Zotero、Python 库与字体是外部依赖，
遵循各自许可；本发布不包含它们的二进制、字体文件或学校模板原文件。
本项目不是 Microsoft、OfficeCLI、Zotero 或中国人民大学的官方产品。
