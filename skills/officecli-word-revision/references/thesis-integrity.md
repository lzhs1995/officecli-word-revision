# 显式整稿检查契约

`scripts/thesis_integrity.py` 只读 DOCX，覆盖内容块顺序/边界、节页码/链接、跨 story 域与书签、原生脚注、对象—题注绑定、全篇活书目实际输出顺序和指定图表的结构风险。不会自动移动内容、猜题注、修改元数据或触发 Word/NLM。

## 运行合成正例

从本 skill 根执行，选一个尚不存在的输出目录：

```bash
python3 scripts/make_thesis_fixture.py --out-dir /absolute/run/synthetic-example
python3 scripts/thesis_integrity.py \
  --docx /absolute/run/synthetic-example/synthetic-thesis.docx \
  --contract /absolute/run/synthetic-example/contract.json \
  --output /absolute/run/synthetic-example/independent-report.json
```

生成器给出八个内容块、三节、三目录、图表各一、同节不同页的两个脚注及英文/中文各一条合成书目。Zotero形状的域仅用于解析测试，不是原生 Zotero 插入或 Refresh 的证据。正式应用前由作者建立真实文稿契约，不能把例子的条数和样式当成学校要求。最小样例不启用奇偶页不同，也不验证人大双面页码外侧规则。

这是结构解析正例，不是人大完整排版模板；契约也不是自动排版指令。实际稿件仍由明确的profile/适配器落实所需属性，再检查其保存结果。

反例见 `scripts/test_thesis_integrity.py`：错误重启页码、链接未断开、手工目录、空REF、页脚引用失效、相邻/嵌套/简单域、脚注覆盖、错误对象、缺SEQ、图片固定行距、双单位缩进、表格固定行高、实际书目顺序和拼音同首字母等。测试不启动应用或网络。

也可直接生成反例文件：`python3 scripts/make_thesis_fixture.py --case appendix-restart --out-dir /absolute/run/bad-restart`。
另有 `missing-reference`、`bibliography-order`、`clipped-image`。生成器退出0表示已生成并核实所选案例的预期结果；
每个目录的 `expectation.json` 保留预期错误。对反例单独运行检查器退出2；生成器成功不等于反例通过检查。

## 接入已有 pipeline

对一个 `scenario=thesis_format` job 增加可选根字段：

```json
"thesis_integrity": {
  "path": "/absolute/run/contract.json",
  "sha256": "填入该契约文件实际SHA256"
}
```

契约为 [thesis-integrity.schema.json](thesis-integrity.schema.json) 的1.0版。至少声明一个显式内容块；未配置的领域逐项列在 `omitted_contract_areas`，不能用空契约产生“全稿已验”的错觉。书签锚点只解析正文，文本锚点必须精确且唯一；同一标题出现在目录缓存中时，使用专用书签可避免歧义。

`blocks` 按期望物理顺序列出 `{id, anchor, start}`。`start` 为 `any/new_page/new_section`；`new_page` 兼容正确的分页与分节，不强制后者。`sections` 用1基节号显式声明页码、方向、节起点、首页不同及六类页眉页脚的链接。`page_number.format=roman` 接受大小写罗马制；只有有明确依据时使用 `lowerRoman/upperRoman`。`start=continue` 拒绝隐含在副本里的 `w:start` 重启值。

`fields` 可声明三目录数量/标签、活引文下限、唯一书目及域不得锁定。无论数量约束是否提供，都解析全部已读 Word story 的真实域，检查字段边界、REF目标和Zotero域内JSON。PAGEREF允许零长度定位书签。多次引用同一文献允许，单个引文域内重复URI拒绝。

按自定义题注样式生成图表目录时，在 `fields.toc` 提供 `figure_styles` 和 `table_styles` 字符串数组，
例如 `["Figure Caption"]` 对应 `TOC \t "Figure Caption,4"`。样式选择与 `\c` 标签若冲突或同时跨两类，则报错；
不根据目录当前可见标题猜它是哪一种域。

`footnotes` 声明位置、重编号、最小数量及上标要求。默认位置与省略属性按有效值处理；每节覆盖不能漏核。脚注编号制默认 continuous，若要求每页重编号就应在有效属性中得到 eachPage。

样例为每节显式设置页底和eachPage，并保留schema顺序；原生保存后仍须重新检查。只检查第一个脚注的编号不能证明重编号制度。正式文档不应为获得同样的示例页数而插入多余分页。

`captions` 每条绑定具体 caption/object 锚点、类别、标签、章号和序号，检查实际显示、SEQ、引用及配对方向。启用 `keep_pair` 时也检查对象和题注之间没有正文/其他对象或断开的空段链。选择 `caption_sequence_complete=true` 后，契约声明的每章每类序号必须从1连续；只检查局部对象时不要启用。检查器未独立证明契约枚举了全部图表，也不判断首次引用的学术顺序。

`bibliography` 绑定标题、语言组顺序及逐条准确可见文本。每条有显示作者、完整排序键、排序键来源、年份、题名及题名排序键。工具按这些键排序后与实际活书目显示比较；不根据正文汉字猜语言，不改作者姓名。排序键来源是待作者核实的声明，机器不能证明多音姓读音或文献元数据真实。

可选 `bibliography.entries_separate_paragraphs=true` 另外要求每条书目为独立实际段落。一个 `w:t` 中的换行字符或同段手动换行不能替代段落。省略时仅检查文字/顺序，不声称覆盖条目段落布局；最终悬挂缩进与断行仍须检查原生PDF。

`objects` 只检查声明的图表规则：嵌入图、固定行距风险、节版心尺寸，以及显式选择表格的缩进、对齐、行高、表头重复与加粗。条件表格样式另列渲染复核提示；该检测不宣称完整实现 Word 的条件样式/主题级联，也不替代图内字号和跨页可读性审阅。

## 返回值与缓存

`status=FAIL`、退出2：契约中的机械要求失败。`PASS` 或 `REVIEW`、退出0：机械要求通过；REVIEW保留实际审美/样式待核项。`all_pass` 只指机械项，不表示论文最终接受。报告包含输入SHA、规范化契约SHA、检查范围、遗漏领域及 `not_checked`；不覆盖已有输出。

pipeline 启用此契约后，将报告放入 `thesis_format_qa.json` 的 `thesis_integrity`，失败进入 adapter 的最终失败列表。文件原始SHA绑定在job里，初始化和调用前均核实；同路径换内容会拒绝，更新绑定后重新计算缓存。脚本、schema、原生服务及依赖也纳入实现指纹，不能复用旧检查器的缓存PASS。

`rendered_document_qa` 可另外绑定实际 PDF 字形区域和页眉页脚几何契约。它只检查明确选择的页面/区域；空页规则或没有匹配到目标文字不能算字体检查成功。全文原生分页、目录页码、视觉、文献支持及NLM审议由各自真实证据承担。

`word_field_snapshot.snapshot(saved_docx, contract, observed_pages)` 与 `compare(first, second)` 可供原生工作器比较两次已保存结果：保留原始域代码/缓存，单独解析自动目录书签的实际显示位置和文字。同文不同位置、缓存页码变化、内部链接错指或目标丢失均不能通过。只消除已证明的自动名称差异，常规书签名不归一化；此函数不启动Word，也不能证明调用者给的页数来自真实执行。
