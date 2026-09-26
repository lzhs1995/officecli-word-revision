# 电子附录与版本化交付

本仓库只提供通用策略和合成示例。实际论文、数据附件、账号绑定和私有执行日志不随经验工程发布。

## 纸面材料与电子材料

先列正文理解和学校提交必需的说明、表格和证明，确保它们出现在可提交文件中。补充机器表、代码、环境和可公开材料可以放电子附录；外部链接能否替代某份提交附件须有学校要求或用户决定，不能自行假定。

主稿与独立附录分别记录Word/PDF、页数、索引、版本和验收范围。正文引用“附录B表B-3（电子附录v1.0，文件及定位……）”，避免只写“见GitHub”。独立卷用卷内页码，不把它和主稿物理页码混为一谈。

## 推荐仓库结构

```text
README.md                   材料范围、读者入口、运行次序、受限数据获取途径
CITATION.cff                固定版本引用信息；有真实DOI才填写DOI
LICENSE                     代码许可
docs/materials-license.md   文本、图形和数据各自的许可/范围
environment/                依赖及版本，必要时精确锁文件
src/                        通用脚本
examples/synthetic/         明确标注的合成输入和可再生结果
manifest.json               本版本文件集合、大小、SHA256、逻辑角色
CHANGELOG.md                变更及对论文定位/结论的影响
```

大二进制可放对应Release附件；不能假定GitHub自动生成的源码zip包含外置Release资产或LFS真实大文件。论文引用固定tag/commit和具体资产，并保存下载后的SHA256。tag和普通Release可被有权限者修改；若采用GitHub不可变Release功能，应核实际启用状态。长期引用可另存正式归档或真实DOI，不编造标识符。

一个通用成员清单可以如下；哈希必须由实际文件计算，示例字符不是可用校验值：

```json
{
  "version": "v1.0.0",
  "commit": "发布时的真实提交SHA",
  "scope": "synthetic examples only",
  "members": [
    {"path": "examples/synthetic/table.csv", "bytes": 123,
     "sha256": "实际64位SHA256", "role": "synthetic-table"}
  ]
}
```

清单放包内时不要要求它记录自己的哈希；将清单哈希及ZIP哈希放外部冻结回执，或者明确清单不在其自身成员集合中。封包后重新打开ZIP，核路径唯一、预期集合、无额外/缺失成员、每件解压大小、CRC和SHA256。数目相同不证明成员相同，目录里存在文件也不证明已打入ZIP。

最终冻结后增加外部独立核收回执，通常不应为了加入回执而不断重封同一实物包。文件确有变化则新版本、新清单、新SHA，并按影响范围重核；旧版本和失败证据保留。

## 复现说明要准确

区分“文件找到”“数值相容”“唯一实证链已追溯”“全部模型重新估计”。代码能运行不等于历史模型已完整重估；表格存在不等于收敛和标准误通过。公开样例要显式标注合成，避免被当作真实研究结果。

GitHub官方参考：[About releases](https://docs.github.com/en/repositories/releasing-projects-on-github/about-releases)、[Immutable releases](https://docs.github.com/en/repositories/releasing-projects-on-github/about-immutable-releases)。访问时间及本任务保存的来源哈希见规则矩阵；引用外部功能并不说明本仓库已经启用它。
