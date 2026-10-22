# 批注事务管理

先执行技能入口要求的自检。从文档工作目录调用；将 `<skill_dir>` 替换为主入口定义的真实目录。本入口不依赖旧 `comment.py` 的手工补锚点流程，每次只接受一个动作。

```console
python -I -S -B "<skill_dir>/scripts/manage_comments.py" ./input.docx ./action.json ./edited.docx
```

使用 `manage_comments(source, specification, destination=None)` 获取动作结果字典。变更必须输出同扩展名的新 DOCX/DOTX，禁止输入别名、链接输出或数字签名包；先在副本中规划，检查结构、线程元数据和 XML 值保真后原子打包。失败保留源及已有输出。

## 查询

将 `action.json` 设为：

```json
{"action": "list"}
```

查询不提供输出文档参数，结果写到标准输出：

```console
python -I -S -B "<skill_dir>/scripts/manage_comments.py" ./input.docx ./action.json
```

返回 `mode`（`none/classic/modern`）、`mutable`、`diagnostics` 及已读取批注的 `id/text/author/parent_id/paragraph_index`。`mutable` 仅说明支持的修改子集可用，不表示所有动作均可执行。审阅诊断；未知或损坏元数据不等于完整查询成功，无法安全解析的 XML 会直接失败。

## 按文字新增并自动锚定

```json
{
  "action": "add",
  "paragraph_index": 0,
  "expected_text": "需要核实费用上限",
  "anchor_text": "需要核实费用上限",
  "text": "请核对本项预算。",
  "author": "复核人",
  "initials": "FH"
}
```

`paragraph_index` 只计算正文直属 `w:p`，从 0 开始；`expected_text` 核对整段原文，`anchor_text` 在选定段落中必须唯一，且恰好覆盖一个或多个完整连续普通 run。不得为子串拆 run，不跨段，不自动统一空白或标点。段落索引与原文明确消除不同段落之间的定位歧义。

拒绝字段、修订、超链接、既有批注/书签/权限等活动范围、未知容器和不明确的边缘 XML 空白。首次新增会建立最小四部件元数据、正文范围及引用；不安装字体，也不声明阅读器中的批注显示已经验证。

## 修改、删除与回复

每个动作都核验整数 ID 及批注原文：

```json
{"action": "update", "id": 0, "expected_text": "请核对本项预算。", "text": "请补充预算依据。"}
```

```json
{"action": "reply", "id": 0, "expected_text": "请补充预算依据。", "text": "已补充材料。", "author": "经办人"}
```

```json
{"action": "delete", "id": 1, "expected_text": "已补充材料。"}
```

- 更新只支持单段、单个普通文字节点的批注，保留其 ID、作者和其他非目标内容。新文字非空且为单行 XML 文本。
- 删除同步移除定义、正文起点/终点/独立引用 run 及对应元数据；有回复的父批注必须先删除回复，不静默级联。保留空部件及关系，不删除未知依赖。
- 回复只支持完整一致的四部件现代根批注，最多一级；新范围嵌套在父范围内。不能回复回复，也不声称支持任意云线程、提及、人员、解决状态工作流。

## 兼容与限制

经典 `comments.xml` 仅支持严格单段查询、更新和删除，不自动升级以新增/回复。现代四部件检查 ID、paraId、durableId 唯一及一一对应、父存在、无环、自父和正文锚点配对；不完整或未知语义拒绝变更。

含复杂文字、多个段落、表格内或页眉页脚批注、未知扩展、额外人员/云元数据或非标准依赖时，查询会报告诊断，修改拒绝。未知非目标 ZIP 部件保持字节不变；这不是任意已有批注的无损编辑器。

完成后重新查询并运行 `<skill_dir>/scripts/office/validate.py`，核对目标批注和正文范围，核对非目标 XML 值及部件字节。基础校验器不代替本入口的四部件元数据检查；以上均不等同于 XSD、真实 Word/WPS 或视觉验收。
