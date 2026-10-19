# 批注与修订的 OOXML 编辑规则

## 放置批注锚点

使用 `scripts/comment.py` 维护 `comments.xml`、`commentsExtended.xml`、`commentsIds.xml`、`commentsExtensible.xml`、关系和内容类型。不要只生成批注部件而遗漏正文锚点。

按脚本分配的 ID，在目标段落中添加以下标记；将其放在 `w:p` 的直接子节点中，禁止放入 `w:r`。

```xml
<w:commentRangeStart w:id="0"/>
<w:r><w:t>被批注的文字</w:t></w:r>
<w:commentRangeEnd w:id="0"/>
<w:r>
  <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
  <w:commentReference w:id="0"/>
</w:r>
```

添加回复时，使用 `--parent`，并按脚本输出将回复锚点嵌入父批注范围。不要复用不属于目标批注的 ID。

## 标记插入和删除

在 run 外包裹 `w:ins` 或 `w:del`；提供 `w:id`、`w:author`、`w:date`。在删除范围内使用 `w:delText`，不要继续使用 `w:t`。

```xml
<w:del w:id="1" w:author="Editor" w:date="2026-01-01T00:00:00Z">
  <w:r><w:delText>旧文字</w:delText></w:r>
</w:del>
<w:ins w:id="2" w:author="Editor" w:date="2026-01-01T00:00:00Z">
  <w:r><w:t>新文字</w:t></w:r>
</w:ins>
```

处理删除的段落标记时，将 `w:del` 放入 `w:pPr/w:rPr`，位于该 `rPr` 的其他子节点之前；按 XSD 顺序检查。理解其接受后的语义是与下一段合并；不要仅清空文字而声称段落已被删除。

使用 `scripts/office/validate.py edited.docx --original input.docx --author Editor` 检查遗漏的未追踪正文文字修改。将示例路径按 skill 根目录解析；不要声称该检查覆盖全部部件、格式变更或段落合并语义。

保留原作者的修订内容和元数据；对其他作者的插入做删除时，在其 `w:ins` 内嵌套自己的 `w:del`。恢复其删除时，在原 `w:del` 后添加自己的 `w:ins`；不要直接重写其他作者的修订历史。

禁止将提取器的 `accepted` 文本视图声称为保存后的净稿；完整修订接受和拒绝不在本 skill 的支持范围内。
