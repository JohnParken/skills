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

不要自动处理删除的段落标记、移动修订或格式修订；保留已有结构并报告超出支持范围。禁止仅清空文字而声称段落已被删除。

使用文末修订对比命令检查遗漏的未追踪正文文字修改。以读取 `SKILL.md` 的真实位置定位脚本，将文件放在宿主的下载工作目录，使用相对文件参数；不要声称该检查覆盖全部部件、格式变更或段落合并语义。

保留原作者的修订内容和元数据；对其他作者的插入做删除时，在其 `w:ins` 内嵌套自己的 `w:del`。恢复其删除时，在原 `w:del` 后添加自己的 `w:ins`；不要直接重写其他作者的修订历史。

禁止将提取器的 `accepted` 文本视图声称为保存后的净稿；完整修订接受和拒绝不在本 skill 的支持范围内。

## 使用低层辅助工具

先按 [执行环境约束](../SKILL.md#执行环境约束与首步自检) 完成定位及自检，再从下载工作目录调用脚本。下文仅在已核实 skill 相对位置为 `docx` 时适用；修改前保留源文件，完成后执行基础检查，禁止为照抄示例切换目录：

```console
python -I -S -B docx/scripts/comment.py ./input.docx "核对费用上限" -o ./annotated.docx
python -I -S -B docx/scripts/merge_runs.py ./input.docx -o ./merged.docx
python -I -S -B docx/scripts/office/validate.py ./edited.docx --original ./input.docx --author Editor
```

添加批注定义后，先补正文锚点再验收。当前基础检查仅支持有起点、终点和引用的范围批注，不支持无范围的点批注，也不验证完整线程回复语义；不要把生成批注部件声称为已验证阅读器中的显示效果。

仅在明确理解原结构时合并相邻且格式相同的 run；该脚本会改写 `word/document.xml`、移除校对标记和 run 的 rsid，不处理其他部件。不要为了让受控替换通过而盲目合并复杂结构。

保留命名空间声明、元素顺序、关系引用和 `xml:space`；禁止重新排版整个 XML 或删除未知扩展。使用 `scripts/office/xmlio.py` 中的受限解析入口，不要直接用默认解析器处理外来 XML。

修订对比仅比较正文文字及有限的插入/删除标记，不检查所有格式、段落合并、移动、页眉页脚、字段、制表符或换行语义。指定 `--original` 不会豁免结构错误，校验器不提供自动修复或 XSD 检查。
