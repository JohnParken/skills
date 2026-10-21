# 受控结构编辑输入约定

先执行技能入口要求的自检，保留输入文件，从确认的文档工作目录执行：

```console
python -I -S -B docx/scripts/edit_structure.py ./input.docx ./operations.json ./edited.docx
python -I -S -B docx/scripts/office/validate.py ./edited.docx
```

顶层只接受非空 `operations` 列表。各操作按原始 XML 定位，全部验证通过后才应用。`paragraph_index` 从 0 开始，只计算 `w:body` 的直属 `w:p`，不计表格或容器内段落；`run_index` 从 0 开始，计算锚点段落的直属 `w:r`。同时提供完全一致的 `expected_text`，run 格式修改另提供 `expected_run_text`，避免索引漂移时改错内容。

## 插入段落

例如输入的第一个直属段落原文为“现有正文”：

```json
{
  "operations": [
    {
      "type": "insert_paragraph",
      "paragraph_index": 0,
      "expected_text": "现有正文",
      "position": "after",
      "text": "补充说明：“中文与 English 均可保留。”",
      "style": "BodyText"
    }
  ]
}
```

`position` 只接受 `before` 或 `after`；`text` 为非空单行 XML 文字。`style` 可省略；指定时必须为目标文档中真实存在且 `w:type="paragraph"` 的段落样式 ID，不接受字符、表格或编号样式。省略时使用文档默认样式，不复制原段落的编号、格式、书签或节属性，不把样式显示名称当成样式 ID。

## 修改普通 run 格式

例如第一个直属段落只有一个普通 run，原文为“现有正文”：

```json
{
  "operations": [
    {
      "type": "format_run",
      "paragraph_index": 0,
      "expected_text": "现有正文",
      "run_index": 0,
      "expected_run_text": "现有正文",
      "format": {
        "bold": true,
        "italic": false,
        "size_pt": 16,
        "color": "000000",
        "font_east_asia": "仿宋",
        "font_latin": "Times New Roman"
      }
    }
  ]
}
```

`format` 必须非空，仅支持上述字段；字号为 1–200pt 的半点增量，颜色为六位 RGB 十六进制。显式 `false` 会关闭相应属性；指定颜色或字体时清除对应主题覆盖，未指定的属性保留。不安装或嵌入字体，不承诺字体可用或视觉显示结果。

## 拒绝边界与验收

禁止操作空段落、表格内段落、书签、复杂 run、字段、批注/移动范围、修订/节属性或未知容器锚点。相同 run 或相同段落插入位置重复操作时停止，不猜测顺序；多个不同锚点的索引均以原文为准。

本入口不增删表格行、不替换图片、不改页眉页脚或整个段落样式；不得通过删除复杂结构让操作通过。使用独立输出且保持 DOCX/DOTX 扩展名不变；输入/输出不能经符号链接或硬链接指向同一文件。

实际输出前执行受限解析、已实现结构规则和 XML 节点/属性值保真核对，失败时保留原文件及已有输出。成功后对照操作核对文字和实际 XML 格式值，核对未修改 ZIP 部件内容字节；仍不等同于 XSD、国家标准、真实阅读器或视觉验收。
