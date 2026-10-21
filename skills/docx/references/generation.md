# 生成器输入约定

从 skill 根目录调用 `python -I -S -B scripts/create_docx.py ../docx-work/document.json ../docx-work/output.docx`，将 `../docx-work/` 替换为实际的独立工作目录。相对图片路径以输入 JSON 所在目录为基准。需要 Python API 时调用 `create_document(spec, output, base_dir=None)`。禁止传入未声明的字段，生成器会拒绝未知字段、错误类型和不支持的块。

## 顶层字段

- 使用必需的 `blocks` 数组按顺序描述内容。
- 使用 `page.width_mm`、`page.height_mm`、`page.margin_mm` 设置单节页面；默认 A4、25 mm 页边距。横向页面交换宽高。
- 使用 `header`、`footer` 字符串设置普通页眉页脚；使用 `page_numbers: true` 插入 PAGE 字段。不要把字段缓存值当作已计算的页码。
- 使用默认 Arial/宋体或 run 的 `font` 字体名称；不要声称运行环境提供了该字体。

## 内容块

| type | 字段 | 约束 |
|---|---|---|
| heading | `text`, `level` | level 为 1–3。 |
| paragraph | `text` 或 `runs`, `align` | 不要同时给 text/runs；align 为 left、center、right、justify。 |
| list | `items`, `ordered` | items 为字符串数组；仅单层项目符号或编号。 |
| table | `rows`, `widths_mm`, `header` | rows 为规则二维字符串数组；可指定各列宽度，header 控制首行表头。 |
| image | `path`, `width_mm`, `height_mm`, `alt` | 仅嵌入 PNG/JPEG；明确给定显示宽高。 |
| page_break | 无其他字段 | 插入手动分页。 |

在 `runs` 中使用 `text`、`bold`、`italic`、`underline`、`font`、`size_pt`、`color`；颜色使用不带井号的六位十六进制 RGB。让 XML 库处理文本转义，禁止预先将 `&` 转成 `&amp;`。

将文本字段限制为单行，禁止插入控制字符；多段文字拆成多个 paragraph 块。提供 image 的 `alt` 描述（允许空字符串），不要依赖图片尺寸自动推断。

按以下正反例构造输入；不要把自然语言需求直接当作生成器字段：

| 需求 | 使用 | 不要使用 |
|---|---|---|
| 加粗文字 | `{"text": "结论", "bold": true}` | `{"text": "结论", "bold": "true"}` |
| 两段内容 | 两个 paragraph 块 | 在单个 text 中嵌入换行 |
| 填写公司名称 | 单个 run 中完整的 `{{ company }}` | 将占位符拆分到多个 run |

```json
{
  "header": "季度进展",
  "footer": "内部资料",
  "page_numbers": true,
  "blocks": [
    {"type": "heading", "level": 1, "text": "项目报告"},
    {"type": "paragraph", "runs": [
      {"text": "结论：", "bold": true},
      {"text": "按期完成", "color": "008000", "size_pt": 12}
    ]},
    {"type": "list", "ordered": false, "items": ["完成复核", "安排交付"]},
    {"type": "table", "header": true, "widths_mm": [80, 80],
      "rows": [["项目", "状态"], ["验证", "通过"]]}
  ]
}
```

将输入 JSON 和文档输出放在任务工作目录。不要在 skill 中保存用户数据、生成产物或临时图片；输出失败时保留已有目标文件。
