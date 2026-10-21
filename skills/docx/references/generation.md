# 生成器输入约定

## 目录

- [顶层字段](#顶层字段)
- [默认格式与页面](#默认格式与页面)
- [样式字段与覆盖顺序](#样式字段与覆盖顺序)
- [内容块](#内容块)
- [表格配置与示例](#表格配置与示例)
- [显式覆盖示例](#显式覆盖示例)
- [中文 JSON 与引号](#中文-json-与引号)

先按 [执行环境约束](../SKILL.md#执行环境约束与首步自检) 定位脚本并完成自检。在宿主的下载工作目录调用生成器，输入输出使用 `./document.json`、`./output.docx` 等相对路径；不要切换到 skill 源码目录。脚本路径使用已核实的 skill 相对位置，禁止拼接掩码。相对图片路径以输入 JSON 所在目录为基准。需要 Python API 时调用 `create_document(spec, output, base_dir=None)`，同样使用相对文件参数。禁止传入未声明的字段，生成器会拒绝未知字段、错误类型和不支持的块。

## 顶层字段

- 使用必需的 `blocks` 数组按顺序描述内容。
- 使用可选 `page` 设置单节页面；默认 A4、上下 25.4mm、左右 31.8mm 页边距。横向页面交换宽高。
- 使用可选 `styles` 设置正文、四级标题及表格正文/表头的样式；未提供的字段保留对应默认值。
- 使用可选 `table_format` 设置所有表格的布局；允许单张表格使用 `format` 覆盖。
- 使用 `header`、`footer` 字符串设置普通页眉页脚；使用 `page_numbers: true` 插入 PAGE 字段。不要把字段缓存值当作已计算的页码。

## 默认格式与页面

省略 `page`、`styles` 和局部格式时，生成以下公文排版预设；不要将其等同于完整公文标准认证。

| 对象 | 中文字体 | 字号 | 对齐 | 加粗 | 首行缩进 | 行距 |
|---|---|---|---|---|---|---|
| `heading1` | 方正小标宋简体 | 18pt | 居中 | 否 | 0 字 | 1.15 倍 |
| `heading2` | 黑体 | 16pt | 左对齐 | 否 | 0 字 | 1.15 倍 |
| `heading3` | 楷体 | 16pt | 左对齐 | 是 | 0 字 | 1.15 倍 |
| `heading4` | 仿宋 | 16pt | 左对齐 | 否 | 0 字 | 1.15 倍 |
| `body` | 仿宋 | 16pt | 左对齐 | 否 | 2 字 | 固定 28pt |

使用 Times New Roman 作为默认西文字体；使用零段前、段后间距。通过 `BodyText` 段落样式应用正文格式，不要插入空格模拟首行缩进。让标题、表格、列表、图片、分页及页眉页脚使用各自样式或基础样式，避免继承正文的缩进和固定行距。基础样式使用仿宋 16pt、Times New Roman、1.15 倍行距；保留列表自身的编号缩进。表格正文和表头分别使用独立的 `TableBody`、`TableHeader` 段落样式。

仅记录字体名称；不要安装、下载或嵌入字体，不要声称已验证阅读端字体可用。只为新建文档应用这些默认值；受控编辑保持已有文档样式。

| `page` 字段 | 单位 | 默认值与规则 |
|---|---|---|
| `width_mm`、`height_mm` | mm | 210、297；保留单节页面。 |
| `margin_top_mm`、`margin_bottom_mm` | mm | 25.4、25.4。 |
| `margin_left_mm`、`margin_right_mm` | mm | 31.8、31.8。 |
| `margin_mm` | mm | 可选，统一覆盖四边；仍允许用单边字段再次覆盖。 |

按“单边字段 > `margin_mm` > 默认单边值”计算边距。例如 `{"margin_mm": 20, "margin_top_mm": 25.4}` 表示顶部 25.4mm，其余三边 20mm。不要将厘米值误写为毫米。所有提供的字段均须有效，即使某字段随后被更高优先级覆盖；边距不得为负，换算后须保留正的可用宽高。

按扣除实际左右边距后的宽度限制表格和图片；默认可用宽度约为 146.4mm。省略 `widths_mm` 时均分可用宽度，显式列宽之和不得超过可用宽度；图片还须满足扣除上下边距后的可用高度。不要通过超出页面可用尺寸的值模拟浮动对象。

## 样式字段与覆盖顺序

仅在 `styles` 中使用 `body`、`heading1`、`heading2`、`heading3`、`heading4`、`table_body`、`table_header`；不要传入任意 Word 样式 ID。每一项均为部分字段对象。标题及普通段落可使用 `format` 对象覆盖下列字段；表格文字的额外字段和覆盖方式见下文：

| 字段 | 类型及范围 | 用途 |
|---|---|---|
| `font_east_asia` | 非空单行字符串，最多 64 字符 | 中文字体名称。 |
| `font_latin` | 非空单行字符串，最多 64 字符 | 西文字体名称。 |
| `size_pt` | 数值，1–200 | 字号，单位 pt。 |
| `bold` | 布尔值 | `false` 显式取消继承的加粗。 |
| `align` | `left`、`center`、`right`、`justify` | 段落对齐。 |
| `first_line_chars` | 数值，0–20 | 首行缩进字符数；0 取消首行缩进。 |
| `line_spacing_pt` | 数值，1–400 | 固定行距，单位 pt；不接受倍数或自动行距字符串。 |
| `space_before_pt`、`space_after_pt` | 数值，0–400 | 段前、段后间距，单位 pt。 |

按字段依次覆盖：内置默认值 → 对应 `styles` 项 → 当前块 `format` → 块级 `align`（仅对齐）→ run 字符字段（仅文字格式）。省略字段表示继承；禁止使用 `null` 表示默认值。不要用字符串、布尔值代替数值，不要使用 NaN、Infinity、负值或未知字段。

使用 run 的 `font_east_asia`、`font_latin` 分别控制中西文字体。保留 run 的旧字段 `font`：它同时指定中西文字体；同一 run 中的专用字体字段再覆盖相应字体。例如 `{"text": "中文 ABC", "font": "Arial", "font_east_asia": "仿宋"}` 使用仿宋显示中文、Arial 显示西文。不要将 run 的 `font` 放入 `styles` 或 `format`。

允许 run 的 `bold`、`italic`、`underline` 显式设为 `false`；不要通过省略字段尝试取消已继承的格式。仅覆盖已声明字段，不要假定修改字号会自动调整行距或首行字符缩进。

## 内容块

| type | 字段 | 约束 |
|---|---|---|
| heading | `text`, `level`, `format`, `align` | 明确提供 level，范围为 1–4；format 和 align 可选。 |
| paragraph | `text` 或 `runs`, `format`, `align` | 不要同时给 text/runs；align 为 left、center、right、justify。 |
| list | `items`, `ordered` | items 为字符串数组；仅单层项目符号或编号。 |
| table | `rows`, `widths_mm`, `header`, `format`, `body_format`, `header_format` | rows 为规则二维字符串数组；可指定列宽；header 默认 false；其他格式字段均可省略。 |
| image | `path`, `width_mm`, `height_mm`, `alt` | 仅嵌入 PNG/JPEG；明确给定显示宽高。 |
| page_break | 无其他字段 | 插入手动分页。 |

在 `runs` 中使用 `text`、`bold`、`italic`、`underline`、`font`、`font_east_asia`、`font_latin`、`size_pt`、`color`；颜色使用不带井号的六位十六进制 RGB。让 XML 库处理文本转义，禁止预先将 `&` 转成 `&amp;`。

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
    {"type": "table", "header": true, "widths_mm": [70, 70],
      "rows": [["项目", "状态"], ["验证", "通过"]]}
  ]
}
```

## 表格配置与示例

参照 GB/T 9704—2012 第 5.2.2 条的一般字体字号原则使用仿宋 16pt；将边框、留白、加粗、对齐、行距与跨页控制标为项目约定，不要声称这些具体参数均由国标规定。保留当前页面边距；不要将本预设宣称为整份文档的国标符合性检查。国标第 8 章的横排表格页码和单双页表头朝向需要另外的分节与分页能力，当前不提供。

为 `styles.table_body` 和 `styles.table_header` 分别设置表格文字样式；不要让它们继承 `styles.body` 或相互继承。默认均为仿宋 16pt、西文 Times New Roman、黑色、首行缩进 0、段前段后 0、1.15 倍行距；正文不加粗且左对齐，表头加粗且居中。

在表格文字样式中使用上文的字体、字号、加粗、对齐、缩进、段距及固定行距字段，并允许以下额外字段：

| 字段 | 类型及范围 | 用法 |
|---|---|---|
| `color` | 六位十六进制 RGB，不带 `#` | 默认 `"000000"`。 |
| `line_spacing_multiple` | 数值，0.5–10 | 默认 1.15；使用倍数行距。 |

在同一对象中仅提供 `line_spacing_pt` 或 `line_spacing_multiple` 之一；不要同时传入。后续层显式指定一种行距时替换前一层的行距模式，例如从固定 24pt 改为 1.5 倍。上述额外字段仅适用于表格文字样式；普通段落、标题的 `format` 不接受它们。

按“内置表格文字默认值 → `styles.table_body` / `styles.table_header` → 当前表格的 `body_format` / `header_format`”逐字段覆盖。正文和表头分别计算，不要将 `body_format` 当作整张表格的通用文字覆盖；需要统一字体时分别设置两者。保持单元格为字符串，不要在 `rows` 中传入 run、段落或单元格对象。

在顶层 `table_format` 和表格块的 `format` 中使用同一组布局字段；按“内置布局默认值 → 顶层 `table_format` → 单表 `format`”覆盖：

| 字段 | 类型及范围 | 默认值 |
|---|---|---|
| `align` | `left`、`center`、`right`；整张表的位置 | `left` |
| `vertical_align` | `top`、`center`、`bottom`；单元格内容位置 | `center` |
| `border_style` | `single` 或 `none`；作用于内外框 | `single` |
| `border_width_pt` | 数值，0.25–12；单位 pt | 0.5 |
| `border_color` | 六位十六进制 RGB，不带 `#` | `"000000"` |
| `cell_fill` | 六位十六进制 RGB 或 `"none"`；所有单元格底色 | `"none"` |
| `cell_margin_top_mm`、`cell_margin_bottom_mm` | 数值，0–50；单位 mm | 各 1 |
| `cell_margin_left_mm`、`cell_margin_right_mm` | 数值，0–50；单位 mm | 各 1.5 |
| `repeat_header` | 布尔值；仅在 `header: true` 时生效 | `true` |
| `allow_row_split` | 布尔值；是否允许单行跨页拆分 | `false` |

- 使用 `body_format.align` / `header_format.align` 控制文字水平对齐，使用 `format.align` 控制整张表的位置；不要混淆两者。
- 省略 `widths_mm` 时均分页面可用宽度；指定列宽时保留原有规则。列宽包含单元格留白，换算到文档单位后左右留白之和须小于每列宽度。需要窄列时显式减小留白，不要依赖阅读器自动挤压表格。
- 保持自动行高；不要传入 `row_height` 等未声明字段，不要以固定行高压缩内容。重复表头和避免拆行是分页请求，不保证超高行保持完整，也不验证实际页数。
- 仅在 `header: true` 时将第一行作为表头；关闭 `repeat_header` 仅关闭重复，不取消表头格式。省略 `header` 时，不要自动识别第一行。
- 使用 `border_style: "none"` 关闭边框，使用 `cell_fill: "none"` 清除全局底色；不要使用 `null` 表示关闭。所有提供的字段均须有效，即使之后被覆盖或当前不生效；拒绝未知字段、错误类型和非有限数值。

保留上文原有中英文示例的内容。需要覆盖表格样式时，使用下面的完整中英混排示例；其中的 14pt 是显式覆盖，默认仍为 16pt：

```json
{
  "styles": {
    "table_body": {"size_pt": 14, "line_spacing_multiple": 1.2},
    "table_header": {"size_pt": 14, "bold": true}
  },
  "table_format": {"border_width_pt": 0.75, "cell_fill": "F2F2F2"},
  "blocks": [
    {"type": "heading", "level": 1, "text": "项目清单 / Project list"},
    {"type": "table", "header": true, "widths_mm": [70, 70],
      "format": {"cell_fill": "none", "cell_margin_top_mm": 0, "repeat_header": false},
      "body_format": {"line_spacing_pt": 22},
      "header_format": {"font_east_asia": "黑体", "bold": false},
      "rows": [["项目 / Item", "状态 / Status"], ["材料复核", "已完成 / Done"]]}
  ]
}
```

核对该例：表格文字 14pt；正文使用固定 22pt 行距，表头仍为 1.15 倍行距；表头改为黑体、不加粗，继续居中；边框为 0.75pt、无底色、顶部留白为 0；表头不跨页重复。普通标题和正文仍沿用原有公文预设，不要把单表覆盖保存为下一份文档的默认值。

## 显式覆盖示例

保留默认字体，分别覆盖正文、一级标题、单个段落和局部文字。将下面的完整 JSON 保存为独立输入；不要与其他完整输入直接拼接：

```json
{
  "page": {"margin_mm": 20, "margin_top_mm": 25.4},
  "styles": {
    "body": {"size_pt": 14, "line_spacing_pt": 24},
    "heading1": {"size_pt": 20, "align": "left"}
  },
  "blocks": [
    {"type": "heading", "level": 1, "text": "项目报告 / Project report",
      "format": {"size_pt": 18}, "align": "center"},
    {"type": "heading", "level": 2, "text": "一、工作进展"},
    {"type": "heading", "level": 3, "text": "（一）阶段成果"},
    {"type": "heading", "level": 4, "text": "1. 复核情况"},
    {"type": "paragraph", "text": "负责人说：“已完成复核。” Status: reviewed."},
    {"type": "paragraph", "format": {"first_line_chars": 0, "bold": true},
      "runs": [
        {"text": "说明："},
        {"text": "此句取消加粗，西文使用 Arial。", "bold": false, "font_latin": "Arial"}
      ]}
  ]
}
```

核对生成结果：顶部边距 25.4mm、其余三边 20mm；一级标题按块覆盖为 18pt 居中；普通正文为 14pt、固定 24pt；最后一段取消首行缩进，第二个 run 取消加粗并仅覆盖西文字体。二至四级标题继续使用各自默认值。不要把一次调用的覆盖保存成下一份文档的新默认值。

## 中文 JSON 与引号

使用 UTF-8 编码保存输入。保留英文半角结构符号 `"`、`:`、`,`、`{}`、`[]`；中文文字、中文标点和中文引号可以作为字符串内容。不要翻译字段名和 `type` 等枚举值，不要加入注释或尾随逗号。

使用下面的合法段落块；正文中的中文引号无需转义，正文中的英文双引号使用 `\"` 转义：

```json
{"type": "paragraph", "text": "负责人说：“请在周五前完成复核。”"}
```

```json
{"type": "paragraph", "text": "负责人说：\"请在周五前完成复核。\""}
```

将这些段落块放入完整输入的 `blocks` 数组，不要直接把单个段落块作为生成器的顶层输入。

不要使用下列写法；它们是错误示例，不是可运行 JSON：

```text
{“type”: “paragraph”, “text”: “中文内容”}
{"type": "paragraph", "text": "负责人说："请完成复核。""}
```

第一例使用了中文语法引号；第二例没有转义正文中的英文双引号。不要把正文中合法的中文引号全部替换掉。出现 `JSONDecodeError` 时，根据报错的行号、列号检查该位置及其前面的引号、冒号和逗号；仅凭“有中文引号”无法确定原始输入的具体错误。

需要避免手工转义时，在下载工作目录执行以下 Python 示例，使用字典构造输入后序列化；不要手工拼接 JSON 字符串：

```python
import json
from pathlib import Path

spec = {"blocks": [
    {"type": "heading", "level": 1, "text": "项目进展报告"},
    {"type": "paragraph", "text": '负责人说：“请完成复核。” 文档状态为 "已确认"。'}
]}
work = Path(".")
(work / "document.json").write_text(
    json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8"
)
```

使用 `ensure_ascii=False` 使中文在 JSON 文件中直接可读；它不负责修复已有非法 JSON，英文双引号等必要字符仍会由序列化器正确转义。生成器使用标准 JSON 解析器，不会自动修复错误引号。

将输入 JSON 和文档输出放在任务工作目录。不要在 skill 中保存用户数据、生成产物或临时图片；输出失败时保留已有目标文件。
