# 生成器输入约定

## 目录

- [预检与集成验收](#预检与集成验收)
- [主题与文档标题](#主题与文档标题)
- [顶层字段](#顶层字段)
- [默认格式与页面](#默认格式与页面)
- [样式字段与覆盖顺序](#样式字段与覆盖顺序)
- [内容块](#内容块)
- [表格配置与示例](#表格配置与示例)
- [显式覆盖示例](#显式覆盖示例)
- [中文 JSON 与引号](#中文-json-与引号)

先按 [执行环境约束](../SKILL.md#执行环境约束与首步自检) 定位脚本并完成自检。在宿主的下载工作目录调用生成器，输入输出使用 `./document.json`、`./output.docx` 等相对路径；不要切换到 skill 源码目录。脚本和内部资源均基于入口定义的 `<skill_dir>` 定位，执行前替换为真实目录。相对图片路径以输入 JSON 所在目录为基准。需要 Python API 时调用 `create_document(spec, output, base_dir=None, verify=True, require_no_placeholders=False)`，输入输出及 `base_dir` 同样使用任务目录内的相对路径；API 模块从 `<skill_dir>/scripts/` 加载。禁止传入未声明的字段，生成器会拒绝未知字段、错误类型和不支持的块。

## 预检与集成验收

先按主入口执行环境自检。以下命令中的 `<skill_dir>` 替换为已确认的真实目录；输入输出位于任务工作目录：

```console
python -I -S -B "<skill_dir>/scripts/create_docx.py" ./document.json --preflight
python -I -S -B "<skill_dir>/scripts/create_docx.py" ./document.json ./output.docx --verify --require-no-placeholders
```

`--preflight` 只接受输入路径，检查字段、图片引用、生成 XML 和共享预算，不写文档或执行写入探针；不是结构或视觉验收。以上两条为可选路径：普通生成直接用第二条，无需先重复预检。`--verify` 在临时 ZIP 写完后读取全部成员并检查 CRC、结构及正文文字序列、表格/图片数量、页眉页脚文字，再按输入核对排版意图：块顺序与分页符、标题样式与大纲、列表实例/层级/编号定义、样式及显式文字格式、页面尺寸/边距、表格行列/网格宽度/单元格内容归属/边框/留白/跨页标记、图片显示尺寸/替代文字及引用媒体的 SHA-256。通过后原子发布输出，失败保留已有输出。`--require-no-placeholders` 隐含开启验收，检查跨 run 的双花括号标记；正文需要字面花括号时省略该参数。验收遇到未验证警告同样拒绝发布。成功摘要增加 `layout_intent: true`；这是生成子集的输入与 XML 对照，不计算页码，不做 XSD、字体可用性或视觉验证。

API 默认 `verify=False` 保持兼容；按本 Skill 流程生成交付时显式设为 `True`。所有生成调用在写出前检查共享 XML 预算，避免输出自身校验器拒绝的超限文件。调用 `preflight_document(spec, base_dir=None)` 获取预检摘要，返回块数、部件数和未压缩大小。使用 Python 字典可直接调用 API，无需先写 JSON。CLI 拒绝重复 JSON 键，未知字段在可确定时提示相近字段名；非法字符诊断包含字段路径、从 1 起的字符位置和 Unicode 码位，不静默清洗内容。

生成失败时执行 [生成失败修正与有限重试](../SKILL.md#生成失败修正与有限重试)：由助手按诊断修正可确定的输入或调用错误，首次之外最多重试 2 次，每次重新完整验收；脚本本身仍为单次执行。安全、资源、环境或能力边界受阻时停止，不绕过校验，也不将旧输出交付为本次结果。

## 主题与文档标题

使用 `theme` 选择 `official`（默认）、`business` 或 `academic`；显式样式覆盖优先于主题。主题仅为项目排版预设，不代表公文或学术标准认证。编辑已有文档不应用主题。

| 主题 | 正文与西文字体 | 正文排版 | 标题 |
|---|---|---|---|
| official | 仿宋 / Times New Roman | 保持 16pt、固定 28pt、首行 2 字 | 保持原有四级标题 |
| business | 微软雅黑 / Calibri | 10.5pt、1.25 倍、段后 4pt、无首行缩进 | 黑体、深蓝、章节左对齐 |
| academic | 宋体 / Times New Roman | 10.5pt、1.25 倍、段后 4pt、首行 2 字 | 黑体、黑色、章节左对齐 |

主题同时设置表格正文和表头的基础字体、字号、行距；后者保留加粗及居中。字体只记录名称，不安装或嵌入。`title` 和 `subtitle` 块支持 `text`、`format`、`align`，不接受 `level`，不写大纲级别；使用 `styles.title`、`styles.subtitle` 覆盖，默认分别为 22pt 和 14pt。章节继续使用原有 heading 1–4，不改变旧输入的语义。

商务和学术主题的页眉页脚使用独立的 `Header`、`Footer` 段落样式，采用对应主题基础中西文字体、10.5pt、1.25 倍行距、黑色、不加粗、左对齐、无首行缩进、段前段后 0pt。正文的 `styles.body` 覆盖不会改变这些样式；它们是内置 Word 样式，不是新增 JSON 样式字段。纯页码页脚及页码字段结果同样继承 `Footer`。公文主题保留原有基础样式：仿宋 / Times New Roman、16pt、1.15 倍行距、段前段后 0pt。省略页眉页脚且未启用页码时不生成相应部件；显式空字符串保留空部件。字体是否可用及页码实际值仍须在阅读器中核对。

```json
{"theme":"business","blocks":[
  {"type":"title","text":"项目报告"},
  {"type":"subtitle","text":"季度进展"},
  {"type":"heading","level":1,"text":"一、概况"},
  {"type":"paragraph","text":"正文内容。"}
]}
```

商务和学术列表使用内置 `ListParagraph` 样式，基础中西文字体与主题一致，10.5pt、1.25 倍行距、段前 0pt、段后 4pt、黑色、不加粗、左对齐；不写首行缩进属性，保留编号定义的悬挂缩进。正文覆盖不影响列表。公文列表仍使用原有基础样式，`ordered`、`list_id` 与 `level` 的语义不变；不增加列表样式输入字段。

## 顶层字段

- 使用可选 `theme` 选择命名主题；使用可选布尔值 `deduplicate_images` 复用相同图片字节，默认 `false`。
- 每次生成或预检独立缓存规范化路径的图片字节及尺寸头；同一路径读取一次，不跨调用保存缓存。缓存与包内去重独立，`false` 仍生成独立媒体，各引用的尺寸、替代文字和绘图 ID 独立。累计缓存和待写媒体仍受 128 MiB 预算约束，单图仍最多 32 MiB，最终包预算还包括 XML。
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

仅在 `styles` 中使用 `title`、`subtitle`、`body`、`heading1`、`heading2`、`heading3`、`heading4`、`table_body`、`table_header`；不要传入任意 Word 样式 ID。每一项均为部分字段对象。标题及普通段落可使用 `format` 对象覆盖下列字段；表格文字的默认值和覆盖方式见下文：

| 字段 | 类型及范围 | 用途 |
|---|---|---|
| `font_east_asia` | 非空单行字符串，最多 64 字符 | 中文字体名称。 |
| `font_latin` | 非空单行字符串，最多 64 字符 | 西文字体名称。 |
| `size_pt` | 数值，1–200 | 字号，单位 pt。 |
| `bold` | 布尔值 | `false` 显式取消继承的加粗。 |
| `align` | `left`、`center`、`right`、`justify` | 段落对齐。 |
| `first_line_chars` | 数值，0–20 | 首行缩进字符数；0 取消首行缩进。 |
| `line_spacing_pt` | 数值，1–400 | 固定行距，单位 pt；不接受倍数或自动行距字符串。 |
| `line_spacing_multiple` | 数值，0.5–10 | 正文、标题及表格均支持倍数行距；与同层固定行距互斥。 |
| `color` | 六位十六进制 RGB，不带 `#` | 正文、标题及表格均支持文字颜色。 |
| `space_before_pt`、`space_after_pt` | 数值，0–400 | 段前、段后间距，单位 pt。 |

按字段依次覆盖：内置主题默认值 → 对应 `styles` 项 → 当前块 `format` → 块级 `align`（仅对齐）→ run 字符字段（仅文字格式）。省略字段表示继承；禁止使用 `null` 表示默认值。不要用字符串、布尔值代替数值，不要使用 NaN、Infinity、负值或未知字段。

普通段落与标题通过 `BodyText`、`Heading1` 至 `Heading4` 继承对应样式；仅将块和 run 显式指定的格式写为直接属性。后续修改 Word 样式可影响未局部覆盖的属性；显式覆盖即使与样式当前值相同也保留。

使用 run 的 `font_east_asia`、`font_latin` 分别控制中西文字体。保留 run 的旧字段 `font`：它同时指定中西文字体；同一 run 中的专用字体字段再覆盖相应字体。例如 `{"text": "中文 ABC", "font": "Arial", "font_east_asia": "仿宋"}` 使用仿宋显示中文、Arial 显示西文。不要将 run 的 `font` 放入 `styles` 或 `format`。

允许 run 的 `bold`、`italic`、`underline` 显式设为 `false`；不要通过省略字段尝试取消已继承的格式。仅覆盖已声明字段，不要假定修改字号会自动调整行距或首行字符缩进。

## 内容块

| type | 字段 | 约束 |
|---|---|---|
| title / subtitle | `text`, `format`, `align` | 文档标题与副标题，不接受 level，不占大纲层级。 |
| heading | `text`, `level`, `format`, `align` | 明确提供 level，范围为 1–4；format 和 align 可选。 |
| paragraph | `text` 或 `runs`, `format`, `align` | 不要同时给 text/runs；align 为 left、center、right、justify。 |
| list | `items`, `ordered`, `list_id`, `level` | items 为字符串数组；旧单层用法不变，命名实例可按明确层级续号。 |
| table | `rows`, `widths_mm`, `header`, `header_rows`, `format`, `body_format`, `header_format` | rows 为规则二维数组，单元格为字符串或受限 text/image/borders 对象；可指定列宽及前置表头行数。 |
| image | `path`, `width_mm`, `height_mm`, `alt` | 仅嵌入 PNG/JPEG；至少给定宽或高，缺少的尺寸由尺寸头等比例推算。 |
| page_break | 无其他字段 | 插入手动分页。 |

在 `runs` 中使用 `text`、`bold`、`italic`、`underline`、`font`、`font_east_asia`、`font_latin`、`size_pt`、`color`；颜色使用不带井号的六位十六进制 RGB。让 XML 库处理文本转义，禁止预先将 `&` 转成 `&amp;`。

将文本字段限制为单行，禁止插入控制字符；多段文字拆成多个 paragraph 块。提供 image 的 `alt` 描述（允许空字符串），只给宽或高时读取 PNG IHDR（含头部 CRC）或 JPEG SOF 尺寸头推算比例；两项都给时保持显式尺寸。尺寸头检查不等于像素解码，不按 EXIF 方向旋转。

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

在顶层 `table_format` 或单表 `format` 使用以下便捷字段；同层显式字段优先于预设默认值：

| 字段 | 规则 |
|---|---|
| `preset` | `grid`、`three_line`、`zebra`；省略保持原有框线表。显式选择预设重设六边的样式与粗细；同层全局边框字段再覆盖预设，逐边字段优先。 |
| `column_sizing` | `equal`（默认）或 `content`；显式 `widths_mm` 始终优先。内容估算按中西文字符权重平方根分配，保留每列留白及至少 8mm 内容宽；无法满足则报错。不是字体测量或阅读器 AutoFit。 |
| `band_fill` | 六位 RGB 或 `none`；每隔一个数据行设置底色，表头不计入数据行。单元格条纹覆盖表格统一底色；显式 `none` 可关闭。 |

`three_line` 默认顶底线 1.5pt、无竖线和内部横线；已声明表头且存在数据行时，在最后表头行之后加 0.75pt 分隔线。显式单元格共享边优先，原有冲突检测继续生效。`zebra` 默认无框线，第二个及后续偶数数据行底色为 `F2F4F7`。需要表头样式必须显式声明 `header` 或 `header_rows`，不猜测第一行。


参照 GB/T 9704—2012 第 5.2.2 条的一般字体字号原则使用仿宋 16pt；将边框、留白、加粗、对齐、行距与跨页控制标为项目约定，不要声称这些具体参数均由国标规定。保留当前页面边距；不要将本预设宣称为整份文档的国标符合性检查。国标第 8 章的横排表格页码和单双页表头朝向需要另外的分节与分页能力，当前不提供。

为 `styles.table_body` 和 `styles.table_header` 分别设置表格文字样式；不要让它们继承 `styles.body` 或相互继承。默认均为仿宋 16pt、西文 Times New Roman、黑色、首行缩进 0、段前段后 0、1.15 倍行距；正文不加粗且左对齐，表头加粗且居中。

表格文字与普通正文共用上述格式字段；以下列出颜色及倍数行距的表格默认值：

| 字段 | 类型及范围 | 用法 |
|---|---|---|
| `color` | 六位十六进制 RGB，不带 `#` | 默认 `"000000"`。 |
| `line_spacing_multiple` | 数值，0.5–10 | 默认 1.15；使用倍数行距。 |

在同一对象中仅提供 `line_spacing_pt` 或 `line_spacing_multiple` 之一；不要同时传入。后续层显式指定一种行距时替换前一层的行距模式，例如从固定 24pt 改为 1.5 倍。普通段落、标题和表格文字样式均接受颜色及倍数行距。

按“内置表格文字默认值 → `styles.table_body` / `styles.table_header` → 当前表格的 `body_format` / `header_format`”逐字段覆盖。正文和表头分别计算，不要将 `body_format` 当作整张表格的通用文字覆盖；需要统一字体时分别设置两者。单元格默认用字符串，或按 [列表与表格细节](generation-details.md) 使用严格的 text/image/borders 对象；不接受任意 run、段落或合并布局。

将公共文字格式保存在 `TableBody` / `TableHeader`，各单元格仅引用样式；仅写出单表显式 `body_format` / `header_format` 的直接属性。显式覆盖即使与全局样式相同也保留，`false` 及固定/倍数行距切换仍有效。后续在 Word 中修改公共样式可影响未局部覆盖的属性；不要仅检查单元格直接属性来判断最终格式。

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
| `repeat_header` | 布尔值；对声明的所有前置表头行生效 | `true` |
| `allow_row_split` | 布尔值；是否允许单行跨页拆分 | `false` |

- 使用 `body_format.align` / `header_format.align` 控制文字水平对齐，使用 `format.align` 控制整张表的位置；不要混淆两者。
- 省略 `widths_mm` 时均分页面可用宽度；指定列宽时保留原有规则。列宽包含单元格留白，换算到文档单位后左右留白之和须小于每列宽度。需要窄列时显式减小留白，不要依赖阅读器自动挤压表格。
- 保持自动行高；不要传入 `row_height` 等未声明字段，不要以固定行高压缩内容。重复表头和避免拆行是分页请求，不保证超高行保持完整，也不验证实际页数。
- `header: true` 将第一行作为表头，`header_rows` 可明确指定连续前置表头数量；关闭 `repeat_header` 仅关闭重复，不取消表头格式。未声明表头时，不自动识别第一行。
- 使用 `border_style: "none"` 关闭边框，使用 `cell_fill: "none"` 清除全局底色；不要使用 `null` 表示关闭。所有提供的字段均须有效，即使之后被覆盖或当前不生效；拒绝未知字段、错误类型和非有限数值。

使用 `format.borders` 逐边设置表格边框，使用单元格对象的 `borders` 覆盖其 top/left/bottom/right；按 [列表与表格细节](generation-details.md) 核对继承、共享边冲突、层级续号、`header_rows` 和单元格图片规则。

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
