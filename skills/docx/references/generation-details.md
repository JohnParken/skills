# 列表与表格细节

本文件补充 [生成器输入约定](generation.md)。先执行技能入口的自检；全部能力只生成 ZIP/XML，不计算阅读器最终编号、字体、图像像素或页面布局。

## 跨说明段续号与层级

```json
{
  "blocks": [
    {"type": "list", "list_id": "工作清单", "level": 0, "ordered": true, "items": ["第一项"]},
    {"type": "list", "list_id": "工作清单", "level": 1, "ordered": true, "items": ["子项甲"]},
    {"type": "paragraph", "text": "说明段不重置列表实例。"},
    {"type": "list", "list_id": "工作清单", "level": 1, "ordered": true, "items": ["子项乙"]}
  ]
}
```

- 同一 `list_id` 复用同一个 `numId`，`level` 为整数 0–8；说明段后继续使用相同 ID/层级，不新建编号实例。实际显示值由兼容阅读器计算，文本提取器不伪造编号。
- 命名实例首块必须为 0 级且 `items` 非空；后续不能跳过父层级，`ordered` 必须一致。ID 为非空单行字符串，最多 64 字符，无边缘空白，可用中文。
- 使用新的 ID 建立独立实例以重启；不提供任意起始编号、自定义编号样式或复杂重启命令。子级按父层级重启定义生成；不要承诺所有阅读器的具体编号结果。
- 不提供 `list_id` 时保留旧的独立单层行为；此时 `level` 只能为 0。不跨不同 ID 推断父子关系。

## 多行连续表头与每边边框

```json
{
  "table_format": {"borders": {"top": {"color": "123456", "width_pt": 1}}},
  "blocks": [
    {"type": "table", "header_rows": 2,
     "format": {"borders": {"top": {"width_pt": 1.5}, "insideH": {"style": "none"}}},
     "rows": [["报表", "状态"], ["项目", "结果"], ["复核", "完成"]]}
  ]
}
```

`header_rows` 为整数 0–总行数，仅表示连续前置行；应用表头样式，且在 `repeat_header: true` 时对这些行逐一写入重复表头请求。关闭重复仍保留表头样式，不保证实际分页。旧 `header` 对应 0/1 行；同时声明时两者必须一致，否则拒绝。

在顶层 `table_format.borders` 或表格 `format.borders` 中使用 `top/left/bottom/right/insideH/insideV`，每边可含：

| 字段 | 要求 |
|---|---|
| `style` | `single` 或 `none`。 |
| `width_pt` | 0.25–12pt 的有限数值。 |
| `color` | 不带 `#` 的六位 RGB 十六进制。 |

先计算原有标量 `border_style/border_width_pt/border_color`，再覆盖各边；顶层和单表各边按字段递进合并。例如上例 top 使用 1.5pt、颜色 `123456`，其他未指定字段继续继承；insideH 无线。

## 单元格独立边框与共享边

在单元格对象上声明 `borders`，仅接受 `top/left/bottom/right`，每边的 `style/width_pt/color` 合法值与上表相同。例如：

```json
{"type": "table", "rows": [[
  {"text": "重点", "borders": {"top": {"width_pt": 1.5}, "right": {"color": "AA0000"}}},
  "普通单元格"
]]}
```

先解析完整的表格边框，再按字段覆盖单元格边：外侧继承对应 top/left/bottom/right，内部上下边继承 insideH、左右边继承 insideV。仅指定颜色时继续继承线型和宽度；要从隐藏边恢复实线，显式指定 `style: "single"`。空 `borders: {}` 不添加单元格边框，单边空对象如 `top: {}` 则显式使用该处继承值。

共享边属于两个相邻单元格：单方指定时，将相同的最终边框同步写入另一方对应边；因此上例右边同时也是第二格的左边。双方指定时须产生相同的最终 XML 边框，否则拒绝并保留已有输出，不依赖阅读器冲突优先级。宽度按 1/8pt 四舍五入后比较，`style: "none"` 写为 `nil`，此时宽度、颜色不参与实际边框比较。

对象仍须至少有 `text/image` 之一；空格可用 `text: ""`，不接受仅 `borders` 的对象。拒绝 insideH/insideV、对角线、start/end、未知字段、无效数值或不规则行。支持表头、图片单元格和普通规则表；不扩展到合并单元格、嵌套表格或已有复杂表格编辑。不使用此字段时保持原生成输出。

## 单元格文字与 inline 图片

字符串单元格继续可用；也可将单个单元格写成如下字段片段，先准备真实存在的 `./photo.png`，再放入规则 `rows` 数组：

```json
{"text": "说明文字", "image": {"path": "./photo.png", "width_mm": 20, "height_mm": 15, "alt": "现成图片"}}
```

对象至少含 `text/image` 之一；可仅文字或仅图片，不接受未知字段、任意块或多个图片数组。图片位于同一普通单元格段落的文字之后，使用单个 inline DrawingML，不环绕、不浮动、不转码。

图片仍须为现成 PNG/JPEG，明确给定尺寸和 `alt`，相对路径以 JSON 所在目录为准。宽度不得超过列宽扣除左右留白后的可用宽度，高度不得超过页面可用高度扣除上下单元格留白；不计算文字与图片组合后的实际换行和整行高度。

正文图片与单元格图片分别分配唯一媒体名、关系和 drawing ID，不覆盖输入图片；不声明已解码或显示正确。生成后核对关系、内容类型、尺寸及 ZIP 完整性。

带图片的单元格超出 `edit_tables.py` 纯文本行增删的支持范围；该入口会明确拒绝，不通过克隆图片 run 或删除图片让编辑通过。媒体替换仍按其独立哈希和共享引用契约执行。
