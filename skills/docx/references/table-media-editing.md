# 表格行与图片编辑

先执行技能入口的环境自检；使用原文件及同扩展名的新输出，失败时保留二者。输入必须通过已实现结构检查，数字签名包拒绝编辑，不删除签名来绕过限制。

## 规则表格数据行

从文档工作目录执行，按实际技能相对位置调整 `docx/`：

```console
python -I -S -B docx/scripts/edit_tables.py ./input.docx ./row.json ./edited.docx
```

每次只接受一个 `operation`。索引从 0 开始，只计算正文 `w:body` 的直属表格；行索引按目标表格的原始直属行计算。

```json
{
  "operation": {
    "type": "insert_row",
    "table_index": 0,
    "row_index": 1,
    "expected_rows": 3,
    "expected_text": ["甲", "完成"],
    "position": "after",
    "cells": ["新项目", "已验收"]
  }
}
```

- `expected_rows` 为原始总行数；`expected_text` 必须匹配锚点行所有单元格的原始文字。
- `insert_row` 在锚点数据行 `before/after` 插入，克隆其行、单元格、段落及 run 格式，再完整替换每列文字；允许显式空字符串。
- 删除使用 `delete_row`，不提供 `position/cells`。禁止增删前置连续的重复表头标记行；禁止删除最后剩余数据行，但可删除有其他数据行时的末尾行。
- 仅识别实际 `tblHeader` 标记，不从文字或 `TableHeader` 样式猜测未标记的业务表头；如业务首行没有此标记，调用方仍须明确选择数据行。
- 只接受有一个显式 `tblGrid` 的规则矩形表，每个单元格恰有一个普通段落/run/text；所有行都须在此范围内。拒绝任何合并、嵌套、复杂内容、未知/引用属性，以及活动字段、批注、书签或权限范围。不会修复或展开合并网格。

## 同格式图片替换

```console
python -I -S -B docx/scripts/replace_image.py ./input.docx ./new.png ./selection.json ./edited.docx
```

`selection.json` 必须提供：

| 字段 | 要求 |
|---|---|
| `media_part` | 明确指定 `word/media/` 下的 PNG/JPG/JPEG 成员名。 |
| `expected_sha256` | 原媒体内容字节的 64 位十六进制 SHA-256，不是整个 DOCX 的哈希。 |
| `expected_relationships` | 指向该部件的内部 image 关系条数，正整数。 |
| `expected_references` | 使用这些关系 ID 的实际 DrawingML picture 引用次数，正整数。 |
| `all_references` | 可选布尔值；任一计数大于 1 时必须显式为 `true`。 |

可用标准库从原文档获取媒体哈希，不要猜测：

```python
import hashlib
import zipfile

with zipfile.ZipFile("./input.docx") as package:
    digest = hashlib.sha256(package.read("word/media/image1.png")).hexdigest()
```

关系条数不等于绘图次数；一个关系可以被多处引用。替换共享媒体会影响全部已核验引用，不承诺只改页面上的某一张图。不支持拆分引用、新增关系或格式转换。

核对原图、新图扩展、内容类型、魔数及 PNG IEND/JPEG EOI；不解码像素或声称验证图片显示。拒绝孤立/非 image 关系、未知扩展对关系 ID 的使用、VML、未知绘图容器及混合外链/嵌入引用。

不改任何 XML、关系、尺寸、裁剪或布局，所有非目标 ZIP 成员内容字节保持不变。图片原始宽高比变化可能影响显示效果，须由调用方准备符合需求的图片；不要声称完成视觉验收。

## 验收

```console
python -I -S -B docx/scripts/office/validate.py ./edited.docx
python -I -S -B docx/scripts/read_docx.py ./edited.docx --format json -o ./edited-structure.json
```

对照操作检查表格原行/表头顺序及新增文字，核对未修改部件字节。图片替换应仅目标 media 内容变化，所有其他成员逐字节相同。以上不替代 XSD、真实阅读器、图像解码或视觉验收。
