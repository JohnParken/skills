"""Create a small, standards-based DOCX package from a JSON block specification."""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
import tempfile
import unicodedata
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
XML = "http://www.w3.org/XML/1998/namespace"

ET.register_namespace("w", W)
ET.register_namespace("r", R)
ET.register_namespace("wp", WP)
ET.register_namespace("a", A)
ET.register_namespace("pic", PIC)
ET.register_namespace("", CT)

_CONTENT_TYPES = {
    "document": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "styles": "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
    "numbering": "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    "header": "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    "footer": "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
}
_REL_TYPES = {
    "officeDocument": f"{R}/officeDocument",
    "styles": f"{R}/styles",
    "numbering": f"{R}/numbering",
    "header": f"{R}/header",
    "footer": f"{R}/footer",
    "image": f"{R}/image",
}
_MAX_PAGE_MM = 558.8  # 22 inches, the Word page-size limit.
_MAX_MARGIN_MM = 250.0
_MAX_IMAGE_BYTES = 32 * 1024 * 1024
_MAX_XML_BYTES = 16 * 1024 * 1024
_MAX_PACKAGE_BYTES = 128 * 1024 * 1024
_MAX_PACKAGE_PARTS = 4096
_COLOR_RE = re.compile(r"^[0-9A-Fa-f]{6}$")


def _q(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def _object(value: Any, label: str, allowed: set[str], required: set[str] = frozenset()) -> dict:
    if not isinstance(value, dict):
        raise ValueError(f"{label} must be an object")
    unknown = set(value) - allowed
    if unknown:
        raise ValueError(f"{label} has unknown field(s): {', '.join(map(str, sorted(unknown, key=str)))}")
    missing = required - set(value)
    if missing:
        raise ValueError(f"{label} is missing field(s): {', '.join(sorted(missing))}")
    return value


def _number(value: Any, label: str, low: float, high: float, *, allow_zero: bool = False) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"{label} must be a number")
    result = float(value)
    if not math.isfinite(result) or result > high or result < low or (not allow_zero and result == 0):
        raise ValueError(f"{label} must be between {low:g} and {high:g}")
    return result


def _text(value: Any, label: str, *, allow_empty: bool = True) -> str:
    if not isinstance(value, str):
        raise ValueError(f"{label} must be a string")
    if not allow_empty and not value:
        raise ValueError(f"{label} must not be empty")
    for char in value:
        point = ord(char)
        if unicodedata.category(char) == "Cc" or not (
            point in (0x9, 0xA, 0xD)
            or 0x20 <= point <= 0xD7FF
            or 0xE000 <= point <= 0xFFFD
            or 0x10000 <= point <= 0x10FFFF
        ):
            raise ValueError(f"{label} contains an unsupported control or XML character")
    return value


def _twips(mm: float) -> int:
    return int(mm * 1440.0 / 25.4 + 0.5)


def _emu(mm: float) -> int:
    return int(mm * 36000.0 + 0.5)


def _image(path_value: Any, width: float, height: float, base_dir: Path, label: str) -> dict:
    rel_path = _text(path_value, f"{label}.path", allow_empty=False)
    path = Path(rel_path)
    if not path.is_absolute():
        path = base_dir / path
    path = path.resolve()
    if path.stat().st_size > _MAX_IMAGE_BYTES:
        raise ValueError(f"{label}.path exceeds the 32 MiB image limit")
    with path.open("rb") as image_file:
        data = image_file.read(_MAX_IMAGE_BYTES + 1)
    if len(data) > _MAX_IMAGE_BYTES:
        raise ValueError(f"{label}.path exceeds the 32 MiB image limit")
    if data.startswith(b"\x89PNG\r\n\x1a\n") and data.endswith(b"\x00\x00\x00\x00IEND\xaeB\x60\x82"):
        extension = "png"
    elif data.startswith(b"\xff\xd8") and b"\xff\xd9" in data[2:]:
        extension = "jpg"
    else:
        raise ValueError(f"{label}.path must reference an existing PNG or JPEG image")
    return {
        "type": "image", "image_path": path, "image_data": data,
        "extension": extension, "cx": _emu(width), "cy": _emu(height),
        "alt": _text(label, "image alt"),
    }


def _validate_run(value: Any, label: str) -> dict:
    run = _object(value, label,
                  {"text", "bold", "italic", "underline", "font", "size_pt", "color"},
                  {"text"})
    result = {"text": _text(run["text"], f"{label}.text")}
    for key in ("bold", "italic", "underline"):
        if key in run:
            if not isinstance(run[key], bool):
                raise ValueError(f"{label}.{key} must be a boolean")
            result[key] = run[key]
    if "font" in run:
        result["font"] = _text(run["font"], f"{label}.font", allow_empty=False)
        if len(result["font"]) > 64:
            raise ValueError(f"{label}.font must be at most 64 characters")
    if "size_pt" in run:
        result["size_pt"] = _number(run["size_pt"], f"{label}.size_pt", 1, 200)
    if "color" in run:
        color = _text(run["color"], f"{label}.color", allow_empty=False)
        if not _COLOR_RE.fullmatch(color):
            raise ValueError(f"{label}.color must be six hexadecimal RGB characters without #")
        result["color"] = color.upper()
    return result


def _validate_spec(spec: Any, base_dir: Path) -> dict:
    source = _object(spec, "spec", {"page", "header", "footer", "page_numbers", "blocks"}, {"blocks"})
    page = _object(source.get("page", {}), "page", {"width_mm", "height_mm", "margin_mm"})
    width = _number(page.get("width_mm", 210), "page.width_mm", 1, _MAX_PAGE_MM)
    height = _number(page.get("height_mm", 297), "page.height_mm", 1, _MAX_PAGE_MM)
    margin = _number(page.get("margin_mm", 25), "page.margin_mm", 0, _MAX_MARGIN_MM, allow_zero=True)
    if margin * 2 >= min(width, height):
        raise ValueError("page.margin_mm must leave positive width and height for content")
    page_twips = (_twips(width), _twips(height), _twips(margin))
    available_width = page_twips[0] - 2 * page_twips[2]
    available_height = page_twips[1] - 2 * page_twips[2]

    header = _text(source["header"], "header") if "header" in source else None
    footer = _text(source["footer"], "footer") if "footer" in source else None
    page_numbers = source.get("page_numbers", False)
    if not isinstance(page_numbers, bool):
        raise ValueError("page_numbers must be a boolean")

    blocks_value = source["blocks"]
    if not isinstance(blocks_value, list):
        raise ValueError("blocks must be an array")
    blocks: list[dict] = []
    image_paths: list[Path] = []
    for index, value in enumerate(blocks_value):
        label = f"blocks[{index}]"
        if not isinstance(value, dict):
            raise ValueError(f"{label} must be an object")
        kind = value.get("type")
        if not isinstance(kind, str):
            raise ValueError(f"{label}.type must be a string")
        if kind == "heading":
            block = _object(value, label, {"type", "text", "level"}, {"type", "text", "level"})
            level = block["level"]
            if isinstance(level, bool) or not isinstance(level, int) or level not in (1, 2, 3):
                raise ValueError(f"{label}.level must be 1, 2, or 3")
            blocks.append({"type": kind, "text": _text(block["text"], f"{label}.text"), "level": level})
        elif kind == "paragraph":
            block = _object(value, label, {"type", "text", "runs", "align"}, {"type"})
            if ("text" in block) == ("runs" in block):
                raise ValueError(f"{label} must contain exactly one of text or runs")
            align = block.get("align", "left")
            if align not in ("left", "center", "right", "justify"):
                raise ValueError(f"{label}.align must be left, center, right, or justify")
            if "text" in block:
                content: Any = _text(block["text"], f"{label}.text")
            else:
                if not isinstance(block["runs"], list):
                    raise ValueError(f"{label}.runs must be an array")
                content = [_validate_run(run, f"{label}.runs[{i}]") for i, run in enumerate(block["runs"])]
            blocks.append({"type": kind, "content": content, "align": align})
        elif kind == "list":
            block = _object(value, label, {"type", "items", "ordered"}, {"type", "items", "ordered"})
            if not isinstance(block["ordered"], bool):
                raise ValueError(f"{label}.ordered must be a boolean")
            if not isinstance(block["items"], list):
                raise ValueError(f"{label}.items must be an array")
            items = [_text(item, f"{label}.items[{i}]") for i, item in enumerate(block["items"])]
            blocks.append({"type": kind, "items": items, "ordered": block["ordered"]})
        elif kind == "table":
            block = _object(value, label, {"type", "rows", "widths_mm", "header"}, {"type", "rows"})
            if not isinstance(block["rows"], list) or not block["rows"]:
                raise ValueError(f"{label}.rows must be a non-empty array")
            rows = []
            columns = None
            for row_index, row in enumerate(block["rows"]):
                if not isinstance(row, list) or not row:
                    raise ValueError(f"{label}.rows[{row_index}] must be a non-empty array")
                if columns is None:
                    columns = len(row)
                elif len(row) != columns:
                    raise ValueError(f"{label}.rows must be rectangular")
                rows.append([_text(cell, f"{label}.rows[{row_index}][{col}]") for col, cell in enumerate(row)])
            assert columns is not None
            if "header" in block and not isinstance(block["header"], bool):
                raise ValueError(f"{label}.header must be a boolean")
            if "widths_mm" in block:
                widths = block["widths_mm"]
                if not isinstance(widths, list) or len(widths) != columns:
                    raise ValueError(f"{label}.widths_mm must contain one width per column")
                widths_twips = [_twips(_number(v, f"{label}.widths_mm[{i}]", 0.1, _MAX_PAGE_MM))
                                for i, v in enumerate(widths)]
                if sum(widths_twips) > available_width:
                    raise ValueError(f"{label}.widths_mm exceeds the page content width")
            else:
                quotient, remainder = divmod(available_width, columns)
                widths_twips = [quotient + (1 if i < remainder else 0) for i in range(columns)]
            blocks.append({"type": kind, "rows": rows, "widths_twips": widths_twips,
                           "header": block.get("header", False)})
        elif kind == "image":
            block = _object(value, label, {"type", "path", "width_mm", "height_mm", "alt"},
                            {"type", "path", "width_mm", "height_mm", "alt"})
            image_width = _number(block["width_mm"], f"{label}.width_mm", 0.1, _MAX_PAGE_MM)
            image_height = _number(block["height_mm"], f"{label}.height_mm", 0.1, _MAX_PAGE_MM)
            if _twips(image_width) > available_width or _twips(image_height) > available_height:
                raise ValueError(f"{label} dimensions exceed the page content area")
            image = _image(block["path"], image_width, image_height, base_dir, block["alt"])
            image_paths.append(image["image_path"])
            blocks.append(image)
        elif kind == "page_break":
            _object(value, label, {"type"}, {"type"})
            blocks.append({"type": kind})
        else:
            raise ValueError(f"{label}.type has unsupported value {kind!r}")
    return {
        "page_twips": page_twips, "blocks": blocks,
        "header": header, "footer": footer,
        "page_numbers": page_numbers, "image_paths": image_paths,
    }


def _add_text_run(parent: ET.Element, text: str, formatting: dict | None = None) -> None:
    run = ET.SubElement(parent, _q(W, "r"))
    if formatting:
        props = ET.SubElement(run, _q(W, "rPr"))
        if "font" in formatting:
            ET.SubElement(props, _q(W, "rFonts"), {
                _q(W, key): formatting["font"] for key in ("ascii", "hAnsi", "eastAsia", "cs")
            })
        if formatting.get("bold"):
            ET.SubElement(props, _q(W, "b"))
        if formatting.get("italic"):
            ET.SubElement(props, _q(W, "i"))
        if "color" in formatting:
            ET.SubElement(props, _q(W, "color"), {_q(W, "val"): formatting["color"]})
        if "size_pt" in formatting:
            half_points = str(int(formatting["size_pt"] * 2 + 0.5))
            ET.SubElement(props, _q(W, "sz"), {_q(W, "val"): half_points})
            ET.SubElement(props, _q(W, "szCs"), {_q(W, "val"): half_points})
        if formatting.get("underline"):
            ET.SubElement(props, _q(W, "u"), {_q(W, "val"): "single"})
    text_element = ET.SubElement(run, _q(W, "t"))
    if text[:1].isspace() or text[-1:].isspace():
        text_element.set(_q(XML, "space"), "preserve")
    text_element.text = text


def _paragraph(parent: ET.Element, *, text: str | None = None, runs: list[dict] | None = None,
               style: str | None = None, align: str | None = None,
               num_id: int | None = None, bold: bool = False) -> ET.Element:
    paragraph = ET.SubElement(parent, _q(W, "p"))
    if style or align or num_id is not None:
        props = ET.SubElement(paragraph, _q(W, "pPr"))
        if style:
            ET.SubElement(props, _q(W, "pStyle"), {_q(W, "val"): style})
        if num_id is not None:
            num_props = ET.SubElement(props, _q(W, "numPr"))
            ET.SubElement(num_props, _q(W, "ilvl"), {_q(W, "val"): "0"})
            ET.SubElement(num_props, _q(W, "numId"), {_q(W, "val"): str(num_id)})
        if align:
            value = "both" if align == "justify" else align
            ET.SubElement(props, _q(W, "jc"), {_q(W, "val"): value})
    if text is not None:
        _add_text_run(paragraph, text, {"bold": True} if bold else None)
    elif runs is not None:
        for run in runs:
            _add_text_run(paragraph, run["text"], run)
    return paragraph


def _styles_xml() -> bytes:
    styles = ET.Element(_q(W, "styles"))
    defaults = ET.SubElement(styles, _q(W, "docDefaults"))
    run_default = ET.SubElement(defaults, _q(W, "rPrDefault"))
    run_props = ET.SubElement(run_default, _q(W, "rPr"))
    ET.SubElement(run_props, _q(W, "rFonts"), {
        _q(W, "ascii"): "Arial", _q(W, "hAnsi"): "Arial",
        _q(W, "eastAsia"): "宋体", _q(W, "cs"): "Arial",
    })
    ET.SubElement(run_props, _q(W, "sz"), {_q(W, "val"): "22"})
    ET.SubElement(run_props, _q(W, "szCs"), {_q(W, "val"): "22"})
    paragraph_default = ET.SubElement(defaults, _q(W, "pPrDefault"))
    default_ppr = ET.SubElement(paragraph_default, _q(W, "pPr"))
    ET.SubElement(default_ppr, _q(W, "spacing"), {
        _q(W, "after"): "120", _q(W, "line"): "276", _q(W, "lineRule"): "auto",
    })

    normal = ET.SubElement(styles, _q(W, "style"), {
        _q(W, "type"): "paragraph", _q(W, "default"): "1", _q(W, "styleId"): "Normal",
    })
    ET.SubElement(normal, _q(W, "name"), {_q(W, "val"): "Normal"})
    ET.SubElement(normal, _q(W, "qFormat"))
    for level, size in ((1, 32), (2, 28), (3, 24)):
        heading = ET.SubElement(styles, _q(W, "style"), {
            _q(W, "type"): "paragraph", _q(W, "styleId"): f"Heading{level}",
        })
        ET.SubElement(heading, _q(W, "name"), {_q(W, "val"): f"heading {level}"})
        ET.SubElement(heading, _q(W, "basedOn"), {_q(W, "val"): "Normal"})
        ET.SubElement(heading, _q(W, "next"), {_q(W, "val"): "Normal"})
        ET.SubElement(heading, _q(W, "qFormat"))
        ppr = ET.SubElement(heading, _q(W, "pPr"))
        ET.SubElement(ppr, _q(W, "keepNext"))
        ET.SubElement(ppr, _q(W, "outlineLvl"), {_q(W, "val"): str(level - 1)})
        rpr = ET.SubElement(heading, _q(W, "rPr"))
        ET.SubElement(rpr, _q(W, "b"))
        ET.SubElement(rpr, _q(W, "sz"), {_q(W, "val"): str(size)})
        ET.SubElement(rpr, _q(W, "szCs"), {_q(W, "val"): str(size)})

    default_char = ET.SubElement(styles, _q(W, "style"), {
        _q(W, "type"): "character", _q(W, "default"): "1", _q(W, "styleId"): "DefaultParagraphFont",
    })
    ET.SubElement(default_char, _q(W, "name"), {_q(W, "val"): "Default Paragraph Font"})
    comment_ref = ET.SubElement(styles, _q(W, "style"), {
        _q(W, "type"): "character", _q(W, "styleId"): "CommentReference",
    })
    ET.SubElement(comment_ref, _q(W, "name"), {_q(W, "val"): "Comment Reference"})
    ET.SubElement(comment_ref, _q(W, "basedOn"), {_q(W, "val"): "DefaultParagraphFont"})
    ET.SubElement(comment_ref, _q(W, "semiHidden"))
    ET.SubElement(comment_ref, _q(W, "unhideWhenUsed"))
    return ET.tostring(styles, encoding="UTF-8", xml_declaration=True)


def _numbering_xml(ordered_ids: list[int]) -> bytes:
    numbering = ET.Element(_q(W, "numbering"))
    for abstract_id, ordered in ((0, False), (1, True)):
        abstract = ET.SubElement(numbering, _q(W, "abstractNum"), {
            _q(W, "abstractNumId"): str(abstract_id),
        })
        ET.SubElement(abstract, _q(W, "multiLevelType"), {_q(W, "val"): "singleLevel"})
        level = ET.SubElement(abstract, _q(W, "lvl"), {_q(W, "ilvl"): "0"})
        ET.SubElement(level, _q(W, "start"), {_q(W, "val"): "1"})
        ET.SubElement(level, _q(W, "numFmt"), {
            _q(W, "val"): "decimal" if ordered else "bullet",
        })
        ET.SubElement(level, _q(W, "lvlText"), {
            _q(W, "val"): "%1." if ordered else "•",
        })
        ET.SubElement(level, _q(W, "lvlJc"), {_q(W, "val"): "left"})
        ppr = ET.SubElement(level, _q(W, "pPr"))
        tabs = ET.SubElement(ppr, _q(W, "tabs"))
        ET.SubElement(tabs, _q(W, "tab"), {_q(W, "val"): "num", _q(W, "pos"): "720"})
        ET.SubElement(ppr, _q(W, "ind"), {
            _q(W, "left"): "720", _q(W, "hanging"): "360",
        })
    for num_id, ordered in enumerate(ordered_ids, start=1):
        num = ET.SubElement(numbering, _q(W, "num"), {_q(W, "numId"): str(num_id)})
        ET.SubElement(num, _q(W, "abstractNumId"), {_q(W, "val"): "1" if ordered else "0"})
    return ET.tostring(numbering, encoding="UTF-8", xml_declaration=True)


def _table(parent: ET.Element, block: dict) -> None:
    table = ET.SubElement(parent, _q(W, "tbl"))
    props = ET.SubElement(table, _q(W, "tblPr"))
    total = sum(block["widths_twips"])
    ET.SubElement(props, _q(W, "tblW"), {_q(W, "w"): str(total), _q(W, "type"): "dxa"})
    borders = ET.SubElement(props, _q(W, "tblBorders"))
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        ET.SubElement(borders, _q(W, edge), {
            _q(W, "val"): "single", _q(W, "sz"): "4", _q(W, "space"): "0",
            _q(W, "color"): "auto",
        })
    ET.SubElement(props, _q(W, "tblLayout"), {_q(W, "type"): "fixed"})
    grid = ET.SubElement(table, _q(W, "tblGrid"))
    for width in block["widths_twips"]:
        ET.SubElement(grid, _q(W, "gridCol"), {_q(W, "w"): str(width)})
    for row_index, values in enumerate(block["rows"]):
        row = ET.SubElement(table, _q(W, "tr"))
        if row_index == 0 and block["header"]:
            row_props = ET.SubElement(row, _q(W, "trPr"))
            ET.SubElement(row_props, _q(W, "tblHeader"), {_q(W, "val"): "true"})
        for text, width in zip(values, block["widths_twips"]):
            cell = ET.SubElement(row, _q(W, "tc"))
            cell_props = ET.SubElement(cell, _q(W, "tcPr"))
            ET.SubElement(cell_props, _q(W, "tcW"), {
                _q(W, "w"): str(width), _q(W, "type"): "dxa",
            })
            _paragraph(cell, text=text, bold=row_index == 0 and block["header"])


def _add_picture(parent: ET.Element, block: dict, drawing_id: int) -> None:
    run = ET.SubElement(parent, _q(W, "r"))
    drawing = ET.SubElement(run, _q(W, "drawing"))
    inline = ET.SubElement(drawing, _q(WP, "inline"), {
        "distT": "0", "distB": "0", "distL": "0", "distR": "0",
    })
    ET.SubElement(inline, _q(WP, "extent"), {"cx": str(block["cx"]), "cy": str(block["cy"])})
    ET.SubElement(inline, _q(WP, "docPr"), {
        "id": str(drawing_id), "name": f"Picture {drawing_id}", "descr": block["alt"],
    })
    frame = ET.SubElement(inline, _q(WP, "cNvGraphicFramePr"))
    ET.SubElement(frame, _q(A, "graphicFrameLocks"), {"noChangeAspect": "1"})
    graphic = ET.SubElement(inline, _q(A, "graphic"))
    graphic_data = ET.SubElement(graphic, _q(A, "graphicData"), {
        "uri": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    })
    picture = ET.SubElement(graphic_data, _q(PIC, "pic"))
    nonvisual = ET.SubElement(picture, _q(PIC, "nvPicPr"))
    ET.SubElement(nonvisual, _q(PIC, "cNvPr"), {
        "id": str(drawing_id), "name": f"image{drawing_id}.{block['extension']}",
        "descr": block["alt"],
    })
    ET.SubElement(nonvisual, _q(PIC, "cNvPicPr"))
    fill = ET.SubElement(picture, _q(PIC, "blipFill"))
    ET.SubElement(fill, _q(A, "blip"), {_q(R, "embed"): block["rid"]})
    stretch = ET.SubElement(fill, _q(A, "stretch"))
    ET.SubElement(stretch, _q(A, "fillRect"))
    shape = ET.SubElement(picture, _q(PIC, "spPr"))
    transform = ET.SubElement(shape, _q(A, "xfrm"))
    ET.SubElement(transform, _q(A, "off"), {"x": "0", "y": "0"})
    ET.SubElement(transform, _q(A, "ext"), {"cx": str(block["cx"]), "cy": str(block["cy"])})
    geometry = ET.SubElement(shape, _q(A, "prstGeom"), {"prst": "rect"})
    ET.SubElement(geometry, _q(A, "avLst"))


def _related_text_xml(root_name: str, text: str, *, page_numbers: bool = False) -> bytes:
    root = ET.Element(_q(W, root_name))
    paragraph = ET.SubElement(root, _q(W, "p"))
    if text:
        _add_text_run(paragraph, text)
    if page_numbers:
        _add_text_run(paragraph, " ")
        begin = ET.SubElement(ET.SubElement(paragraph, _q(W, "r")), _q(W, "fldChar"))
        begin.set(_q(W, "fldCharType"), "begin")
        begin.set(_q(W, "dirty"), "true")
        instruction = ET.SubElement(ET.SubElement(paragraph, _q(W, "r")), _q(W, "instrText"))
        instruction.set(_q(XML, "space"), "preserve")
        instruction.text = " PAGE "
        separate = ET.SubElement(ET.SubElement(paragraph, _q(W, "r")), _q(W, "fldChar"))
        separate.set(_q(W, "fldCharType"), "separate")
        _add_text_run(paragraph, "1")
        end = ET.SubElement(ET.SubElement(paragraph, _q(W, "r")), _q(W, "fldChar"))
        end.set(_q(W, "fldCharType"), "end")
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def _document_xml(spec: dict) -> tuple[bytes, list[dict], list[bool]]:
    root = ET.Element(_q(W, "document"))
    body = ET.SubElement(root, _q(W, "body"))
    ordered_ids: list[bool] = []
    image_blocks: list[dict] = []
    next_num_id = 1
    next_drawing_id = 1
    for block in spec["blocks"]:
        kind = block["type"]
        if kind == "heading":
            _paragraph(body, text=block["text"], style=f"Heading{block['level']}")
        elif kind == "paragraph":
            if isinstance(block["content"], str):
                _paragraph(body, text=block["content"], align=block["align"])
            else:
                _paragraph(body, runs=block["content"], align=block["align"])
        elif kind == "list":
            ordered_ids.append(block["ordered"])
            block["num_id"] = next_num_id
            next_num_id += 1
            for item in block["items"]:
                _paragraph(body, text=item, num_id=block["num_id"])
        elif kind == "table":
            _table(body, block)
        elif kind == "image":
            paragraph = ET.SubElement(body, _q(W, "p"))
            _add_picture(paragraph, block, next_drawing_id)
            block["drawing_id"] = next_drawing_id
            image_blocks.append(block)
            next_drawing_id += 1
        elif kind == "page_break":
            paragraph = ET.SubElement(body, _q(W, "p"))
            run = ET.SubElement(paragraph, _q(W, "r"))
            ET.SubElement(run, _q(W, "br"), {_q(W, "type"): "page"})

    section = ET.SubElement(body, _q(W, "sectPr"))
    if spec["header"] is not None:
        ET.SubElement(section, _q(W, "headerReference"), {
            _q(W, "type"): "default", _q(R, "id"): spec["header_rid"],
        })
    if spec["footer"] is not None or spec["page_numbers"]:
        ET.SubElement(section, _q(W, "footerReference"), {
            _q(W, "type"): "default", _q(R, "id"): spec["footer_rid"],
        })
    width, height, margin = spec["page_twips"]
    size_attrs = {_q(W, "w"): str(width), _q(W, "h"): str(height)}
    if width > height:
        size_attrs[_q(W, "orient")] = "landscape"
    ET.SubElement(section, _q(W, "pgSz"), size_attrs)
    ET.SubElement(section, _q(W, "pgMar"), {
        _q(W, "top"): str(margin), _q(W, "right"): str(margin),
        _q(W, "bottom"): str(margin), _q(W, "left"): str(margin),
        _q(W, "header"): str(min(720, margin)), _q(W, "footer"): str(min(720, margin)),
        _q(W, "gutter"): "0",
    })
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True), image_blocks, ordered_ids


def _relationship_xml(relationships: list[tuple[str, str, str]]) -> bytes:
    root = ET.Element(_q(REL, "Relationships"))
    for rel_id, rel_type, target in relationships:
        ET.SubElement(root, _q(REL, "Relationship"), {
            "Id": rel_id, "Type": rel_type, "Target": target,
        })
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def _content_types_xml(extensions: set[str], overrides: list[tuple[str, str]]) -> bytes:
    root = ET.Element(_q(CT, "Types"))
    ET.SubElement(root, _q(CT, "Default"), {
        "Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml",
    })
    ET.SubElement(root, _q(CT, "Default"), {"Extension": "xml", "ContentType": "application/xml"})
    for extension in sorted(extensions):
        content_type = "image/png" if extension == "png" else "image/jpeg"
        ET.SubElement(root, _q(CT, "Default"), {
            "Extension": extension, "ContentType": content_type,
        })
    for name, content_type in overrides:
        ET.SubElement(root, _q(CT, "Override"), {
            "PartName": f"/{name}", "ContentType": content_type,
        })
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def _package_parts(spec: dict) -> dict[str, bytes]:
    relationships: list[tuple[str, str, str]] = []

    def add_relationship(rel_type: str, target: str) -> str:
        rel_id = f"rId{len(relationships) + 1}"
        relationships.append((rel_id, _REL_TYPES[rel_type], target))
        return rel_id

    add_relationship("styles", "styles.xml")
    add_relationship("numbering", "numbering.xml")
    if spec["header"] is not None:
        spec["header_rid"] = add_relationship("header", "header1.xml")
    if spec["footer"] is not None or spec["page_numbers"]:
        spec["footer_rid"] = add_relationship("footer", "footer1.xml")

    for index, block in enumerate(spec["blocks"], start=1):
        if block["type"] == "image":
            name = f"image{index}.{block['extension']}"
            block["media_name"] = name
            block["rid"] = add_relationship("image", f"media/{name}")

    document_xml, image_blocks, ordered_ids = _document_xml(spec)
    parts = {
        "[Content_Types].xml": _content_types_xml(
            {block["extension"] for block in image_blocks},
            [("word/document.xml", _CONTENT_TYPES["document"]),
             ("word/styles.xml", _CONTENT_TYPES["styles"]),
             ("word/numbering.xml", _CONTENT_TYPES["numbering"])],
        ),
        "_rels/.rels": _relationship_xml([
            ("rId1", _REL_TYPES["officeDocument"], "word/document.xml"),
        ]),
        "word/document.xml": document_xml,
        "word/_rels/document.xml.rels": _relationship_xml(relationships),
        "word/styles.xml": _styles_xml(),
        "word/numbering.xml": _numbering_xml(ordered_ids),
    }
    if spec["header"] is not None:
        parts["word/header1.xml"] = _related_text_xml("hdr", spec["header"])
        parts["[Content_Types].xml"] = _content_types_xml(
            {block["extension"] for block in image_blocks},
            [("word/document.xml", _CONTENT_TYPES["document"]),
             ("word/styles.xml", _CONTENT_TYPES["styles"]),
             ("word/numbering.xml", _CONTENT_TYPES["numbering"]),
             ("word/header1.xml", _CONTENT_TYPES["header"])],
        )
    if spec["footer"] is not None or spec["page_numbers"]:
        parts["word/footer1.xml"] = _related_text_xml(
            "ftr", spec["footer"] or "", page_numbers=spec["page_numbers"],
        )
        overrides = [("word/document.xml", _CONTENT_TYPES["document"]),
                     ("word/styles.xml", _CONTENT_TYPES["styles"]),
                     ("word/numbering.xml", _CONTENT_TYPES["numbering"])]
        if spec["header"] is not None:
            overrides.append(("word/header1.xml", _CONTENT_TYPES["header"]))
        overrides.append(("word/footer1.xml", _CONTENT_TYPES["footer"]))
        parts["[Content_Types].xml"] = _content_types_xml(
            {block["extension"] for block in image_blocks}, overrides,
        )
    for block in image_blocks:
        parts[f"word/media/{block['media_name']}"] = block["image_data"]
    if len(parts) > _MAX_PACKAGE_PARTS:
        raise ValueError("generated DOCX exceeds the 4096-part limit")
    for name, content in parts.items():
        if name.endswith((".xml", ".rels")) and len(content) > _MAX_XML_BYTES:
            raise ValueError(f"generated XML part {name} exceeds the 16 MiB limit")
    if sum(map(len, parts.values())) > _MAX_PACKAGE_BYTES:
        raise ValueError("generated DOCX exceeds the 128 MiB package limit")
    return parts


def create_document(spec: dict, output: str | Path, base_dir: str | Path | None = None) -> Path:
    """Create a DOCX at output, resolving relative image paths under base_dir."""
    output_path = Path(output)
    if output_path.suffix.lower() != ".docx":
        raise ValueError("output path must use the .docx extension")
    image_base = Path(base_dir) if base_dir is not None else Path.cwd()
    validated = _validate_spec(spec, image_base)
    resolved_output = output_path.resolve()
    if resolved_output in validated["image_paths"]:
        raise ValueError("output must not overwrite an image used by the document")
    parts = _package_parts(validated)

    parent = output_path.parent
    parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(prefix=f".{output_path.name}.", suffix=".tmp", dir=parent)
    try:
        with os.fdopen(descriptor, "w+b") as file:
            with zipfile.ZipFile(file, "w", compression=zipfile.ZIP_DEFLATED) as package:
                for name, content in parts.items():
                    package.writestr(name, content)
            if os.fstat(file.fileno()).st_size > _MAX_PACKAGE_BYTES:
                raise ValueError("generated DOCX exceeds the 128 MiB package limit")
            file.flush()
            os.fsync(file.fileno())
        os.replace(temporary_name, output_path)
    except BaseException:
        try:
            os.unlink(temporary_name)
        except FileNotFoundError:
            pass
        raise
    return output_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="JSON document specification")
    parser.add_argument("output", type=Path, help="output DOCX path")
    args = parser.parse_args(argv)
    try:
        if args.input.resolve() == args.output.resolve():
            parser.error("output must not refer to the input JSON file")
        spec = json.loads(args.input.read_text(encoding="utf-8"))
        create_document(spec, args.output, base_dir=args.input.resolve().parent)
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as exc:
        parser.error(str(exc))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


