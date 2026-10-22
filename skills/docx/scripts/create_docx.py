"""Create a small, standards-based DOCX package from a JSON block specification."""

from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import math
import os
import re
import stat
import struct
import sys
import tempfile
import unicodedata
import zipfile
import zlib
from pathlib import Path
from typing import Any
from xml.parsers import expat
from xml.etree import ElementTree as ET

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from office import xmlio

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
_ALIGNMENTS = {"left", "center", "right", "justify"}
_STYLE_FIELDS = {
    "font_east_asia", "font_latin", "size_pt", "bold", "align",
    "first_line_chars", "line_spacing_pt", "space_before_pt", "space_after_pt",
    "line_spacing_multiple", "color",
}
_STYLE_DEFAULTS = {
    "body": {
        "font_east_asia": "仿宋", "font_latin": "Times New Roman", "size_pt": 16,
        "bold": False, "align": "left", "first_line_chars": 2,
        "line_spacing_pt": 28, "space_before_pt": 0, "space_after_pt": 0,
    },
    "heading1": {
        "font_east_asia": "方正小标宋简体", "font_latin": "Times New Roman", "size_pt": 18,
        "bold": False, "align": "center", "first_line_chars": 0,
        "space_before_pt": 0, "space_after_pt": 0,
    },
    "heading2": {
        "font_east_asia": "黑体", "font_latin": "Times New Roman", "size_pt": 16,
        "bold": False, "align": "left", "first_line_chars": 0,
        "space_before_pt": 0, "space_after_pt": 0,
    },
    "heading3": {
        "font_east_asia": "楷体", "font_latin": "Times New Roman", "size_pt": 16,
        "bold": True, "align": "left", "first_line_chars": 0,
        "space_before_pt": 0, "space_after_pt": 0,
    },
    "heading4": {
        "font_east_asia": "仿宋", "font_latin": "Times New Roman", "size_pt": 16,
        "bold": False, "align": "left", "first_line_chars": 0,
        "space_before_pt": 0, "space_after_pt": 0,
    },
}
_STYLE_DEFAULTS.update({
    "title": {**_STYLE_DEFAULTS["heading1"], "size_pt": 22, "space_after_pt": 12},
    "subtitle": {**_STYLE_DEFAULTS["heading1"], "size_pt": 14, "space_after_pt": 8},
})
_TABLE_STYLE_DEFAULTS = {
    "table_body": {
        "font_east_asia": "仿宋", "font_latin": "Times New Roman", "size_pt": 16,
        "color": "000000", "bold": False, "align": "left", "first_line_chars": 0,
        "line_spacing_multiple": 1.15, "space_before_pt": 0, "space_after_pt": 0,
    },
    "table_header": {
        "font_east_asia": "仿宋", "font_latin": "Times New Roman", "size_pt": 16,
        "color": "000000", "bold": True, "align": "center", "first_line_chars": 0,
        "line_spacing_multiple": 1.15, "space_before_pt": 0, "space_after_pt": 0,
    },
}
_TABLE_LAYOUT_DEFAULTS = {
    "align": "left", "vertical_align": "center", "border_style": "single",
    "border_width_pt": 0.5, "border_color": "000000", "cell_fill": "none",
    "cell_margin_top_mm": 1, "cell_margin_bottom_mm": 1,
    "cell_margin_left_mm": 1.5, "cell_margin_right_mm": 1.5,
    "repeat_header": True, "allow_row_split": False,
    "preset": "grid", "column_sizing": "equal", "band_fill": "none",
}
_TABLE_BORDER_EDGES = {"top", "bottom", "left", "right", "insideH", "insideV"}


def _q(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def _object(value: Any, label: str, allowed: set[str], required: set[str] = frozenset()) -> dict:
    if not isinstance(value, dict):
        raise ValueError(f"{label} must be an object")
    unknown = set(value) - allowed
    if unknown:
        hints = []
        for key in sorted(unknown, key=str):
            matches = difflib.get_close_matches(str(key), sorted(allowed), n=1, cutoff=0.7)
            if matches:
                hints.append(f"{key!r}: did you mean {matches[0]!r}?")
        suffix = " (" + "; ".join(hints) + ")" if hints else ""
        raise ValueError(f"{label} has unknown field(s): {', '.join(map(str, sorted(unknown, key=str)))}{suffix}")
    missing = required - set(value)
    if missing:
        raise ValueError(f"{label} is missing field(s): {', '.join(sorted(missing))}")
    return value


def _number(value: Any, label: str, low: float, high: float, *, allow_zero: bool = False) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"{label} must be a number")
    try:
        result = float(value)
    except (OverflowError, ValueError):
        raise ValueError(f"{label} must be a finite number") from None
    if not math.isfinite(result) or result > high or result < low or (not allow_zero and result == 0):
        raise ValueError(f"{label} must be between {low:g} and {high:g}")
    return result


def _single_line_text(value: Any, label: str) -> str:
    text = _text(value, label, allow_empty=False)
    if len(text) > 64:
        raise ValueError(f"{label} must be at most 64 characters")
    if any(char in "\t\n\r\u0085\u2028\u2029" for char in text):
        raise ValueError(f"{label} must be a single-line string")
    return text


def _validate_style_fields(value: Any, label: str) -> dict:
    fields = _object(value, label, _STYLE_FIELDS)
    if "line_spacing_pt" in fields and "line_spacing_multiple" in fields:
        raise ValueError(f"{label} cannot set both line_spacing_pt and line_spacing_multiple")
    result: dict[str, Any] = {}
    for key, field_value in fields.items():
        field_label = f"{label}.{key}"
        if key in ("font_east_asia", "font_latin"):
            result[key] = _single_line_text(field_value, field_label)
        elif key == "size_pt":
            result[key] = _number(field_value, field_label, 1, 200)
        elif key == "bold":
            if not isinstance(field_value, bool):
                raise ValueError(f"{field_label} must be a boolean")
            result[key] = field_value
        elif key == "align":
            if not isinstance(field_value, str) or field_value not in _ALIGNMENTS:
                raise ValueError(f"{field_label} must be left, center, right, or justify")
            result[key] = field_value
        elif key == "first_line_chars":
            result[key] = _number(field_value, field_label, 0, 20, allow_zero=True)
        elif key == "line_spacing_pt":
            result[key] = _number(field_value, field_label, 1, 400)
        elif key == "line_spacing_multiple":
            result[key] = _number(field_value, field_label, 0.5, 10)
        elif key == "color":
            color = _text(field_value, field_label, allow_empty=False)
            if not _COLOR_RE.fullmatch(color):
                raise ValueError(f"{field_label} must be six hexadecimal RGB characters without #")
            result[key] = color.upper()
        elif key in ("space_before_pt", "space_after_pt"):
            result[key] = _number(field_value, field_label, 0, 400, allow_zero=True)
    return result


def _merge_table_text_format(base: dict, layer: dict) -> dict:
    result = base.copy()
    if "line_spacing_pt" in layer:
        result.pop("line_spacing_multiple", None)
    elif "line_spacing_multiple" in layer:
        result.pop("line_spacing_pt", None)
    result.update(layer)
    return result


def _validate_table_text_fields(value: Any, label: str) -> dict:
    return _validate_style_fields(value, label)


def _validate_table_layout(value: Any, label: str) -> dict:
    fields = _object(value, label, set(_TABLE_LAYOUT_DEFAULTS) | {"borders"})
    result: dict[str, Any] = {}
    for key, field_value in fields.items():
        field_label = f"{label}.{key}"
        if key == "borders":
            result[key] = _validate_borders(field_value, field_label)
        elif key in {"preset", "column_sizing"}:
            choices = {"grid", "three_line", "zebra"} if key == "preset" else {"equal", "content"}
            if not isinstance(field_value, str) or field_value not in choices:
                raise ValueError(f"{field_label} must be one of {', '.join(sorted(choices))}")
            result[key] = field_value
        elif key == "align":
            if not isinstance(field_value, str) or field_value not in {"left", "center", "right"}:
                raise ValueError(f"{field_label} must be left, center, or right")
            result[key] = field_value
        elif key == "vertical_align":
            if not isinstance(field_value, str) or field_value not in {"top", "center", "bottom"}:
                raise ValueError(f"{field_label} must be top, center, or bottom")
            result[key] = field_value
        elif key == "border_style":
            if not isinstance(field_value, str) or field_value not in {"single", "none"}:
                raise ValueError(f"{field_label} must be single or none")
            result[key] = field_value
        elif key == "border_width_pt":
            result[key] = _number(field_value, field_label, 0.25, 12)
        elif key in {"border_color", "cell_fill", "band_fill"}:
            text = _text(field_value, field_label, allow_empty=False)
            if key != "border_color" and text == "none":
                result[key] = text
            elif _COLOR_RE.fullmatch(text):
                result[key] = text.upper()
            else:
                allowed = "six hexadecimal RGB characters without #" if key == "border_color" \
                    else "six hexadecimal RGB characters or 'none'"
                raise ValueError(f"{field_label} must be {allowed}")
        elif key.startswith("cell_margin_"):
            result[key] = _number(field_value, field_label, 0, 50, allow_zero=True)
        elif key in {"repeat_header", "allow_row_split"}:
            if not isinstance(field_value, bool):
                raise ValueError(f"{field_label} must be a boolean")
            result[key] = field_value
    if "preset" in result:
        preset = result["preset"]
        defaults = {"border_style": "single" if preset == "grid" else "none",
                    "band_fill": "F2F4F7" if preset == "zebra" else "none",
                    "borders": {edge: {"style": "single" if preset == "grid" else "none", "width_pt": 0.5}
                                for edge in _TABLE_BORDER_EDGES}}
        if preset == "three_line":
            for edge in ("top", "bottom"):
                defaults["borders"][edge] = {"style": "single", "width_pt": 1.5}
        for fields in defaults["borders"].values():
            if "border_style" in result:
                fields["style"] = result["border_style"]
            if "border_width_pt" in result:
                fields["width_pt"] = result["border_width_pt"]
        result = _merge_table_layout(defaults, result)
    return result


def _validate_borders(value: Any, label: str) -> dict:
    edges = _object(value, label, _TABLE_BORDER_EDGES)
    result = {}
    for edge, value in edges.items():
        fields = _object(value, f"{label}.{edge}", {"style", "width_pt", "color"})
        border = {}
        if "style" in fields:
            style = fields["style"]
            if not isinstance(style, str) or style not in {"single", "none"}:
                raise ValueError(f"{label}.{edge}.style must be single or none")
            border["style"] = style
        if "width_pt" in fields:
            border["width_pt"] = _number(fields["width_pt"], f"{label}.{edge}.width_pt", 0.25, 12)
        if "color" in fields:
            color = _text(fields["color"], f"{label}.{edge}.color", allow_empty=False)
            if not _COLOR_RE.fullmatch(color):
                raise ValueError(f"{label}.{edge}.color must be six hexadecimal RGB characters")
            border["color"] = color.upper()
        result[edge] = border
    return result


def _merge_table_layout(base: dict, layer: dict) -> dict:
    result = {key: value for key, value in base.items() if key != "borders"}
    result.update({key: value for key, value in layer.items() if key != "borders"})
    if "borders" in base or "borders" in layer:
        borders = {edge: fields.copy() for edge, fields in base.get("borders", {}).items()}
        for edge, fields in layer.get("borders", {}).items():
            borders.setdefault(edge, {}).update(fields)
        result["borders"] = borders
    return result


def _resolved_border(layout: dict, edge: str) -> dict:
    result = {"style": layout["border_style"], "width_pt": layout["border_width_pt"],
              "color": layout["border_color"]}
    result.update(layout.get("borders", {}).get(edge, {}))
    return result


def _border_attributes(border: dict) -> dict:
    if border["style"] == "none":
        return {_q(W, "val"): "nil"}
    return {_q(W, "val"): "single", _q(W, "sz"): str(int(border["width_pt"] * 8 + 0.5)),
            _q(W, "space"): "0", _q(W, "color"): border["color"]}


def _cell_borders(rows: list, layout: dict, label: str) -> None:
    """Resolve explicit cell edges and make shared boundaries unambiguous.

    Mirror a single explicit override to its neighbour. Two explicit overrides
    must encode the same border; never rely on a reader's border conflict rules.
    Cells without any explicit or mirrored override keep the old XML output.
    """
    row_count, column_count = len(rows), len(rows[0])
    resolved = {}
    for row_index, row in enumerate(rows):
        for column, cell in enumerate(row):
            if not isinstance(cell, dict):
                continue
            for edge, fields in cell.get("borders", {}).items():
                table_edge = edge
                if edge == "top" and row_index > 0 or edge == "bottom" and row_index < row_count - 1:
                    table_edge = "insideH"
                elif edge == "left" and column > 0 or edge == "right" and column < column_count - 1:
                    table_edge = "insideV"
                border = _resolved_border(layout, table_edge)
                border.update(fields)
                resolved[row_index, column, edge] = border
    for (row_index, column, edge), border in list(resolved.items()):
        adjacent = {"top": (row_index - 1, column, "bottom"),
                    "bottom": (row_index + 1, column, "top"),
                    "left": (row_index, column - 1, "right"),
                    "right": (row_index, column + 1, "left")}[edge]
        other_row, other_column, _ = adjacent
        if not (0 <= other_row < row_count and 0 <= other_column < column_count):
            continue
        if adjacent in resolved and _border_attributes(resolved[adjacent]) != _border_attributes(border):
            raise ValueError(f"{label}.rows[{row_index}][{column}].borders.{edge} conflicts with the adjacent cell border")
        resolved[adjacent] = border.copy()
    for (row_index, column, edge), border in resolved.items():
        cell = rows[row_index][column]
        if isinstance(cell, str):
            cell = {"text": cell, "image": None}
            rows[row_index][column] = cell
        cell.setdefault("borders", {})[edge] = border


def _theme_base_format(theme: str) -> dict:
    """Share non-official theme typography across body and related parts."""
    return {"font_east_asia": "微软雅黑" if theme == "business" else "宋体",
            "font_latin": "Calibri" if theme == "business" else "Times New Roman",
            "size_pt": 10.5, "line_spacing_multiple": 1.25}


def _validated_styles(value: Any, theme: str = "official") -> dict[str, dict]:
    if value is None:
        raise ValueError("styles must be an object")
    entries = _object(value, "styles", set(_STYLE_DEFAULTS) | set(_TABLE_STYLE_DEFAULTS))
    result = {name: defaults.copy() for name, defaults in _STYLE_DEFAULTS.items()}
    result.update({name: defaults.copy() for name, defaults in _TABLE_STYLE_DEFAULTS.items()})
    if theme != "official":
        for name, defaults in result.items():
            defaults.pop("line_spacing_pt", None)
            defaults.update({**_theme_base_format(theme), "space_after_pt": 4, "first_line_chars": 0})
            if name in {"title", "subtitle"} or name.startswith("heading"):
                defaults.update({"font_east_asia": "黑体", "bold": True,
                                 "size_pt": {"title": 22, "subtitle": 12, "heading1": 16,
                                             "heading2": 14, "heading3": 12, "heading4": 11}[name],
                                 "color": "1A365D" if theme == "business" else "000000"})
                defaults["align"] = "center" if name in {"title", "subtitle"} else "left"
            elif name == "body" and theme == "academic":
                defaults["first_line_chars"] = 2
    for name, fields in entries.items():
        if name in _TABLE_STYLE_DEFAULTS:
            result[name] = _merge_table_text_format(
                result[name], _validate_table_text_fields(fields, f"styles.{name}"),
            )
        else:
            result[name] = _merge_table_text_format(
                result[name], _validate_style_fields(fields, f"styles.{name}"))
    return result


def _normalize_run_fonts(run: dict) -> dict:
    normalized = {key: value for key, value in run.items() if key != "font"}
    if "font" in run:
        normalized["font_east_asia"] = run["font"]
        normalized["font_latin"] = run["font"]
    if "font_east_asia" in run:
        normalized["font_east_asia"] = run["font_east_asia"]
    if "font_latin" in run:
        normalized["font_latin"] = run["font_latin"]
    return normalized


def _text(value: Any, label: str, *, allow_empty: bool = True) -> str:
    if not isinstance(value, str):
        raise ValueError(f"{label} must be a string")
    if not allow_empty and not value:
        raise ValueError(f"{label} must not be empty")
    for index, char in enumerate(value):
        point = ord(char)
        if unicodedata.category(char) == "Cc" or not (
            point in (0x9, 0xA, 0xD)
            or 0x20 <= point <= 0xD7FF
            or 0xE000 <= point <= 0xFFFD
            or 0x10000 <= point <= 0x10FFFF
        ):
            hint = "; split multiline content into paragraph blocks" if char in "\t\r\n" else ""
            raise ValueError(f"{label} contains an unsupported control or XML character "
                             f"U+{point:04X} at character {index + 1}{hint}")
    return value


def _content_widths(rows: list, total: int, padding: int, label: str) -> list[int]:
    """Estimate widths without font metrics; keep a usable minimum per column."""
    columns = len(rows[0])
    minimum = padding + _twips(8)
    if columns * minimum > total:
        raise ValueError(f"{label}: content column sizing needs fewer columns or smaller cell margins")
    weights = []
    for column in range(columns):
        longest = 1
        for row in rows:
            cell = row[column]
            text = cell if isinstance(cell, str) else cell.get("text", "") if isinstance(cell, dict) else ""
            if isinstance(text, str):
                longest = max(longest, sum(2 if unicodedata.east_asian_width(char) in {"W", "F"}
                                           else 1 for char in text))
        weights.append(math.sqrt(longest))
    remaining = total - minimum * columns
    weight_total = sum(weights)
    widths = [minimum + int(remaining * weight / weight_total) for weight in weights]
    for index in range(total - sum(widths)):
        widths[index % columns] += 1
    return widths


def _twips(mm: float) -> int:
    return int(mm * 1440.0 / 25.4 + 0.5)


def _emu(mm: float) -> int:
    return int(mm * 36000.0 + 0.5)


def _image_size(data: bytes, label: str) -> tuple[int, int]:
    """Read PNG IHDR or JPEG SOF dimensions without decoding image pixels."""
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        if len(data) >= 33 and data[8:16] == b"\x00\x00\x00\rIHDR":
            if zlib.crc32(data[12:29]) & 0xFFFFFFFF == int.from_bytes(data[29:33], "big"):
                width, height = struct.unpack(">II", data[16:24])
                if width and height:
                    return width, height
    elif data.startswith(b"\xff\xd8"):
        cursor = 2
        while cursor + 4 <= len(data):
            if data[cursor] != 0xFF:
                break
            while cursor < len(data) and data[cursor] == 0xFF:
                cursor += 1
            if cursor >= len(data):
                break
            marker = data[cursor]
            cursor += 1
            if marker in {0xD9, 0xDA, 0x00}:
                break
            if marker == 0x01 or 0xD0 <= marker <= 0xD7:
                continue
            length = int.from_bytes(data[cursor:cursor + 2], "big")
            if length < 2 or cursor + length > len(data):
                break
            if marker in {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF} and length >= 8:
                height, width = struct.unpack(">HH", data[cursor + 3:cursor + 7])
                if width and height:
                    return width, height
            cursor += length
    raise ValueError(f"{label}: cannot infer image ratio from a valid PNG/JPEG dimension header; provide both dimensions")


class ImageResources:
    """One operation's immutable image bytes; never cache across documents."""

    def __init__(self, deduplicate: bool):
        self.paths = {}
        self.payloads = {}
        self.loaded_bytes = 0
        self.packaged_bytes = 0
        self.packaged_digests = set()
        self.deduplicate = deduplicate

    def load(self, path: Path, label: str) -> dict:
        if path not in self.paths:
            if not stat.S_ISREG(path.stat().st_mode):
                raise ValueError(f"{label}.path must reference a regular image file")
            with path.open("rb") as image_file:
                info = os.fstat(image_file.fileno())
                if not stat.S_ISREG(info.st_mode):
                    raise ValueError(f"{label}.path must reference a regular image file")
                if info.st_size > _MAX_IMAGE_BYTES:
                    raise ValueError(f"{label}.path exceeds the 32 MiB image limit")
                if self.loaded_bytes + info.st_size > _MAX_PACKAGE_BYTES:
                    raise ValueError(f"{label}.path exceeds the 128 MiB image resource budget")
                # Bound by the observed file size, not a 32 MiB allocation for
                # every tiny image. Reject growth/truncation during this read.
                data = image_file.read(info.st_size + 1)
                if len(data) != info.st_size:
                    raise ValueError(f"{label}.path changed size while reading")
            if data.startswith(b"\x89PNG\r\n\x1a\n") and data.endswith(b"\x00\x00\x00\x00IEND\xaeB\x60\x82"):
                extension = "png"
            elif data.startswith(b"\xff\xd8") and b"\xff\xd9" in data[2:]:
                extension = "jpg"
            else:
                raise ValueError(f"{label}.path must reference an existing PNG or JPEG image")
            digest = hashlib.sha256(data).digest()
            resource = self.payloads.get(digest)
            if resource is None:
                resource = {"data": data, "extension": extension, "digest": digest}
                self.payloads[digest] = resource
                self.loaded_bytes += len(data)
            self.paths[path] = resource
        resource = self.paths[path]
        if not self.deduplicate or resource["digest"] not in self.packaged_digests:
            self.packaged_bytes += len(resource["data"])
            if self.packaged_bytes > _MAX_PACKAGE_BYTES:
                raise ValueError(f"{label}.path exceeds the 128 MiB package image budget")
            self.packaged_digests.add(resource["digest"])
        return resource


def _image(path_value: Any, width: float | None, height: float | None, base_dir: Path,
           alt: Any, label: str, resources: ImageResources) -> dict:
    alt_text = _text(alt, f"{label}.alt")
    rel_path = _text(path_value, f"{label}.path", allow_empty=False)
    path = Path(rel_path)
    if not path.is_absolute():
        path = base_dir / path
    path = path.resolve()
    resource = resources.load(path, label)
    if width is None or height is None:
        if "dimensions" not in resource:
            resource["dimensions"] = _image_size(resource["data"], label)
        pixels_wide, pixels_high = resource["dimensions"]
        if width is None:
            width = height * pixels_wide / pixels_high
        else:
            height = width * pixels_high / pixels_wide
    return {
        "type": "image", "image_path": path, "image_data": resource["data"],
        "image_digest": resource["digest"],
        "extension": resource["extension"], "cx": _emu(width), "cy": _emu(height),
        "alt": alt_text,
    }


def _prepare_image(fields: dict, base_dir: Path, label: str, max_width: int, max_height: int,
                   resources: ImageResources) -> dict:
    if "width_mm" not in fields and "height_mm" not in fields:
        raise ValueError(f"{label} requires width_mm and/or height_mm")
    width = _number(fields["width_mm"], f"{label}.width_mm", 0.1, _MAX_PAGE_MM) if "width_mm" in fields else None
    height = _number(fields["height_mm"], f"{label}.height_mm", 0.1, _MAX_PAGE_MM) if "height_mm" in fields else None
    if width is not None and _emu(width) > max_width or height is not None and _emu(height) > max_height:
        raise ValueError(f"{label} dimensions exceed the cell width/page content area")
    image = _image(fields["path"], width, height, base_dir, fields["alt"], label, resources)
    if image["cx"] > max_width or image["cy"] > max_height or min(image["cx"], image["cy"]) < _emu(0.1):
        raise ValueError(f"{label} dimensions exceed the cell width/page content area or are too small")
    return image


def _validate_run(value: Any, label: str) -> dict:
    run = _object(value, label,
                  {"text", "bold", "italic", "underline", "font", "font_east_asia",
                   "font_latin", "size_pt", "color"},
                  {"text"})
    result = {"text": _text(run["text"], f"{label}.text")}
    for key in ("bold", "italic", "underline"):
        if key in run:
            if not isinstance(run[key], bool):
                raise ValueError(f"{label}.{key} must be a boolean")
            result[key] = run[key]
    for key in ("font", "font_east_asia", "font_latin"):
        if key in run:
            result[key] = _single_line_text(run[key], f"{label}.{key}")
    if "size_pt" in run:
        result["size_pt"] = _number(run["size_pt"], f"{label}.size_pt", 1, 200)
    if "color" in run:
        color = _text(run["color"], f"{label}.color", allow_empty=False)
        if not _COLOR_RE.fullmatch(color):
            raise ValueError(f"{label}.color must be six hexadecimal RGB characters without #")
        result["color"] = color.upper()
    return result


def _validate_spec(spec: Any, base_dir: Path) -> dict:
    source = _object(spec, "spec", {
        "page", "styles", "table_format", "header", "footer", "page_numbers", "blocks", "theme", "deduplicate_images",
    }, {"blocks"})
    page = _object(source.get("page", {}), "page", {
        "width_mm", "height_mm", "margin_mm", "margin_top_mm", "margin_bottom_mm",
        "margin_left_mm", "margin_right_mm",
    })
    width = _number(page.get("width_mm", 210), "page.width_mm", 1, _MAX_PAGE_MM)
    height = _number(page.get("height_mm", 297), "page.height_mm", 1, _MAX_PAGE_MM)
    legacy_margin = _number(page["margin_mm"], "page.margin_mm", 0, _MAX_MARGIN_MM, allow_zero=True) \
        if "margin_mm" in page else None
    margin_defaults = {"top": 25.4, "bottom": 25.4, "left": 31.8, "right": 31.8}
    margins = {}
    for side in ("top", "bottom", "left", "right"):
        key = f"margin_{side}_mm"
        if key in page:
            supplied = _number(page[key], f"page.{key}", 0, _MAX_MARGIN_MM, allow_zero=True)
        else:
            supplied = legacy_margin if legacy_margin is not None else margin_defaults[side]
        margins[side] = supplied
    page_twips = {
        "width": _twips(width), "height": _twips(height),
        **{side: _twips(value) for side, value in margins.items()},
    }
    available_width = page_twips["width"] - page_twips["left"] - page_twips["right"]
    available_height = page_twips["height"] - page_twips["top"] - page_twips["bottom"]
    if available_width <= 0 or available_height <= 0:
        raise ValueError("page margins must leave positive width and height for content after rounding")
    theme = source.get("theme", "official")
    if not isinstance(theme, str) or theme not in {"official", "business", "academic"}:
        raise ValueError("theme must be official, business, or academic")
    styles = _validated_styles(source.get("styles", {}), theme)
    table_layout = _merge_table_layout(
        _TABLE_LAYOUT_DEFAULTS,
        _validate_table_layout(source["table_format"], "table_format") if "table_format" in source else {},
    )

    header = _text(source["header"], "header") if "header" in source else None
    footer = _text(source["footer"], "footer") if "footer" in source else None
    page_numbers = source.get("page_numbers", False)
    if not isinstance(page_numbers, bool):
        raise ValueError("page_numbers must be a boolean")
    deduplicate_images = source.get("deduplicate_images", False)
    if not isinstance(deduplicate_images, bool):
        raise ValueError("deduplicate_images must be a boolean")
    resources = ImageResources(deduplicate_images)

    blocks_value = source["blocks"]
    if not isinstance(blocks_value, list):
        raise ValueError("blocks must be an array")
    blocks: list[dict] = []
    image_paths: list[Path] = []
    shared_lists = {}
    for index, value in enumerate(blocks_value):
        label = f"blocks[{index}]"
        if not isinstance(value, dict):
            raise ValueError(f"{label} must be an object")
        kind = value.get("type")
        if not isinstance(kind, str):
            raise ValueError(f"{label}.type must be a string")
        if kind in {"title", "subtitle"}:
            block = _object(value, label, {"type", "text", "format", "align"}, {"type", "text"})
            formatting = _validate_style_fields(block.get("format", {}), f"{label}.format")
            if "align" in block:
                formatting.update(_validate_style_fields({"align": block["align"]}, f"{label}.format"))
            blocks.append({"type": kind, "text": _text(block["text"], f"{label}.text"), "format": formatting})
        elif kind == "heading":
            block = _object(value, label, {"type", "text", "level", "format", "align"},
                            {"type", "text", "level"})
            level = block["level"]
            if isinstance(level, bool) or not isinstance(level, int) or level not in (1, 2, 3, 4):
                raise ValueError(f"{label}.level must be 1, 2, 3, or 4")
            block_format = _validate_style_fields(block["format"], f"{label}.format") \
                if "format" in block else {}
            align = block.get("align")
            if "align" in block and (not isinstance(align, str) or align not in _ALIGNMENTS):
                raise ValueError(f"{label}.align must be left, center, right, or justify")
            blocks.append({"type": kind, "text": _text(block["text"], f"{label}.text"),
                           "level": level, "format": block_format, "align": align})
        elif kind == "paragraph":
            block = _object(value, label, {"type", "text", "runs", "format", "align"}, {"type"})
            if ("text" in block) == ("runs" in block):
                raise ValueError(f"{label} must contain exactly one of text or runs")
            align = block.get("align")
            if "align" in block and (not isinstance(align, str) or align not in _ALIGNMENTS):
                raise ValueError(f"{label}.align must be left, center, right, or justify")
            block_format = _validate_style_fields(block["format"], f"{label}.format") \
                if "format" in block else {}
            if "text" in block:
                content: Any = _text(block["text"], f"{label}.text")
            else:
                if not isinstance(block["runs"], list):
                    raise ValueError(f"{label}.runs must be an array")
                content = [_validate_run(run, f"{label}.runs[{i}]") for i, run in enumerate(block["runs"])]
            blocks.append({"type": kind, "content": content, "format": block_format, "align": align})
        elif kind == "list":
            block = _object(value, label, {"type", "items", "ordered", "list_id", "level"},
                            {"type", "items", "ordered"})
            if not isinstance(block["ordered"], bool):
                raise ValueError(f"{label}.ordered must be a boolean")
            if not isinstance(block["items"], list):
                raise ValueError(f"{label}.items must be an array")
            items = [_text(item, f"{label}.items[{i}]") for i, item in enumerate(block["items"])]
            level = block.get("level", 0)
            if type(level) is not int or not 0 <= level <= 8:
                raise ValueError(f"{label}.level must be an integer from 0 to 8")
            list_id = None
            if "list_id" in block:
                list_id = _single_line_text(block["list_id"], f"{label}.list_id")
                if list_id != list_id.strip():
                    raise ValueError(f"{label}.list_id cannot contain edge whitespace")
                if not items:
                    raise ValueError(f"{label}: a named list must contain at least one item")
                if list_id not in shared_lists:
                    if level != 0:
                        raise ValueError(f"{label}: a named list must begin at level 0")
                    shared_lists[list_id] = {"ordered": block["ordered"], "last_level": level}
                previous = shared_lists[list_id]
                if previous["ordered"] != block["ordered"]:
                    raise ValueError(f"{label}: ordered must stay consistent for one list_id")
                if level > previous["last_level"] + 1:
                    raise ValueError(f"{label}: a named list cannot skip an intervening parent level")
                if items:
                    previous["last_level"] = level
            elif level != 0:
                raise ValueError(f"{label}: a nonzero level requires list_id")
            blocks.append({"type": kind, "items": items, "ordered": block["ordered"],
                           "list_id": list_id, "level": level})
        elif kind == "table":
            block = _object(value, label, {
                "type", "rows", "widths_mm", "header", "header_rows", "format", "body_format", "header_format",
            }, {"type", "rows"})
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
                rows.append(list(row))
            assert columns is not None
            if "header" in block and not isinstance(block["header"], bool):
                raise ValueError(f"{label}.header must be a boolean")
            header_rows = block.get("header_rows", 1 if block.get("header", False) else 0)
            if type(header_rows) is not int or not 0 <= header_rows <= len(rows):
                raise ValueError(f"{label}.header_rows must be an integer from 0 to the row count")
            if "header" in block and "header_rows" in block and header_rows != int(block["header"]):
                raise ValueError(f"{label}.header and header_rows conflict; choose one header declaration")
            layout_overrides = _validate_table_layout(block["format"], f"{label}.format") \
                if "format" in block else {}
            body_overrides = _validate_table_text_fields(block["body_format"], f"{label}.body_format") \
                if "body_format" in block else {}
            header_overrides = _validate_table_text_fields(block["header_format"], f"{label}.header_format") \
                if "header_format" in block else {}
            resolved_layout = _merge_table_layout(table_layout, layout_overrides)
            if "widths_mm" in block:
                widths = block["widths_mm"]
                if not isinstance(widths, list) or len(widths) != columns:
                    raise ValueError(f"{label}.widths_mm must contain one width per column")
                widths_twips = [_twips(_number(v, f"{label}.widths_mm[{i}]", 0.1, _MAX_PAGE_MM))
                                for i, v in enumerate(widths)]
                if sum(widths_twips) > available_width:
                    raise ValueError(f"{label}.widths_mm exceeds the page content width")
            else:
                if resolved_layout["column_sizing"] == "content":
                    padding = _twips(resolved_layout["cell_margin_left_mm"]) + _twips(resolved_layout["cell_margin_right_mm"])
                    widths_twips = _content_widths(rows, available_width, padding, label)
                else:
                    quotient, remainder = divmod(available_width, columns)
                    widths_twips = [quotient + (1 if i < remainder else 0) for i in range(columns)]
            horizontal_padding = _twips(resolved_layout["cell_margin_left_mm"]) \
                + _twips(resolved_layout["cell_margin_right_mm"])
            if any(width - horizontal_padding <= 0 for width in widths_twips):
                raise ValueError(f"{label} cell margins must leave positive content width in every column")
            parsed_rows = []
            vertical_padding = _twips(resolved_layout["cell_margin_top_mm"]) \
                + _twips(resolved_layout["cell_margin_bottom_mm"])
            for row_index, row in enumerate(rows):
                parsed = []
                for column, cell in enumerate(row):
                    cell_label = f"{label}.rows[{row_index}][{column}]"
                    if isinstance(cell, str):
                        parsed.append(_text(cell, cell_label))
                        continue
                    fields = _object(cell, cell_label, {"text", "image", "borders"})
                    if not ({"text", "image"} & fields.keys()):
                        raise ValueError(f"{cell_label} must contain text and/or image")
                    text = _text(fields.get("text", ""), f"{cell_label}.text")
                    normalized = {"text": text, "image": None}
                    if "borders" in fields:
                        _object(fields["borders"], f"{cell_label}.borders", {"top", "left", "bottom", "right"})
                        normalized["borders"] = _validate_borders(fields["borders"], f"{cell_label}.borders")
                    if "image" in fields:
                        image_label = f"{cell_label}.image"
                        image_fields = _object(fields["image"], image_label,
                                               {"path", "width_mm", "height_mm", "alt"},
                                               {"path", "alt"})
                        image = _prepare_image(image_fields, base_dir, image_label,
                                               (widths_twips[column] - horizontal_padding) * 635,
                                               (available_height - vertical_padding) * 635, resources)
                        image_paths.append(image["image_path"])
                        normalized["image"] = image
                    parsed.append(normalized)
                parsed_rows.append(parsed)
            _cell_borders(parsed_rows, resolved_layout, label)
            if resolved_layout["preset"] == "three_line" and 0 < header_rows < len(parsed_rows):
                for column in range(columns):
                    for row_index, edge in ((header_rows - 1, "bottom"), (header_rows, "top")):
                        cell = parsed_rows[row_index][column]
                        if isinstance(cell, str):
                            cell = {"text": cell, "image": None}
                            parsed_rows[row_index][column] = cell
                        cell.setdefault("borders", {}).setdefault(edge, {
                            "style": "single", "width_pt": 0.75, "color": resolved_layout["border_color"],
                        })
            blocks.append({"type": kind, "rows": parsed_rows, "widths_twips": widths_twips,
                           "header": header_rows > 0, "header_rows": header_rows, "format": resolved_layout,
                           "body_format": body_overrides, "header_format": header_overrides})
        elif kind == "image":
            block = _object(value, label, {"type", "path", "width_mm", "height_mm", "alt"},
                            {"type", "path", "alt"})
            image = _prepare_image(block, base_dir, label, available_width * 635, available_height * 635,
                                   resources)
            image_paths.append(image["image_path"])
            blocks.append(image)
        elif kind == "page_break":
            _object(value, label, {"type"}, {"type"})
            blocks.append({"type": kind})
        else:
            raise ValueError(f"{label}.type has unsupported value {kind!r}")
    return {
        "page_twips": page_twips, "available_width_twips": available_width,
        "available_height_twips": available_height, "styles": styles, "table_format": table_layout,
        "blocks": blocks,
        "header": header, "footer": footer,
        "page_numbers": page_numbers, "image_paths": image_paths,
        "deduplicate_images": deduplicate_images,
        "list_format": None if theme == "official" else {
            **_theme_base_format(theme), "bold": False, "color": "000000", "align": "left",
            "space_before_pt": 0, "space_after_pt": 4,
        },
        "related_format": None if theme == "official" else {
            **_theme_base_format(theme), "bold": False, "color": "000000", "align": "left",
            "first_line_chars": 0, "space_before_pt": 0, "space_after_pt": 0,
        },
    }


def _half_points(value: float) -> str:
    return str(int(value * 2 + 0.5))


def _points_to_twips(value: float) -> str:
    return str(int(value * 20 + 0.5))


def _characters_to_hundredths(value: float) -> str:
    return str(int(value * 100 + 0.5))


def _append_character_properties(parent: ET.Element, formatting: dict) -> None:
    font_attrs = {}
    if "font_latin" in formatting:
        for key in ("ascii", "hAnsi", "cs"):
            font_attrs[_q(W, key)] = formatting["font_latin"]
    if "font_east_asia" in formatting:
        font_attrs[_q(W, "eastAsia")] = formatting["font_east_asia"]
    if font_attrs:
        ET.SubElement(parent, _q(W, "rFonts"), font_attrs)
    for key, tag in (("bold", "b"), ("italic", "i")):
        if key in formatting:
            ET.SubElement(parent, _q(W, tag), {_q(W, "val"): "1" if formatting[key] else "0"})
    if "color" in formatting:
        ET.SubElement(parent, _q(W, "color"), {_q(W, "val"): formatting["color"]})
    if "size_pt" in formatting:
        half_points = _half_points(formatting["size_pt"])
        ET.SubElement(parent, _q(W, "sz"), {_q(W, "val"): half_points})
        ET.SubElement(parent, _q(W, "szCs"), {_q(W, "val"): half_points})
    if "underline" in formatting:
        ET.SubElement(parent, _q(W, "u"), {
            _q(W, "val"): "single" if formatting["underline"] else "none",
        })


def _add_text_run(parent: ET.Element, text: str, formatting: dict | None = None) -> None:
    run = ET.SubElement(parent, _q(W, "r"))
    if formatting:
        props = ET.Element(_q(W, "rPr"))
        _append_character_properties(props, formatting)
        if len(props):
            run.append(props)
    text_element = ET.SubElement(run, _q(W, "t"))
    if text[:1].isspace() or text[-1:].isspace():
        text_element.set(_q(XML, "space"), "preserve")
    text_element.text = text


def _append_paragraph_properties(props: ET.Element, formatting: dict, *, outline_level: int | None = None,
                                 keep_next: bool = False) -> None:
    if keep_next:
        ET.SubElement(props, _q(W, "keepNext"))
    if "space_before_pt" in formatting or "space_after_pt" in formatting \
            or "line_spacing_pt" in formatting or "line_spacing_multiple" in formatting:
        attrs = {}
        if "space_before_pt" in formatting:
            attrs[_q(W, "before")] = _points_to_twips(formatting["space_before_pt"])
        if "space_after_pt" in formatting:
            attrs[_q(W, "after")] = _points_to_twips(formatting["space_after_pt"])
        if "line_spacing_pt" in formatting:
            attrs[_q(W, "line")] = _points_to_twips(formatting["line_spacing_pt"])
            attrs[_q(W, "lineRule")] = "exact"
        elif "line_spacing_multiple" in formatting:
            attrs[_q(W, "line")] = str(int(formatting["line_spacing_multiple"] * 240 + 0.5))
            attrs[_q(W, "lineRule")] = "auto"
        ET.SubElement(props, _q(W, "spacing"), attrs)
    if "first_line_chars" in formatting:
        ET.SubElement(props, _q(W, "ind"), {
            _q(W, "firstLineChars"): _characters_to_hundredths(formatting["first_line_chars"]),
        })
    if "align" in formatting:
        value = "both" if formatting["align"] == "justify" else formatting["align"]
        ET.SubElement(props, _q(W, "jc"), {_q(W, "val"): value})
    if outline_level is not None:
        ET.SubElement(props, _q(W, "outlineLvl"), {_q(W, "val"): str(outline_level)})


def _paragraph(parent: ET.Element, *, text: str | None = None, runs: list[dict] | None = None,
               style: str | None = None, formatting: dict | None = None,
               num_id: int | None = None, bold: bool = False,
               outline_level: int | None = None, keep_next: bool = False,
               list_level: int = 0) -> ET.Element:
    paragraph = ET.SubElement(parent, _q(W, "p"))
    if style or formatting or num_id is not None or outline_level is not None or keep_next:
        props = ET.SubElement(paragraph, _q(W, "pPr"))
        if style:
            ET.SubElement(props, _q(W, "pStyle"), {_q(W, "val"): style})
        if keep_next:
            ET.SubElement(props, _q(W, "keepNext"))
        if num_id is not None:
            num_props = ET.SubElement(props, _q(W, "numPr"))
            ET.SubElement(num_props, _q(W, "ilvl"), {_q(W, "val"): str(list_level)})
            ET.SubElement(num_props, _q(W, "numId"), {_q(W, "val"): str(num_id)})
        _append_paragraph_properties(props, formatting or {}, outline_level=outline_level)
    if text is not None:
        _add_text_run(paragraph, text, formatting if formatting is not None else ({"bold": True} if bold else None))
    elif runs is not None:
        for run in runs:
            run_format = dict(formatting or {})
            run_format.update(_normalize_run_fonts(run))
            run_format.pop("text", None)
            _add_text_run(paragraph, run["text"], run_format)
    return paragraph


def _styles_xml(style_formats: dict[str, dict], related_format: dict | None = None,
                list_format: dict | None = None) -> bytes:
    styles = ET.Element(_q(W, "styles"))
    defaults = ET.SubElement(styles, _q(W, "docDefaults"))
    run_default = ET.SubElement(defaults, _q(W, "rPrDefault"))
    run_props = ET.SubElement(run_default, _q(W, "rPr"))
    normal_format = {
        "font_east_asia": "仿宋", "font_latin": "Times New Roman", "size_pt": 16, "bold": False,
    }
    _append_character_properties(run_props, normal_format)
    paragraph_default = ET.SubElement(defaults, _q(W, "pPrDefault"))
    default_ppr = ET.SubElement(paragraph_default, _q(W, "pPr"))
    ET.SubElement(default_ppr, _q(W, "spacing"), {
        _q(W, "before"): "0", _q(W, "after"): "0",
        _q(W, "line"): "276", _q(W, "lineRule"): "auto",
    })

    normal = ET.SubElement(styles, _q(W, "style"), {
        _q(W, "type"): "paragraph", _q(W, "default"): "1", _q(W, "styleId"): "Normal",
    })
    ET.SubElement(normal, _q(W, "name"), {_q(W, "val"): "Normal"})
    ET.SubElement(normal, _q(W, "qFormat"))
    normal_ppr = ET.SubElement(normal, _q(W, "pPr"))
    _append_paragraph_properties(normal_ppr, {
        "space_before_pt": 0, "space_after_pt": 0,
    })
    normal_spacing = normal_ppr.find(_q(W, "spacing"))
    normal_spacing.set(_q(W, "line"), "276")
    normal_spacing.set(_q(W, "lineRule"), "auto")
    normal_ppr.append(ET.Element(_q(W, "jc"), {_q(W, "val"): "left"}))
    normal_rpr = ET.SubElement(normal, _q(W, "rPr"))
    _append_character_properties(normal_rpr, normal_format)

    if related_format is not None:
        for style_id in ("Header", "Footer"):
            style = ET.SubElement(styles, _q(W, "style"), {
                _q(W, "type"): "paragraph", _q(W, "styleId"): style_id,
            })
            ET.SubElement(style, _q(W, "name"), {_q(W, "val"): style_id})
            ET.SubElement(style, _q(W, "basedOn"), {_q(W, "val"): "Normal"})
            _append_paragraph_properties(ET.SubElement(style, _q(W, "pPr")), related_format)
            _append_character_properties(ET.SubElement(style, _q(W, "rPr")), related_format)

    if list_format is not None:
        style = ET.SubElement(styles, _q(W, "style"), {
            _q(W, "type"): "paragraph", _q(W, "styleId"): "ListParagraph",
        })
        ET.SubElement(style, _q(W, "name"), {_q(W, "val"): "List Paragraph"})
        ET.SubElement(style, _q(W, "basedOn"), {_q(W, "val"): "Normal"})
        _append_paragraph_properties(ET.SubElement(style, _q(W, "pPr")), list_format)
        _append_character_properties(ET.SubElement(style, _q(W, "rPr")), list_format)

    body = ET.SubElement(styles, _q(W, "style"), {
        _q(W, "type"): "paragraph", _q(W, "styleId"): "BodyText",
    })
    ET.SubElement(body, _q(W, "name"), {_q(W, "val"): "Body Text"})
    ET.SubElement(body, _q(W, "basedOn"), {_q(W, "val"): "Normal"})
    ET.SubElement(body, _q(W, "next"), {_q(W, "val"): "BodyText"})
    ET.SubElement(body, _q(W, "qFormat"))
    body_ppr = ET.SubElement(body, _q(W, "pPr"))
    _append_paragraph_properties(body_ppr, style_formats["body"])
    body_rpr = ET.SubElement(body, _q(W, "rPr"))
    _append_character_properties(body_rpr, style_formats["body"])

    for style_key, style_name in (("table_body", "TableBody"), ("table_header", "TableHeader")):
        table_style = ET.SubElement(styles, _q(W, "style"), {
            _q(W, "type"): "paragraph", _q(W, "styleId"): style_name,
        })
        ET.SubElement(table_style, _q(W, "name"), {
            _q(W, "val"): "Table Body" if style_key == "table_body" else "Table Header",
        })
        ET.SubElement(table_style, _q(W, "basedOn"), {_q(W, "val"): "Normal"})
        ET.SubElement(table_style, _q(W, "next"), {_q(W, "val"): "Normal"})
        ET.SubElement(table_style, _q(W, "qFormat"))
        table_ppr = ET.SubElement(table_style, _q(W, "pPr"))
        _append_paragraph_properties(table_ppr, style_formats[style_key])
        table_rpr = ET.SubElement(table_style, _q(W, "rPr"))
        _append_character_properties(table_rpr, style_formats[style_key])

    for key, style_id in (("title", "Title"), ("subtitle", "Subtitle")):
        style = ET.SubElement(styles, _q(W, "style"), {
            _q(W, "type"): "paragraph", _q(W, "styleId"): style_id,
        })
        ET.SubElement(style, _q(W, "name"), {_q(W, "val"): style_id})
        ET.SubElement(style, _q(W, "basedOn"), {_q(W, "val"): "Normal"})
        ET.SubElement(style, _q(W, "next"), {_q(W, "val"): "BodyText"})
        _append_paragraph_properties(ET.SubElement(style, _q(W, "pPr")), style_formats[key], keep_next=True)
        _append_character_properties(ET.SubElement(style, _q(W, "rPr")), style_formats[key])

    for level in range(1, 5):
        format_name = f"heading{level}"
        style_name = f"Heading{level}"
        heading = ET.SubElement(styles, _q(W, "style"), {
            _q(W, "type"): "paragraph", _q(W, "styleId"): style_name,
        })
        ET.SubElement(heading, _q(W, "name"), {_q(W, "val"): f"heading {level}"})
        ET.SubElement(heading, _q(W, "basedOn"), {_q(W, "val"): "Normal"})
        ET.SubElement(heading, _q(W, "next"), {_q(W, "val"): "Normal"})
        ET.SubElement(heading, _q(W, "qFormat"))
        ppr = ET.SubElement(heading, _q(W, "pPr"))
        _append_paragraph_properties(ppr, style_formats[format_name],
                                     outline_level=level - 1, keep_next=True)
        rpr = ET.SubElement(heading, _q(W, "rPr"))
        _append_character_properties(rpr, style_formats[format_name])

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


def _numbering_xml(instances: list[dict]) -> bytes:
    numbering = ET.Element(_q(W, "numbering"))
    definitions = [(0, False, False), (1, True, False)]
    for ordered in (False, True):
        if any(instance["ordered"] == ordered and instance["multilevel"] for instance in instances):
            definitions.append((3 if ordered else 2, ordered, True))
    for abstract_id, ordered, multilevel in definitions:
        abstract = ET.SubElement(numbering, _q(W, "abstractNum"), {
            _q(W, "abstractNumId"): str(abstract_id),
        })
        ET.SubElement(abstract, _q(W, "multiLevelType"), {
            _q(W, "val"): "multilevel" if multilevel else "singleLevel",
        })
        for index in range(9 if multilevel else 1):
            level = ET.SubElement(abstract, _q(W, "lvl"), {_q(W, "ilvl"): str(index)})
            ET.SubElement(level, _q(W, "start"), {_q(W, "val"): "1"})
            ET.SubElement(level, _q(W, "numFmt"), {
                _q(W, "val"): "decimal" if ordered else "bullet",
            })
            if multilevel and index:
                ET.SubElement(level, _q(W, "lvlRestart"), {_q(W, "val"): str(index)})
            label = ".".join(f"%{position + 1}" for position in range(index + 1)) + "." if ordered else "•"
            ET.SubElement(level, _q(W, "lvlText"), {_q(W, "val"): label})
            ET.SubElement(level, _q(W, "lvlJc"), {_q(W, "val"): "left"})
            ppr = ET.SubElement(level, _q(W, "pPr"))
            tabs = ET.SubElement(ppr, _q(W, "tabs"))
            left = str(720 * (index + 1))
            ET.SubElement(tabs, _q(W, "tab"), {_q(W, "val"): "num", _q(W, "pos"): left})
            ET.SubElement(ppr, _q(W, "ind"), {_q(W, "left"): left, _q(W, "hanging"): "360"})
    for num_id, instance in enumerate(instances, start=1):
        num = ET.SubElement(numbering, _q(W, "num"), {_q(W, "numId"): str(num_id)})
        abstract_id = int(instance["ordered"]) + (2 if instance["multilevel"] else 0)
        ET.SubElement(num, _q(W, "abstractNumId"), {_q(W, "val"): str(abstract_id)})
    return ET.tostring(numbering, encoding="UTF-8", xml_declaration=True)


def _table(parent: ET.Element, block: dict, next_drawing_id: int) -> list[dict]:
    table = ET.SubElement(parent, _q(W, "tbl"))
    props = ET.SubElement(table, _q(W, "tblPr"))
    layout = block["format"]
    total = sum(block["widths_twips"])
    ET.SubElement(props, _q(W, "tblW"), {_q(W, "w"): str(total), _q(W, "type"): "dxa"})
    ET.SubElement(props, _q(W, "jc"), {_q(W, "val"): layout["align"]})
    borders = ET.SubElement(props, _q(W, "tblBorders"))
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        ET.SubElement(borders, _q(W, edge), _border_attributes(_resolved_border(layout, edge)))
    if layout["cell_fill"] != "none":
        ET.SubElement(props, _q(W, "shd"), {
            _q(W, "val"): "clear", _q(W, "fill"): layout["cell_fill"],
        })
    ET.SubElement(props, _q(W, "tblLayout"), {_q(W, "type"): "fixed"})
    cell_margins = ET.SubElement(props, _q(W, "tblCellMar"))
    for side in ("top", "left", "bottom", "right"):
        ET.SubElement(cell_margins, _q(W, side), {
            _q(W, "w"): str(_twips(layout[f"cell_margin_{side}_mm"])),
            _q(W, "type"): "dxa",
        })
    grid = ET.SubElement(table, _q(W, "tblGrid"))
    for width in block["widths_twips"]:
        ET.SubElement(grid, _q(W, "gridCol"), {_q(W, "w"): str(width)})
    images = []
    for row_index, values in enumerate(block["rows"]):
        row = ET.SubElement(table, _q(W, "tr"))
        is_header = row_index < block["header_rows"]
        repeat_header = is_header and layout["repeat_header"]
        if not layout["allow_row_split"] or repeat_header:
            row_props = ET.SubElement(row, _q(W, "trPr"))
            if not layout["allow_row_split"]:
                ET.SubElement(row_props, _q(W, "cantSplit"), {_q(W, "val"): "true"})
            if repeat_header:
                ET.SubElement(row_props, _q(W, "tblHeader"), {_q(W, "val"): "true"})
        for content, width in zip(values, block["widths_twips"]):
            cell = ET.SubElement(row, _q(W, "tc"))
            cell_props = ET.SubElement(cell, _q(W, "tcPr"))
            ET.SubElement(cell_props, _q(W, "tcW"), {
                _q(W, "w"): str(width), _q(W, "type"): "dxa",
            })
            if isinstance(content, dict) and content.get("borders"):
                cell_edges = ET.SubElement(cell_props, _q(W, "tcBorders"))
                for edge in ("top", "left", "bottom", "right"):
                    if edge in content["borders"]:
                        ET.SubElement(cell_edges, _q(W, edge), _border_attributes(content["borders"][edge]))
            if not is_header and (row_index - block["header_rows"]) % 2 == 1 and layout["band_fill"] != "none":
                ET.SubElement(cell_props, _q(W, "shd"), {
                    _q(W, "val"): "clear", _q(W, "fill"): layout["band_fill"],
                })
            ET.SubElement(cell_props, _q(W, "vAlign"), {_q(W, "val"): layout["vertical_align"]})
            style_name = "TableHeader" if is_header else "TableBody"
            text_format = block["header_format"] if is_header else block["body_format"]
            text = content if isinstance(content, str) else content["text"]
            paragraph = _paragraph(cell, text=text, style=style_name, formatting=text_format)
            if isinstance(content, dict) and content["image"] is not None:
                image = content["image"]
                _add_picture(paragraph, image, next_drawing_id)
                images.append(image)
                next_drawing_id += 1
    return images


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


def _related_text_xml(root_name: str, text: str, *, page_numbers: bool = False,
                      style: str | None = None) -> bytes:
    root = ET.Element(_q(W, root_name))
    paragraph = _paragraph(root, style=style)
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


def _document_xml(spec: dict) -> tuple[bytes, list[dict], list[dict]]:
    root = ET.Element(_q(W, "document"))
    body = ET.SubElement(root, _q(W, "body"))
    instances: list[dict] = []
    image_blocks: list[dict] = []
    next_num_id = 1
    next_drawing_id = 1
    shared_ids = {}
    for block in spec["blocks"]:
        kind = block["type"]
        if kind in {"title", "subtitle"}:
            _paragraph(body, text=block["text"], style="Title" if kind == "title" else "Subtitle",
                       formatting=block["format"])
        elif kind == "heading":
            formatting = block["format"].copy()
            if block["align"] is not None:
                formatting["align"] = block["align"]
            _paragraph(body, text=block["text"], style=f"Heading{block['level']}",
                       formatting=formatting)
        elif kind == "paragraph":
            formatting = block["format"].copy()
            if block["align"] is not None:
                formatting["align"] = block["align"]
            if isinstance(block["content"], str):
                _paragraph(body, text=block["content"], style="BodyText", formatting=formatting)
            else:
                _paragraph(body, runs=block["content"], style="BodyText", formatting=formatting)
        elif kind == "list":
            identifier = block["list_id"]
            if identifier is None or identifier not in shared_ids:
                block["num_id"] = next_num_id
                instances.append({"ordered": block["ordered"], "multilevel": identifier is not None})
                if identifier is not None:
                    shared_ids[identifier] = next_num_id
                next_num_id += 1
            else:
                block["num_id"] = shared_ids[identifier]
            for item in block["items"]:
                _paragraph(body, text=item, num_id=block["num_id"], list_level=block["level"],
                           style="ListParagraph" if spec["list_format"] is not None else None)
        elif kind == "table":
            images = _table(body, block, next_drawing_id)
            image_blocks.extend(images)
            next_drawing_id += len(images)
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
    page = spec["page_twips"]
    width, height = page["width"], page["height"]
    size_attrs = {_q(W, "w"): str(width), _q(W, "h"): str(height)}
    if width > height:
        size_attrs[_q(W, "orient")] = "landscape"
    ET.SubElement(section, _q(W, "pgSz"), size_attrs)
    ET.SubElement(section, _q(W, "pgMar"), {
        _q(W, "top"): str(page["top"]), _q(W, "right"): str(page["right"]),
        _q(W, "bottom"): str(page["bottom"]), _q(W, "left"): str(page["left"]),
        _q(W, "header"): str(min(720, page["top"])), _q(W, "footer"): str(min(720, page["bottom"])),
        _q(W, "gutter"): "0",
    })
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True), image_blocks, instances


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
    media = {}

    def add_relationship(rel_type: str, target: str) -> str:
        rel_id = f"rId{len(relationships) + 1}"
        relationships.append((rel_id, _REL_TYPES[rel_type], target))
        return rel_id

    def add_image(image: dict, name: str) -> None:
        digest = image["image_digest"]
        if spec["deduplicate_images"] and digest in media:
            image["media_name"], image["rid"] = media[digest]
        else:
            image["media_name"] = name
            image["rid"] = add_relationship("image", f"media/{name}")
            media[digest] = (name, image["rid"])

    add_relationship("styles", "styles.xml")
    add_relationship("numbering", "numbering.xml")
    if spec["header"] is not None:
        spec["header_rid"] = add_relationship("header", "header1.xml")
    if spec["footer"] is not None or spec["page_numbers"]:
        spec["footer_rid"] = add_relationship("footer", "footer1.xml")

    for index, block in enumerate(spec["blocks"], start=1):
        if block["type"] == "image":
            name = f"image{index}.{block['extension']}"
            add_image(block, name)
        elif block["type"] == "table":
            for row_index, row in enumerate(block["rows"], 1):
                for column_index, cell in enumerate(row, 1):
                    if not isinstance(cell, dict) or cell["image"] is None:
                        continue
                    image = cell["image"]
                    name = f"image{index}_r{row_index}_c{column_index}.{image['extension']}"
                    add_image(image, name)

    document_xml, image_blocks, instances = _document_xml(spec)
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
        "word/styles.xml": _styles_xml(spec["styles"], spec["related_format"], spec["list_format"]),
        "word/numbering.xml": _numbering_xml(instances),
    }
    if spec["header"] is not None:
        parts["word/header1.xml"] = _related_text_xml(
            "hdr", spec["header"], style="Header" if spec["related_format"] is not None else None,
        )
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
            style="Footer" if spec["related_format"] is not None else None,
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


def _check_generated_budget(parts: dict[str, bytes]) -> None:
    budget = xmlio.XmlBudget()
    for name, data in parts.items():
        if name.endswith((".xml", ".rels")):
            try:
                xmlio.fromstring(data, budget=budget)
            except ValueError as exc:
                raise ValueError(f"generated package {name}: {exc}") from exc


def preflight_document(spec: dict, base_dir: str | Path | None = None) -> dict:
    """Check the full specification and generated XML budgets without writing."""
    validated = _validate_spec(spec, Path(base_dir) if base_dir is not None else Path.cwd())
    parts = _package_parts(validated)
    _check_generated_budget(parts)
    return {"status": "preflight-ok", "blocks": len(validated["blocks"]),
            "parts": len(parts), "uncompressed_bytes": sum(map(len, parts.values())),
            "visual_validation": False, "xsd_validation": False}


def _verify_content(spec: dict, roots: dict) -> None:
    expected = []
    images = tables = 0
    for block in spec["blocks"]:
        kind = block["type"]
        if kind in {"title", "subtitle", "heading"}:
            expected.append(block["text"])
        elif kind == "paragraph":
            expected.extend([block["content"]] if isinstance(block["content"], str)
                            else [run["text"] for run in block["content"]])
        elif kind == "list":
            expected.extend(block["items"])
        elif kind == "table":
            tables += 1
            for row in block["rows"]:
                for cell in row:
                    expected.append(cell if isinstance(cell, str) else cell["text"])
                    images += int(isinstance(cell, dict) and cell["image"] is not None)
        elif kind == "image":
            images += 1
    document = roots["word/document.xml"]
    actual = [node.text or "" for node in document.iter(_q(W, "t"))]
    if actual != expected:
        raise ValueError("Generated body text does not match input content coverage")
    if len(list(document.iter(_q(W, "tbl")))) != tables or len(list(document.iter(_q(W, "drawing")))) != images:
        raise ValueError("Generated table/image counts do not match input content coverage")
    for key, name in (("header", "word/header1.xml"), ("footer", "word/footer1.xml")):
        if name in roots:
            actual_text = "".join(node.text or "" for node in roots[name].iter(_q(W, "t")))
            expected_text = (spec[key] or "") + (" 1" if key == "footer" and spec["page_numbers"] else "")
            if actual_text != expected_text:
                raise ValueError(f"Generated {key} text does not match input content coverage")


def create_document(spec: dict, output: str | Path, base_dir: str | Path | None = None, *,
                    verify: bool = False, require_no_placeholders: bool = False) -> Path:
    """Create a DOCX at output, resolving relative image paths under base_dir."""
    output_path = Path(output)
    if not isinstance(verify, bool) or not isinstance(require_no_placeholders, bool):
        raise ValueError("verify and require_no_placeholders must be booleans")
    if output_path.suffix.lower() != ".docx":
        raise ValueError("output path must use the .docx extension")
    image_base = Path(base_dir) if base_dir is not None else Path.cwd()
    validated = _validate_spec(spec, image_base)
    resolved_output = output_path.resolve()
    if resolved_output in validated["image_paths"]:
        raise ValueError("output must not overwrite an image used by the document")
    parts = _package_parts(validated)
    _check_generated_budget(parts)

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
        if output_path.exists():
            mode = stat.S_IMODE(output_path.stat().st_mode)
        else:
            umask = os.umask(0)
            os.umask(umask)
            mode = 0o666 & ~umask
        os.chmod(temporary_name, mode)
        if verify or require_no_placeholders:
            from office.archiveio import open_archive
            from office.validators.structure import StructureValidator
            from office.validators.generation import verify_generation

            with open_archive(Path(temporary_name)) as archive:
                validator = StructureValidator.from_archive(
                    archive, extension=".docx", require_no_placeholders=require_no_placeholders,
                )
                if not validator.validate(quiet=True) or validator.warnings:
                    raise ValueError("Generated document failed verification: "
                                     + "; ".join((validator.errors + validator.warnings)[:5]))
                _verify_content(validated, validator.roots)
                verify_generation(validated, validator)
        os.replace(temporary_name, output_path)
    except BaseException:
        try:
            os.unlink(temporary_name)
        except FileNotFoundError:
            pass
        raise
    return output_path


def _runtime_check() -> bool:
    checks_passed = True
    version = ".".join(map(str, sys.version_info[:3]))
    print(f"python={version} required>=3.10 supported={sys.version_info >= (3, 10)}")
    if sys.version_info < (3, 10):
        checks_passed = False
    try:
        payload = b"docx-runtime-check"
        if zlib.decompress(zlib.compress(payload)) != payload:
            checks_passed = False
            print("zlib=failed")
        else:
            print("zlib=ok")
    except Exception:
        checks_passed = False
        print("zlib=failed")
    try:
        match = re.search(r"(\d+)\.(\d+)\.(\d+)", expat.EXPAT_VERSION)
        expat_version = tuple(map(int, match.groups())) if match else (0, 0, 0)
        if expat_version < (2, 6, 0):
            checks_passed = False
        print(f"expat={expat.EXPAT_VERSION} required>=2.6.0 supported={expat_version >= (2, 6, 0)}")
        ET.fromstring(b"<runtime-check><ok/></runtime-check>")
    except Exception:
        checks_passed = False

    writable = False
    try:
        with tempfile.TemporaryFile(mode="w+b", dir=".") as probe:
            probe.write(payload)
            probe.seek(0)
            writable = probe.read() == payload
    except Exception:
        writable = False
    print(f"output_dir=. writable={writable}")
    if not checks_passed or not writable:
        print("check-failed: runtime environment check failed", file=sys.stderr)
        return False
    print("check-ok")
    return True


def _json_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"Duplicate JSON field {key!r}; keep exactly one value")
        result[key] = value
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, nargs="?", help="JSON document specification")
    parser.add_argument("output", type=Path, nargs="?", help="output DOCX path")
    parser.add_argument("--check", action="store_true", help="check runtime support and current-directory write access")
    parser.add_argument("--preflight", action="store_true", help="validate input and resource budgets without writing")
    parser.add_argument("--verify", action="store_true", help="check ZIP CRC, structure and input coverage before publishing")
    parser.add_argument("--require-no-placeholders", action="store_true", help="also verify and reject remaining {{...}} markers")
    args = parser.parse_args(argv)
    if args.check:
        if args.input is not None or args.output is not None or args.preflight or args.verify or args.require_no_placeholders:
            parser.error("--check does not accept positional arguments or generation options")
        return 0 if _runtime_check() else 1
    if args.preflight and (args.output is not None or args.verify or args.require_no_placeholders):
        parser.error("--preflight accepts input only, without output or verification options")
    if args.input is None or (args.output is None and not args.preflight):
        parser.error("input is required; output is required unless --check or --preflight is used")
    try:
        if args.output is not None and args.input.resolve() == args.output.resolve():
            parser.error("output must not refer to the input JSON file")
        spec = json.loads(args.input.read_text(encoding="utf-8"), object_pairs_hook=_json_object)
        if args.preflight:
            print(json.dumps(preflight_document(spec, args.input.resolve().parent), ensure_ascii=False))
        else:
            verified = args.verify or args.require_no_placeholders
            create_document(spec, args.output, base_dir=args.input.resolve().parent,
                            verify=verified, require_no_placeholders=args.require_no_placeholders)
            if verified:
                print(json.dumps({"status": "verified", "output": args.output.as_posix(),
                                  "blocks": len(spec["blocks"]), "zip_crc": True,
                                  "structure": True, "content_coverage": True,
                                  "layout_intent": True,
                                  "visual_validation": False, "xsd_validation": False}, ensure_ascii=False))
    except (OSError, UnicodeError, ValueError, RuntimeError, zipfile.BadZipFile) as exc:
        parser.error(str(exc))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


