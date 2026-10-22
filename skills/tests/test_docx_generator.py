"""Focused tests for the stdlib DOCX generator and its JSON CLI."""

import json
import os
import stat
import struct
import subprocess
import sys
import tempfile
import unittest
import zipfile
import zlib
from pathlib import Path
from unittest import mock
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "docx" / "scripts" / "create_docx.py"
PYTHON = Path(sys.executable)
sys.path.insert(0, str(SCRIPT.parent))
import create_docx  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def q(namespace, local):
    return f"{{{namespace}}}{local}"


def png_chunk(kind, payload):
    body = kind + payload
    return struct.pack(">I", len(payload)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)


def tiny_png():
    return (
        b"\x89PNG\r\n\x1a\n"
        + png_chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 6, 0, 0, 0))
        + png_chunk(b"IDAT", zlib.compress(b"\x00\x20\x80\xe0\xff"))
        + png_chunk(b"IEND", b"")
    )


def xml_part(path, name):
    with zipfile.ZipFile(path) as package:
        return ET.fromstring(package.read(name))


def style_by_id(styles, style_id):
    return next(style for style in styles.findall(q(W, "style"))
                if style.get(q(W, "styleId")) == style_id)


class DocxGeneratorTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-generator-")
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name)
        self.env = dict(os.environ)
        self.env["PATH"] = ""
        self.env["PYTHONDONTWRITEBYTECODE"] = "1"

    def run_cli(self, input_path, output_path):
        return subprocess.run(
            [str(PYTHON), "-I", "-S", "-B", str(SCRIPT), str(input_path), str(output_path)],
            cwd=self.work, env=self.env, capture_output=True, text=True,
        )

    def test_cli_writes_content_styles_relationships_and_media(self):
        source_dir = self.work / "source"
        source_dir.mkdir()
        image = source_dir / "pixel.png"
        image.write_bytes(tiny_png())
        spec = {
            "page": {"width_mm": 210, "height_mm": 297, "margin_mm": 25},
            "header": "页眉 & 注释",
            "footer": "内部资料",
            "page_numbers": True,
            "blocks": [
                {"type": "heading", "text": "季度报告", "level": 1},
                {"type": "paragraph", "align": "center", "runs": [
                    {"text": "已完成：", "bold": True},
                    {"text": "验证通过", "italic": True, "underline": True,
                     "font": "Arial", "size_pt": 12, "color": "00AA00"},
                ]},
                {"type": "list", "ordered": True, "items": ["第一项", "第二项"]},
                {"type": "list", "ordered": False, "items": ["补充项"]},
                {"type": "table", "header": True, "widths_mm": [80, 80],
                 "rows": [["项目", "状态"], ["校验", "通过"]]},
                {"type": "image", "path": "pixel.png", "width_mm": 30,
                 "height_mm": 20, "alt": "一个像素"},
                {"type": "page_break"},
            ],
        }
        spec_path = source_dir / "spec.json"
        spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
        output = self.work / "generated.docx"
        result = self.run_cli(spec_path, output)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

        with zipfile.ZipFile(output) as package:
            self.assertIsNone(package.testzip())
            names = set(package.namelist())
            required = {
                "[Content_Types].xml", "_rels/.rels", "word/document.xml",
                "word/_rels/document.xml.rels", "word/styles.xml", "word/numbering.xml",
                "word/header1.xml", "word/footer1.xml", "word/media/image6.png",
            }
            self.assertTrue(required <= names)
            roots = {}
            for name in names:
                if name.endswith((".xml", ".rels")) or name == "[Content_Types].xml":
                    roots[name] = ET.fromstring(package.read(name))
            document = roots["word/document.xml"]
            texts = [node.text or "" for node in document.iter(q(W, "t"))]
            self.assertIn("季度报告", texts)
            self.assertIn("已完成：", texts)
            self.assertIn("验证通过", texts)
            self.assertIn("校验", texts)
            self.assertEqual(len(list(document.iter(q(W, "tbl")))), 1)
            self.assertEqual(len(list(document.iter(q(W, "drawing")))), 1)
            self.assertEqual(len(list(document.iter(q(W, "br")))), 1)
            self.assertEqual(len(list(document.iter(q(W, "numPr")))), 3)
            body = document.find(q(W, "body"))
            self.assertFalse(any(child.tag == q(W, "r") for child in body))
            image_paragraph = next(
                paragraph for paragraph in document.iter(q(W, "p"))
                if paragraph.find(f"./{q(W, 'r')}/{q(W, 'drawing')}") is not None
            )
            self.assertEqual(image_paragraph.find(q(W, "r")) is not None, True)
            formatted_run = next(
                run for run in document.iter(q(W, "r"))
                if any(node.text == "验证通过" for node in run.findall(q(W, "t")))
            )
            run_props = formatted_run.find(q(W, "rPr"))
            self.assertEqual([node.tag.rsplit("}", 1)[-1] for node in run_props],
                             ["rFonts", "i", "color", "sz", "szCs", "u"])
            self.assertEqual(run_props.find(q(W, "rFonts")).get(q(W, "ascii")), "Arial")
            self.assertEqual(run_props.find(q(W, "rFonts")).get(q(W, "hAnsi")), "Arial")
            self.assertEqual(run_props.find(q(W, "rFonts")).get(q(W, "eastAsia")), "Arial")
            self.assertEqual(run_props.find(q(W, "rFonts")).get(q(W, "cs")), "Arial")
            self.assertTrue(any(True for _ in document.iter(q(W, "b"))))
            table_props = document.find(f".//{q(W, 'tblPr')}")
            self.assertEqual([node.tag.rsplit("}", 1)[-1] for node in table_props],
                             ["tblW", "jc", "tblBorders", "tblLayout", "tblCellMar"])
            page_size = document.find(f".//{q(W, 'pgSz')}")
            self.assertEqual(page_size.get(q(W, "w")), "11906")
            self.assertEqual(page_size.get(q(W, "h")), "16838")
            page_margins = document.find(f".//{q(W, 'pgMar')}")
            self.assertEqual([page_margins.get(q(W, side)) for side in ("top", "right", "bottom", "left")],
                             ["1417"] * 4)

            styles = roots["word/styles.xml"]
            fonts = styles.find(f".//{q(W, 'docDefaults')}/{q(W, 'rPrDefault')}/{q(W, 'rPr')}/{q(W, 'rFonts')}")
            self.assertEqual(fonts.get(q(W, "ascii")), "Times New Roman")
            self.assertEqual(fonts.get(q(W, "eastAsia")), "仿宋")
            self.assertTrue(any(style.get(q(W, "styleId")) == "CommentReference"
                                for style in styles.findall(q(W, "style"))))
            numbering = roots["word/numbering.xml"]
            formats = [node.get(q(W, "val")) for node in numbering.findall(f".//{q(W, 'numFmt')}")]
            self.assertEqual(formats, ["bullet", "decimal"])

            header_text = "".join(node.text or "" for node in roots["word/header1.xml"].iter(q(W, "t")))
            footer = roots["word/footer1.xml"]
            footer_text = "".join(node.text or "" for node in footer.iter(q(W, "t")))
            self.assertEqual(header_text, "页眉 & 注释")
            self.assertIn("内部资料", footer_text)
            self.assertEqual([node.text for node in footer.iter(q(W, "instrText"))], [" PAGE "])

            relationships = roots["word/_rels/document.xml.rels"]
            relationship_map = {rel.get("Id"): rel.get("Target") for rel in relationships}
            for node in document.iter():
                rel_id = node.get(q(R, "id"))
                if rel_id:
                    self.assertIn(rel_id, relationship_map)
                    self.assertIn(f"word/{relationship_map[rel_id]}", names)
            content_types = roots["[Content_Types].xml"]
            self.assertTrue(any(node.get("Extension") == "png" and node.get("ContentType") == "image/png"
                                for node in content_types.findall(q(CT, "Default"))))
            self.assertEqual(package.read("word/media/image6.png"), image.read_bytes())

    def test_api_rejects_bad_specs_without_replacing_existing_output(self):
        output = self.work / "preserved.docx"
        original = b"existing output must survive"
        output.write_bytes(original)
        invalid_specs = (
            {"extra": True, "blocks": []},
            {"page": {"width_mm": float("inf")}, "blocks": []},
            {"blocks": [{"type": "unsupported"}]},
            {"blocks": [{"type": "paragraph", "text": "bad\x01text"}]},
            {"blocks": [{"type": "page_break", "extra": 1}]},
        )
        for spec in invalid_specs:
            with self.subTest(spec=spec):
                with self.assertRaises(ValueError):
                    create_docx.create_document(spec, output)
                self.assertEqual(output.read_bytes(), original)

        missing_image = {"blocks": [{
            "type": "image", "path": "missing.png", "width_mm": 10,
            "height_mm": 10, "alt": "missing",
        }]}
        with self.assertRaises(OSError):
            create_docx.create_document(missing_image, output, base_dir=self.work)
        self.assertEqual(output.read_bytes(), original)

        oversized = self.work / "oversized.png"
        with oversized.open("wb") as image_file:
            image_file.truncate(create_docx._MAX_IMAGE_BYTES + 1)
        large_image = {"blocks": [{
            "type": "image", "path": "oversized.png", "width_mm": 10,
            "height_mm": 10, "alt": "oversized",
        }]}
        with self.assertRaisesRegex(ValueError, "32 MiB"):
            create_docx.create_document(large_image, output, base_dir=self.work)
        self.assertEqual(output.read_bytes(), original)

        for suffix in (".dotx", ".pdf"):
            alternate = self.work / f"preserved{suffix}"
            alternate.write_bytes(original)
            with self.assertRaisesRegex(ValueError, r"\.docx"):
                create_docx.create_document({"blocks": []}, alternate)
            self.assertEqual(alternate.read_bytes(), original)

    def test_official_styles_and_bodytext_are_serialized_and_applied(self):
        output = self.work / "official.docx"
        create_docx.create_document({"blocks": [
            {"type": "heading", "text": "第一层", "level": 1},
            {"type": "heading", "text": "第二层", "level": 2},
            {"type": "heading", "text": "第三层", "level": 3},
            {"type": "heading", "text": "第四层", "level": 4},
            {"type": "paragraph", "text": "正文首行缩进"},
        ]}, output)
        styles = xml_part(output, "word/styles.xml")

        heading_defaults = {
            1: ("方正小标宋简体", "36", "0", "center"),
            2: ("黑体", "32", "0", "left"),
            3: ("楷体", "32", "1", "left"),
            4: ("仿宋", "32", "0", "left"),
        }
        for level, (east_asia, size, bold, alignment) in heading_defaults.items():
            with self.subTest(level=level):
                style = style_by_id(styles, f"Heading{level}")
                rpr = style.find(q(W, "rPr"))
                fonts = rpr.find(q(W, "rFonts"))
                self.assertEqual(fonts.get(q(W, "eastAsia")), east_asia)
                self.assertEqual(fonts.get(q(W, "ascii")), "Times New Roman")
                self.assertEqual(rpr.find(q(W, "sz")).get(q(W, "val")), size)
                self.assertEqual(rpr.find(q(W, "b")).get(q(W, "val")), bold)
                ppr = style.find(q(W, "pPr"))
                self.assertEqual(ppr.find(q(W, "jc")).get(q(W, "val")), alignment)
                self.assertEqual(ppr.find(q(W, "outlineLvl")).get(q(W, "val")), str(level - 1))
                self.assertEqual(ppr.find(q(W, "ind")).get(q(W, "firstLineChars")), "0")
                self.assertIsNotNone(ppr.find(q(W, "keepNext")))
                self.assertIsNone(ppr.find(q(W, "spacing")).get(q(W, "lineRule")))

        body_style = style_by_id(styles, "BodyText")
        body_ppr = body_style.find(q(W, "pPr"))
        self.assertEqual(body_ppr.find(q(W, "spacing")).get(q(W, "line")), "560")
        self.assertEqual(body_ppr.find(q(W, "spacing")).get(q(W, "lineRule")), "exact")
        self.assertEqual(body_ppr.find(q(W, "ind")).get(q(W, "firstLineChars")), "200")
        body_rpr = body_style.find(q(W, "rPr"))
        self.assertEqual(body_rpr.find(q(W, "rFonts")).get(q(W, "eastAsia")), "仿宋")
        self.assertEqual(body_rpr.find(q(W, "sz")).get(q(W, "val")), "32")

        defaults = styles.find(q(W, "docDefaults"))
        normal_rpr = defaults.find(f"./{q(W, 'rPrDefault')}/{q(W, 'rPr')}")
        self.assertEqual(normal_rpr.find(q(W, "rFonts")).get(q(W, "ascii")), "Times New Roman")
        self.assertEqual(normal_rpr.find(q(W, "rFonts")).get(q(W, "eastAsia")), "仿宋")
        self.assertEqual(normal_rpr.find(q(W, "sz")).get(q(W, "val")), "32")
        self.assertEqual(normal_rpr.find(q(W, "b")).get(q(W, "val")), "0")
        default_ppr = defaults.find(f"./{q(W, 'pPrDefault')}/{q(W, 'pPr')}")
        self.assertEqual(default_ppr.find(q(W, "spacing")).get(q(W, "line")), "276")
        self.assertEqual(default_ppr.find(q(W, "spacing")).get(q(W, "lineRule")), "auto")
        self.assertEqual(default_ppr.find(q(W, "spacing")).get(q(W, "before")), "0")
        self.assertEqual(default_ppr.find(q(W, "spacing")).get(q(W, "after")), "0")
        self.assertIsNone(default_ppr.find(q(W, "ind")))
        self.assertIsNone(style_by_id(styles, "Normal").find(f"./{q(W, 'pPr')}/{q(W, 'ind')}"))

        document = xml_part(output, "word/document.xml")
        paragraphs = document.findall(f"./{q(W, 'body')}/{q(W, 'p')}")
        self.assertEqual([p.find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}").get(q(W, "val"))
                          for p in paragraphs],
                         ["Heading1", "Heading2", "Heading3", "Heading4", "BodyText"])
        body_direct_ppr = paragraphs[-1].find(q(W, "pPr"))
        self.assertEqual([node.tag for node in body_direct_ppr], [q(W, "pStyle")])
        for paragraph in paragraphs:
            self.assertIsNone(paragraph.find(f"./{q(W, 'r')}/{q(W, 'rPr')}"))
            self.assertIsNone(paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'spacing')}"))
            self.assertIsNone(paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'ind')}"))

    def test_global_styles_leave_only_explicit_local_overrides_in_document(self):
        output = self.work / "inherited.docx"
        create_docx.create_document({
            "styles": {
                "body": {"font_east_asia": "楷体", "size_pt": 15, "bold": True},
                "heading3": {"font_latin": "Calibri", "size_pt": 20},
            },
            "blocks": [
                {"type": "paragraph", "runs": [{"text": "继承"},
                 {"text": "局部", "bold": False, "font_east_asia": "宋体"}]},
                {"type": "heading", "level": 3, "text": "标题", "format": {"align": "right"}},
                {"type": "paragraph", "text": "等值覆盖", "format": {"size_pt": 15}},
            ],
        }, output)
        document = xml_part(output, "word/document.xml")
        paragraphs = document.findall(f"./{q(W, 'body')}/{q(W, 'p')}")
        runs = paragraphs[0].findall(q(W, "r"))
        self.assertIsNone(runs[0].find(q(W, "rPr")))
        local = runs[1].find(q(W, "rPr"))
        self.assertEqual(local.find(q(W, "b")).get(q(W, "val")), "0")
        self.assertEqual(local.find(q(W, "rFonts")).attrib, {q(W, "eastAsia"): "宋体"})
        self.assertIsNone(local.find(q(W, "sz")))
        self.assertEqual(paragraphs[1].find(f"./{q(W, 'pPr')}/{q(W, 'jc')}").get(q(W, "val")), "right")
        self.assertIsNone(paragraphs[1].find(f"./{q(W, 'r')}/{q(W, 'rPr')}"))
        self.assertEqual(paragraphs[2].find(f".//{q(W, 'sz')}").get(q(W, "val")), "30")
        styles = xml_part(output, "word/styles.xml")
        self.assertEqual(style_by_id(styles, "BodyText").find(f".//{q(W, 'sz')}").get(q(W, "val")), "30")
        self.assertEqual(style_by_id(styles, "Heading3").find(f".//{q(W, 'sz')}").get(q(W, "val")), "40")

    @unittest.skipUnless(os.name == "posix", "POSIX file mode semantics")
    def test_generation_uses_umask_for_new_files_and_preserves_existing_modes(self):
        spec = self.work / "mode.json"
        spec.write_text('{"blocks": []}', encoding="utf-8")
        for mask, expected in ((0o022, 0o644), (0o077, 0o600)):
            with self.subTest(umask=oct(mask)):
                output = self.work / f"mode-{mask}.docx"
                result = subprocess.run([str(PYTHON), "-I", "-S", "-B", str(SCRIPT),
                                         str(spec), str(output)], env=self.env, cwd=self.work,
                                        umask=mask, capture_output=True, text=True)
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertEqual(stat.S_IMODE(output.stat().st_mode), expected)
        for mode in (0o640, 0o600):
            output.chmod(mode)
            create_docx.create_document({"blocks": []}, output)
            self.assertEqual(stat.S_IMODE(output.stat().st_mode), mode)

    def test_permission_failure_preserves_output_and_cleans_staging_file(self):
        output = self.work / "preserved.docx"
        output.write_bytes(b"keep existing output")
        before = set(self.work.iterdir())
        with mock.patch.object(create_docx.os, "chmod", side_effect=OSError("permission failure")):
            with self.assertRaises(OSError):
                create_docx.create_document({"blocks": []}, output)
        self.assertEqual(output.read_bytes(), b"keep existing output")
        self.assertEqual(set(self.work.iterdir()), before)

    def test_image_diagnostics_use_field_location_and_preserve_alt_text(self):
        image = self.work / "invalid.png"
        image.write_bytes(b"not image data")
        output = self.work / "image.docx"
        for path, alt, expected in (("invalid.png", "示意图", r"blocks\[0\]\.path"),
                                    (123, "", r"blocks\[0\]\.path"),
                                    ("invalid.png", None, r"blocks\[0\]\.alt")):
            with self.subTest(path=path, alt=alt):
                with self.assertRaisesRegex(ValueError, expected):
                    create_docx.create_document({"blocks": [{"type": "image", "path": path,
                        "width_mm": 10, "height_mm": 10, "alt": alt}]}, output, self.work)
        image.write_bytes(tiny_png())
        for alt in ("", "示意图 & 中文"):
            create_docx.create_document({"blocks": [{"type": "image", "path": image.name,
                "width_mm": 10, "height_mm": 10, "alt": alt}]}, output, self.work)
            document = xml_part(output, "word/document.xml")
            for local_name in ("docPr", "cNvPr"):
                node = next(node for node in document.iter() if node.tag.endswith("}" + local_name))
                self.assertEqual(node.get("descr"), alt)

    def test_default_margins_drive_table_width_and_other_paragraphs_keep_normal_inheritance(self):
        output = self.work / "paragraph-inheritance.docx"
        create_docx.create_document({
            "header": "页眉",
            "footer": "页脚",
            "blocks": [
                {"type": "paragraph", "text": "普通正文"},
                {"type": "list", "ordered": True, "items": ["列表项"]},
                {"type": "table", "rows": [["表格单元"]]},
            ],
        }, output)
        document = xml_part(output, "word/document.xml")
        section = document.find(f".//{q(W, 'sectPr')}")
        margins = section.find(q(W, "pgMar"))
        self.assertEqual([margins.get(q(W, side)) for side in ("top", "right", "bottom", "left")],
                         ["1440", "1803", "1440", "1803"])
        self.assertEqual(margins.get(q(W, "header")), "720")
        self.assertEqual(margins.get(q(W, "footer")), "720")
        table_width = document.find(f".//{q(W, 'tblW')}").get(q(W, "w"))
        self.assertEqual(table_width, str(11906 - 2 * 1803))

        paragraphs = list(document.iter(q(W, "p")))
        body_text_paragraphs = [paragraph for paragraph in paragraphs
                                if paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}") is not None
                                and paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}").get(q(W, "val"))
                                == "BodyText"]
        self.assertEqual(len(body_text_paragraphs), 1)
        self.assertEqual(body_text_paragraphs[0].find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}").get(q(W, "val")),
                         "BodyText")
        list_paragraph = next(paragraph for paragraph in paragraphs
                              if paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'numPr')}") is not None)
        self.assertIsNone(list_paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}") )
        self.assertIsNone(list_paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'ind')}"))
        numbering = xml_part(output, "word/numbering.xml")
        numbering_indent = numbering.find(f".//{q(W, 'lvl')}/{q(W, 'pPr')}/{q(W, 'ind')}")
        self.assertEqual(numbering_indent.get(q(W, "left")), "720")
        self.assertEqual(numbering_indent.get(q(W, "hanging")), "360")
        table_paragraph = document.find(f".//{q(W, 'tc')}/{q(W, 'p')}")
        self.assertEqual(table_paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}").get(q(W, "val")),
                         "TableBody")
        self.assertIsNone(table_paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'spacing')}"))
        table_style = style_by_id(xml_part(output, "word/styles.xml"), "TableBody")
        self.assertEqual(table_style.find(f"./{q(W, 'pPr')}/{q(W, 'spacing')}").get(q(W, "lineRule")),
                         "auto")

        header = xml_part(output, "word/header1.xml")
        footer = xml_part(output, "word/footer1.xml")
        for related in (header, footer):
            paragraph = related.find(q(W, "p"))
            self.assertIsNone(paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}"))
            self.assertIsNone(paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'ind')}"))

    def test_global_block_and_run_overrides_preserve_bilingual_text_and_explicit_off_values(self):
        output = self.work / "overrides.docx"
        source = {
            "styles": {
                "body": {"font_east_asia": "楷体", "font_latin": "Calibri", "size_pt": 15,
                         "bold": False, "align": "justify", "first_line_chars": 2,
                         "line_spacing_pt": 28, "space_before_pt": 6, "space_after_pt": 8},
                "heading2": {"font_east_asia": "黑体", "size_pt": 19},
            },
            "blocks": [
                {"type": "paragraph", "align": "center",
                 "format": {"font_east_asia": "仿宋", "font_latin": "Aptos", "size_pt": 14,
                            "bold": True, "align": "right", "first_line_chars": 0,
                            "line_spacing_pt": 26, "space_before_pt": 0, "space_after_pt": 0},
                 "runs": [
                     {"text": " 引号：“你好” and ‘hello’ ", "font": "Arial",
                      "font_east_asia": "宋体", "font_latin": "Courier New", "size_pt": 12,
                      "bold": False, "italic": False, "underline": False},
                     {"text": "后续文字", "font_east_asia": "仿宋"},
                 ]},
                {"type": "heading", "text": "局部标题", "level": 2,
                 "format": {"font_east_asia": "楷体", "size_pt": 20, "bold": True,
                            "align": "right", "first_line_chars": 2, "line_spacing_pt": 30},
                 "align": "left"},
            ],
        }
        original = json.loads(json.dumps(source, ensure_ascii=False))
        create_docx.create_document(source, output)
        self.assertEqual(source, original)
        document = xml_part(output, "word/document.xml")
        paragraphs = document.findall(f"./{q(W, 'body')}/{q(W, 'p')}")
        body_ppr = paragraphs[0].find(q(W, "pPr"))
        self.assertEqual(body_ppr.find(q(W, "jc")).get(q(W, "val")), "center")
        self.assertEqual(body_ppr.find(q(W, "spacing")).get(q(W, "before")), "0")
        self.assertEqual(body_ppr.find(q(W, "spacing")).get(q(W, "after")), "0")
        self.assertEqual(body_ppr.find(q(W, "spacing")).get(q(W, "line")), "520")
        self.assertEqual(body_ppr.find(q(W, "spacing")).get(q(W, "lineRule")), "exact")
        self.assertEqual(body_ppr.find(q(W, "ind")).get(q(W, "firstLineChars")), "0")
        body_runs = paragraphs[0].findall(q(W, "r"))
        first_rpr = body_runs[0].find(q(W, "rPr"))
        first_fonts = first_rpr.find(q(W, "rFonts"))
        self.assertEqual(first_fonts.get(q(W, "eastAsia")), "宋体")
        self.assertEqual(first_fonts.get(q(W, "ascii")), "Courier New")
        self.assertEqual(first_rpr.find(q(W, "sz")).get(q(W, "val")), "24")
        self.assertEqual(first_rpr.find(q(W, "b")).get(q(W, "val")), "0")
        self.assertEqual(first_rpr.find(q(W, "i")).get(q(W, "val")), "0")
        self.assertEqual(first_rpr.find(q(W, "u")).get(q(W, "val")), "none")
        self.assertEqual("".join(node.text or "" for node in paragraphs[0].iter(q(W, "t"))),
                         " 引号：“你好” and ‘hello’ 后续文字")
        second_fonts = body_runs[1].find(f"./{q(W, 'rPr')}/{q(W, 'rFonts')}")
        self.assertEqual(second_fonts.get(q(W, "eastAsia")), "仿宋")
        self.assertEqual(second_fonts.get(q(W, "ascii")), "Aptos")
        self.assertEqual(body_runs[1].find(f"./{q(W, 'rPr')}/{q(W, 'b')}").get(q(W, "val")), "1")

        heading_style = style_by_id(xml_part(output, "word/styles.xml"), "Heading2")
        self.assertEqual(heading_style.find(f"./{q(W, 'rPr')}/{q(W, 'sz')}").get(q(W, "val")), "38")
        heading_ppr = paragraphs[1].find(q(W, "pPr"))
        self.assertEqual(heading_ppr.find(q(W, "jc")).get(q(W, "val")), "left")
        self.assertEqual(heading_ppr.find(q(W, "ind")).get(q(W, "firstLineChars")), "200")
        self.assertEqual(heading_ppr.find(q(W, "spacing")).get(q(W, "line")), "600")
        heading_rpr = paragraphs[1].find(f"./{q(W, 'r')}/{q(W, 'rPr')}")
        self.assertEqual(heading_rpr.find(q(W, "sz")).get(q(W, "val")), "40")
        self.assertEqual(heading_rpr.find(q(W, "b")).get(q(W, "val")), "1")

    def test_asymmetric_margins_override_legacy_and_bound_tables_and_images(self):
        output = self.work / "asymmetric.docx"
        base_spec = {
            "page": {"width_mm": 100, "height_mm": 100, "margin_mm": 10,
                     "margin_top_mm": 5, "margin_bottom_mm": 15,
                     "margin_left_mm": 20, "margin_right_mm": 10},
            "header": "页眉", "footer": "页脚",
            "blocks": [{"type": "table", "rows": [["正好占满内容宽度"]]}],
        }
        create_docx.create_document(base_spec, output)
        document = xml_part(output, "word/document.xml")
        section = document.find(f".//{q(W, 'sectPr')}")
        margins = section.find(q(W, "pgMar"))
        self.assertEqual([margins.get(q(W, side)) for side in ("top", "right", "bottom", "left")],
                         ["283", "567", "850", "1134"])
        self.assertEqual(margins.get(q(W, "header")), "283")
        self.assertEqual(margins.get(q(W, "footer")), "720")
        self.assertEqual(document.find(f".//{q(W, 'tblW')}").get(q(W, "w")),
                         str(create_docx._twips(100) - create_docx._twips(20) - create_docx._twips(10)))

        too_wide = json.loads(json.dumps(base_spec))
        too_wide["blocks"][0]["widths_mm"] = [70.1]
        with self.assertRaisesRegex(ValueError, "content width"):
            create_docx.create_document(too_wide, self.work / "too-wide.docx")

        too_wide_image = {"page": base_spec["page"], "blocks": [{
            "type": "image", "path": "unused.png", "width_mm": 70.1, "height_mm": 10, "alt": "wide",
        }]}
        with self.assertRaisesRegex(ValueError, "content area"):
            create_docx.create_document(too_wide_image, self.work / "wide-image.docx", base_dir=self.work)
        too_tall_image = {"page": base_spec["page"], "blocks": [{
            "type": "image", "path": "unused.png", "width_mm": 10, "height_mm": 80.1, "alt": "tall",
        }]}
        with self.assertRaisesRegex(ValueError, "content area"):
            create_docx.create_document(too_tall_image, self.work / "tall-image.docx", base_dir=self.work)

    def test_default_table_styles_and_word_table_properties(self):
        output = self.work / "table-defaults.docx"
        create_docx.create_document({"blocks": [{
            "type": "table", "header": True,
            "rows": [["项目", "Status"], ["校验", "Ready"]],
        }]}, output)

        styles = xml_part(output, "word/styles.xml")
        for style_id, bold, alignment in (("TableBody", "0", "left"), ("TableHeader", "1", "center")):
            with self.subTest(style=style_id):
                style = style_by_id(styles, style_id)
                self.assertEqual(style.find(q(W, "basedOn")).get(q(W, "val")), "Normal")
                self.assertNotEqual(style.find(q(W, "basedOn")).get(q(W, "val")), "TableBody")
                rpr = style.find(q(W, "rPr"))
                fonts = rpr.find(q(W, "rFonts"))
                self.assertEqual(fonts.get(q(W, "eastAsia")), "仿宋")
                self.assertEqual(fonts.get(q(W, "ascii")), "Times New Roman")
                self.assertEqual(rpr.find(q(W, "sz")).get(q(W, "val")), "32")
                self.assertEqual(rpr.find(q(W, "color")).get(q(W, "val")), "000000")
                self.assertEqual(rpr.find(q(W, "b")).get(q(W, "val")), bold)
                ppr = style.find(q(W, "pPr"))
                self.assertEqual(ppr.find(q(W, "jc")).get(q(W, "val")), alignment)
                self.assertEqual(ppr.find(q(W, "ind")).get(q(W, "firstLineChars")), "0")
                spacing = ppr.find(q(W, "spacing"))
                self.assertEqual([spacing.get(q(W, attr)) for attr in ("before", "after", "line", "lineRule")],
                                 ["0", "0", "276", "auto"])

        document = xml_part(output, "word/document.xml")
        table = document.find(q(W, "body")).find(q(W, "tbl"))
        props = table.find(q(W, "tblPr"))
        self.assertEqual([node.tag.rsplit("}", 1)[-1] for node in props],
                         ["tblW", "jc", "tblBorders", "tblLayout", "tblCellMar"])
        self.assertEqual(props.find(q(W, "jc")).get(q(W, "val")), "left")
        self.assertIsNone(props.find(q(W, "shd")))
        borders = props.find(q(W, "tblBorders"))
        self.assertEqual([node.tag.rsplit("}", 1)[-1] for node in borders],
                         ["top", "left", "bottom", "right", "insideH", "insideV"])
        for edge in borders:
            self.assertEqual(edge.get(q(W, "val")), "single")
            self.assertEqual(edge.get(q(W, "sz")), "4")
            self.assertEqual(edge.get(q(W, "space")), "0")
            self.assertEqual(edge.get(q(W, "color")), "000000")
        margins = props.find(q(W, "tblCellMar"))
        self.assertEqual({node.tag.rsplit("}", 1)[-1]: (node.get(q(W, "w")), node.get(q(W, "type")))
                          for node in margins},
                         {"top": ("57", "dxa"), "bottom": ("57", "dxa"),
                          "left": ("85", "dxa"), "right": ("85", "dxa")})

        rows = table.findall(q(W, "tr"))
        self.assertEqual(len(rows), 2)
        for index, row in enumerate(rows):
            row_props = row.find(q(W, "trPr"))
            self.assertEqual(row_props.find(q(W, "cantSplit")).get(q(W, "val")), "true")
            self.assertIsNone(row_props.find(q(W, "trHeight")))
            self.assertEqual(row_props.find(q(W, "tblHeader")) is not None, index == 0)
            for cell in row.findall(q(W, "tc")):
                cell_props = cell.find(q(W, "tcPr"))
                self.assertEqual([node.tag.rsplit("}", 1)[-1] for node in cell_props], ["tcW", "vAlign"])
                self.assertEqual(cell_props.find(q(W, "vAlign")).get(q(W, "val")), "center")
                paragraph = cell.find(q(W, "p"))
                pstyle = paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}")
                self.assertEqual(pstyle.get(q(W, "val")), "TableHeader" if index == 0 else "TableBody")

    def test_table_header_flag_and_repeat_header_are_independent(self):
        output = self.work / "table-header-flags.docx"
        create_docx.create_document({"blocks": [
            {"type": "table", "rows": [["普通首行"], ["正文"]]},
            {"type": "table", "header": True,
             "format": {"repeat_header": False, "allow_row_split": True},
             "rows": [["视觉表头"], ["正文"]]},
        ]}, output)
        tables = xml_part(output, "word/document.xml").findall(f".//{q(W, 'tbl')}")
        first, second = tables
        self.assertEqual(first.find(f"./{q(W, 'tr')}/{q(W, 'trPr')}/{q(W, 'tblHeader')}"), None)
        self.assertIsNotNone(first.find(f"./{q(W, 'tr')}/{q(W, 'trPr')}/{q(W, 'cantSplit')}"))
        self.assertEqual(first.find(f"./{q(W, 'tr')}/{q(W, 'tc')}/{q(W, 'p')}/{q(W, 'pPr')}/{q(W, 'pStyle')}")
                         .get(q(W, "val")), "TableBody")
        self.assertIsNone(second.find(f"./{q(W, 'tr')}/{q(W, 'trPr')}"))
        self.assertIsNone(second.find(f".//{q(W, 'tblHeader')}"))
        self.assertEqual(second.find(f"./{q(W, 'tr')}/{q(W, 'tc')}/{q(W, 'p')}/{q(W, 'pPr')}/{q(W, 'pStyle')}")
                         .get(q(W, "val")), "TableHeader")

    def test_table_global_and_local_format_precedence_and_false_values(self):
        output = self.work / "table-precedence.docx"
        spec = {
            "table_format": {
                "align": "right", "vertical_align": "bottom", "border_style": "single",
                "border_width_pt": 1, "border_color": "a0b0c0", "cell_fill": "00ff00",
                "cell_margin_top_mm": 2, "cell_margin_bottom_mm": 3,
                "cell_margin_left_mm": 1, "cell_margin_right_mm": 1,
                "repeat_header": True, "allow_row_split": False,
            },
            "styles": {
                "table_body": {"size_pt": 14, "color": "aa11bb", "align": "right",
                               "first_line_chars": 2, "line_spacing_pt": 24},
                "table_header": {"font_east_asia": "楷体", "size_pt": 17,
                                 "color": "123456", "bold": True},
            },
            "blocks": [{
                "type": "table", "header": True, "widths_mm": [100],
                "format": {"align": "center", "vertical_align": "top", "border_style": "none",
                           "border_width_pt": 2, "border_color": "abcdef", "cell_fill": "none",
                           "cell_margin_top_mm": 0, "repeat_header": False, "allow_row_split": True},
                "body_format": {"size_pt": 12, "color": "fFaAbB", "align": "left",
                                "first_line_chars": 0, "line_spacing_multiple": 1.5},
                "header_format": {"font_east_asia": "黑体", "size_pt": 11, "color": "abcdef",
                                  "bold": False, "line_spacing_pt": 18},
                "rows": [["列名"], ["中文 and English"]],
            }],
        }
        original = json.loads(json.dumps(spec, ensure_ascii=False))
        create_docx.create_document(spec, output)
        self.assertEqual(spec, original)

        document = xml_part(output, "word/document.xml")
        table = document.find(f".//{q(W, 'tbl')}")
        props = table.find(q(W, "tblPr"))
        self.assertEqual(props.find(q(W, "jc")).get(q(W, "val")), "center")
        self.assertIsNone(props.find(q(W, "shd")))
        margins = props.find(q(W, "tblCellMar"))
        self.assertEqual(margins.find(q(W, "top")).get(q(W, "w")), "0")
        borders = props.find(q(W, "tblBorders"))
        for edge in borders:
            self.assertEqual(edge.get(q(W, "val")), "nil")
            self.assertIsNone(edge.get(q(W, "sz")))
        for row in table.findall(q(W, "tr")):
            for cell in row.findall(q(W, "tc")):
                self.assertEqual(cell.find(f"./{q(W, 'tcPr')}/{q(W, 'vAlign')}").get(q(W, "val")), "top")
        self.assertIsNone(table.find(f".//{q(W, 'tblHeader')}"))
        self.assertIsNone(table.find(f".//{q(W, 'cantSplit')}"))

        rows = table.findall(q(W, "tr"))
        header_p = rows[0].find(f"./{q(W, 'tc')}/{q(W, 'p')}")
        header_ppr = header_p.find(q(W, "pPr"))
        self.assertEqual(header_ppr.find(q(W, "pStyle")).get(q(W, "val")), "TableHeader")
        self.assertIsNone(header_ppr.find(q(W, "jc")))
        header_style = style_by_id(xml_part(output, "word/styles.xml"), "TableHeader")
        self.assertEqual(header_style.find(f"./{q(W, 'pPr')}/{q(W, 'jc')}").get(q(W, "val")), "center")
        self.assertEqual(header_ppr.find(q(W, "spacing")).get(q(W, "line")), "360")
        self.assertEqual(header_ppr.find(q(W, "spacing")).get(q(W, "lineRule")), "exact")
        header_rpr = header_p.find(f"./{q(W, 'r')}/{q(W, 'rPr')}")
        self.assertEqual(header_rpr.find(q(W, "rFonts")).get(q(W, "eastAsia")), "黑体")
        self.assertEqual(header_rpr.find(q(W, "sz")).get(q(W, "val")), "22")
        self.assertEqual(header_rpr.find(q(W, "b")).get(q(W, "val")), "0")
        self.assertEqual(header_rpr.find(q(W, "color")).get(q(W, "val")), "ABCDEF")

        body_p = rows[1].find(f"./{q(W, 'tc')}/{q(W, 'p')}")
        body_ppr = body_p.find(q(W, "pPr"))
        self.assertEqual(body_ppr.find(q(W, "pStyle")).get(q(W, "val")), "TableBody")
        self.assertEqual(body_ppr.find(q(W, "jc")).get(q(W, "val")), "left")
        self.assertEqual(body_ppr.find(q(W, "ind")).get(q(W, "firstLineChars")), "0")
        self.assertEqual(body_ppr.find(q(W, "spacing")).get(q(W, "line")), "360")
        self.assertEqual(body_ppr.find(q(W, "spacing")).get(q(W, "lineRule")), "auto")
        body_rpr = body_p.find(f"./{q(W, 'r')}/{q(W, 'rPr')}")
        self.assertEqual(body_rpr.find(q(W, "sz")).get(q(W, "val")), "24")
        self.assertEqual(body_rpr.find(q(W, "color")).get(q(W, "val")), "FFAABB")

    def test_table_line_spacing_mode_switches_in_both_directions(self):
        exact_global = self.work / "exact-global-auto-local.docx"
        create_docx.create_document({
            "styles": {"table_body": {"line_spacing_pt": 24}},
            "blocks": [{"type": "table", "rows": [["auto override"]],
                        "body_format": {"line_spacing_multiple": 1.5}}],
        }, exact_global)
        auto_local = xml_part(exact_global, "word/document.xml").find(f".//{q(W, 'pPr')}/{q(W, 'spacing')}")
        self.assertEqual(auto_local.get(q(W, "line")), "360")
        self.assertEqual(auto_local.get(q(W, "lineRule")), "auto")
        exact_global_style = style_by_id(xml_part(exact_global, "word/styles.xml"), "TableBody")
        self.assertEqual(exact_global_style.find(f"./{q(W, 'pPr')}/{q(W, 'spacing')}").get(q(W, "lineRule")),
                         "exact")

        auto_global = self.work / "auto-global-exact-local.docx"
        create_docx.create_document({
            "styles": {"table_body": {"line_spacing_multiple": 2}},
            "blocks": [{"type": "table", "rows": [["exact override"]],
                        "body_format": {"line_spacing_pt": 16}}],
        }, auto_global)
        exact_local = xml_part(auto_global, "word/document.xml").find(f".//{q(W, 'pPr')}/{q(W, 'spacing')}")
        self.assertEqual(exact_local.get(q(W, "line")), "320")
        self.assertEqual(exact_local.get(q(W, "lineRule")), "exact")
        auto_global_style = style_by_id(xml_part(auto_global, "word/styles.xml"), "TableBody")
        self.assertEqual(auto_global_style.find(f"./{q(W, 'pPr')}/{q(W, 'spacing')}").get(q(W, "lineRule")),
                         "auto")

    def test_invalid_table_contract_preserves_existing_output(self):
        output = self.work / "preserved-invalid-table.docx"
        original = b"existing destination survives invalid table formatting"
        table_block = {"type": "table", "header": False, "rows": [["x"]], "widths_mm": [100]}
        invalid_specs = [
            {"table_format": None, "blocks": []},
            {"table_format": {"unknown": 1}, "blocks": []},
            {"table_format": {"align": None}, "blocks": []},
            {"table_format": {"vertical_align": "baseline"}, "blocks": []},
            {"table_format": {"border_style": "none", "border_width_pt": 0.1}, "blocks": []},
            {"table_format": {"border_color": None}, "blocks": []},
            {"table_format": {"cell_fill": "#ffffff"}, "blocks": []},
            {"table_format": {"cell_margin_left_mm": 50.1}, "blocks": []},
            {"table_format": {"cell_margin_right_mm": float("nan")}, "blocks": []},
            {"table_format": {"repeat_header": 0}, "blocks": []},
            {"table_format": {"allow_row_split": None}, "blocks": []},
            {"styles": {"table_body": None}, "blocks": []},
            {"styles": {"table_body": {"color": None}}, "blocks": []},
            {"styles": {"table_body": {"color": "GGGGGG"}}, "blocks": []},
            {"styles": {"table_body": {"line_spacing_multiple": 0.49}}, "blocks": []},
            {"styles": {"table_body": {"line_spacing_pt": 20, "line_spacing_multiple": 1.5}}, "blocks": []},
            {"styles": {"table_header": {"font_east_asia": None}}, "blocks": []},
            {"styles": {"table_body": {"size_pt": None}}, "blocks": []},
            {"blocks": [{"type": "paragraph", "text": "x", "format": {"color": "#123456"}}]},
            {"blocks": [{**table_block, "format": None}]},
            {"blocks": [{**table_block, "format": {"border_style": "none", "border_width_pt": 0}}]},
            {"table_format": {"border_color": "nothex"},
             "blocks": [{**table_block, "format": {"border_color": "abcdef"}}]},
            {"styles": {"table_body": {"color": "bad", "size_pt": 20}},
             "blocks": [{**table_block, "body_format": {"size_pt": 12}}]},
            {"blocks": [{**table_block, "body_format": {"unknown": False}}]},
            {"blocks": [{**table_block, "header_format": {"bold": None}}]},
            {"blocks": [{"type": "table", "widths_mm": [0.1], "rows": [["x"]]}]},
        ]
        for index, spec in enumerate(invalid_specs):
            with self.subTest(index=index, spec=spec):
                output.write_bytes(original)
                with self.assertRaises(ValueError):
                    create_docx.create_document(spec, output)
                self.assertEqual(output.read_bytes(), original)

    def test_table_formats_do_not_mutate_or_leak_into_ordinary_paragraphs(self):
        custom_spec = {
            "styles": {"table_body": {"font_east_asia": "自定义表格字体", "color": "aabbcc"}},
            "table_format": {"cell_fill": "abcdef"},
            "blocks": [
                {"type": "paragraph", "text": "ordinary"},
                {"type": "table", "rows": [["table"]],
                 "body_format": {"font_latin": "Table Latin", "color": "112233"}},
            ],
        }
        original = json.loads(json.dumps(custom_spec, ensure_ascii=False))
        custom = self.work / "table-custom.docx"
        default = self.work / "table-default-after-custom.docx"
        create_docx.create_document(custom_spec, custom)
        self.assertEqual(custom_spec, original)
        create_docx.create_document({"blocks": [{"type": "table", "rows": [["default"]]}]}, default)

        custom_doc = xml_part(custom, "word/document.xml")
        body_paragraph = custom_doc.find(f"./{q(W, 'body')}/{q(W, 'p')}")
        self.assertEqual(body_paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}").get(q(W, "val")), "BodyText")
        ordinary_rpr = body_paragraph.find(f"./{q(W, 'r')}/{q(W, 'rPr')}")
        self.assertIsNone(ordinary_rpr)
        self.assertEqual(custom_doc.find(f".//{q(W, 'tblPr')}/{q(W, 'shd')}").get(q(W, "fill")), "ABCDEF")
        custom_cell_p = custom_doc.find(f".//{q(W, 'tc')}/{q(W, 'p')}")
        self.assertEqual(custom_cell_p.find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}").get(q(W, "val")), "TableBody")
        custom_rpr = custom_cell_p.find(f"./{q(W, 'r')}/{q(W, 'rPr')}")
        self.assertIsNone(custom_rpr.find(q(W, "rFonts")).get(q(W, "eastAsia")))
        self.assertEqual(custom_rpr.find(q(W, "rFonts")).get(q(W, "ascii")), "Table Latin")
        custom_style = style_by_id(xml_part(custom, "word/styles.xml"), "TableBody")
        self.assertEqual(custom_style.find(f"./{q(W, 'rPr')}/{q(W, 'rFonts')}").get(q(W, "eastAsia")),
                         "自定义表格字体")
        self.assertEqual(custom_rpr.find(q(W, "color")).get(q(W, "val")), "112233")
        default_doc = xml_part(default, "word/document.xml")
        self.assertIsNone(default_doc.find(f".//{q(W, 'tblPr')}/{q(W, 'shd')}"))
        default_rpr = default_doc.find(f".//{q(W, 'tc')}/{q(W, 'p')}/{q(W, 'r')}/{q(W, 'rPr')}")
        self.assertIsNone(default_rpr)
        default_style = style_by_id(xml_part(default, "word/styles.xml"), "TableBody")
        self.assertEqual(default_style.find(f"./{q(W, 'rPr')}/{q(W, 'rFonts')}").get(q(W, "eastAsia")), "仿宋")
        self.assertEqual(default_style.find(f"./{q(W, 'rPr')}/{q(W, 'color')}").get(q(W, "val")), "000000")

    def test_invalid_new_contract_values_preserve_existing_output(self):
        output = self.work / "preserved-new-contract.docx"
        original = b"existing destination survives invalid formatting"
        invalid_specs = [
            {"styles": None, "blocks": []},
            {"styles": {"unknown": {}}, "blocks": []},
            {"styles": {"body": {"font_latin": "line\nbreak"}}, "blocks": []},
            {"styles": {"body": {"font_latin": "x" * 65}}, "blocks": []},
            {"styles": {"body": {"size_pt": True}}, "blocks": []},
            {"styles": {"body": {"bold": 0}}, "blocks": []},
            {"styles": {"body": {"align": []}}, "blocks": []},
            {"styles": {"body": {"first_line_chars": -0.1}}, "blocks": []},
            {"styles": {"body": {"line_spacing_pt": 0}}, "blocks": []},
            {"styles": {"body": {"space_before_pt": float("nan")}}, "blocks": []},
            {"styles": {"body": {"space_after_pt": None}}, "blocks": []},
            {"blocks": [{"type": "paragraph", "text": "x", "format": None}]},
            {"blocks": [{"type": "paragraph", "text": "x", "format": {"italic": True}}]},
            {"blocks": [{"type": "paragraph", "text": "x", "align": None}]},
            {"blocks": [{"type": "heading", "text": "x", "level": 5}]},
            {"blocks": [{"type": "paragraph", "runs": [{"text": "x", "font_east_asia": None}]}]},
            {"page": {"margin_mm": -1, "margin_top_mm": 0, "margin_bottom_mm": 0,
                       "margin_left_mm": 0, "margin_right_mm": 0}, "blocks": []},
            {"page": {"margin_mm": 10, "margin_right_mm": float("inf")}, "blocks": []},
            {"page": {"width_mm": 10, "height_mm": 10, "margin_top_mm": 5,
                       "margin_bottom_mm": 5}, "blocks": []},
        ]
        for index, spec in enumerate(invalid_specs):
            with self.subTest(spec=spec):
                output.write_bytes(original)
                with self.assertRaises(ValueError):
                    create_docx.create_document(spec, output)
                self.assertEqual(output.read_bytes(), original)

    def test_style_configuration_does_not_leak_between_calls(self):
        custom = {"styles": {"body": {"font_east_asia": "自定义字体", "font_latin": "Custom Latin",
                                       "size_pt": 13, "bold": True}},
                  "blocks": [{"type": "paragraph", "text": "custom"}]}
        custom_original = json.loads(json.dumps(custom, ensure_ascii=False))
        custom_output = self.work / "custom.docx"
        default_output = self.work / "default-after-custom.docx"
        create_docx.create_document(custom, custom_output)
        create_docx.create_document({"blocks": [{"type": "paragraph", "text": "default"}]}, default_output)
        self.assertEqual(custom, custom_original)
        custom_styles = xml_part(custom_output, "word/styles.xml")
        default_styles = xml_part(default_output, "word/styles.xml")
        self.assertEqual(style_by_id(custom_styles, "BodyText").find(
            f"./{q(W, 'rPr')}/{q(W, 'rFonts')}").get(q(W, "eastAsia")), "自定义字体")
        default_body = style_by_id(default_styles, "BodyText")
        self.assertEqual(default_body.find(f"./{q(W, 'rPr')}/{q(W, 'rFonts')}").get(q(W, "eastAsia")), "仿宋")
        self.assertEqual(default_body.find(f"./{q(W, 'rPr')}/{q(W, 'sz')}").get(q(W, "val")), "32")

    def test_cli_error_preserves_existing_output(self):
        spec_path = self.work / "invalid.json"
        spec_path.write_text('{"blocks":[{"type":"page_break","extra":true}]}', encoding="utf-8")
        output = self.work / "preserved.docx"
        original = b"keep this file"
        output.write_bytes(original)
        result = self.run_cli(spec_path, output)
        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("unknown field", result.stderr)
        self.assertEqual(output.read_bytes(), original)


if __name__ == "__main__":
    unittest.main()
