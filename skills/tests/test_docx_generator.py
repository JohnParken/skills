"""Focused tests for the stdlib DOCX generator and its JSON CLI."""

import json
import os
import struct
import subprocess
import sys
import tempfile
import unittest
import zipfile
import zlib
from pathlib import Path
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
            self.assertTrue(any(True for _ in document.iter(q(W, "b"))))
            table_props = document.find(f".//{q(W, 'tblPr')}")
            self.assertEqual([node.tag.rsplit("}", 1)[-1] for node in table_props],
                             ["tblW", "tblBorders", "tblLayout"])
            page_size = document.find(f".//{q(W, 'pgSz')}")
            self.assertEqual(page_size.get(q(W, "w")), "11906")
            self.assertEqual(page_size.get(q(W, "h")), "16838")

            styles = roots["word/styles.xml"]
            fonts = styles.find(f".//{q(W, 'docDefaults')}/{q(W, 'rPrDefault')}/{q(W, 'rPr')}/{q(W, 'rFonts')}")
            self.assertEqual(fonts.get(q(W, "ascii")), "Arial")
            self.assertEqual(fonts.get(q(W, "eastAsia")), "宋体")
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
