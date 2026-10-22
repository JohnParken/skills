"""Shared multilevel lists, repeated headers and cell images use real OOXML."""

import copy
import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "docx/scripts"))
from create_docx import create_document
from office.package import unpack_document
from office.validators.structure import StructureValidator
from test_docx_generator import tiny_png

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def q(name):
    return f"{{{W}}}{name}"


class GenerationDetailsTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-generation-details-")
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name)
        self.output = self.work / "output.docx"
        self.image = self.work / "pixel.png"
        self.image.write_bytes(tiny_png())

    def generate(self, spec):
        create_document(spec, self.output, base_dir=self.work)
        with zipfile.ZipFile(self.output) as package:
            self.assertIsNone(package.testzip())
            return {name: package.read(name) for name in package.namelist()}

    def reject(self, spec):
        self.output.write_bytes(b"keep output")
        before = copy.deepcopy(spec)
        with self.assertRaises((ValueError, OSError)):
            create_document(spec, self.output, base_dir=self.work)
        self.assertEqual(self.output.read_bytes(), b"keep output")
        self.assertEqual(spec, before)

    def list_block(self, identifier="工作清单", level=0, **updates):
        block = {"type": "list", "ordered": True, "list_id": identifier,
                 "level": level, "items": [f"层级 {level}"]}
        block.update(updates)
        return block

    def image_cell(self, **updates):
        image = {"path": "pixel.png", "width_mm": 20, "height_mm": 15, "alt": "单元格图片"}
        image.update(updates)
        return {"text": "中文图片说明", "image": image}

    def test_shared_list_continues_after_description_and_new_id_restarts(self):
        spec = {"blocks": [self.list_block(), self.list_block(level=1),
                           {"type": "paragraph", "text": "这是一段不编号的说明。"},
                           self.list_block(level=1, items=["继续子层级"]),
                           self.list_block(identifier="另一清单"),
                           {"type": "list", "ordered": True, "items": ["旧式独立列表"]}]}
        parts = self.generate(spec)
        document = ET.fromstring(parts["word/document.xml"])
        numbering = ET.fromstring(parts["word/numbering.xml"])
        references = [(node.find(q("numId")).get(q("val")), node.find(q("ilvl")).get(q("val")))
                      for node in document.iter(q("numPr"))]
        self.assertEqual(references[:3], [("1", "0"), ("1", "1"), ("1", "1")])
        self.assertEqual(references[3:], [("2", "0"), ("3", "0")])
        instances = {node.get(q("numId")): node.find(q("abstractNumId")).get(q("val"))
                     for node in numbering.findall(q("num"))}
        self.assertEqual(instances["1"], instances["2"])
        self.assertNotEqual(instances["1"], instances["3"])
        abstract = next(node for node in numbering.findall(q("abstractNum"))
                        if node.get(q("abstractNumId")) == instances["1"])
        self.assertEqual([node.get(q("ilvl")) for node in abstract.findall(q("lvl"))],
                         list(map(str, range(9))))
        unpacked = unpack_document(self.output, self.work / "list-parts")
        self.assertTrue(StructureValidator(unpacked, extension=".docx").validate())

    def test_invalid_shared_list_identifiers_levels_and_order_are_rejected(self):
        for block in (self.list_block(identifier=""), self.list_block(identifier=True),
                      self.list_block(identifier=" padded "), self.list_block(identifier="x\ny"),
                      self.list_block(level=True), self.list_block(level=-1), self.list_block(level=9),
                      self.list_block(level=1), self.list_block(items=[]),
                      {"type": "list", "ordered": True, "level": 1, "items": ["no parent"]}):
            with self.subTest(block=block):
                self.reject({"blocks": [block]})
        self.reject({"blocks": [self.list_block(), self.list_block(ordered=False)]})
        self.reject({"blocks": [self.list_block(), self.list_block(level=2)]})

    def test_all_nine_levels_and_bullet_instances_are_explicit(self):
        spec = {"blocks": [self.list_block(level=level) for level in range(9)]
                + [self.list_block(identifier="项目符号", ordered=False),
                   self.list_block(identifier="项目符号", ordered=False, level=1)]}
        parts = self.generate(spec)
        document = ET.fromstring(parts["word/document.xml"])
        levels = [node.get(q("val")) for node in document.iter(q("ilvl"))]
        self.assertEqual(levels, list(map(str, range(9))) + ["0", "1"])
        unpacked = unpack_document(self.output, self.work / "nine-levels")
        self.assertTrue(StructureValidator(unpacked, extension=".docx").validate())

    def test_header_rows_apply_style_and_repeat_markers_independently(self):
        parts = self.generate({"blocks": [
            {"type": "table", "header_rows": 2, "rows": [["第一表头"], ["第二表头"], ["数据"]]},
            {"type": "table", "header_rows": 2, "format": {"repeat_header": False},
             "rows": [["A"], ["B"], ["C"]]},
        ]})
        tables = list(ET.fromstring(parts["word/document.xml"]).iter(q("tbl")))
        self.assertEqual(len(list(tables[0].iter(q("tblHeader")))), 2)
        self.assertEqual(len(list(tables[1].iter(q("tblHeader")))), 0)
        for table in tables:
            styles = [row.find(f"./{q('tc')}/{q('p')}/{q('pPr')}/{q('pStyle')}").get(q("val"))
                      for row in table.findall(q("tr"))]
            self.assertEqual(styles, ["TableHeader", "TableHeader", "TableBody"])

    def test_invalid_or_conflicting_header_counts_preserve_output(self):
        table = {"type": "table", "rows": [["head"], ["data"]]}
        for count in (True, -1, 3, 1.5, "2", None):
            with self.subTest(count=count):
                self.reject({"blocks": [{**table, "header_rows": count}]})
        self.reject({"blocks": [{**table, "header": True, "header_rows": 2}]})
        self.reject({"blocks": [{**table, "header": False, "header_rows": 1}]})
        self.generate({"blocks": [{**table, "header": True, "header_rows": 1}]})

    def test_cell_images_use_unique_drawings_media_and_relationships(self):
        parts = self.generate({"blocks": [
            {"type": "image", "path": "pixel.png", "width_mm": 20, "height_mm": 10, "alt": "正文图片"},
            {"type": "table", "header_rows": 1, "widths_mm": [60, 60],
             "rows": [["名称", "图片"], [self.image_cell(), {"image": self.image_cell()["image"]}],
                      ["纯文字", {"text": ""}]]},
        ]})
        doc = ET.fromstring(parts["word/document.xml"])
        blips = list(doc.iter(f"{{{A}}}blip"))
        drawings = list(doc.iter(f"{{{WP}}}docPr"))
        self.assertEqual(len(blips), 3)
        self.assertEqual(len({node.get(f"{{{R}}}embed") for node in blips}), 3)
        self.assertEqual(len({node.get("id") for node in drawings}), 3)
        self.assertEqual([node.get("descr") for node in drawings], ["正文图片", "单元格图片", "单元格图片"])
        self.assertEqual(len([name for name in parts if name.startswith("word/media/")]), 3)
        self.assertIn("中文图片说明", [node.text for node in doc.iter(q("t"))])
        table = next(doc.iter(q("tbl")))
        self.assertEqual(len(list(table.iter(q("drawing")))), 2)
        unpacked = unpack_document(self.output, self.work / "image-parts")
        self.assertTrue(StructureValidator(unpacked, extension=".docx").validate())

    def test_invalid_cell_objects_and_images_preserve_output(self):
        for cell in ({}, {"text": 1}, {"unknown": "x"}, {"image": None},
                     self.image_cell(width_mm=40), self.image_cell(height_mm=400),
                     self.image_cell(width_mm=True), self.image_cell(alt=None),
                     self.image_cell(path="missing.png")):
            with self.subTest(cell=cell):
                self.reject({"blocks": [{"type": "table", "widths_mm": [30], "rows": [[cell]]}]})
        self.image.write_bytes(b"not image data")
        self.reject({"blocks": [{"type": "table", "rows": [[self.image_cell()]]}]})

    def test_cell_image_diagnostics_identify_the_nested_field(self):
        self.image.write_bytes(b"invalid PNG")
        for updates, field in (({}, "path"), ({"path": 123}, "path"), ({"alt": None}, "alt")):
            with self.subTest(updates=updates):
                spec = {"blocks": [{"type": "table", "rows": [[self.image_cell(**updates)]]}]}
                with self.assertRaisesRegex(ValueError, rf"blocks\[0\]\.rows\[0\]\[0\]\.image\.{field}"):
                    create_document(spec, self.output, base_dir=self.work)

    def test_cell_image_input_cannot_be_overwritten_by_output(self):
        image_alias = self.work / "image-output.docx"
        image_alias.symlink_to(self.image)
        before = self.image.read_bytes()
        with self.assertRaisesRegex(ValueError, "overwrite an image"):
            create_document({"blocks": [{"type": "table", "rows": [[self.image_cell()]]}]},
                            image_alias, base_dir=self.work)
        self.assertEqual(self.image.read_bytes(), before)

    def test_each_border_merges_global_and_local_overrides_without_leaking(self):
        spec = {"table_format": {"border_color": "111111", "borders": {
                    "top": {"color": "abcdef", "width_pt": 2}, "insideV": {"style": "none"}}},
                "blocks": [{"type": "table", "rows": [["边框"]], "format": {
                    "borders": {"top": {"width_pt": 1}, "bottom": {"color": "ff0000"}}}}]}
        original = copy.deepcopy(spec)
        parts = self.generate(spec)
        self.assertEqual(spec, original)
        borders = next(ET.fromstring(parts["word/document.xml"]).iter(q("tblBorders")))
        self.assertEqual(borders.find(q("top")).get(q("color")), "ABCDEF")
        self.assertEqual(borders.find(q("top")).get(q("sz")), "8")
        self.assertEqual(borders.find(q("bottom")).get(q("color")), "FF0000")
        self.assertEqual(borders.find(q("left")).get(q("color")), "111111")
        self.assertEqual(borders.find(q("insideV")).get(q("val")), "nil")
        parts = self.generate({"blocks": [{"type": "table", "rows": [["默认"]]}]})
        borders = next(ET.fromstring(parts["word/document.xml"]).iter(q("tblBorders")))
        self.assertTrue(all(node.get(q("color")) == "000000" for node in borders))

    def test_invalid_borders_are_rejected(self):
        for borders in (None, {"diagonal": {}}, {"top": {"style": "double"}},
                        {"bottom": {"width_pt": 0}}, {"left": {"width_pt": True}},
                        {"right": {"color": "#ffffff"}}, {"top": {"unknown": 1}}):
            with self.subTest(borders=borders):
                self.reject({"table_format": {"borders": borders}, "blocks": []})

    def test_complex_spec_is_unchanged_and_cli_needs_no_executables(self):
        spec = {"blocks": [self.list_block(), self.list_block(level=1),
                           {"type": "table", "header_rows": 1,
                            "rows": [["表头"], [self.image_cell()]]}]}
        original = copy.deepcopy(spec)
        self.generate(spec)
        self.assertEqual(spec, original)
        (self.work / "input.json").write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
        result = subprocess.run([sys.executable, "-I", "-S", "-B", str(ROOT / "docx/scripts/create_docx.py"),
                                 "input.json", "cli.docx"], cwd=self.work,
                                env=dict(os.environ, PATH=""), capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue((self.work / "cli.docx").is_file())


if __name__ == "__main__":
    unittest.main()
