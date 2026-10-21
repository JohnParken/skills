"""Independent cell edges resolve deterministically without reader conflicts."""

import copy
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


def q(name):
    return f"{{{W}}}{name}"


class CellBordersTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-cell-borders-")
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name)
        self.output = self.work / "output.docx"

    def generate(self, rows, **fields):
        spec = {"blocks": [{"type": "table", "rows": rows, **fields}]}
        original = copy.deepcopy(spec)
        create_document(spec, self.output, base_dir=self.work)
        self.assertEqual(spec, original)
        with zipfile.ZipFile(self.output) as package:
            self.assertIsNone(package.testzip())
            document = ET.fromstring(package.read("word/document.xml"))
        return [[cell.find(q("tcPr")) for cell in row.findall(q("tc"))]
                for row in document.iter(q("tr"))]

    def border(self, cell, edge):
        node = cell.find(q("tcBorders") + "/" + q(edge))
        return None if node is None else {key.split("}")[-1]: value for key, value in node.attrib.items()}

    def reject(self, rows, **fields):
        self.output.write_bytes(b"existing output")
        with self.assertRaises(ValueError):
            self.generate(rows, **fields)
        self.assertEqual(self.output.read_bytes(), b"existing output")

    def test_default_and_empty_overrides_add_no_cell_border_xml(self):
        for cell in ("文本", {"text": "文本"}, {"text": "文本", "borders": {}}):
            with self.subTest(cell=cell):
                self.assertIsNone(self.generate([[cell]])[0][0].find(q("tcBorders")))

    def test_outer_edges_inherit_final_table_edges_and_override_fields(self):
        cells = self.generate([[{"text": "甲", "borders": {
            "top": {"color": "aabbcc"}, "left": {"width_pt": 2},
            "bottom": {"style": "none"}, "right": {}}}]],
            format={"border_width_pt": 1, "borders": {"top": {"width_pt": 1.5},
                    "left": {"color": "123456"}}})
        self.assertEqual(self.border(cells[0][0], "top"), {"val": "single", "sz": "12", "space": "0", "color": "AABBCC"})
        self.assertEqual(self.border(cells[0][0], "left")["color"], "123456")
        self.assertEqual(self.border(cells[0][0], "left")["sz"], "16")
        self.assertEqual(self.border(cells[0][0], "bottom"), {"val": "nil"})
        self.assertEqual(self.border(cells[0][0], "right")["sz"], "8")
        self.assertEqual([node.tag.split("}")[-1] for node in cells[0][0]], ["tcW", "tcBorders", "vAlign"])

    def test_shared_horizontal_and_vertical_edges_are_mirrored(self):
        cells = self.generate([[{"text": "表头", "borders": {"right": {"color": "abcdef"},
                    "bottom": {"width_pt": 2}}}, "表头乙"], ["数据甲", "数据乙"]],
                    header_rows=1, format={"borders": {"insideH": {"color": "123456"},
                                                        "insideV": {"width_pt": 1.5}}})
        self.assertEqual(self.border(cells[0][0], "right"), self.border(cells[0][1], "left"))
        self.assertEqual(self.border(cells[0][1], "left")["sz"], "12")
        self.assertEqual(self.border(cells[0][0], "bottom"), self.border(cells[1][0], "top"))
        self.assertEqual(self.border(cells[1][0], "top")["color"], "123456")
        self.assertIsNone(cells[1][1].find(q("tcBorders")))

    def test_reverse_side_and_none_overrides_are_mirrored(self):
        cells = self.generate([["甲", "乙"], ["丙", {"text": "丁", "borders": {
            "top": {"style": "none"}, "left": {"color": "123456"}}}]])
        self.assertEqual(self.border(cells[1][1], "top"), {"val": "nil"})
        self.assertEqual(self.border(cells[0][1], "bottom"), {"val": "nil"})
        self.assertEqual(self.border(cells[1][0], "right"), self.border(cells[1][1], "left"))

    def test_matching_encoded_shared_borders_are_allowed(self):
        for first, second in (({"width_pt": 0.30}, {"width_pt": 0.31}),
                              ({"style": "none", "color": "123456"}, {"style": "none"})):
            with self.subTest(first=first, second=second):
                cells = self.generate([[{"text": "甲", "borders": {"right": first}},
                                        {"text": "乙", "borders": {"left": second}}]])
                self.assertEqual(self.border(cells[0][0], "right"), self.border(cells[0][1], "left"))

    def test_conflicting_shared_edges_refuse_and_preserve_existing_output(self):
        self.reject([[{"text": "甲", "borders": {"right": {"color": "123456"}}},
                      {"text": "乙", "borders": {"left": {"color": "abcdef"}}}]])
        self.reject([[{"text": "甲", "borders": {"bottom": {"style": "none"}}}],
                     [{"text": "乙", "borders": {"top": {"style": "single"}}}]])

    def test_invalid_edges_types_values_and_borders_only_are_rejected(self):
        invalid = (None, [], {"insideH": {}}, {"diagonal": {}}, {"top": None},
                   {"top": {"style": "double"}}, {"top": {"width_pt": True}},
                   {"top": {"width_pt": float("nan")}}, {"top": {"width_pt": float("inf")}},
                   {"top": {"width_pt": 0.24}}, {"top": {"width_pt": 12.1}},
                   {"top": {"color": "#123456"}}, {"top": {"color": "１２３４５６"}},
                   {"top": {"space": 1}}, {"start": {"style": "none"}})
        for borders in invalid:
            with self.subTest(borders=borders):
                self.reject([[{"text": "甲", "borders": borders}]])
        self.reject([[{"borders": {"top": {}}}]])
        self.reject([[{"text": "甲", "borders": {"top": {}}}], ["乙", "丙"]])

    def test_cell_images_headers_and_tables_do_not_leak_overrides(self):
        (self.work / "pixel.png").write_bytes(tiny_png())
        image_cell = {"image": {"path": "pixel.png", "width_mm": 10, "height_mm": 10, "alt": "图"},
                      "borders": {"top": {"width_pt": 12}, "bottom": {"width_pt": 0.25}}}
        cells = self.generate([[image_cell], ["数据"]], header_rows=1)
        self.assertEqual(self.border(cells[0][0], "top")["sz"], "96")
        self.assertEqual(self.border(cells[1][0], "top")["sz"], "2")
        unpacked = unpack_document(self.output, self.work / "unpacked")
        self.assertTrue(StructureValidator(unpacked, extension=".docx").validate())
        self.assertIsNone(self.generate([["新表"]])[0][0].find(q("tcBorders")))


if __name__ == "__main__":
    unittest.main()
