"""Regression tests for the DOCX generator's decided optimizations."""

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "docx" / "scripts"
sys.path.insert(0, str(SCRIPTS))
import create_docx
from test_docx_generator import png_chunk, tiny_png

W = create_docx.W
WP = create_docx.WP
A = create_docx.A
R = create_docx.R


def q(namespace, local):
    return f"{{{namespace}}}{local}"


def xml_part(path, name):
    with zipfile.ZipFile(path) as package:
        return ET.fromstring(package.read(name))


def text_in(node):
    return "".join(part.text or "" for part in node.iter(q(W, "t")))


def png_with_text(label):
    image = tiny_png()
    return image[:-12] + png_chunk(b"tEXt", b"Label\x00" + label) + image[-12:]


class DocxOptimizationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-optimization-test-")
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name)
        self.output = self.work / "output.docx"

    def generate(self, spec, output=None, *, verify=True):
        return create_docx.create_document(
            spec, output or self.output, self.work, verify=verify,
        )

    def assert_rejected_atomically(self, spec, patcher):
        sentinel = b"preserve this existing document"
        self.output.write_bytes(sentinel)
        before = {path.name for path in self.work.iterdir()}
        with patcher:
            with self.assertRaises(ValueError):
                self.generate(spec)
        self.assertEqual(self.output.read_bytes(), sentinel)
        self.assertEqual({path.name for path in self.work.iterdir()}, before)

    def patch_document_xml(self, mutate):
        original = create_docx._document_xml
        changed = []

        def corrupt(spec):
            data, images, lists = original(spec)
            root = ET.fromstring(data)
            changed.append(mutate(root))
            return ET.tostring(root, encoding="UTF-8", xml_declaration=True), images, lists

        return mock.patch.object(create_docx, "_document_xml", side_effect=corrupt), changed

    def patch_styles_xml(self, mutate):
        original = create_docx._styles_xml
        changed = []

        def corrupt(*args, **kwargs):
            root = ET.fromstring(original(*args, **kwargs))
            changed.append(mutate(root))
            return ET.tostring(root, encoding="UTF-8", xml_declaration=True)

        return mock.patch.object(create_docx, "_styles_xml", side_effect=corrupt), changed

    def assert_document_fault_rejected(self, spec, mutate):
        patcher, changed = self.patch_document_xml(mutate)
        self.assert_rejected_atomically(spec, patcher)
        self.assertEqual(changed, [True])

    def assert_styles_fault_rejected(self, spec, mutate):
        patcher, changed = self.patch_styles_xml(mutate)
        self.assert_rejected_atomically(spec, patcher)
        self.assertEqual(changed, [True])

    def test_themed_lists_use_list_style_without_body_overrides_or_extra_indent(self):
        for theme, east_asia, latin in (
            ("business", "微软雅黑", "Calibri"),
            ("academic", "宋体", "Times New Roman"),
        ):
            with self.subTest(theme=theme):
                output = self.work / f"{theme}.docx"
                spec = {
                    "theme": theme,
                    "styles": {"body": {
                        "font_east_asia": "楷体", "font_latin": "Arial", "size_pt": 18,
                        "line_spacing_pt": 22, "space_after_pt": 16, "first_line_chars": 2,
                    }},
                    "blocks": [
                        {"type": "paragraph", "text": "正文继承覆盖"},
                        {"type": "list", "ordered": True, "list_id": "steps", "items": ["一级"]},
                        {"type": "list", "ordered": True, "list_id": "steps", "level": 1,
                         "items": ["二级"]},
                        {"type": "list", "ordered": True, "list_id": "steps", "items": ["续号"]},
                        {"type": "list", "ordered": False, "list_id": "notes", "items": ["另一列表"]},
                        {"type": "list", "ordered": False, "items": ["独立列表"]},
                        {"type": "list", "ordered": False, "items": ["再次独立"]},
                        {"type": "list", "ordered": False, "items": ["第三个独立列表"]},
                    ],
                }
                self.generate(spec, output)
                document = xml_part(output, "word/document.xml")
                styles = xml_part(output, "word/styles.xml")
                style = next(node for node in styles.findall(q(W, "style"))
                             if node.get(q(W, "styleId")) == "ListParagraph")
                style_ppr = style.find(q(W, "pPr"))
                style_rpr = style.find(q(W, "rPr"))
                self.assertIsNone(style_ppr.find(q(W, "ind")))
                self.assertEqual(style_ppr.find(q(W, "spacing")).get(q(W, "line")), "300")
                self.assertEqual(style_ppr.find(q(W, "spacing")).get(q(W, "lineRule")), "auto")
                self.assertEqual(style_ppr.find(q(W, "spacing")).get(q(W, "after")), "80")
                self.assertEqual(style_rpr.find(q(W, "sz")).get(q(W, "val")), "21")
                fonts = style_rpr.find(q(W, "rFonts"))
                self.assertEqual(fonts.get(q(W, "eastAsia")), east_asia)
                self.assertEqual(fonts.get(q(W, "ascii")), latin)

                paragraphs = [node for node in document.iter(q(W, "p"))
                              if node.find(f"./{q(W, 'pPr')}/{q(W, 'numPr')}") is not None]
                self.assertEqual(len(paragraphs), 7)
                self.assertEqual(
                    [node.find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}").get(q(W, "val"))
                     for node in paragraphs],
                    ["ListParagraph"] * 7,
                )
                self.assertTrue(all(node.find(f"./{q(W, 'pPr')}/{q(W, 'ind')}") is None
                                    for node in paragraphs))
                list_levels = [node.find(f"./{q(W, 'pPr')}/{q(W, 'numPr')}/{q(W, 'ilvl')}")
                               .get(q(W, "val")) for node in paragraphs]
                self.assertEqual(list_levels, ["0", "1", "0", "0", "0", "0", "0"])
                list_ids = [node.find(f"./{q(W, 'pPr')}/{q(W, 'numPr')}/{q(W, 'numId')}")
                            .get(q(W, "val")) for node in paragraphs]
                self.assertEqual(list_ids[0], list_ids[1])
                self.assertEqual(list_ids[0], list_ids[2])
                self.assertNotEqual(list_ids[0], list_ids[3])
                self.assertNotEqual(list_ids[3], list_ids[4])
                self.assertNotEqual(list_ids[4], list_ids[5])
                self.assertNotEqual(list_ids[5], list_ids[6])

                body_paragraph = next(node for node in document.iter(q(W, "p"))
                                      if text_in(node) == "正文继承覆盖")
                self.assertEqual(body_paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}")
                                 .get(q(W, "val")), "BodyText")
                body_style = next(node for node in styles.findall(q(W, "style"))
                                  if node.get(q(W, "styleId")) == "BodyText")
                self.assertEqual(body_style.find(f"./{q(W, 'rPr')}/{q(W, 'sz')}")
                                 .get(q(W, "val")), "36")

        official = self.work / "official.docx"
        self.generate({"blocks": [{"type": "list", "ordered": True, "items": ["默认列表"]}]}, official)
        official_document = xml_part(official, "word/document.xml")
        official_list = next(node for node in official_document.iter(q(W, "p"))
                             if node.find(f"./{q(W, 'pPr')}/{q(W, 'numPr')}") is not None)
        self.assertIsNone(official_list.find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}"))

    def test_table_formats_live_in_shared_styles_and_keep_explicit_overrides(self):
        output = self.work / "table-formats.docx"
        self.generate({
            "styles": {
                "table_body": {"font_east_asia": "楷体", "font_latin": "Arial", "size_pt": 13,
                               "bold": True, "align": "right", "first_line_chars": 1,
                               "line_spacing_pt": 20, "space_after_pt": 5},
                "table_header": {"font_east_asia": "黑体", "font_latin": "Calibri", "size_pt": 12,
                                 "bold": True, "align": "center", "line_spacing_multiple": 1.2},
            },
            "blocks": [{
                "type": "table", "header": True, "rows": [["列名"], ["数据"]],
                "body_format": {"size_pt": 13, "bold": False, "align": "right",
                                "line_spacing_multiple": 1.5},
                "header_format": {"size_pt": 12, "bold": False, "align": "center",
                                  "line_spacing_pt": 18},
            }],
        }, output)
        styles = xml_part(output, "word/styles.xml")
        body_style = next(node for node in styles.findall(q(W, "style"))
                          if node.get(q(W, "styleId")) == "TableBody")
        header_style = next(node for node in styles.findall(q(W, "style"))
                            if node.get(q(W, "styleId")) == "TableHeader")
        self.assertEqual(body_style.find(f"./{q(W, 'rPr')}/{q(W, 'sz')}").get(q(W, "val")), "26")
        self.assertEqual(body_style.find(f"./{q(W, 'rPr')}/{q(W, 'rFonts')}")
                         .get(q(W, "eastAsia")), "楷体")
        self.assertEqual(body_style.find(f"./{q(W, 'pPr')}/{q(W, 'spacing')}")
                         .get(q(W, "lineRule")), "exact")
        self.assertEqual(header_style.find(f"./{q(W, 'pPr')}/{q(W, 'spacing')}")
                         .get(q(W, "lineRule")), "auto")

        table = xml_part(output, "word/document.xml").find(f".//{q(W, 'tbl')}")
        header = table.find(f"./{q(W, 'tr')}/{q(W, 'tc')}/{q(W, 'p')}")
        body = table.findall(f"./{q(W, 'tr')}/{q(W, 'tc')}/{q(W, 'p')}")[1]
        self.assertEqual(header.find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}").get(q(W, "val")),
                         "TableHeader")
        self.assertEqual(body.find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}").get(q(W, "val")),
                         "TableBody")
        header_rpr = header.find(f"./{q(W, 'r')}/{q(W, 'rPr')}")
        body_rpr = body.find(f"./{q(W, 'r')}/{q(W, 'rPr')}")
        self.assertEqual(header_rpr.find(q(W, "sz")).get(q(W, "val")), "24")
        self.assertEqual(header_rpr.find(q(W, "b")).get(q(W, "val")), "0")
        self.assertIsNone(header_rpr.find(q(W, "rFonts")))
        self.assertEqual(body_rpr.find(q(W, "sz")).get(q(W, "val")), "26")
        self.assertEqual(body_rpr.find(q(W, "b")).get(q(W, "val")), "0")
        self.assertIsNone(body_rpr.find(q(W, "rFonts")))
        header_spacing = header.find(f"./{q(W, 'pPr')}/{q(W, 'spacing')}")
        body_spacing = body.find(f"./{q(W, 'pPr')}/{q(W, 'spacing')}")
        self.assertEqual((header_spacing.get(q(W, "line")), header_spacing.get(q(W, "lineRule"))),
                         ("360", "exact"))
        self.assertEqual((body_spacing.get(q(W, "line")), body_spacing.get(q(W, "lineRule"))),
                         ("360", "auto"))
        self.assertIsNone(body_spacing.get(q(W, "after")))
        self.assertIsNone(body.find(f"./{q(W, 'pPr')}/{q(W, 'ind')}"))

    def test_image_reads_are_shared_per_call_but_media_and_drawings_stay_independent(self):
        image_path = self.work / "pixel.png"
        image_path.write_bytes(tiny_png())
        image = lambda path, width, height, alt: {
            "type": "image", "path": path, "width_mm": width, "height_mm": height, "alt": alt,
        }
        references = [
            image("pixel.png", 10, 5, "第一引用"),
            image("./pixel.png", 12, 4, "第二引用"),
            image(image_path.as_posix(), 20, 10, "第三引用"),
        ]
        original_open = Path.open
        for deduplicate, expected_media in ((False, 3), (True, 1)):
            with self.subTest(deduplicate=deduplicate):
                opened = []

                def count_open(path, *args, **kwargs):
                    if path.resolve() == image_path.resolve():
                        opened.append(path)
                    return original_open(path, *args, **kwargs)

                output = self.work / f"images-{deduplicate}.docx"
                with mock.patch.object(Path, "open", new=count_open):
                    self.generate({"deduplicate_images": deduplicate,
                                   "blocks": references}, output)
                self.assertEqual(len(opened), 1)
                with zipfile.ZipFile(output) as package:
                    media = [name for name in package.namelist() if name.startswith("word/media/")]
                    self.assertEqual(len(media), expected_media)
                    self.assertTrue(all(package.read(name) == image_path.read_bytes() for name in media))
                document = xml_part(output, "word/document.xml")
                extents = list(document.iter(q(WP, "extent")))
                self.assertEqual([(node.get("cx"), node.get("cy")) for node in extents], [
                    ("360000", "180000"), ("432000", "144000"), ("720000", "360000"),
                ])
                properties = list(document.iter(q(WP, "docPr")))
                self.assertEqual([node.get("descr") for node in properties],
                                 ["第一引用", "第二引用", "第三引用"])
                self.assertEqual(len({node.get("id") for node in properties}), 3)
                embeds = [node.get(q(R, "embed")) for node in document.iter(q(A, "blip"))]
                self.assertEqual(len(set(embeds)), expected_media)

    def test_preflight_reads_one_normalized_image_path_once(self):
        image_path = self.work / "pixel.png"
        image_path.write_bytes(tiny_png())
        spec = {"blocks": [
            {"type": "image", "path": "pixel.png", "width_mm": 10, "height_mm": 10, "alt": "甲"},
            {"type": "image", "path": "./pixel.png", "width_mm": 11, "height_mm": 9, "alt": "乙"},
        ]}
        opened = []
        original_open = Path.open

        def count_open(path, *args, **kwargs):
            if path.resolve() == image_path.resolve():
                opened.append(path)
            return original_open(path, *args, **kwargs)

        with mock.patch.object(Path, "open", new=count_open):
            receipt = create_docx.preflight_document(spec, self.work)
        self.assertEqual(receipt["status"], "preflight-ok")
        self.assertEqual(len(opened), 1)
        self.assertFalse(self.output.exists())

    def test_image_cache_is_per_call_and_does_not_skip_dimension_or_size_limits(self):
        image_path = self.work / "pixel.png"
        first_bytes = tiny_png()
        second_bytes = png_with_text(b"changed between calls")
        image_path.write_bytes(first_bytes)
        spec = {"blocks": [{"type": "image", "path": "pixel.png", "width_mm": 10,
                            "height_mm": 10, "alt": "图"}]}
        first = self.work / "first.docx"
        second = self.work / "second.docx"
        self.generate(spec, first)
        image_path.write_bytes(second_bytes)
        self.generate(spec, second)
        with zipfile.ZipFile(first) as package:
            first_media = package.read("word/media/image1.png")
        with zipfile.ZipFile(second) as package:
            second_media = package.read("word/media/image1.png")
        self.assertEqual(first_media, first_bytes)
        self.assertEqual(second_media, second_bytes)
        self.assertNotEqual(first_media, second_media)

        too_wide = {"blocks": [
            {"type": "image", "path": "pixel.png", "width_mm": 10, "height_mm": 10, "alt": "正常"},
            {"type": "image", "path": "./pixel.png", "width_mm": 150, "height_mm": 10, "alt": "越界"},
        ]}
        sentinel = b"existing output survives image bounds failure"
        self.output.write_bytes(sentinel)
        before = {path.name for path in self.work.iterdir()}
        with self.assertRaisesRegex(ValueError, "dimensions exceed"):
            self.generate(too_wide)
        self.assertEqual(self.output.read_bytes(), sentinel)
        self.assertEqual({path.name for path in self.work.iterdir()}, before)

        oversized = self.work / "oversized.png"
        with oversized.open("wb") as image_file:
            image_file.truncate(create_docx._MAX_IMAGE_BYTES + 1)
        too_large = {"blocks": [{"type": "image", "path": oversized.name, "width_mm": 1,
                                 "height_mm": 1, "alt": "超限"}]}
        before = {path.name for path in self.work.iterdir()}
        with self.assertRaisesRegex(ValueError, "32 MiB"):
            self.generate(too_large)
        self.assertEqual(self.output.read_bytes(), sentinel)
        self.assertEqual({path.name for path in self.work.iterdir()}, before)

    def test_reused_images_still_count_toward_the_cumulative_package_limit_atomically(self):
        image_path = self.work / "large.png"
        image_bytes = tiny_png()
        image_path.write_bytes(image_bytes)
        spec = {"blocks": [
            {"type": "image", "path": "large.png", "width_mm": 1, "height_mm": 1, "alt": str(index)}
            for index in range(5)
        ]}
        sentinel = b"existing output survives package limit"
        self.output.write_bytes(sentinel)
        before = {path.name for path in self.work.iterdir()}
        opened = []
        original_open = Path.open

        def count_open(path, *args, **kwargs):
            if path.resolve() == image_path.resolve():
                opened.append(path)
            return original_open(path, *args, **kwargs)

        with mock.patch.object(Path, "open", new=count_open):
            with mock.patch.object(create_docx, "_MAX_PACKAGE_BYTES", len(image_bytes) * 4):
                with self.assertRaisesRegex(ValueError, "128 MiB"):
                    self.generate(spec, verify=False)
        self.assertEqual(len(opened), 1)
        self.assertEqual(self.output.read_bytes(), sentinel)
        self.assertEqual({path.name for path in self.work.iterdir()}, before)

    def test_verify_accepts_a_mixed_theme_document_with_intended_layout(self):
        image_path = self.work / "pixel.png"
        image_path.write_bytes(tiny_png())
        spec = {
            "theme": "business",
            "page": {"margin_top_mm": 20, "margin_bottom_mm": 22,
                     "margin_left_mm": 25, "margin_right_mm": 25},
            "blocks": [
                {"type": "title", "text": "综合报告"},
                {"type": "heading", "level": 1, "text": "概况"},
                {"type": "paragraph", "text": "中英文 mixed text"},
                {"type": "list", "ordered": True, "list_id": "items", "items": ["第一项"]},
                {"type": "list", "ordered": True, "list_id": "items", "level": 1,
                 "items": ["子项"]},
                {"type": "page_break"},
                {"type": "table", "header": True, "widths_mm": [40, 80],
                 "rows": [["类别", "内容"], ["版本", "v1"]]},
                {"type": "image", "path": "pixel.png", "width_mm": 10,
                 "height_mm": 5, "alt": "示意图"},
            ],
        }
        self.generate(spec)
        self.assertGreater(self.output.stat().st_size, 0)
        with zipfile.ZipFile(self.output) as package:
            self.assertIsNone(package.testzip())

    def test_verify_rejects_semantically_wrong_but_well_formed_document_parts(self):
        page_break_spec = {"blocks": [
            {"type": "paragraph", "text": "前文"}, {"type": "page_break"},
            {"type": "paragraph", "text": "后文"},
        ]}

        def delete_page_break(root):
            node = next(node for node in root.iter(q(W, "br"))
                        if node.get(q(W, "type")) == "page")
            parent = next(parent for parent in root.iter() if node in list(parent))
            parent.remove(node)
            return True

        self.assert_document_fault_rejected(page_break_spec, delete_page_break)

        heading_spec = {"blocks": [{"type": "heading", "level": 1, "text": "章节"}]}

        def wrong_heading_style(root):
            paragraph = next(node for node in root.iter(q(W, "p")) if text_in(node) == "章节")
            paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'pStyle')}").set(q(W, "val"), "BodyText")
            return True

        self.assert_document_fault_rejected(heading_spec, wrong_heading_style)

        def wrong_heading_outline(root):
            style = next(node for node in root.findall(q(W, "style"))
                         if node.get(q(W, "styleId")) == "Heading1")
            style.find(f"./{q(W, 'pPr')}/{q(W, 'outlineLvl')}").set(q(W, "val"), "1")
            return True

        self.assert_styles_fault_rejected(heading_spec, wrong_heading_outline)

        named_lists = {"blocks": [
            {"type": "list", "ordered": True, "items": ["甲"]},
            {"type": "list", "ordered": True, "items": ["乙"]},
        ]}

        def wrong_list_instance(root):
            paragraphs = [node for node in root.iter(q(W, "p"))
                          if node.find(f"./{q(W, 'pPr')}/{q(W, 'numPr')}") is not None]
            first_id = paragraphs[0].find(f"./{q(W, 'pPr')}/{q(W, 'numPr')}/{q(W, 'numId')}")
            second_id = paragraphs[1].find(f"./{q(W, 'pPr')}/{q(W, 'numPr')}/{q(W, 'numId')}")
            second_id.set(q(W, "val"), first_id.get(q(W, "val")))
            return True

        self.assert_document_fault_rejected(named_lists, wrong_list_instance)

        def wrong_list_level(root):
            paragraph = next(node for node in root.iter(q(W, "p")) if text_in(node) == "子项")
            paragraph.find(f"./{q(W, 'pPr')}/{q(W, 'numPr')}/{q(W, 'ilvl')}")\
                .set(q(W, "val"), "0")
            return True

        self.assert_document_fault_rejected({"blocks": [
            {"type": "list", "ordered": True, "list_id": "nested", "items": ["根"]},
            {"type": "list", "ordered": True, "list_id": "nested", "level": 1, "items": ["子项"]},
        ]}, wrong_list_level)

        table_spec = {"blocks": [{"type": "table", "widths_mm": [40, 80],
                                   "rows": [["左格", "右格"], ["甲", "乙"]]}]}

        def swap_table_text(root):
            cells = list(root.iter(q(W, "tc")))[:2]
            left = next(cells[0].iter(q(W, "t")))
            right = next(cells[1].iter(q(W, "t")))
            left.text, right.text = right.text, left.text
            return True

        self.assert_document_fault_rejected(table_spec, swap_table_text)

        def change_column_widths(root):
            grid_widths = list(root.iter(q(W, "gridCol")))
            original_widths = [node.get(q(W, "w")) for node in grid_widths]
            for node, value in zip(grid_widths, reversed(original_widths)):
                node.set(q(W, "w"), value)
            for row in root.iter(q(W, "tr")):
                cell_widths = [cell.find(f"./{q(W, 'tcPr')}/{q(W, 'tcW')}")
                               for cell in row.findall(q(W, "tc"))]
                original = [node.get(q(W, "w")) for node in cell_widths]
                for node, value in zip(cell_widths, reversed(original)):
                    node.set(q(W, "w"), value)
            return True

        self.assert_document_fault_rejected(table_spec, change_column_widths)

        def remove_table_row(root):
            table = root.find(f".//{q(W, 'tbl')}")
            table.remove(table.findall(q(W, "tr"))[-1])
            return True

        self.assert_document_fault_rejected(table_spec, remove_table_row)

        image_path = self.work / "pixel.png"
        image_path.write_bytes(tiny_png())
        image_spec = {"blocks": [{"type": "image", "path": "pixel.png", "width_mm": 10,
                                  "height_mm": 5, "alt": "示意图"}]}

        def change_picture_extent(root):
            extent = next(root.iter(q(WP, "extent")))
            extents = [extent, *list(root.iter(q(A, "ext")))]
            for node in extents:
                node.set("cx", str(int(node.get("cx")) // 2))
                node.set("cy", str(int(node.get("cy")) // 2))
            return True

        self.assert_document_fault_rejected(image_spec, change_picture_extent)

        def change_picture_alt(root):
            properties = next(root.iter(q(WP, "docPr")))
            properties.set("descr", "错误替代文字")
            return True

        self.assert_document_fault_rejected(image_spec, change_picture_alt)

        image_two_path = self.work / "second.png"
        image_two_path.write_bytes(png_with_text(b"second image"))
        media_spec = {"blocks": [
            {"type": "image", "path": "pixel.png", "width_mm": 10, "height_mm": 5, "alt": "第一图"},
            {"type": "image", "path": "second.png", "width_mm": 10, "height_mm": 5, "alt": "第二图"},
        ]}

        def swap_image_relationships(root):
            embeds = list(root.iter(q(A, "blip")))
            left, right = embeds[0].get(q(R, "embed")), embeds[1].get(q(R, "embed"))
            embeds[0].set(q(R, "embed"), right)
            embeds[1].set(q(R, "embed"), left)
            return True

        self.assert_document_fault_rejected(media_spec, swap_image_relationships)

        margin_spec = {"page": {"margin_left_mm": 20}, "blocks": []}

        def change_page_margin(root):
            margins = next(root.iter(q(W, "pgMar")))
            margins.set(q(W, "left"), str(int(margins.get(q(W, "left"))) + 20))
            return True

        self.assert_document_fault_rejected(margin_spec, change_page_margin)

        def change_style_size(root):
            body = next(node for node in root.findall(q(W, "style"))
                        if node.get(q(W, "styleId")) == "BodyText")
            body.find(f"./{q(W, 'rPr')}/{q(W, 'sz')}").set(q(W, "val"), "34")
            return True

        self.assert_styles_fault_rejected({"blocks": [{"type": "paragraph", "text": "样式"}]},
                                          change_style_size)


if __name__ == "__main__":
    unittest.main()
