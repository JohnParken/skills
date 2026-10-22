"""User-facing generation, atomic verification and resource budget regressions."""

import contextlib
import io
import json
import os
import struct
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "docx/scripts"
sys.path.insert(0, str(SCRIPTS))
import create_docx
from office import validate, xmlio
from office.archiveio import open_archive
from office.validators.structure import StructureValidator
from test_docx_generator import tiny_png, png_chunk

W = create_docx.W


def w(name):
    return f"{{{W}}}{name}"


class DocxExperienceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-ux-")
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name)
        self.output = self.work / "output.docx"
        self.env = dict(os.environ, PATH="", PYTHONDONTWRITEBYTECODE="1")
        (self.work / "pixel.png").write_bytes(tiny_png())

    def cli(self, *args):
        return subprocess.run([sys.executable, "-I", "-S", "-B", str(SCRIPTS / "create_docx.py"),
                               *map(str, args)], cwd=self.work, env=self.env,
                              capture_output=True, text=True, check=False)

    def spec_file(self, spec):
        source = self.work / "input.json"
        source.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
        return source

    def generate(self, spec, **kwargs):
        return create_docx.create_document(spec, self.output, self.work, verify=True, **kwargs)

    def part(self, name="word/document.xml"):
        with zipfile.ZipFile(self.output) as archive:
            return ET.fromstring(archive.read(name))

    def assert_rejected(self, spec, pattern, **kwargs):
        self.output.write_bytes(b"existing output")
        before = set(self.work.iterdir())
        with self.assertRaisesRegex(ValueError, pattern):
            self.generate(spec, **kwargs)
        self.assertEqual(self.output.read_bytes(), b"existing output")
        self.assertEqual(set(self.work.iterdir()), before)

    def test_integrated_cli_verifies_mixed_content_and_reports_limits(self):
        spec = {"theme": "business", "header": "页眉", "footer": "页脚", "page_numbers": True,
                "blocks": [{"type": "title", "text": "测试报告"},
                           {"type": "subtitle", "text": "中文 & English"},
                           {"type": "paragraph", "runs": [{"text": "引号\""}, {"text": "“正文”", "bold": True}]},
                           {"type": "table", "header": True, "rows": [["项目", "状态"], ["测试", "通过"]]},
                           {"type": "image", "path": "pixel.png", "width_mm": 10, "alt": "图"},
                           {"type": "list", "ordered": True, "items": ["甲", "乙"]}]}
        result = self.cli(self.spec_file(spec), "output.docx", "--verify", "--require-no-placeholders")
        self.assertEqual(result.returncode, 0, result.stderr)
        receipt = json.loads(result.stdout)
        self.assertTrue(receipt["content_coverage"])
        self.assertTrue(receipt["zip_crc"])
        self.assertTrue(receipt["layout_intent"])
        self.assertFalse(receipt["visual_validation"])
        self.assertFalse(receipt["xsd_validation"])
        with zipfile.ZipFile(self.output) as archive:
            self.assertIsNone(archive.testzip())
        self.assertEqual(len(self.part().findall(".//" + w("tbl"))), 1)

    def test_preflight_does_not_create_output_or_probe_files(self):
        source = self.spec_file({"blocks": [{"type": "paragraph", "text": "预检"}]})
        before = set(self.work.iterdir())
        result = self.cli(source, "--preflight")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)["status"], "preflight-ok")
        self.assertEqual(set(self.work.iterdir()), before)

    def test_preflight_and_option_conflicts_fail(self):
        source = self.spec_file({"blocks": []})
        for options in (("--preflight", "out.docx"), ("--preflight", "--verify"),
                        ("--check", "--verify"), ("--preflight", "--require-no-placeholders")):
            with self.subTest(options=options):
                self.assertEqual(self.cli(source, *options).returncode, 2)

    def test_unfilled_placeholders_preserve_output(self):
        self.assert_rejected({"blocks": [{"type": "paragraph", "runs": [
            {"text": "{{ 公"}, {"text": "司 }}"}]}]}, "Unfilled placeholder", require_no_placeholders=True)

    def test_verification_detects_dropped_text_before_publishing(self):
        original = create_docx._document_xml

        def dropped(spec):
            data, images, lists = original(spec)
            return data.replace("正文".encode(), "遗漏".encode()), images, lists

        with mock.patch.object(create_docx, "_document_xml", side_effect=dropped):
            self.assert_rejected({"blocks": [{"type": "paragraph", "text": "正文"}]}, "content coverage")

    def test_generated_budget_is_shared_and_preserves_output_without_verify(self):
        self.output.write_bytes(b"old")
        with mock.patch.object(xmlio, "MAX_XML_ELEMENTS", 10):
            with self.assertRaisesRegex(ValueError, "generated package.*elements"):
                create_docx.create_document({"blocks": []}, self.output)
        self.assertEqual(self.output.read_bytes(), b"old")

    def test_real_long_document_is_rejected_before_success_or_output(self):
        self.assert_rejected({"blocks": [{"type": "paragraph", "text": "x"}] * 41000}, "elements limit")

    def test_archive_validation_does_not_extract(self):
        self.generate({"blocks": [{"type": "paragraph", "text": "内存校验"}]})
        with mock.patch.object(validate.tempfile, "TemporaryDirectory", side_effect=AssertionError("extraction")):
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(validate.main([str(self.output)]), 0)

    def test_archive_validation_reads_media_crc(self):
        self.generate({"blocks": [{"type": "image", "path": "pixel.png", "width_mm": 10, "height_mm": 10, "alt": "图"}]})
        broken = self.work / "broken.docx"
        with zipfile.ZipFile(self.output) as source, zipfile.ZipFile(broken, "w", zipfile.ZIP_STORED) as target:
            for name in source.namelist():
                target.writestr(name, source.read(name))
        with zipfile.ZipFile(broken) as archive:
            entry = next(info for info in archive.infolist() if info.filename.startswith("word/media/"))
        data = bytearray(broken.read_bytes())
        name_length, extra_length = struct.unpack_from("<HH", data, entry.header_offset + 26)
        data[entry.header_offset + 30 + name_length + extra_length] ^= 1
        broken.write_bytes(data)
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            self.assertEqual(validate.main([str(broken)]), 1)

    def test_directory_and_archive_rules_agree(self):
        self.generate({"blocks": [{"type": "table", "header": True, "rows": [["表头"], ["值"]]}]})
        unpacked = self.work / "unpacked"
        from office.helpers import safe_extract
        with open_archive(self.output) as archive:
            safe_extract(archive, unpacked)
            memory = StructureValidator.from_archive(archive, extension=".docx")
            self.assertTrue(memory.validate(quiet=True), memory.errors)
        directory = StructureValidator(unpacked, extension=".docx")
        self.assertTrue(directory.validate(quiet=True), directory.errors)
        self.assertEqual(memory.files, directory.files)
        self.assertEqual(memory.warnings, directory.warnings)

    def test_errors_locate_characters_and_suggest_field_names(self):
        self.assert_rejected({"blocks": [{"type": "paragraph", "text": "甲\n乙"}]}, r"blocks\[0\].text.*U\+000A.*character 2.*paragraph blocks")
        self.assert_rejected({"blocks": [], "styels": {}}, "did you mean 'styles'")

    def test_allowed_invisible_characters_remain_exact(self):
        text = "甲\u200b\ufeff\u00a0乙"
        self.generate({"blocks": [{"type": "paragraph", "text": text}]})
        self.assertEqual(next(self.part().iter(w("t"))).text, text)

    def test_cli_rejects_duplicate_json_fields(self):
        source = self.work / "input.json"
        source.write_text('{"blocks":[],"blocks":[]}', encoding="utf-8")
        result = self.cli(source, "output.docx")
        self.assertEqual(result.returncode, 2)
        self.assertIn("Duplicate JSON field", result.stderr)
        self.assertFalse(self.output.exists())

    def test_runtime_diagnostics_identify_old_expat(self):
        output = io.StringIO()
        with mock.patch.object(create_docx.expat, "EXPAT_VERSION", "2.5.0"):
            with mock.patch.object(create_docx.tempfile, "TemporaryFile", side_effect=OSError()):
                with contextlib.redirect_stdout(output), contextlib.redirect_stderr(io.StringIO()):
                    self.assertFalse(create_docx._runtime_check())
        self.assertIn("expat=2.5.0 required>=2.6.0 supported=False", output.getvalue())
        self.assertIn("python=", output.getvalue())

    def test_themes_override_without_leaking_and_titles_are_not_outline_headings(self):
        for theme in ("business", "academic", "official"):
            with self.subTest(theme=theme):
                self.generate({"theme": theme, "styles": {"body": {"size_pt": 13}},
                               "blocks": [{"type": "title", "text": "文档标题"},
                                          {"type": "heading", "level": 1, "text": "章节"}]})
                styles = self.part("word/styles.xml")
                body = next(node for node in styles if node.get(w("styleId")) == "BodyText")
                self.assertEqual(body.find("./" + w("rPr") + "/" + w("sz")).get(w("val")), "26")
                spacing = body.find("./" + w("pPr") + "/" + w("spacing"))
                self.assertEqual(spacing.get(w("lineRule")), "exact" if theme == "official" else "auto")
                title = next(node for node in styles if node.get(w("styleId")) == "Title")
                self.assertIsNone(title.find(".//" + w("outlineLvl")))
        self.assert_rejected({"theme": "unknown", "blocks": []}, "theme must")

    def test_paragraph_multiple_spacing_color_and_mode_overrides(self):
        self.generate({"styles": {"body": {"line_spacing_multiple": 1.5, "color": "123abc"}},
                       "blocks": [{"type": "paragraph", "text": "正文", "format": {"line_spacing_pt": 20}}]})
        styles = self.part("word/styles.xml")
        body = next(node for node in styles if node.get(w("styleId")) == "BodyText")
        self.assertEqual(body.find(".//" + w("spacing")).get(w("line")), "360")
        self.assertEqual(body.find(".//" + w("color")).get(w("val")), "123ABC")
        self.assertEqual(self.part().find(".//" + w("spacing")).get(w("lineRule")), "exact")
        self.assert_rejected({"styles": {"body": {"line_spacing_pt": 20, "line_spacing_multiple": 1.5}}, "blocks": []}, "cannot set both")

    def test_content_column_sizing_and_explicit_width_precedence(self):
        block = {"type": "table", "format": {"column_sizing": "content"},
                 "rows": [["序号", "详细说明"], ["1", "说明" * 50]]}
        self.generate({"blocks": [block]})
        widths = [int(node.get(w("w"))) for node in self.part().iter(w("gridCol"))]
        self.assertGreater(widths[1], widths[0])
        self.assertEqual(sum(widths), create_docx._twips(210) - 2 * create_docx._twips(31.8))
        self.generate({"blocks": [{**block, "widths_mm": [40, 60]}]})
        self.assertEqual([node.get(w("w")) for node in self.part().iter(w("gridCol"))],
                         [str(create_docx._twips(40)), str(create_docx._twips(60))])

    def test_three_line_and_zebra_presets(self):
        rows = [["表头"], ["数据一"], ["数据二"]]
        self.generate({"blocks": [{"type": "table", "header": True, "format": {"preset": "three_line"}, "rows": rows}]})
        root = self.part()
        borders = root.find(".//" + w("tblBorders"))
        self.assertEqual(borders.find(w("top")).get(w("sz")), "12")
        self.assertEqual(borders.find(w("insideV")).get(w("val")), "nil")
        self.assertEqual(root.find(".//" + w("tcBorders") + "/" + w("bottom")).get(w("sz")), "6")
        self.generate({"blocks": [{"type": "table", "header": True, "format": {"preset": "zebra"}, "rows": rows}]})
        fills = self.part().findall(".//" + w("tcPr") + "/" + w("shd"))
        self.assertEqual([node.get(w("fill")) for node in fills], ["F2F4F7"])

    def test_image_ratio_inference_and_opt_in_deduplication(self):
        image = {"type": "image", "path": "pixel.png", "width_mm": 10, "alt": "图"}
        self.generate({"deduplicate_images": True, "blocks": [image, image]})
        with zipfile.ZipFile(self.output) as archive:
            self.assertEqual(len([name for name in archive.namelist() if name.startswith("word/media/")]), 1)
        extents = self.part().findall(f".//{{{create_docx.WP}}}extent")
        self.assertEqual(len(extents), 2)
        self.assertTrue(all(node.get("cx") == node.get("cy") for node in extents))
        self.assertEqual(len({node.get("id") for node in self.part().iter(f"{{{create_docx.WP}}}docPr")}), 2)

    def test_image_height_inference_and_invalid_header(self):
        data = tiny_png()
        # A dimension header is sufficient for ratio inference, not pixel decoding.
        data = data[:8] + png_chunk(b"IHDR", struct.pack(">IIBBBBB", 2, 1, 8, 6, 0, 0, 0)) + data[33:]
        (self.work / "wide.png").write_bytes(data)
        self.generate({"blocks": [{"type": "image", "path": "wide.png", "height_mm": 10, "alt": "图"}]})
        extent = self.part().find(f".//{{{create_docx.WP}}}extent")
        self.assertEqual(int(extent.get("cx")), 2 * int(extent.get("cy")))
        (self.work / "invalid.png").write_bytes(b"\x89PNG\r\n\x1a\n\x00\x00\x00\x00IEND\xaeB\x60\x82")
        self.assert_rejected({"blocks": [{"type": "image", "path": "invalid.png", "width_mm": 10, "alt": "图"}]}, "cannot infer image ratio")

    def test_jpeg_dimension_header_and_inferred_page_limits(self):
        # SOF0 with 20 x 10 dimensions; only header parsing is promised.
        data = b"\xff\xd8\xff\xc0" + struct.pack(">HBHHB", 11, 8, 10, 20, 1) + b"\x01\x11\x00\xff\xd9"
        (self.work / "header.jpg").write_bytes(data)
        self.generate({"blocks": [{"type": "image", "path": "header.jpg", "width_mm": 20, "alt": "JPEG"}]})
        extent = self.part().find(f".//{{{create_docx.WP}}}extent")
        self.assertEqual(int(extent.get("cx")), 2 * int(extent.get("cy")))
        self.assert_rejected({"blocks": [{"type": "image", "path": "header.jpg", "height_mm": 100, "alt": "JPEG"}]}, "dimensions exceed")

    def test_preset_overrides_inherited_preset_and_explicit_edges_win(self):
        self.generate({"table_format": {"preset": "three_line"}, "blocks": [{
            "type": "table", "format": {"preset": "zebra", "borders": {"left": {"style": "single", "width_pt": 2}}},
            "rows": [["甲"], ["乙"]]}]})
        borders = self.part().find(".//" + w("tblBorders"))
        self.assertEqual(borders.find(w("top")).get(w("val")), "nil")
        self.assertEqual(borders.find(w("left")).get(w("sz")), "16")

    def test_packed_xml_budget_remains_cumulative(self):
        self.generate({"blocks": []})
        with open_archive(self.output) as archive:
            validator = StructureValidator.from_archive(archive, extension=".docx")
            with mock.patch.object(xmlio, "MAX_XML_ELEMENTS", 10):
                self.assertFalse(validator.validate(quiet=True))
            self.assertTrue(any("shared elements" in error for error in validator.errors))

    def test_self_check_documentation_keeps_scripts_entry(self):
        instructions = (ROOT / "docx/SKILL.md").read_text(encoding="utf-8")
        self.assertIn("python -I -S -B scripts/create_docx.py --check", instructions)
        self.assertNotIn("docx/scripts/create_docx.py --check", instructions)

    def test_verification_checks_media_bytes_after_valid_zip_and_crc(self):
        original = create_docx._package_parts

        def wrong_media(spec):
            parts = original(spec)
            media = next(name for name in parts if name.startswith("word/media/"))
            data = parts[media]
            parts[media] = data[:-12] + png_chunk(b"tEXt", b"wrong image") + data[-12:]
            return parts

        with mock.patch.object(create_docx, "_package_parts", side_effect=wrong_media):
            self.assert_rejected({"blocks": [{"type": "image", "path": "pixel.png",
                                              "width_mm": 10, "alt": "图"}]}, "layout mismatch.*media")

    def test_verification_preserves_cell_membership_not_only_flat_text_order(self):
        original = create_docx._document_xml

        def move_text(spec):
            data, images, lists = original(spec)
            root = ET.fromstring(data)
            cells = root.findall(".//" + w("tc"))
            source = cells[1].find(w("p"))
            run = source.find(w("r"))
            source.remove(run)
            cells[0].find(w("p")).append(run)
            return ET.tostring(root, encoding="UTF-8", xml_declaration=True), images, lists

        with mock.patch.object(create_docx, "_document_xml", side_effect=move_text):
            self.assert_rejected({"blocks": [{"type": "table", "rows": [["甲", "乙"]]}]},
                                 "layout mismatch")

    def test_verification_reports_missing_generation_parts_without_traceback(self):
        original = create_docx._package_parts

        def missing_styles(spec):
            parts = original(spec)
            del parts["word/styles.xml"]
            return parts

        with mock.patch.object(create_docx, "_package_parts", side_effect=missing_styles):
            self.assert_rejected({"blocks": []}, "failed verification")

    def assert_related_style(self, part, style_id, font_east_asia, font_latin):
        paragraph = self.part(part).find(w("p"))
        self.assertEqual(paragraph.find(f"{w('pPr')}/{w('pStyle')}").get(w("val")), style_id)
        styles = self.part("word/styles.xml")
        style = next(node for node in styles if node.get(w("styleId")) == style_id)
        fonts = style.find(f"{w('rPr')}/{w('rFonts')}")
        self.assertEqual(fonts.get(w("eastAsia")), font_east_asia)
        for script in ("ascii", "hAnsi", "cs"):
            self.assertEqual(fonts.get(w(script)), font_latin)
        self.assertEqual(style.find(f"{w('rPr')}/{w('sz')}").get(w("val")), "21")
        self.assertEqual(style.find(f"{w('rPr')}/{w('b')}").get(w("val")), "0")
        self.assertEqual(style.find(f"{w('rPr')}/{w('color')}").get(w("val")), "000000")
        props = style.find(w("pPr"))
        self.assertEqual(props.find(w("spacing")).attrib,
                         {w("before"): "0", w("after"): "0", w("line"): "300", w("lineRule"): "auto"})
        self.assertEqual(props.find(w("ind")).get(w("firstLineChars")), "0")
        self.assertEqual(props.find(w("jc")).get(w("val")), "left")
        # Text and field-result runs inherit their common paragraph style.
        self.assertTrue(list(paragraph.iter(w("r"))))
        self.assertFalse(list(paragraph.iter(w("rPr"))))

    def test_business_and_academic_headers_footers_follow_theme(self):
        for theme, east_asia, latin in (("business", "微软雅黑", "Calibri"),
                                        ("academic", "宋体", "Times New Roman")):
            with self.subTest(theme=theme):
                self.generate({"theme": theme, "header": "页眉 Header", "footer": "页脚 Footer",
                               "page_numbers": True, "blocks": [{"type": "paragraph", "text": "正文"}]})
                self.assert_related_style("word/header1.xml", "Header", east_asia, latin)
                self.assert_related_style("word/footer1.xml", "Footer", east_asia, latin)
                footer = self.part("word/footer1.xml")
                self.assertEqual([node.text for node in footer.iter(w("instrText"))], [" PAGE "])
                self.assertEqual([node.text for node in footer.iter(w("t"))], ["页脚 Footer", " ", "1"])

    def test_page_number_only_footer_uses_themed_style(self):
        for theme, east_asia, latin in (("business", "微软雅黑", "Calibri"),
                                        ("academic", "宋体", "Times New Roman")):
            with self.subTest(theme=theme):
                self.generate({"theme": theme, "page_numbers": True, "blocks": []})
                self.assert_related_style("word/footer1.xml", "Footer", east_asia, latin)
                with zipfile.ZipFile(self.output) as archive:
                    self.assertNotIn("word/header1.xml", archive.namelist())

    def test_body_overrides_do_not_leak_into_related_theme_styles(self):
        for theme, east_asia, latin in (("business", "微软雅黑", "Calibri"),
                                        ("academic", "宋体", "Times New Roman")):
            with self.subTest(theme=theme):
                self.generate({"theme": theme, "header": "页眉", "footer": "页脚", "styles": {
                    "body": {"font_east_asia": "楷体", "font_latin": "Arial", "size_pt": 24,
                             "bold": True, "align": "justify", "line_spacing_pt": 40,
                             "first_line_chars": 3, "space_before_pt": 12, "space_after_pt": 18}}, "blocks": []})
                self.assert_related_style("word/header1.xml", "Header", east_asia, latin)
                self.assert_related_style("word/footer1.xml", "Footer", east_asia, latin)

    def test_official_related_format_and_cross_call_isolation_are_preserved(self):
        for theme in ("business", "academic", "official"):
            self.generate({"theme": theme, "header": "页眉", "footer": "页脚", "page_numbers": True, "blocks": []})
        for part in ("word/header1.xml", "word/footer1.xml"):
            self.assertIsNone(self.part(part).find(f"{w('p')}/{w('pPr')}"))
        styles = self.part("word/styles.xml")
        normal = next(node for node in styles if node.get(w("styleId")) == "Normal")
        self.assertEqual(normal.find(f"{w('rPr')}/{w('rFonts')}").get(w("eastAsia")), "仿宋")
        self.assertEqual(normal.find(f"{w('rPr')}/{w('sz')}").get(w("val")), "32")
        self.assertEqual(normal.find(f"{w('pPr')}/{w('spacing')}").get(w("line")), "276")

    def test_empty_related_text_and_omitted_parts_remain_distinct(self):
        self.generate({"theme": "academic", "header": "", "footer": "", "blocks": []})
        for part, style_id in (("word/header1.xml", "Header"), ("word/footer1.xml", "Footer")):
            paragraph = self.part(part).find(w("p"))
            self.assertEqual(paragraph.find(f"{w('pPr')}/{w('pStyle')}").get(w("val")), style_id)
            self.assertFalse(list(paragraph.iter(w("t"))))
        self.generate({"theme": "academic", "blocks": []})
        with zipfile.ZipFile(self.output) as archive:
            self.assertNotIn("word/header1.xml", archive.namelist())
            self.assertNotIn("word/footer1.xml", archive.namelist())


if __name__ == "__main__":
    unittest.main()
