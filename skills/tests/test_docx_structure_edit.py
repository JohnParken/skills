"""Controlled structure edits must preserve all non-target package content."""

import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "docx/scripts"))
from create_docx import create_document
from edit_structure import edit_structure
from office import xmlio
from test_docx_stdlib_core import W, package_parts, write_package


class StructureEditTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-structure-test-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source, self.output = self.root / "source.docx", self.root / "edited.docx"
        create_document({"blocks": [
            {"type": "paragraph", "runs": [{"text": "中文"}, {"text": " English", "italic": True}]},
            {"type": "paragraph", "text": "下一段"},
            {"type": "table", "rows": [["表格保持原样"]]},
        ]}, self.source)

    def insertion(self, **updates):
        operation = {"type": "insert_paragraph", "paragraph_index": 0, "expected_text": "中文 English",
                     "position": "after", "text": "新插入段落 & “引号”", "style": "BodyText"}
        operation.update(updates)
        return operation

    def formatting(self, **updates):
        operation = {"type": "format_run", "paragraph_index": 0, "expected_text": "中文 English",
                     "run_index": 1, "expected_run_text": " English",
                     "format": {"bold": True, "italic": False, "size_pt": 12.5,
                                "color": "0aBcDf", "font_east_asia": "仿宋", "font_latin": "Arial"}}
        operation.update(updates)
        return operation

    def run_edit(self, *operations):
        return edit_structure(self.source, self.output, {"operations": list(operations)})

    def rejected(self, *operations):
        before = self.source.read_bytes()
        self.output.write_bytes(b"keep")
        with self.assertRaises(ValueError):
            self.run_edit(*operations)
        self.assertEqual(self.source.read_bytes(), before)
        self.assertEqual(self.output.read_bytes(), b"keep")

    def test_insert_and_format_simultaneously_using_original_indices(self):
        self.run_edit(self.insertion(), self.formatting(),
                      self.insertion(paragraph_index=1, expected_text="下一段", position="before", text="说明段"))
        with zipfile.ZipFile(self.source) as before, zipfile.ZipFile(self.output) as after:
            self.assertIsNone(after.testzip())
            self.assertEqual(set(before.namelist()), set(after.namelist()))
            for name in before.namelist():
                if name != "word/document.xml":
                    self.assertEqual(before.read(name), after.read(name), name)
            original = xmlio.fromstring(before.read("word/document.xml"))
            edited = xmlio.fromstring(after.read("word/document.xml"))
        body = edited.find(f"{{{W}}}body")
        paragraphs = body.findall(f"{{{W}}}p")
        self.assertEqual(["".join(node.text or "" for node in p.iter(f"{{{W}}}t")) for p in paragraphs],
                         ["中文 English", "新插入段落 & “引号”", "说明段", "下一段"])
        properties = paragraphs[0].findall(f"{{{W}}}r")[1].find(f"{{{W}}}rPr")
        values = {node.tag: node.get(f"{{{W}}}val") for node in properties}
        self.assertEqual(values[f"{{{W}}}b"], "1")
        self.assertEqual(values[f"{{{W}}}i"], "0")
        self.assertEqual(values[f"{{{W}}}sz"], "25")
        self.assertEqual(values[f"{{{W}}}color"], "0ABCDF")
        self.assertEqual(xmlio.ET.tostring(next(original.iter(f"{{{W}}}tbl"))),
                         xmlio.ET.tostring(next(edited.iter(f"{{{W}}}tbl"))))

    def test_bad_inputs_and_unknown_styles_preserve_existing_files(self):
        for operation in (
            self.insertion(paragraph_index=True), self.insertion(paragraph_index=-1),
            self.insertion(expected_text="错位"), self.insertion(position=[]),
            self.insertion(text="bad\ntext"), self.insertion(style="MissingStyle"),
            self.insertion(type=[]), self.insertion(unknown=True),
            self.formatting(run_index=True), self.formatting(expected_run_text="wrong"),
            self.formatting(format={"bold": "false"}), self.formatting(format={"size_pt": float("nan")}),
            self.formatting(format={"size_pt": 10 ** 1000}),
            self.formatting(format={"size_pt": 12.25}), self.formatting(format={"color": "#123456"}),
            self.formatting(format={"font_latin": "bad\tfont"}), self.formatting(format={"unknown": 1}),
        ):
            with self.subTest(operation=operation):
                self.rejected(operation)
        self.rejected(self.insertion(), self.insertion())
        self.rejected(self.formatting(), self.formatting())

    def test_complex_anchors_and_cross_paragraph_ranges_are_rejected(self):
        for body in (
            '<w:p><w:bookmarkStart w:id="1" w:name="a"/><w:r><w:t>anchor</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>',
            '<w:p><w:r><w:tab/><w:t>anchor</w:t></w:r></w:p>',
            '<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>anchor</w:t></w:r></w:p>',
            '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p>'
            '<w:p><w:r><w:t>anchor</w:t></w:r></w:p>'
            '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
            '<w:p><w:moveToRangeStart w:id="1"/></w:p><w:p><w:r><w:t>anchor</w:t></w:r></w:p>'
            '<w:p><w:moveToRangeEnd w:id="1"/></w:p>',
        ):
            with self.subTest(body=body):
                write_package(self.source, package_parts(f'<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>'))
                index = 1 if body.count('<w:p>') > 1 else 0
                self.rejected(self.insertion(paragraph_index=index, expected_text="anchor", style=None))

    def test_format_replaces_theme_overrides_and_handles_non_w_prefixes(self):
        document = (f'<q:document xmlns:q="{W}"><q:body><q:p><q:r><q:rPr>'
                    '<q:rFonts q:asciiTheme="minorHAnsi" q:eastAsiaTheme="minorEastAsia"/>'
                    '<q:color q:val="000000" q:themeColor="accent1"/>'
                    '</q:rPr><q:t>anchor</q:t></q:r></q:p></q:body></q:document>')
        write_package(self.source, package_parts(document))
        self.run_edit(self.formatting(expected_text="anchor", run_index=0, expected_run_text="anchor"))
        with zipfile.ZipFile(self.output) as package:
            tree = xmlio.fromstring(package.read("word/document.xml"))
            fonts = next(tree.iter(f"{{{W}}}rFonts"))
            color = next(tree.iter(f"{{{W}}}color"))
            self.assertNotIn(f"{{{W}}}asciiTheme", fonts.attrib)
            self.assertNotIn(f"{{{W}}}eastAsiaTheme", fonts.attrib)
            self.assertNotIn(f"{{{W}}}themeColor", color.attrib)
            self.assertEqual(color.get(f"{{{W}}}val"), "0ABCDF")

    def test_default_namespace_and_existing_attribute_qname_are_preserved(self):
        document = (f'<document xmlns="{W}" xmlns:q="{W}"><body><p><r><rPr>'
                    '<color q:val="000000" q:themeColor="accent1"/>'
                    '</rPr><t>anchor</t></r></p></body></document>')
        write_package(self.source, package_parts(document))
        self.run_edit(self.formatting(expected_text="anchor", run_index=0, expected_run_text="anchor"))
        with zipfile.ZipFile(self.output) as package:
            dom = xmlio.parse_dom(package.read("word/document.xml"))
            color = dom.getElementsByTagNameNS(W, "color")[0]
            self.assertEqual(color.getAttributeNS(W, "val"), "0ABCDF")
            self.assertEqual(color.getAttributeNodeNS(W, "val").name, "q:val")

    def test_cross_paragraph_bookmark_permission_and_custom_xml_ranges_are_rejected(self):
        for start, end in (("bookmarkStart", "bookmarkEnd"), ("permStart", "permEnd"),
                           ("customXmlInsRangeStart", "customXmlInsRangeEnd")):
            with self.subTest(start=start):
                document = (f'<w:document xmlns:w="{W}"><w:body><w:p><w:{start} w:id="7"/></w:p>'
                            '<w:p><w:r><w:t>anchor</w:t></w:r></w:p>'
                            f'<w:p><w:{end} w:id="7"/></w:p></w:body></w:document>')
                write_package(self.source, package_parts(document))
                self.rejected(self.insertion(paragraph_index=1, expected_text="anchor"))
                self.rejected(self.formatting(paragraph_index=1, expected_text="anchor",
                                              run_index=0, expected_run_text="anchor"))

    def test_same_file_aliases_and_extension_changes_are_rejected(self):
        with self.assertRaises(ValueError):
            edit_structure(self.source, self.source, {"operations": [self.insertion()]})
        alias = self.root / "alias.docx"
        alias.symlink_to(self.source)
        with self.assertRaises(ValueError):
            edit_structure(self.source, alias, {"operations": [self.insertion()]})
        with self.assertRaises(ValueError):
            edit_structure(self.source, self.output.with_suffix('.dotx'), {"operations": [self.insertion()]})

    def test_insert_requires_paragraph_style_not_character_or_table_style(self):
        with zipfile.ZipFile(self.source) as package:
            parts = {name: package.read(name) for name in package.namelist()}
        styles = xmlio.parse_dom(parts["word/styles.xml"])
        for kind in ("character", "table", "numbering"):
            node = styles.createElementNS(W, "w:style")
            node.setAttributeNS(W, "w:type", kind)
            node.setAttributeNS(W, "w:styleId", kind + "Only")
            styles.documentElement.appendChild(node)
        parts["word/styles.xml"] = styles.toxml(encoding="UTF-8")
        write_package(self.source, parts)
        for kind in ("character", "table", "numbering"):
            with self.subTest(kind=kind):
                self.rejected(self.insertion(style=kind + "Only"))

    def test_cli_works_without_external_programs(self):
        (self.root / "operations.json").write_text(json.dumps({"operations": [self.insertion()]}), encoding="utf-8")
        result = subprocess.run(
            [sys.executable, "-I", "-S", "-B", str(ROOT / "docx/scripts/edit_structure.py"),
             "source.docx", "operations.json", "edited.docx"],
            cwd=self.root, env=dict(os.environ, PATH=""), capture_output=True, text=True,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertTrue(self.output.is_file())


if __name__ == "__main__":
    unittest.main()
