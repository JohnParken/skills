"""Adversarial fixtures for controlled editing and XML value preservation."""

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "docx/scripts"))
from edit_docx import replace_text
from office import xmlio
from test_docx_stdlib_core import W, package_parts, write_package

XML = "http://www.w3.org/XML/1998/namespace"


class EditBoundaryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-edit-boundaries-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / "input.docx"
        self.output = self.root / "output.docx"

    def document(self, body, *, attrs=""):
        write_package(self.source, package_parts(
            f'<w:document xmlns:w="{W}" xmlns:x="urn:vendor" {attrs}>'
            f'<w:body>{body}</w:body></w:document>'
        ))

    def replace(self, new="中文 & English"):
        return replace_text(self.source, self.output, {
            "replacements": [{"old": "{{name}}", "new": new}],
        })

    def rejected(self, pattern="unsupported|preserv|space"):
        before = self.source.read_bytes()
        self.output.write_bytes(b"keep existing output")
        with self.assertRaisesRegex(ValueError, pattern):
            self.replace()
        self.assertEqual(self.source.read_bytes(), before)
        self.assertEqual(self.output.read_bytes(), b"keep existing output")

    def test_ambiguous_unpreserved_edge_whitespace_is_rejected(self):
        for attrs in ("", 'xml:space="default"'):
            for text in ("  A {{name}} B  ", "\t{{name}}\n", "\r{{name}}"):
                with self.subTest(attrs=attrs, text=text):
                    self.document(f'<w:p><w:r><w:t {attrs}>{text}</w:t></w:r></w:p>')
                    self.rejected()

    def test_explicit_and_inherited_preserve_keep_edge_whitespace(self):
        for document_attrs, paragraph_attrs, text_attrs in (
            ("", "", 'xml:space="preserve"'),
            ('xml:space="preserve"', "", ""),
            ('xml:space="default"', 'xml:space="preserve"', ""),
        ):
            with self.subTest(attrs=(document_attrs, paragraph_attrs, text_attrs)):
                self.document(f'<w:p {paragraph_attrs}><w:r><w:t {text_attrs}>  A {{{{name}}}} B  </w:t></w:r></w:p>',
                              attrs=document_attrs)
                self.replace()
                with zipfile.ZipFile(self.output) as package:
                    tree = xmlio.fromstring(package.read("word/document.xml"))
                    self.assertEqual(next(tree.iter(f"{{{W}}}t")).text, "  A 中文 & English B  ")

    def test_default_override_and_invalid_space_are_rejected(self):
        for value in ("default", "unknown"):
            with self.subTest(value=value):
                self.document(f'<w:p xml:space="{value}"><w:r><w:t> {{{{name}}}} </w:t></w:r></w:p>',
                              attrs='xml:space="preserve"')
                self.rejected()

    def test_new_edge_spaces_and_non_xml_spaces_are_preserved(self):
        self.document('<w:p><w:r><w:t>\u3000{{name}}\u00a0</w:t></w:r></w:p>')
        self.replace(" 新值 ")
        with zipfile.ZipFile(self.output) as package:
            tree = xmlio.fromstring(package.read("word/document.xml"))
            node = next(tree.iter(f"{{{W}}}t"))
            self.assertEqual(node.text, "\u3000 新值 \u00a0")
            self.assertEqual(node.get(f"{{{XML}}}space"), "preserve")

    def test_empty_replacement_and_standalone_declaration_are_supported(self):
        self.document('<w:p><w:r><w:t>{{name}}</w:t></w:r></w:p>')
        with zipfile.ZipFile(self.source) as package:
            parts = {name: package.read(name) for name in package.namelist()}
        parts["word/document.xml"] = b'<?xml version="1.0" standalone="yes"?>' + parts["word/document.xml"]
        write_package(self.source, parts)
        self.replace("")
        with zipfile.ZipFile(self.output) as package:
            dom = xmlio.parse_dom(package.read("word/document.xml"))
            self.assertTrue(dom.standalone)
            self.assertEqual(list(dom.getElementsByTagNameNS(W, "t")[0].childNodes), [])

    def test_move_ranges_are_rejected_in_and_across_paragraphs(self):
        for kind in ("moveFrom", "moveTo"):
            for middle in ('<w:r><w:t>{{name}}</w:t></w:r>',
                           '</w:p><w:p><w:r><w:t>{{name}}</w:t></w:r></w:p><w:p>'):
                with self.subTest(kind=kind, middle=middle):
                    self.document(f'<w:p><w:{kind}RangeStart w:id="7" w:name="move"/>'
                                  f'{middle}<w:{kind}RangeEnd w:id="7"/></w:p>')
                    self.rejected()

    def test_malformed_move_ranges_fail_closed(self):
        for markers in (
            '<w:moveFromRangeStart w:id="1"/>',
            '<w:moveFromRangeEnd w:id="1"/>',
            '<w:moveToRangeStart/><w:moveToRangeEnd/>',
            '<w:moveFromRangeStart w:id="1"/><w:moveFromRangeStart w:id="1"/><w:moveFromRangeEnd w:id="1"/>',
        ):
            with self.subTest(markers=markers):
                self.document(f'<w:p>{markers}</w:p><w:p><w:r><w:t>{{{{name}}}}</w:t></w:r></w:p>')
                self.rejected()

    def test_plain_text_after_movement_range_can_be_edited(self):
        self.document('<w:p><w:moveToRangeStart w:id="1" w:name="move"/>'
                      '<w:r><w:t>移动文字</w:t></w:r><w:moveToRangeEnd w:id="1"/></w:p>'
                      '<w:p><w:r><w:t>{{name}}</w:t></w:r></w:p>')
        self.replace()
        with zipfile.ZipFile(self.output) as package:
            tree = xmlio.fromstring(package.read("word/document.xml"))
            self.assertEqual([node.text for node in tree.iter(f"{{{W}}}t")],
                             ["移动文字", "中文 & English"])

    def test_non_target_xml_values_cannot_change_on_serialization(self):
        for extension in (
            '<x:data x:val="first&#10;second&#9;third&#13;last">vendor</x:data>',
            '<x:data>A&#13;B</x:data>',
            '<w:p x:val="first&#10;second"><w:r><w:t>untouched</w:t></w:r></w:p>',
        ):
            with self.subTest(extension=extension):
                self.document(extension + '<w:p><w:r><w:t>{{name}}</w:t></w:r></w:p>')
                self.rejected()

    def test_target_comments_processing_instructions_and_extensions_are_rejected(self):
        for text, siblings in (
            ('{{name}}<!--keep-->', ""),
            ('{{name}}<?keep data?>', ""),
            ('{{name}}', '<x:binding/>'),
        ):
            with self.subTest(text=text, siblings=siblings):
                self.document(f'<w:p><w:r>{siblings}<w:t>{text}</w:t></w:r></w:p>')
                self.rejected()

    def test_complex_containers_and_target_extension_attributes_are_rejected(self):
        for body in (
            '<w:p><w:sdt><w:sdtContent><w:r><w:t>{{name}}</w:t></w:r></w:sdtContent></w:sdt></w:p>',
            '<w:p><w:customXml><w:r><w:t>{{name}}</w:t></w:r></w:customXml></w:p>',
            '<w:p><x:wrapper><w:r><w:t>{{name}}</w:t></w:r></x:wrapper></w:p>',
            '<w:p><w:r><w:t x:binding="external">{{name}}</w:t></w:r></w:p>',
        ):
            with self.subTest(body=body):
                self.document(body)
                self.rejected()

    def test_opaque_parts_and_unrelated_extensions_remain_unchanged(self):
        self.document('<x:data x:val="keep &amp; 中文">vendor<!--keep--><?keep data?></x:data>'
                      '<w:p><w:r><w:t>{{name}}</w:t></w:r></w:p>')
        self.replace()
        with zipfile.ZipFile(self.source) as before, zipfile.ZipFile(self.output) as after:
            self.assertIsNone(after.testzip())
            self.assertEqual(set(before.namelist()), set(after.namelist()))
            for name in before.namelist():
                if name != "word/document.xml":
                    self.assertEqual(before.read(name), after.read(name))
            dom = xmlio.parse_dom(after.read("word/document.xml"))
            extension = dom.getElementsByTagNameNS("urn:vendor", "data")[0]
            self.assertEqual(extension.getAttributeNS("urn:vendor", "val"), "keep & 中文")
            self.assertEqual([node.nodeType for node in extension.childNodes], [3, 8, 7])


if __name__ == "__main__":
    unittest.main()
