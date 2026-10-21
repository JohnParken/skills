"""Focused stdlib-only tests for bounded DOCX parsing, validation and edits."""
import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "docx" / "scripts"))

from edit_docx import replace_text
from office.package import unpack_document
from office.validators.structure import StructureValidator
from office import xmlio

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def package_parts(document=None, extra=None, content_type=None, rels_extra=""):
    document = document or f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>'
    content_type = content_type or "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    parts = {
        "[Content_Types].xml": f'<Types xmlns="{CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/word/document.xml" ContentType="{content_type}"/></Types>',
        "_rels/.rels": f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/>{rels_extra}</Relationships>',
        "word/document.xml": document,
        "custom.bin": "opaque payload",
    }
    parts.update(extra or {})
    return parts


def write_package(path, parts, *, prefix=b"", duplicate=None):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        if prefix:
            zf.writestr("prefix.bin", prefix)
        for name, value in parts.items():
            zf.writestr(name, value)
        if duplicate:
            zf.writestr(duplicate[0], duplicate[1])


class StdlibCoreTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def unpack(self, name, parts):
        archive = self.root / name
        write_package(archive, parts)
        target = self.root / (name + "-dir")
        unpack_document(archive, target)
        return target

    def test_valid_minimal_package_and_structure(self):
        self.assertTrue(StructureValidator(self.unpack("valid.docx", package_parts()), extension=".docx").validate())

    def test_validator_detects_relationship_ids_types_styles_numbering_comments_and_placeholders(self):
        bad_parts = package_parts(
            document=f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:pPr><w:pStyle w:val="Absent"/></w:pPr><w:r><w:t>{{{{left}}}}</w:t></w:r><w:numPr><w:numId w:val="9"/></w:numPr><w:hyperlink r:id="missing"/></w:p><w:commentRangeEnd w:id="3"/></w:body></w:document>',
            extra={"word/styles.xml": f'<w:styles xmlns:w="{W}"><w:style w:styleId="Same"/><w:style w:styleId="Same"/></w:styles>',
                   "word/numbering.xml": f'<w:numbering xmlns:w="{W}"><w:num w:numId="1"><w:abstractNumId w:val="7"/></w:num></w:numbering>',
                   "word/comments.xml": f'<w:comments xmlns:w="{W}"><w:comment w:id="3"/></w:comments>',
                   "word/_rels/document.xml.rels": f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="{R}/comments" Target="comments.xml"/></Relationships>'},
            content_type="application/xml",
        )
        directory = self.unpack("invalid.docx", bad_parts)
        validator = StructureValidator(directory, extension=".docx", require_no_placeholders=True)
        self.assertFalse(validator.validate())
        report = " ".join(validator.errors)
        for expected in ("content type", "duplicate style ID", "abstractNum", "Unknown style", "Unknown numbering", "Unresolved relationship", "range ends before start", "Unfilled placeholder"):
            self.assertIn(expected, report)

    def test_xml_rejects_utf16_dtd_entity_and_excess_depth(self):
        for payload in (
            '<?xml version="1.0" encoding="UTF-16"?><!DOCTYPE x [<!ENTITY e "boom">]><x>&e;</x>'.encode("utf-16"),
            ("<x>" + "<a>" * xmlio.MAX_XML_DEPTH + "</a>" * xmlio.MAX_XML_DEPTH + "</x>").encode(),
        ):
            with self.assertRaises(ValueError):
                xmlio.fromstring(payload)

    def test_dom_reports_namespace_parse_failures_as_controlled_value_errors(self):
        with self.assertRaisesRegex(ValueError, "Invalid XML"):
            xmlio.parse_dom(b'<w:document/>')

    def test_archive_preflight_rejects_unsafe_duplicate_and_oversize_entries(self):
        for name, kwargs in (("../escape", {}), ("duplicate", {"duplicate": ("custom.bin", b"x")}),
                             ("oversize", {"large": True})):
            archive = self.root / (name.replace("/", "_") + ".docx")
            with zipfile.ZipFile(archive, "w") as zf:
                if name == "../escape":
                    zf.writestr(name, b"x")
                elif name == "duplicate":
                    zf.writestr("custom.bin", b"a")
                    zf.writestr("custom.bin", b"b")
                else:
                    zf.writestr("large.bin", b"x" * (32 * 1024 * 1024 + 1), compress_type=zipfile.ZIP_STORED)
            with self.assertRaises((ValueError, zipfile.BadZipFile)):
                unpack_document(archive, self.root / (archive.stem + "-out"))

    def test_replace_text_preserves_members_prefix_and_xml_declarations(self):
        doc = f'<?xml version="1.0"?><w:document xmlns:w="{W}" xmlns:x="urn:keep"><w:body><w:p><w:r><w:t>{{{{name}}}}</w:t></w:r></w:p></w:body></w:document>'
        source, dest = self.root / "in.docx", self.root / "out.docx"
        write_package(source, package_parts(doc))
        replace_text(source, dest, {"replacements": [{"old": "{{name}}", "new": "Ada", "count": 1}]})
        with zipfile.ZipFile(dest) as zf:
            self.assertIn("custom.bin", zf.namelist())
            raw = zf.read("word/document.xml")
            self.assertIn(b"xmlns:x=\"urn:keep\"", raw)
            self.assertIn(b">Ada<", raw)
            self.assertEqual(zf.read("custom.bin"), b"opaque payload")

    def test_replace_rejects_unsupported_edits_counts_overlaps_and_preserves_files(self):
        cases = [
            (f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>{{{{</w:t></w:r><w:r><w:t>name}}}}</w:t></w:r></w:p></w:body></w:document>', "{{name}}", "X"),
            (f'<w:document xmlns:w="{W}"><w:body><w:p><w:fldSimple><w:r><w:t>{{{{name}}}}</w:t></w:r></w:fldSimple></w:p></w:body></w:document>', "{{name}}", "X"),
            (f'<w:document xmlns:w="{W}"><w:body><w:ins w:id="1" w:author="a"><w:r><w:t>{{{{name}}}}</w:t></w:r></w:ins></w:body></w:document>', "{{name}}", "X"),
        ]
        for i, (doc, old, new) in enumerate(cases):
            source, output = self.root / f"source{i}.docx", self.root / f"out{i}.docx"
            write_package(source, package_parts(doc))
            original_bytes = source.read_bytes()
            output.write_bytes(b"existing output")
            with self.assertRaises(ValueError):
                replace_text(source, output, {"replacements": [{"old": old, "new": new, "count": 1}]})
            self.assertEqual(output.read_bytes(), b"existing output")
            self.assertEqual(source.read_bytes(), original_bytes)
        source, output = self.root / "overlap.docx", self.root / "overlap-out.docx"
        doc = f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>abc</w:t></w:r></w:p></w:body></w:document>'
        write_package(source, package_parts(doc))
        with self.assertRaises(ValueError):
            replace_text(source, output, {"replacements": [{"old": "ab", "new": "X"}, {"old": "bc", "new": "Y"}]})

    def test_wrong_count_and_unfilled_split_placeholder_fail_without_output(self):
        document = f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>Hello Hello</w:t></w:r></w:p><w:p><w:r><w:t>{{{{na</w:t></w:r><w:r><w:t>me}}}}</w:t></w:r></w:p></w:body></w:document>'
        source, output = self.root / "counts.docx", self.root / "out.docx"
        write_package(source, package_parts(document))
        for specification, no_placeholders in (
            ({"replacements": [{"old": "Hello", "new": "Hi"}]}, False),
            ({"replacements": [{"old": "Hello", "new": "Hi", "count": 2}]}, True),
        ):
            with self.assertRaises(ValueError):
                replace_text(source, output, specification, require_no_placeholders=no_placeholders)
            self.assertFalse(output.exists())

    def test_rejects_text_inside_cross_paragraph_fields_and_comment_ranges(self):
        cases = {
            "field": ('<w:p><w:r><w:fldChar w:fldCharType="begin"/>'
                      '<w:instrText> TOC </w:instrText><w:fldChar w:fldCharType="separate"/></w:r></w:p>'
                      '<w:p><w:r><w:t>待替换</w:t></w:r></w:p>'
                      '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'),
            "comment": ('<w:p><w:commentRangeStart w:id="0"/></w:p>'
                        '<w:p><w:r><w:t>待替换</w:t></w:r></w:p>'
                        '<w:p><w:commentRangeEnd w:id="0"/>'
                        '<w:r><w:commentReference w:id="0"/></w:r></w:p>'),
        }
        extra = {
            "word/comments.xml": f'<w:comments xmlns:w="{W}"><w:comment w:id="0" w:author="Reviewer"/></w:comments>',
            "word/_rels/document.xml.rels": f'<Relationships xmlns="{REL}"><Relationship Id="c" Type="{R}/comments" Target="comments.xml"/></Relationships>',
        }
        for kind, body in cases.items():
            for extension in ("docx", "dotx"):
                with self.subTest(kind=kind, extension=extension):
                    source = self.root / f"range-{kind}.{extension}"
                    output = self.root / f"range-{kind}-out.{extension}"
                    content_type = None if extension == "docx" else "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
                    document = f'<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>'
                    write_package(source, package_parts(document, extra if kind == "comment" else None, content_type))
                    before = source.read_bytes()
                    output.write_bytes(b"existing output")
                    with self.assertRaisesRegex(ValueError, "unsupported"):
                        replace_text(source, output, {"replacements": [{"old": "待替换", "new": "已修改"}]})
                    self.assertEqual(source.read_bytes(), before)
                    self.assertEqual(output.read_bytes(), b"existing output")

    def test_nested_fields_end_before_supported_plain_text(self):
        body = ('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p>'
                '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p>'
                '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
                '<w:p><w:r><w:t>保留字段结果</w:t></w:r></w:p>'
                '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
                '<w:p><w:r><w:t>普通文字</w:t></w:r></w:p>')
        document = f'<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>'
        source, output = self.root / "nested.docx", self.root / "nested-out.docx"
        write_package(source, package_parts(document))
        replace_text(source, output, {"replacements": [{"old": "普通文字", "new": "中文 & English"}]})
        with zipfile.ZipFile(output) as package:
            tree = xmlio.fromstring(package.read("word/document.xml"))
            self.assertEqual([node.text for node in tree.iter(f"{{{W}}}t")],
                             ["保留字段结果", "中文 & English"])

    def test_unbalanced_fields_fail_closed_without_replacing_output(self):
        for marker in ('<w:fldChar w:fldCharType="begin"/>',
                       '<w:fldChar w:fldCharType="end"/>',
                       '<w:fldChar w:fldCharType="separate"/>',
                       '<w:fldChar/>'):
            with self.subTest(marker=marker):
                document = (f'<w:document xmlns:w="{W}"><w:body><w:p><w:r>{marker}</w:r></w:p>'
                            '<w:p><w:r><w:t>普通文字</w:t></w:r></w:p></w:body></w:document>')
                source, output = self.root / "broken-field.docx", self.root / "broken-field-out.docx"
                write_package(source, package_parts(document))
                output.write_bytes(b"keep")
                with self.assertRaisesRegex(ValueError, "field"):
                    replace_text(source, output, {"replacements": [{"old": "普通文字", "new": "changed"}]})
                self.assertEqual(output.read_bytes(), b"keep")

    def test_same_rule_overlapping_matches_are_rejected(self):
        for text, old in (("aaa", "aa"), ("中文中文中文", "中文中文")):
            with self.subTest(text=text):
                document = f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:body></w:document>'
                source, output = self.root / "self-overlap.docx", self.root / "self-overlap-out.docx"
                write_package(source, package_parts(document))
                output.write_bytes(b"keep")
                with self.assertRaisesRegex(ValueError, "Overlapping"):
                    replace_text(source, output, {"replacements": [{"old": old, "new": "X"}]})
                self.assertEqual(output.read_bytes(), b"keep")

    def test_adjacent_matches_and_literal_search_remain_supported(self):
        document = f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>[][]</w:t></w:r></w:p></w:body></w:document>'
        source, output = self.root / "adjacent.docx", self.root / "adjacent-out.docx"
        write_package(source, package_parts(document))
        replace_text(source, output, {"replacements": [{"old": "[]", "new": "中文", "count": 2}]})
        with zipfile.ZipFile(output) as package:
            tree = xmlio.fromstring(package.read("word/document.xml"))
            self.assertEqual(next(tree.iter(f"{{{W}}}t")).text, "中文中文")

    def test_duplicate_relationships_and_missing_target_are_reported(self):
        extra = (
            f'<Relationship Id="rId1" Type="{R}/image" Target="word/media/missing.png"/>'
            f'<Relationship Id="rId2" Type="{R}/image" Target="word/media/missing.png"/>'
        )
        validator = StructureValidator(self.unpack("relationships.docx", package_parts(rels_extra=extra)))
        self.assertFalse(validator.validate())
        self.assertIn("duplicate relationship", " ".join(validator.errors))
        self.assertIn("Missing relationship target", " ".join(validator.errors))

    def test_obviously_invalid_word_container_and_run_properties_are_rejected(self):
        document = f'<w:document xmlns:w="{W}"><w:body><w:r><w:rPr><w:bold/></w:rPr><w:t>bad placement</w:t></w:r></w:body></w:document>'
        validator = StructureValidator(self.unpack("bad-body.docx", package_parts(document)))
        self.assertFalse(validator.validate())
        self.assertIn("Invalid body child", " ".join(validator.errors))
        self.assertIn("Invalid run property name", " ".join(validator.errors))


if __name__ == "__main__":
    unittest.main()
