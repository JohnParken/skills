"""Protect low-level comment preparation from known invalid inputs and IDs."""

import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "docx/scripts"))
import comment
from create_docx import create_document
from office.package import unpack_document


class CommentPreflightTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-comment-preflight-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        source = self.root / "input.docx"
        create_document({"blocks": [{"type": "paragraph", "text": "正文"}]}, source)
        self.directory = unpack_document(source, self.root / "package")

    def snapshot(self):
        return {p.relative_to(self.directory).as_posix(): p.read_bytes()
                for p in self.directory.rglob("*") if p.is_file()}

    def test_invalid_inputs_fail_before_any_package_write(self):
        cases = (("bad\x00text", {}), ("text", {"author": "bad\nname"}),
                 ("text", {"author": ""}), ("text", {"comment_id": -1}),
                 ("text", {"comment_id": True}), ("text", {"parent_id": 10 ** 1000}),
                 ("<w:p/>", {"raw": True}), ("&#10;", {"raw": True}))
        for text, keywords in cases:
            with self.subTest(text=text, keywords=keywords):
                before = self.snapshot()
                with self.assertRaises(ValueError):
                    comment.add_comment(self.directory, text, **keywords)
                self.assertEqual(self.snapshot(), before)

    def test_cdata_input_is_rejected_and_existing_opaque_node_values_are_preserved(self):
        before = self.snapshot()
        with self.assertRaises(ValueError):
            comment.add_comment(self.directory, '<![CDATA[“中文” & text]]>', raw=True)
        self.assertEqual(self.snapshot(), before)
        comment.add_comment(self.directory, "first")
        path = self.directory / "word/comments.xml"
        dom = comment.parse_dom(comment.read_xml(path))
        opaque = dom.createElementNS("urn:opaque", "x:data")
        opaque.setAttribute("xmlns:x", "urn:opaque")
        opaque.appendChild(dom.createCDATASection("“keep”"))
        opaque.appendChild(dom.createComment("“keep”"))
        opaque.appendChild(dom.createProcessingInstruction("keep", "“keep”"))
        dom.documentElement.appendChild(opaque)
        path.write_bytes(dom.toxml(encoding="UTF-8"))
        comment.add_comment(self.directory, "A &amp; B &lt;C&gt; &#x2019; 中文", raw=True)
        reopened = comment.parse_dom(comment.read_xml(path))
        node = reopened.getElementsByTagNameNS("urn:opaque", "data")[0]
        self.assertEqual([child.nodeValue for child in node.childNodes], ["“keep”"] * 3)

    def test_invalid_package_roots_fail_before_any_write(self):
        for relative in ("[Content_Types].xml", "word/_rels/document.xml.rels", "word/document.xml"):
            with self.subTest(relative=relative):
                path = self.directory / relative
                original = path.read_bytes()
                path.write_text("<Wrong/>", encoding="utf-8")
                before = self.snapshot()
                with self.assertRaisesRegex(ValueError, "root"):
                    comment.add_comment(self.directory, "new")
                self.assertEqual(self.snapshot(), before)
                path.write_bytes(original)

    def test_minimal_existing_comments_root_remains_usable(self):
        path = self.directory / "word/comments.xml"
        path.write_text('<w:comments xmlns:w="' + comment.NS['w'] + '"/>', encoding="utf-8")
        comment.add_comment(self.directory, "中文与 English")
        dom = comment.parse_dom(comment.read_xml(path))
        self.assertEqual(len(dom.getElementsByTagNameNS(comment.NS['w'], "comment")), 1)

    def test_symlink_metadata_is_rejected_without_touching_external_target(self):
        outside = self.root / "outside-comments.xml"
        outside.write_bytes((comment.TEMPLATE_DIR / 'comments.xml').read_bytes())
        (self.directory / "word/comments.xml").symlink_to(outside)
        before = outside.read_bytes()
        with self.assertRaisesRegex(ValueError, "links"):
            comment.add_comment(self.directory, "new")
        self.assertEqual(outside.read_bytes(), before)

    def test_duplicate_explicit_id_is_rejected_without_partial_writes(self):
        comment.add_comment(self.directory, "first", comment_id=0)
        before = self.snapshot()
        with self.assertRaisesRegex(ValueError, "already exists"):
            comment.add_comment(self.directory, "duplicate", comment_id=0)
        self.assertEqual(self.snapshot(), before)

    def test_metadata_id_collision_fails_before_reply_writes(self):
        with mock.patch.object(comment, "_generate_hex_id", return_value="12345678"):
            comment.add_comment(self.directory, "parent", comment_id=0)
            before = self.snapshot()
            with self.assertRaisesRegex(ValueError, "unique"):
                comment.add_comment(self.directory, "reply", parent_id=0)
            self.assertEqual(self.snapshot(), before)

    def test_unpreservable_existing_xml_fails_without_changes(self):
        path = self.directory / "[Content_Types].xml"
        raw = path.read_text(encoding="utf-8")
        path.write_text(raw.replace("<Types ", '<Types opaque="a&#10;b" ', 1), encoding="utf-8")
        before = self.snapshot()
        with self.assertRaisesRegex(ValueError, "preserve"):
            comment.add_comment(self.directory, "new")
        self.assertEqual(self.snapshot(), before)

    def test_conflicting_metadata_or_equivalent_id_is_rejected_before_writes(self):
        path = self.directory / "word/comments.xml"
        for content in ('<wrong/>',
                        '<q:comments xmlns:q="' + comment.NS['w'] + '" xmlns:w="urn:other"/>',
                        '<w:comments xmlns:w="' + comment.NS['w'] + '"><w:comment w:id="000"/></w:comments>'):
            with self.subTest(content=content):
                path.write_text(content, encoding="utf-8")
                before = self.snapshot()
                with self.assertRaises(ValueError):
                    comment.add_comment(self.directory, "new", comment_id=0)
                self.assertEqual(self.snapshot(), before)


if __name__ == "__main__":
    unittest.main()
