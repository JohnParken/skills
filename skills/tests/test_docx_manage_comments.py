"""Transactional comment lifecycle and adversarial anchor/metadata regression."""

import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "docx/scripts"))
from create_docx import create_document
from manage_comments import COMMENT_PARTS, CT, REL, W, W14, W15, W16CID, W16CEX, manage_comments
from office.xmlio import parse_dom


class ManageCommentsTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-manage-comments-test-")
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name)
        self.source = create_document({"blocks": [
            {"type": "paragraph", "runs": [{"text": "正文"}, {"text": "目标"}, {"text": "范围"}]},
            {"type": "paragraph", "text": "其他内容"},
        ]}, self.work / "input.docx")
        self.output = self.work / "output.docx"
        self.add = {"action": "add", "paragraph_index": 0, "expected_text": "正文目标范围",
                    "anchor_text": "目标范围", "text": "请核实 & 范围", "author": "审阅人", "initials": "审"}

    def parts(self, path=None):
        with zipfile.ZipFile(path or self.source) as package:
            self.assertIsNone(package.testzip())
            return {name: package.read(name) for name in package.namelist()}

    def rewrite(self, updates=None, remove=(), path=None):
        path = path or self.source
        parts = self.parts(path)
        parts.update(updates or {})
        for name in remove:
            parts.pop(name)
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
            for name, data in parts.items():
                package.writestr(name, data)

    def annotated(self):
        result = manage_comments(self.source, self.add, self.output)
        self.assertEqual(result, {"action": "add", "comment_id": 0})
        self.source = self.output
        self.output = self.work / "next.docx"

    def rejected(self, specification, *, source=None, output=None):
        source, output = source or self.source, output or self.output
        original = source.read_bytes()
        output.write_bytes(b"existing output")
        with self.assertRaises(ValueError):
            manage_comments(source, specification, output)
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual(output.read_bytes(), b"existing output")

    def test_full_modern_lifecycle_add_reply_update_delete(self):
        self.annotated()
        listing = manage_comments(self.source, {"action": "list"})
        self.assertEqual((listing["mode"], listing["mutable"]), ("modern", True))
        self.assertEqual(listing["comments"][0]["text"], self.add["text"])
        self.assertEqual(listing["comments"][0]["paragraph_index"], 0)
        before = self.parts()
        reply = {"action": "reply", "id": 0, "expected_text": self.add["text"],
                 "text": "已经复核", "author": "回复人"}
        self.assertEqual(manage_comments(self.source, reply, self.output)["comment_id"], 1)
        self.source, self.output = self.output, self.work / "updated.docx"
        listing = manage_comments(self.source, {"action": "list"})
        self.assertEqual([item["parent_id"] for item in listing["comments"]], [None, 0])
        manage_comments(self.source, {"action": "update", "id": 1, "expected_text": "已经复核", "text": "已复核通过"}, self.output)
        self.source, self.output = self.output, self.work / "delete-reply.docx"
        self.rejected({"action": "delete", "id": 0, "expected_text": self.add["text"]})
        manage_comments(self.source, {"action": "delete", "id": 1, "expected_text": "已复核通过"}, self.output)
        self.source, self.output = self.output, self.work / "delete-root.docx"
        manage_comments(self.source, {"action": "delete", "id": 0, "expected_text": self.add["text"]}, self.output)
        listing = manage_comments(self.output, {"action": "list"})
        self.assertTrue(listing["mutable"], listing["diagnostics"])
        self.assertEqual(listing["comments"], [])
        after = self.parts(self.output)
        self.assertTrue(set(COMMENT_PARTS) <= set(after))
        for name in before:
            if name not in {*COMMENT_PARTS, "word/document.xml"}:
                self.assertEqual(before[name], after[name], name)

    def test_read_only_empty_listing_and_no_document_output(self):
        original = self.source.read_bytes()
        result = manage_comments(self.source, {"action": "list"})
        self.assertEqual((result["mode"], result["comments"], result["mutable"]), ("none", [], True))
        self.assertEqual(self.source.read_bytes(), original)
        self.rejected({"action": "list"})

    def test_complete_run_anchor_does_not_change_run_format_or_content(self):
        before = parse_dom(self.parts()["word/document.xml"])
        original_runs = [node.toxml() for node in before.getElementsByTagNameNS(W, "r")]
        manage_comments(self.source, self.add, self.output)
        after = parse_dom(self.parts(self.output)["word/document.xml"])
        actual_runs = [node.toxml() for node in after.getElementsByTagNameNS(W, "r")
                       if not node.getElementsByTagNameNS(W, "commentReference")]
        self.assertEqual(original_runs, actual_runs)

    def test_ambiguous_partial_and_guard_mismatches_are_rejected(self):
        for changes in ({"anchor_text": "标"}, {"anchor_text": "缺失"},
                        {"expected_text": "错误"}, {"paragraph_index": True},
                        {"paragraph_index": 3}, {"anchor_text": ""}):
            with self.subTest(changes=changes):
                self.rejected({**self.add, **changes})
        self.source = create_document({"blocks": [{"type": "paragraph", "runs": [
            {"text": "重复"}, {"text": "重复"}]}]}, self.work / "ambiguous.docx")
        self.rejected({**self.add, "expected_text": "重复重复", "anchor_text": "重复"})

    def test_add_inside_fields_revisions_ranges_or_complex_containers_is_rejected(self):
        prefix = f'<w:document xmlns:w="{W}"><w:body>'
        suffix = '</w:body></w:document>'
        for body in (
            '<w:p><w:hyperlink><w:r><w:t>正文目标范围</w:t></w:r></w:hyperlink></w:p>',
            '<w:p><w:r><w:fldChar w:fldCharType="begin"/><w:t>正文目标范围</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>',
            '<w:p><w:bookmarkStart w:id="1" w:name="x"/><w:r><w:t>正文目标范围</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>',
            '<w:p><w:sdt><w:sdtContent><w:r><w:t>正文目标范围</w:t></w:r></w:sdtContent></w:sdt></w:p>',
        ):
            with self.subTest(body=body):
                self.rewrite({"word/document.xml": prefix + body + suffix})
                self.rejected(self.add)

    def test_invalid_specs_and_xml_text_are_transactional(self):
        for spec in (None, [], {"action": []}, {"action": "unknown"},
                     {**self.add, "text": "bad\x00"}, {**self.add, "author": "bad\n"},
                     {**self.add, "initials": 1}, {**self.add, "extra": True}):
            with self.subTest(spec=spec):
                self.rejected(spec)

    def test_ids_and_expected_comment_guards_are_enforced(self):
        self.annotated()
        for action in ("update", "delete", "reply"):
            fields = {"text": "新值"} if action in {"update", "reply"} else {}
            if action == "reply":
                fields["author"] = "作者"
            for value, expected in ((False, self.add["text"]), (-1, self.add["text"]),
                                    (2 ** 31, self.add["text"]), (9, self.add["text"]),
                                    (0, "错误原文")):
                with self.subTest(action=action, value=value, expected=expected):
                    self.rejected({"action": action, "id": value, "expected_text": expected, **fields})

    def test_nested_replies_are_rejected(self):
        self.annotated()
        manage_comments(self.source, {"action": "reply", "id": 0, "expected_text": self.add["text"],
                                      "text": "一级回复", "author": "作者"}, self.output)
        self.source, self.output = self.output, self.work / "nested.docx"
        self.rejected({"action": "reply", "id": 1, "expected_text": "一级回复", "text": "二级", "author": "作者"})

    def test_unknown_and_incomplete_modern_metadata_list_but_refuse_mutation(self):
        self.annotated()
        parts = self.parts()
        variants = (
            ({"word/people.xml": '<people xmlns="urn:unknown"/>'}, ()),
            ({"word/commentsUserData.xml": '<metadata/>'}, ()),
            ({"word/commentsExtended.xml": parts["word/commentsExtended.xml"].replace(b'w15:done="0"', b'w15:done="unknown"')}, ()),
            ({"word/commentsExtensible.xml": parts["word/commentsExtensible.xml"].replace(b'/></w16cex:commentsExtensible>', b'/><unknown/></w16cex:commentsExtensible>')}, ()),
            ({}, ("word/commentsIds.xml",)),
        )
        for updates, remove in variants:
            with self.subTest(updates=list(updates), remove=remove):
                with zipfile.ZipFile(self.source, "w", zipfile.ZIP_DEFLATED) as package:
                    for name, data in {**parts, **updates}.items():
                        if name not in remove:
                            package.writestr(name, data)
                result = manage_comments(self.source, {"action": "list"})
                self.assertFalse(result["mutable"], result)
                self.assertTrue(result["diagnostics"])
                self.rejected({"action": "update", "id": 0, "expected_text": self.add["text"], "text": "new"})

    def test_metadata_bijection_parent_cycles_and_id_collisions_are_rejected(self):
        self.annotated()
        original = self.parts()
        ext = original["word/commentsExtended.xml"]
        para_id = parse_dom(ext).getElementsByTagNameNS(W15, "commentEx")[0].getAttributeNS(W15, "paraId")
        for part, replacement in (
            ("word/commentsExtended.xml", ext.replace(b'w15:done="0"', f'w15:done="0" w15:paraIdParent="{para_id}"'.encode())),
            ("word/commentsIds.xml", original["word/commentsIds.xml"].replace(para_id.encode(), b'11111111')),
            ("word/commentsExtensible.xml", original["word/commentsExtensible.xml"].replace(b'w16cex:durableId=', b'w16cex:unknown=')),
        ):
            with self.subTest(part=part):
                with zipfile.ZipFile(self.source, "w") as package:
                    for name, data in original.items():
                        package.writestr(name, replacement if name == part else data)
                self.assertFalse(manage_comments(self.source, {"action": "list"})["mutable"])
                self.rejected({"action": "delete", "id": 0, "expected_text": self.add["text"]})

    def test_generated_hex_ids_do_not_collide_with_existing_ids(self):
        self.annotated()
        original = self.parts()
        existing = int(parse_dom(original["word/comments.xml"]).getElementsByTagNameNS(W, "p")[0].getAttributeNS(W14, "paraId"), 16)
        with mock.patch("manage_comments.secrets.randbelow", return_value=existing - 1):
            self.rejected({"action": "reply", "id": 0, "expected_text": self.add["text"], "text": "reply", "author": "作者"})

    def test_multiple_root_comments_have_unique_mapping_and_duplicate_metadata_is_rejected(self):
        self.annotated()
        second = {**self.add, "paragraph_index": 1, "expected_text": "其他内容", "anchor_text": "其他内容", "text": "第二条"}
        manage_comments(self.source, second, self.output)
        self.source, self.output = self.output, self.work / "duplicate.docx"
        original = self.parts()
        comments = parse_dom(original["word/comments.xml"])
        paragraph_ids = [node.getAttributeNS(W14, "paraId") for node in comments.getElementsByTagNameNS(W, "p")]
        self.assertEqual(len(set(paragraph_ids)), 2)
        identities = parse_dom(original["word/commentsIds.xml"])
        identity_nodes = identities.getElementsByTagNameNS(W16CID, "commentId")
        durable_ids = [node.getAttributeNS(W16CID, "durableId") for node in identity_nodes]
        self.assertEqual(len(set(durable_ids)), 2)
        for part, changed in (
            ("word/comments.xml", original["word/comments.xml"].replace(paragraph_ids[1].encode(), paragraph_ids[0].encode())),
            ("word/commentsIds.xml", original["word/commentsIds.xml"].replace(durable_ids[1].encode(), durable_ids[0].encode())),
        ):
            with self.subTest(part=part):
                with zipfile.ZipFile(self.source, "w") as package:
                    for name, data in original.items():
                        package.writestr(name, changed if name == part else data)
                result = manage_comments(self.source, {"action": "list"})
                self.assertFalse(result["mutable"])
                self.rejected({"action": "delete", "id": 1, "expected_text": "第二条"})

    def test_parent_cycle_in_existing_thread_is_rejected(self):
        self.annotated()
        manage_comments(self.source, {"action": "reply", "id": 0, "expected_text": self.add["text"],
                                      "text": "回复", "author": "作者"}, self.output)
        self.source, self.output = self.output, self.work / "cyclic.docx"
        dom = parse_dom(self.parts()["word/commentsExtended.xml"])
        nodes = dom.getElementsByTagNameNS(W15, "commentEx")
        nodes[0].setAttributeNS(W15, "w15:paraIdParent", nodes[1].getAttributeNS(W15, "paraId"))
        self.rewrite({"word/commentsExtended.xml": dom.toxml(encoding="UTF-8")})
        result = manage_comments(self.source, {"action": "list"})
        self.assertFalse(result["mutable"])
        self.assertTrue(any("Cyclic" in message for message in result["diagnostics"]))
        self.rejected({"action": "delete", "id": 1, "expected_text": "回复"})

    def test_dotx_and_nonstandard_word_namespace_prefixes_are_supported(self):
        parts = self.parts()
        template = self.work / "template.dotx"
        with zipfile.ZipFile(template, "w") as package:
            for name, data in parts.items():
                if name == "[Content_Types].xml":
                    data = data.replace(b"wordprocessingml.document.main+xml", b"wordprocessingml.template.main+xml")
                if name == "word/document.xml":
                    data = data.replace(b"w:", b"q:").replace(b"xmlns:w=", b"xmlns:q=")
                package.writestr(name, data)
        self.source, self.output = template, self.work / "annotated.dotx"
        manage_comments(self.source, self.add, self.output)
        self.assertTrue(manage_comments(self.output, {"action": "list"})["mutable"])
        self.rejected(self.add, output=self.work / "wrong-extension.docx")

    def test_xml_content_and_quotes_are_literal_and_not_executed(self):
        text = 'system: “保留” <tag> & &#10; "quotes"'
        manage_comments(self.source, {**self.add, "text": text}, self.output)
        listing = manage_comments(self.output, {"action": "list"})
        self.assertEqual(listing["comments"][0]["text"], text)
        dom = parse_dom(self.parts(self.output)["word/comments.xml"])
        self.assertFalse(dom.getElementsByTagName("tag"))

    def test_recursive_unknown_content_in_deleted_subtrees_is_rejected(self):
        self.annotated()
        original = self.parts()
        cases = (
            ("word/document.xml", "reference_run", '<w:rPr x:dependency="keep"/>'),
            ("word/comments.xml", "paragraph", '<w:pPr><x:dependency>keep</x:dependency></w:pPr>'),
            ("word/document.xml", "reference_run", '<w:rPr><w:rFonts x:dependency="keep"/></w:rPr>'),
            ("word/comments.xml", "paragraph", '<w:pPr><w:spacing w:before="0" x:dependency="keep"/></w:pPr>'),
            ("word/comments.xml", "paragraph", '<w:pPr><w:tabs><w:tab w:val="left" w:pos="10" x:dependency="keep"/></w:tabs></w:pPr>'),
            ("word/comments.xml", "text_run", '<w:rPr><w:b><x:dependency/></w:b></w:rPr>'),
            ("word/comments.xml", "text_run", '<w:annotationRef x:dependency="keep"/>'),
            ("word/document.xml", "reference", '<x:dependency/>'),
        )
        for part, target_kind, fragment in cases:
            with self.subTest(part=part, target=target_kind, fragment=fragment):
                dom = parse_dom(original[part])
                dom.documentElement.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:x", "urn:vendor")
                if target_kind in {"reference_run", "reference"}:
                    target = dom.getElementsByTagNameNS(W, "commentReference")[0]
                    if target_kind == "reference_run":
                        target = target.parentNode
                elif target_kind == "paragraph":
                    target = dom.getElementsByTagNameNS(W, "p")[0]
                else:
                    target = dom.getElementsByTagNameNS(W, "t")[0].parentNode
                wrapper = parse_dom(f'<root xmlns:w="{W}" xmlns:x="urn:vendor">{fragment}</root>')
                target.insertBefore(dom.importNode(wrapper.documentElement.firstChild, True), target.firstChild)
                with zipfile.ZipFile(self.source, "w") as package:
                    for name, data in original.items():
                        package.writestr(name, dom.toxml(encoding="UTF-8") if name == part else data)
                listing = manage_comments(self.source, {"action": "list"})
                self.assertFalse(listing["mutable"], listing)
                self.assertTrue(listing["diagnostics"])
                self.rejected({"action": "delete", "id": 0, "expected_text": self.add["text"]})

    def test_non_xml_whitespace_in_deleted_formatting_and_metadata_is_rejected(self):
        self.annotated()
        original = self.parts()
        for whitespace in ("\u00a0", "\u3000"):
            for part in ("word/comments.xml", "word/commentsExtended.xml"):
                with self.subTest(whitespace=repr(whitespace), part=part):
                    dom = parse_dom(original[part])
                    if part == "word/comments.xml":
                        target = dom.createElementNS(W, "w:pPr")
                        paragraph = dom.getElementsByTagNameNS(W, "p")[0]
                        paragraph.insertBefore(target, paragraph.firstChild)
                    else:
                        target = dom.getElementsByTagNameNS(W15, "commentEx")[0]
                    target.appendChild(dom.createTextNode(whitespace))
                    with zipfile.ZipFile(self.source, "w") as package:
                        for name, data in original.items():
                            package.writestr(name, dom.toxml(encoding="UTF-8") if name == part else data)
                    listing = manage_comments(self.source, {"action": "list"})
                    self.assertFalse(listing["mutable"], listing)
                    self.assertTrue(listing["diagnostics"])
                    self.rejected({"action": "delete", "id": 0, "expected_text": self.add["text"]})

    def test_known_comment_paragraph_run_and_annotation_reference_properties_remain_supported(self):
        self.annotated()
        parts = self.parts()
        dom = parse_dom(parts["word/comments.xml"])
        paragraph = dom.getElementsByTagNameNS(W, "p")[0]
        fragment = ('<w:pPr><w:spacing w:before="0" w:after="0"/><w:jc w:val="left"/>'
                    '<w:tabs><w:tab w:val="left" w:pos="100"/></w:tabs>'
                    '<w:rPr><w:b w:val="0"/></w:rPr></w:pPr>')
        wrapper = parse_dom(f'<root xmlns:w="{W}">{fragment}</root>')
        paragraph.insertBefore(dom.importNode(wrapper.documentElement.firstChild, True), paragraph.firstChild)
        text_run = dom.getElementsByTagNameNS(W, "t")[0].parentNode
        fragment = '<w:rPr><w:rStyle w:val="CommentReference"/><w:rFonts w:eastAsia="仿宋"/><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        wrapper = parse_dom(f'<root xmlns:w="{W}">{fragment}</root>')
        text_run.insertBefore(dom.importNode(wrapper.documentElement.firstChild, True), text_run.firstChild)
        annotation_run = dom.createElementNS(W, "w:r")
        annotation = dom.createElementNS(W, "w:annotationRef")
        annotation_run.appendChild(annotation)
        paragraph.insertBefore(annotation_run, text_run)
        self.rewrite({"word/comments.xml": dom.toxml(encoding="UTF-8")})
        listing = manage_comments(self.source, {"action": "list"})
        self.assertTrue(listing["mutable"], listing["diagnostics"])
        manage_comments(self.source, {"action": "update", "id": 0, "expected_text": self.add["text"], "text": "保留格式"}, self.output)
        self.source, self.output = self.output, self.work / "known-format-deleted.docx"
        self.assertTrue(manage_comments(self.source, {"action": "list"})["mutable"])
        manage_comments(self.source, {"action": "delete", "id": 0, "expected_text": "保留格式"}, self.output)
        self.assertEqual(manage_comments(self.output, {"action": "list"})["comments"], [])

    def test_classic_single_paragraph_update_delete_and_reply_refusal(self):
        self.annotated()
        parts = self.parts()
        remove = set(COMMENT_PARTS) - {"word/comments.xml"}
        relationships = parse_dom(parts["word/_rels/document.xml.rels"])
        for node in list(relationships.getElementsByTagNameNS(REL, "Relationship")):
            if node.getAttribute("Type") in {COMMENT_PARTS[name][4] for name in remove}:
                node.parentNode.removeChild(node)
        types = parse_dom(parts["[Content_Types].xml"])
        for node in list(types.getElementsByTagNameNS(CT, "Override")):
            if node.getAttribute("PartName").lstrip("/") in remove:
                node.parentNode.removeChild(node)
        self.rewrite({"word/_rels/document.xml.rels": relationships.toxml(encoding="UTF-8"),
                      "[Content_Types].xml": types.toxml(encoding="UTF-8")}, remove=remove)
        self.assertEqual(manage_comments(self.source, {"action": "list"})["mode"], "classic")
        self.rejected({"action": "reply", "id": 0, "expected_text": self.add["text"], "text": "回复", "author": "作者"})
        manage_comments(self.source, {"action": "update", "id": 0, "expected_text": self.add["text"], "text": "经典修改"}, self.output)
        self.source, self.output = self.output, self.work / "classic-deleted.docx"
        manage_comments(self.source, {"action": "delete", "id": 0, "expected_text": "经典修改"}, self.output)
        self.assertEqual(manage_comments(self.output, {"action": "list"})["comments"], [])

    def test_unknown_parts_and_unmodified_comment_metadata_bytes_remain_exact(self):
        types = parse_dom(self.parts()["[Content_Types].xml"])
        default = types.createElementNS(CT, "Default")
        default.setAttribute("Extension", "bin")
        default.setAttribute("ContentType", "application/octet-stream")
        types.documentElement.appendChild(default)
        self.rewrite({"custom/item.bin": b'\x00\xffopaque',
                      "[Content_Types].xml": types.toxml(encoding="UTF-8")})
        self.annotated()
        before = self.parts()
        manage_comments(self.source, {"action": "update", "id": 0, "expected_text": self.add["text"], "text": "仅评论变更"}, self.output)
        after = self.parts(self.output)
        self.assertEqual(set(before), set(after))
        for name in before:
            if name != "word/comments.xml":
                self.assertEqual(before[name], after[name], name)

    def test_unpreservable_non_target_xml_values_fail_without_output_replacement(self):
        dom = self.parts()["word/document.xml"].decode()
        dom = dom.replace('</w:body>', '<x:data xmlns:x="urn:vendor" x:val="a&#10;b"/></w:body>')
        self.rewrite({"word/document.xml": dom})
        self.rejected(self.add)

    def test_signed_packages_and_same_file_or_symlink_outputs_are_rejected(self):
        self.rewrite({"_xmlsignatures/sig1.xml": '<signature xmlns="urn:test"/>'})
        self.assertFalse(manage_comments(self.source, {"action": "list"})["mutable"])
        self.rejected(self.add)
        original = self.source.read_bytes()
        with self.assertRaises(ValueError):
            manage_comments(self.source, self.add, self.source)
        self.assertEqual(self.source.read_bytes(), original)
        alias = self.work / "alias.docx"
        alias.symlink_to(self.source)
        with self.assertRaises(ValueError):
            manage_comments(self.source, self.add, alias)
        self.assertEqual(self.source.read_bytes(), original)

    def test_unknown_external_person_dependencies_are_rejected_before_first_add(self):
        parts = self.parts()
        relationships = parse_dom(parts["word/_rels/document.xml.rels"])
        node = relationships.createElementNS(REL, "Relationship")
        for key, value in (("Id", "cloud"), ("Type", "urn:vendor/people"),
                           ("Target", "https://example.invalid/person"), ("TargetMode", "External")):
            node.setAttribute(key, value)
        relationships.documentElement.appendChild(node)
        self.rewrite({"word/_rels/document.xml.rels": relationships.toxml(encoding="UTF-8")})
        self.assertFalse(manage_comments(self.source, {"action": "list"})["mutable"])
        self.rejected(self.add)

    def test_resolved_comments_and_unknown_anchor_attributes_are_not_replied_to(self):
        self.annotated()
        original = self.parts()
        for part, value in (
            ("word/commentsExtended.xml", original["word/commentsExtended.xml"].replace(b'w15:done="0"', b'w15:done="1"')),
            ("word/document.xml", original["word/document.xml"].replace(b'<w:commentRangeStart ', b'<w:commentRangeStart xmlns:x="urn:vendor" x:binding="cloud" ')),
        ):
            with self.subTest(part=part):
                with zipfile.ZipFile(self.source, "w") as package:
                    for name, data in original.items():
                        package.writestr(name, value if name == part else data)
                self.rejected({"action": "reply", "id": 0, "expected_text": self.add["text"], "text": "reply", "author": "作者"})

    def test_comment_package_without_body_returns_diagnostics_not_index_error(self):
        self.annotated()
        self.rewrite({"word/document.xml": f'<w:document xmlns:w="{W}"/>'})
        result = manage_comments(self.source, {"action": "list"})
        self.assertFalse(result["mutable"])
        self.assertTrue(any("body" in error for error in result["diagnostics"]))
        self.rejected({"action": "update", "id": 0, "expected_text": self.add["text"], "text": "new"})

    def test_cli_lifecycle_and_invalid_json_leave_existing_output(self):
        action = self.work / "action.json"
        action.write_text(json.dumps(self.add, ensure_ascii=False), encoding="utf-8")
        script = ROOT / "docx/scripts/manage_comments.py"
        command = [sys.executable, "-I", "-S", "-B", str(script), str(self.source), str(action)]
        result = subprocess.run([*command, str(self.output)], env=dict(os.environ, PATH=""), capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)["comment_id"], 0)
        action.write_text('{"action":"list"}', encoding="utf-8")
        result = subprocess.run([*command[:5], str(self.output), str(action)], env=dict(os.environ, PATH=""), capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)["comments"][0]["text"], self.add["text"])
        before = self.output.read_bytes()
        action.write_text('{invalid}', encoding="utf-8")
        result = subprocess.run([*command, str(self.output)], env=dict(os.environ, PATH=""), capture_output=True, text=True)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(self.output.read_bytes(), before)


if __name__ == "__main__":
    unittest.main()
