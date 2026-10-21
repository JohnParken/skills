"""Rule-table row editing must preserve original rows and reject complex tables."""

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
from edit_tables import edit_tables
from office import xmlio
from test_docx_stdlib_core import R, REL, W, write_package


class TableEditTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-table-edit-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source, self.output = self.root / "source.docx", self.root / "out.docx"
        create_document({"blocks": [{"type": "paragraph", "text": "表格前文"},
                                    {"type": "table", "header": True,
                                     "rows": [["项目", "状态"], ["甲", "完成"], ["乙", "待办"]]}]}, self.source)

    def operation(self, **updates):
        value = {"type": "insert_row", "table_index": 0, "row_index": 1,
                 "expected_rows": 3, "expected_text": ["甲", "完成"],
                 "position": "after", "cells": ["新项目 & 中文", "已验收"]}
        value.update(updates)
        return value

    def run_edit(self, operation):
        return edit_tables(self.source, self.output, {"operation": operation})

    def rejected(self, operation):
        before = self.source.read_bytes()
        self.output.write_bytes(b"preserve existing")
        with self.assertRaises(ValueError):
            self.run_edit(operation)
        self.assertEqual(self.source.read_bytes(), before)
        self.assertEqual(self.output.read_bytes(), b"preserve existing")

    def mutate(self, callback):
        with zipfile.ZipFile(self.source) as package:
            parts = {name: package.read(name) for name in package.namelist()}
        dom = xmlio.parse_dom(parts["word/document.xml"])
        callback(dom)
        parts["word/document.xml"] = dom.toxml(encoding="UTF-8")
        write_package(self.source, parts)

    def test_insert_clones_data_format_and_preserves_existing_rows_and_members(self):
        self.run_edit(self.operation())
        with zipfile.ZipFile(self.source) as before, zipfile.ZipFile(self.output) as after:
            self.assertIsNone(after.testzip())
            self.assertEqual(set(before.namelist()), set(after.namelist()))
            for name in before.namelist():
                if name != "word/document.xml":
                    self.assertEqual(before.read(name), after.read(name))
            old = xmlio.parse_dom(before.read("word/document.xml"))
            new = xmlio.parse_dom(after.read("word/document.xml"))
        old_rows = old.getElementsByTagNameNS(W, "tr")
        rows = new.getElementsByTagNameNS(W, "tr")
        self.assertEqual(len(rows), 4)
        self.assertEqual([row.toxml() for row in old_rows], [rows[index].toxml() for index in (0, 1, 3)])
        self.assertEqual([node.firstChild.data for node in rows[2].getElementsByTagNameNS(W, "t")],
                         ["新项目 & 中文", "已验收"])
        self.assertEqual(rows[1].getElementsByTagNameNS(W, "tcPr")[0].toxml(),
                         rows[2].getElementsByTagNameNS(W, "tcPr")[0].toxml())
        self.assertEqual(len(rows[2].getElementsByTagNameNS(W, "tblHeader")), 0)

    def test_delete_tail_data_row_preserves_headers(self):
        self.run_edit({"type": "delete_row", "table_index": 0, "row_index": 2,
                       "expected_rows": 3, "expected_text": ["乙", "待办"]})
        with zipfile.ZipFile(self.output) as package:
            dom = xmlio.parse_dom(package.read("word/document.xml"))
        rows = dom.getElementsByTagNameNS(W, "tr")
        self.assertEqual(len(rows), 2)
        self.assertEqual(len(rows[0].getElementsByTagNameNS(W, "tblHeader")), 1)

    def test_header_and_last_remaining_data_row_are_protected(self):
        self.rejected(self.operation(row_index=0, expected_text=["项目", "状态"]))
        self.rejected({"type": "delete_row", "table_index": 0, "row_index": 0,
                       "expected_rows": 3, "expected_text": ["项目", "状态"]})
        create_document({"blocks": [{"type": "table", "header": True,
                                     "rows": [["标题", "状态"], ["甲", "完成"]]}]}, self.source)
        self.rejected({"type": "delete_row", "table_index": 0, "row_index": 1,
                       "expected_rows": 2, "expected_text": ["甲", "完成"]})

    def test_bad_selectors_values_and_unknown_fields_are_rejected(self):
        for updates in ({"table_index": True}, {"row_index": -1}, {"expected_rows": 4},
                        {"expected_text": ["错位", "完成"]}, {"cells": ["少一列"]},
                        {"position": []}, {"cells": ["换\n行", "no"]}, {"extra": True},
                        {"type": []}, {"expected_rows": True}, {"cells": ["a", 1]}):
            with self.subTest(updates=updates):
                self.rejected(self.operation(**updates))

    def test_merges_nested_tables_and_complex_cells_are_rejected(self):
        for markup in ('<w:gridSpan w:val="1"/>', '<w:hMerge w:val="restart"/>',
                       '<w:vMerge w:val="restart"/>'):
            def add(dom):
                dom.getElementsByTagNameNS(W, "tcPr")[0].appendChild(xmlio.parse_dom(
                    f'<w:x xmlns:w="{W}">{markup}</w:x>').documentElement.firstChild)
            self.mutate(add)
            self.rejected(self.operation())
            self._reset_source()
        self.mutate(lambda dom: dom.getElementsByTagNameNS(W, "tc")[2].appendChild(
            dom.getElementsByTagNameNS(W, "tbl")[0].cloneNode(True)))
        self.rejected(self.operation())

    def _reset_source(self):
        create_document({"blocks": [{"type": "table", "header": True,
                                     "rows": [["项目", "状态"], ["甲", "完成"], ["乙", "待办"]]}]}, self.source)

    def test_activity_spanning_table_and_unknown_references_are_rejected(self):
        for start, end in (("bookmarkStart", "bookmarkEnd"), ("permStart", "permEnd"),
                           ("customXmlInsRangeStart", "customXmlInsRangeEnd")):
            self._reset_source()
            def wrap(dom):
                body = dom.getElementsByTagNameNS(W, "body")[0]
                table = dom.getElementsByTagNameNS(W, "tbl")[0]
                for name, before in ((start, True), (end, False)):
                    paragraph = dom.createElementNS(W, "w:p")
                    marker = dom.createElementNS(W, "w:" + name)
                    marker.setAttributeNS(W, "w:id", "7")
                    paragraph.appendChild(marker)
                    body.insertBefore(paragraph, table if before else table.nextSibling)
            self.mutate(wrap)
            self.rejected(self.operation())
        self._reset_source()
        self.mutate(lambda dom: (dom.documentElement.setAttribute("xmlns:x", "urn:vendor"),
                                 dom.getElementsByTagNameNS(W, "tc")[2].setAttributeNS("urn:vendor", "x:binding", "id")))
        self.rejected(self.operation())

    def test_multiple_header_rows_remain_in_order(self):
        self._reset_source()
        def header(dom):
            rows = dom.getElementsByTagNameNS(W, "tr")
            rows[1].getElementsByTagNameNS(W, "trPr")[0].appendChild(
                rows[0].getElementsByTagNameNS(W, "tblHeader")[0].cloneNode(True))
        self.mutate(header)
        self.run_edit(self.operation(row_index=2, expected_text=["乙", "待办"]))
        with zipfile.ZipFile(self.output) as package:
            dom = xmlio.parse_dom(package.read("word/document.xml"))
        self.assertEqual(len(dom.getElementsByTagNameNS(W, "tblHeader")), 2)

    def test_empty_new_cells_are_explicit_and_supported(self):
        self.run_edit(self.operation(cells=["", " 空格 "]))
        with zipfile.ZipFile(self.output) as package:
            dom = xmlio.parse_dom(package.read("word/document.xml"))
        row = dom.getElementsByTagNameNS(W, "tr")[2]
        texts = row.getElementsByTagNameNS(W, "t")
        self.assertEqual(len(texts[0].childNodes), 0)
        self.assertEqual(texts[1].firstChild.data, " 空格 ")

    def test_field_spanning_table_is_rejected(self):
        def field(dom):
            body = dom.getElementsByTagNameNS(W, "body")[0]
            table = dom.getElementsByTagNameNS(W, "tbl")[0]
            before = xmlio.parse_dom(f'<w:p xmlns:w="{W}"><w:r><w:fldChar w:fldCharType="begin"/>'
                                    '<w:instrText>TOC</w:instrText><w:fldChar w:fldCharType="separate"/></w:r></w:p>')
            after = xmlio.parse_dom(f'<w:p xmlns:w="{W}"><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
            body.insertBefore(before.documentElement, table)
            body.insertBefore(after.documentElement, table.nextSibling)
        self.mutate(field)
        self.rejected(self.operation())

    def test_comment_spanning_table_is_rejected(self):
        def comment(dom):
            body = dom.getElementsByTagNameNS(W, "body")[0]
            table = dom.getElementsByTagNameNS(W, "tbl")[0]
            start = xmlio.parse_dom(f'<w:p xmlns:w="{W}"><w:commentRangeStart w:id="0"/></w:p>')
            end = xmlio.parse_dom(f'<w:p xmlns:w="{W}"><w:commentRangeEnd w:id="0"/>'
                                 '<w:r><w:commentReference w:id="0"/></w:r></w:p>')
            body.insertBefore(start.documentElement, table)
            body.insertBefore(end.documentElement, table.nextSibling)
        self.mutate(comment)
        with zipfile.ZipFile(self.source) as package:
            parts = {name: package.read(name) for name in package.namelist()}
        parts["word/comments.xml"] = (f'<w:comments xmlns:w="{W}"><w:comment w:id="0" w:author="reviewer">'
                                      '<w:p><w:r><w:t>范围</w:t></w:r></w:p></w:comment></w:comments>').encode()
        rels = xmlio.parse_dom(parts["word/_rels/document.xml.rels"])
        relation = rels.createElementNS(REL, "rel:Relationship")
        relation.setAttribute("xmlns:rel", REL)
        for key, value in {"Id": "rComment", "Type": R + "/comments", "Target": "comments.xml"}.items():
            relation.setAttribute(key, value)
        rels.documentElement.appendChild(relation)
        parts["word/_rels/document.xml.rels"] = rels.toxml(encoding="UTF-8")
        write_package(self.source, parts)
        self.rejected(self.operation())

    def test_xml_serialization_loss_is_rejected(self):
        self._reset_source()
        with zipfile.ZipFile(self.source) as package:
            parts = {name: package.read(name) for name in package.namelist()}
        parts["word/document.xml"] = parts["word/document.xml"].replace(
            b'<w:body>', b'<w:body><w:p><w:r><w:t>A&#13;B</w:t></w:r></w:p>',
        )
        write_package(self.source, parts)
        self.rejected(self.operation())

    def test_dotx_rows_use_original_template_content_type(self):
        with zipfile.ZipFile(self.source) as package:
            parts = {name: package.read(name) for name in package.namelist()}
        parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
            b'wordprocessingml.document.main+xml', b'wordprocessingml.template.main+xml',
        )
        self.source = self.root / "source.dotx"
        self.output = self.root / "out.dotx"
        write_package(self.source, parts)
        self.run_edit(self.operation())
        with zipfile.ZipFile(self.output) as package:
            self.assertIn(b'wordprocessingml.template.main+xml', package.read("[Content_Types].xml"))

    def test_same_file_alias_and_format_conversion_are_rejected(self):
        for destination in (self.source, self.output.with_suffix(".dotx")):
            with self.assertRaises(ValueError):
                edit_tables(self.source, destination, {"operation": self.operation()})
        alias = self.root / "alias.docx"
        alias.symlink_to(self.source)
        with self.assertRaises(ValueError):
            edit_tables(self.source, alias, {"operation": self.operation()})

    def test_cli_without_external_programs(self):
        (self.root / "operation.json").write_text(json.dumps({"operation": self.operation()}), encoding="utf-8")
        result = subprocess.run([sys.executable, "-I", "-S", "-B", str(ROOT / "docx/scripts/edit_tables.py"),
                                 "source.docx", "operation.json", "out.docx"], cwd=self.root,
                                env=dict(os.environ, PATH=""), capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()
