"""Structured reading preserves explicit XML structure without claiming rendering."""

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
from read_docx import read_docx_structure
import read_docx

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def nodes(node, kind):
    result = [node] if node.get("type") == kind else []
    for child in node.get("children", []):
        result.extend(nodes(child, kind))
    return result


class StructuredReadTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-structured-read-")
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name)
        self.source = self.work / "input.docx"

    def package(self, body, **parts):
        with zipfile.ZipFile(self.source, "w") as package:
            package.writestr("word/document.xml", f'<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:x="urn:vendor"><w:body>{body}</w:body></w:document>')
            for name, value in parts.items():
                package.writestr(name, value)

    def test_explicit_paragraph_run_styles_numbering_and_order(self):
        self.package('<w:p><w:pPr><w:pStyle w:val="Report"/><w:numPr><w:ilvl w:val="2"/><w:numId w:val="7"/></w:numPr></w:pPr>'
                     '<w:r><w:rPr><w:rStyle w:val="Emphasis"/><w:b w:val="0"/><w:rFonts w:eastAsia="仿宋"/></w:rPr>'
                     '<w:t xml:space="preserve">中文 </w:t><w:tab/><w:br/><w:t>English</w:t></w:r></w:p>')
        result = read_docx_structure(self.source)
        paragraph = nodes(result["parts"][0]["root"], "paragraph")[0]
        self.assertEqual(paragraph["style_id"], "Report")
        self.assertEqual(paragraph["numbering"], {"num_id": "7", "level": "2"})
        run = nodes(paragraph, "run")[0]
        self.assertEqual(run["style_id"], "Emphasis")
        self.assertEqual([child["type"] for child in run["children"]], ["properties", "text", "tab", "break", "text"])
        self.assertEqual(nodes(run, "text")[0]["text"], "中文 ")
        properties = run["children"][0]["children"]
        self.assertEqual(properties[1]["attributes"]["w:val"], "0")
        self.assertEqual(properties[2]["attributes"]["w:eastAsia"], "仿宋")
        self.assertEqual(result["semantics"]["styles"], "explicit-only")
        self.assertEqual(result["semantics"]["list_numbers"], "not-calculated")

    def test_grid_merges_row_offsets_and_nested_tables(self):
        self.package('<w:tbl><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>'
                     '<w:tr><w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="2"/></w:trPr><w:tc><w:tcPr>'
                     '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr><w:p/>'
                     '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
                     '</w:tc></w:tr><w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>')
        root = read_docx_structure(self.source)["parts"][0]["root"]
        tables = nodes(root, "table")
        self.assertEqual(len(tables), 2)
        self.assertEqual([node["attributes"]["w:w"] for node in nodes(tables[0], "grid_column")], ["1000", "2000"])
        row = nodes(tables[0], "row")[0]
        self.assertEqual((row["grid_before"], row["grid_after"]), ("1", "2"))
        cells = nodes(tables[0], "cell")
        self.assertEqual((cells[0]["grid_span"], cells[0]["vertical_merge"]), ("2", "restart"))
        self.assertEqual(cells[-1]["vertical_merge"], "continue")
        self.assertEqual(len(nodes(cells[0], "table")), 1)

    def test_revision_views_preserve_wrappers_and_namespace_boundaries(self):
        self.package('<w:p><w:ins w:id="1"><w:r><w:t>I</w:t></w:r></w:ins>'
                     '<w:del w:id="2"><w:r><w:delText>D</w:delText></w:r></w:del>'
                     '<x:ins><w:r><w:t>vendor</w:t></w:r></x:ins></w:p>')
        for mode, expected in (("accepted", ["I", "vendor"]), ("original", ["D", "vendor"]), ("all", ["I", "D", "vendor"])):
            with self.subTest(mode=mode):
                result = read_docx_structure(self.source, revisions=mode)
                root = result["parts"][0]["root"]
                self.assertEqual(result["revisions"], mode)
                self.assertEqual([node["text"] for node in nodes(root, "text")], expected)
                revisions = nodes(root, "revision")
                self.assertEqual(len(revisions), 2)
                self.assertEqual(sum(node.get("omitted") == "revision-view" for node in revisions), 0 if mode == "all" else 1)
                self.assertEqual(nodes(root, "opaque")[0]["namespace"], "urn:vendor")

    def test_hyperlinks_content_controls_and_opaque_nodes_are_not_flattened(self):
        self.package('<w:sdt><w:sdtPr><w:tag w:val="customer"/></w:sdtPr><w:sdtContent>'
                     '<w:p><w:hyperlink r:id="rId9"><w:r><w:t>link</w:t></w:r></w:hyperlink>'
                     '<x:extension x:kind="bound">unknown<!--keep--><?keep data?></x:extension></w:p>'
                     '</w:sdtContent></w:sdt>')
        root = read_docx_structure(self.source)["parts"][0]["root"]
        self.assertEqual(len(nodes(root, "content_control")), 1)
        self.assertEqual(nodes(root, "hyperlink")[0]["attributes"]["r:id"], "rId9")
        opaque = next(node for node in nodes(root, "opaque") if node["tag"] == "x:extension")
        self.assertEqual(opaque["attributes"]["x:kind"], "bound")
        self.assertEqual([node["type"] for node in opaque["children"]], ["text_value", "comment", "processing_instruction"])

    def test_definitions_are_raw_and_related_parts_are_not_page_order(self):
        self.package('<w:p/>', **{
            "word/styles.xml": f'<w:styles xmlns:w="{W}"><w:style w:styleId="A"/></w:styles>',
            "word/numbering.xml": f'<w:numbering xmlns:w="{W}"><w:num w:numId="7"/></w:numbering>',
            "word/header1.xml": f'<w:hdr xmlns:w="{W}"><w:p/></w:hdr>',
            "word/footer2.xml": f'<w:ftr xmlns:w="{W}"><w:p/></w:ftr>',
            "word/footnotes.xml": f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="-1" w:type="separator"><w:p/></w:footnote><w:footnote w:id="1"><w:p/></w:footnote></w:footnotes>',
        })
        result = read_docx_structure(self.source)
        self.assertEqual([part["name"] for part in result["parts"]], ["word/document.xml"])
        self.assertEqual(set(result["definitions"]), {"word/styles.xml", "word/numbering.xml"})
        result = read_docx_structure(self.source, include_related=True)
        self.assertEqual([part["name"] for part in result["parts"]], ["word/document.xml", "word/header1.xml", "word/footer2.xml", "word/footnotes.xml"])
        self.assertEqual(result["semantics"]["part_order"], "main-then-related-not-page-order")
        self.assertEqual(len(nodes(result["parts"][-1]["root"], "note")), 2)

    def test_cli_json_and_default_text_and_failure_preserve_output(self):
        self.package('<w:p><w:r><w:t>中文</w:t></w:r></w:p>')
        script = ROOT / "docx/scripts/read_docx.py"
        def cli(*args):
            return subprocess.run([sys.executable, "-I", "-S", "-B", str(script), str(self.source), *args],
                                  cwd=self.work, env=dict(os.environ, PATH=""), capture_output=True, text=True)
        self.assertEqual(cli().stdout, "中文")
        result = cli("--format", "json")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(result.stdout)["schema_version"], 1)
        before = self.source.read_bytes()
        self.assertNotEqual(cli("--format", "json", "-o", str(self.source)).returncode, 0)
        self.assertEqual(self.source.read_bytes(), before)
        output = self.work / "existing.json"
        output.write_text("keep", encoding="utf-8")
        self.source.write_bytes(b"broken")
        self.assertNotEqual(cli("--format", "json", "-o", str(output)).returncode, 0)
        self.assertEqual(output.read_text(), "keep")

    def test_unsafe_xml_and_invalid_revision_view_are_rejected(self):
        self.package('<w:p/>')
        with self.assertRaises(ValueError):
            read_docx_structure(self.source, revisions="invalid")
        with zipfile.ZipFile(self.source, "w") as package:
            package.writestr("word/document.xml", '<!DOCTYPE x [<!ENTITY e "boom">]><x>&e;</x>')
        with self.assertRaisesRegex(ValueError, "invalid or unsafe XML"):
            read_docx_structure(self.source)

    def test_total_node_budget_fails_without_silent_truncation(self):
        self.package('<w:p><w:r><w:t>完整内容</w:t></w:r></w:p>', **{
            "word/header1.xml": f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>页眉</w:t></w:r></w:p></w:hdr>',
        })
        with mock.patch.object(read_docx, "MAX_STRUCTURED_NODES", 6):
            self.assertTrue(read_docx_structure(self.source)["parts"])
            with self.assertRaisesRegex(ValueError, "total node limit"):
                read_docx_structure(self.source, include_related=True)

    def test_wrong_related_root_and_fake_namespace_document_are_rejected(self):
        self.package('<w:p/>', **{"word/header1.xml": '<x:hdr xmlns:x="urn:vendor"/>'})
        with self.assertRaisesRegex(ValueError, "w:hdr root"):
            read_docx_structure(self.source, include_related=True)
        with zipfile.ZipFile(self.source, "w") as package:
            package.writestr("word/document.xml", '<w:document xmlns:w="urn:vendor"><w:body/></w:document>')
        with self.assertRaisesRegex(ValueError, "w:document root"):
            read_docx_structure(self.source)


if __name__ == "__main__":
    unittest.main()
