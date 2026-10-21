"""Security and semantic preservation for the deliberately plain-run merger."""

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
from merge_runs import merge_runs
from office.package import pack_document, unpack_document
from office.xmlio import parse_dom

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class MergeSecurityTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-merge-security-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = create_document({"blocks": [{"type": "paragraph", "runs": [
            {"text": "A"}, {"text": "B"}]}]}, self.root / "input.docx")
        self.directory = unpack_document(self.source, self.root / "parts")
        self.xml = self.directory / "word/document.xml"

    def document(self, body, attrs=""):
        self.xml.write_text(f'<w:document xmlns:w="{W}" xmlns:x="urn:vendor" {attrs}>'
                            f'<w:body>{body}</w:body></w:document>', encoding="utf-8")

    def unchanged(self, directory=None):
        original = self.xml.read_bytes()
        count, message = merge_runs(str(directory or self.directory))
        self.assertEqual(count, 0)
        self.assertTrue(message.startswith("Error"), message)
        self.assertEqual(self.xml.read_bytes(), original)

    def test_document_and_parent_directory_symlinks_cannot_modify_external_files(self):
        outside = self.root / "outside.xml"
        outside.write_bytes(self.xml.read_bytes())
        self.xml.unlink()
        self.xml.symlink_to(outside)
        original = outside.read_bytes()
        self.unchanged()
        self.assertEqual(outside.read_bytes(), original)
        self.xml.unlink()
        self.xml.write_bytes(original)
        outside_dir = self.root / "outside-word"
        (self.directory / "word").rename(outside_dir)
        (self.directory / "word").symlink_to(outside_dir, target_is_directory=True)
        self.unchanged()
        self.assertEqual(outside.read_bytes(), original)

    def test_symlinked_input_root_is_rejected(self):
        alias = self.root / "alias"
        alias.symlink_to(self.directory, target_is_directory=True)
        self.unchanged(alias)

    def test_foreign_proof_errors_and_rsid_attributes_remain_exact(self):
        for body in (
            '<w:p><x:proofErr x:meaning="keep"/><w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r></w:p>',
            '<w:p><w:r x:rsid="keep"><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r></w:p>',
            '<w:p><w:r><w:t>A</w:t></w:r><w:r x:binding="keep"><w:t>B</w:t></w:r></w:p>',
            '<w:p x:binding="keep"><w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r></w:p>',
            '<w:p><w:pPr><x:binding/></w:pPr><w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r></w:p>',
        ):
            with self.subTest(body=body):
                self.document(body)
                original = self.xml.read_bytes()
                count, message = merge_runs(str(self.directory))
                self.assertFalse(message.startswith("Error"), message)
                self.assertEqual(count, 0)
                self.assertEqual(self.xml.read_bytes(), original)

    def test_inherited_and_overridden_space_are_respected(self):
        for attrs, first_attrs, second_attrs, expected in (
            ('xml:space="preserve"', "", "", "A  B"),
            ('xml:space="preserve"', 'xml:space="default"', "", "A B"),
            ("", 'xml:space="preserve"', 'xml:space="preserve"', "A  B"),
            ("", "", "", "AB"),
        ):
            with self.subTest(attrs=(attrs, first_attrs, second_attrs)):
                self.document(f'<w:p><w:r {first_attrs}><w:t>A </w:t></w:r>'
                              f'<w:r {second_attrs}><w:t> B</w:t></w:r></w:p>', attrs=attrs)
                count, message = merge_runs(str(self.directory))
                self.assertEqual(count, 1, message)
                dom = parse_dom(self.xml.read_bytes())
                texts = dom.getElementsByTagNameNS(W, "t")
                self.assertEqual(len(texts), 1)
                self.assertEqual(texts[0].firstChild.data, expected)

    def test_non_target_character_reference_values_cannot_be_normalized(self):
        for foreign in ('<x:data x:val="first&#10;second&#9;third&#13;last"/>',
                        '<x:data>A&#13;B</x:data>'):
            with self.subTest(foreign=foreign):
                self.document(f'{foreign}<w:p><w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r></w:p>')
                self.unchanged()

    def test_atomic_write_failure_preserves_original_file_and_cleans_probe(self):
        original = self.xml.read_bytes()
        before = set(self.xml.parent.iterdir())
        with mock.patch("office.helpers.os.replace", side_effect=OSError("injected failure")):
            count, message = merge_runs(str(self.directory))
        self.assertEqual(count, 0)
        self.assertTrue(message.startswith("Error"), message)
        self.assertEqual(self.xml.read_bytes(), original)
        self.assertEqual(set(self.xml.parent.iterdir()), before)

    def test_signatures_are_rejected_by_directory_and_packaged_cli(self):
        signature = self.directory / "_xmlsignatures/sig1.xml"
        signature.parent.mkdir()
        signature.write_text('<signature/>', encoding="utf-8")
        self.unchanged()
        signed = pack_document(self.directory, self.root / "signed.docx")
        original = signed.read_bytes()
        output = self.root / "output.docx"
        output.write_bytes(b"keep")
        result = self.cli(signed, "-o", output)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(signed.read_bytes(), original)
        self.assertEqual(output.read_bytes(), b"keep")

    def cli(self, source, *args):
        return subprocess.run([sys.executable, "-I", "-S", "-B", str(ROOT / "docx/scripts/merge_runs.py"),
                               str(source), *map(str, args)], env=dict(os.environ, PATH=""),
                              cwd=self.root, capture_output=True, text=True)

    def test_cli_rejects_link_outputs_and_cross_format_names(self):
        original = self.source.read_bytes()
        output = self.root / "linked.docx"
        output.symlink_to(self.source)
        result = self.cli(self.source, "-o", output)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(self.source.read_bytes(), original)
        output.unlink()
        converted = self.root / "wrong.dotx"
        converted.write_bytes(b"keep")
        result = self.cli(self.source, "-o", converted)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(converted.read_bytes(), b"keep")

    def test_fields_revisions_complex_nodes_and_nested_unknown_formatting_are_unchanged(self):
        for content in (
            '<w:ins w:id="0" w:author="Reviewer"><w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r></w:ins>',
            '<w:r><w:fldChar w:fldCharType="begin"/><w:t>A</w:t></w:r><w:r><w:t>B</w:t><w:fldChar w:fldCharType="end"/></w:r>',
            '<w:r><w:rPr><w:b x:dependency="keep"/></w:rPr><w:t>A</w:t></w:r><w:r><w:rPr><w:b x:dependency="keep"/></w:rPr><w:t>B</w:t></w:r>',
            '<w:r><w:t>A<!--keep--></w:t></w:r><w:r><w:t>B</w:t></w:r>',
        ):
            with self.subTest(content=content):
                self.document('<w:p>' + content + '</w:p>')
                original = self.xml.read_bytes()
                count, message = merge_runs(str(self.directory))
                self.assertEqual(count, 0)
                self.assertFalse(message.startswith("Error"), message)
                self.assertEqual(self.xml.read_bytes(), original)

    def test_plain_merge_remains_valid_and_preserves_non_document_parts(self):
        self.document('<w:p><w:proofErr w:type="spellStart"/>'
                      '<w:r w:rsidR="00000001"><w:rPr><w:b/></w:rPr><w:t>中文</w:t></w:r>'
                      '<w:r w:rsidR="00000002"><w:rPr><w:b/></w:rPr><w:t>English</w:t></w:r>'
                      '<w:proofErr w:type="spellEnd"/></w:p>')
        packaged = pack_document(self.directory, self.root / "plain.docx")
        output = self.root / "merged.docx"
        result = self.cli(packaged, "-o", output)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        with zipfile.ZipFile(packaged) as before, zipfile.ZipFile(output) as after:
            self.assertIsNone(after.testzip())
            self.assertEqual(set(before.namelist()), set(after.namelist()))
            for name in before.namelist():
                if name != "word/document.xml":
                    self.assertEqual(before.read(name), after.read(name))
            dom = parse_dom(after.read("word/document.xml"))
            self.assertEqual(len(dom.getElementsByTagNameNS(W, "r")), 1)
            self.assertEqual(dom.getElementsByTagNameNS(W, "t")[0].firstChild.data, "中文English")
            self.assertFalse(dom.getElementsByTagNameNS(W, "proofErr"))

    def test_relation_only_signature_marker_and_invalid_space_are_rejected(self):
        relationships = self.directory / "_rels/.rels"
        original = relationships.read_text(encoding="utf-8")
        dom = parse_dom(original)
        relation_namespace = "http://schemas.openxmlformats.org/package/2006/relationships"
        node = dom.createElementNS(relation_namespace, "rel:Relationship")
        node.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:rel", relation_namespace)
        for name, value in (("Id", "signature"), ("Type", "urn:test/digital-signature/origin"),
                            ("Target", "https://example.invalid/signature"), ("TargetMode", "External")):
            node.setAttribute(name, value)
        dom.documentElement.appendChild(node)
        relationships.write_bytes(dom.toxml(encoding="UTF-8"))
        self.unchanged()
        relationships.write_text(original, encoding="utf-8")
        self.document('<w:p><w:r><w:t xml:space="invalid">A</w:t></w:r><w:r><w:t>B</w:t></w:r></w:p>')
        self.unchanged()

    def test_cross_paragraph_fields_and_local_namespace_aliases(self):
        self.document('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p>'
                      '<w:p><w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r></w:p>'
                      '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
        original = self.xml.read_bytes()
        count, message = merge_runs(str(self.directory))
        self.assertEqual(count, 0, message)
        self.assertEqual(self.xml.read_bytes(), original)
        self.document('<w:p><w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r></w:p>')
        alias = self.xml.read_text(encoding="utf-8").replace('w:', 'q:').replace('xmlns:w=', 'xmlns:q=')
        self.xml.write_text(alias, encoding="utf-8")
        mode = self.xml.stat().st_mode
        count, message = merge_runs(str(self.directory))
        self.assertEqual(count, 1, message)
        self.assertEqual(self.xml.stat().st_mode, mode)
        dom = parse_dom(self.xml.read_bytes())
        self.assertEqual(dom.getElementsByTagNameNS(W, "t")[0].firstChild.data, "AB")

    def test_packaged_extension_content_type_mismatch_is_rejected(self):
        wrong = self.root / "renamed.dotx"
        wrong.write_bytes(self.source.read_bytes())
        output = self.root / "wrong-output.dotx"
        output.write_bytes(b"keep")
        original = wrong.read_bytes()
        result = self.cli(wrong, "-o", output)
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual(wrong.read_bytes(), original)
        self.assertEqual(output.read_bytes(), b"keep")


if __name__ == "__main__":
    unittest.main()
