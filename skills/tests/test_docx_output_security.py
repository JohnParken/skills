"""Read-only extraction must not follow output links or truncate on failure."""

import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "docx/scripts"))
from create_docx import create_document
from office import helpers
from office import xmlio
from office.validators.structure import StructureValidator


class OutputSecurityTests(unittest.TestCase):
    def test_structure_loading_shares_xml_budget_across_parts(self):
        with tempfile.TemporaryDirectory(prefix="docx-output-security-") as temp:
            work = Path(temp)
            for name in ("[Content_Types].xml", "_rels/.rels", "word/document.xml"):
                path = work / name
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("<root/>", encoding="utf-8")
            validator = StructureValidator(work)
            with mock.patch.object(xmlio, "MAX_XML_ELEMENTS", 2):
                validator._load()
            self.assertTrue(any("shared elements" in error for error in validator.errors), validator.errors)

    def test_text_and_json_output_symlinks_do_not_modify_target(self):
        with tempfile.TemporaryDirectory(prefix="docx-output-security-") as temp:
            work = Path(temp)
            source = create_document({"blocks": [{"type": "paragraph", "text": "安全读取"}]}, work / "input.docx")
            target = work / "unrelated.txt"
            target.write_bytes(b"keep unrelated file")
            output = work / "output.txt"
            output.symlink_to(target)
            for output_format in ("text", "json"):
                with self.subTest(output_format=output_format):
                    result = subprocess.run([sys.executable, "-I", "-S", "-B",
                        str(ROOT / "docx/scripts/read_docx.py"), str(source), "--format", output_format,
                        "-o", str(output)], env=dict(os.environ, PATH=""), capture_output=True, text=True)
                    self.assertNotEqual(result.returncode, 0)
                    self.assertEqual(target.read_bytes(), b"keep unrelated file")
                    self.assertTrue(output.is_symlink())

    def test_atomic_output_write_failure_preserves_old_bytes_and_cleans_temporary(self):
        with tempfile.TemporaryDirectory(prefix="docx-output-security-") as temp:
            work = Path(temp)
            output = work / "output.txt"
            output.write_bytes(b"old bytes")
            with mock.patch.object(helpers.os, "replace", side_effect=OSError("simulated replacement failure")):
                with self.assertRaises(OSError):
                    helpers.atomic_write_bytes(output, "新内容".encode())
            self.assertEqual(output.read_bytes(), b"old bytes")
            self.assertEqual(list(work.iterdir()), [output])


if __name__ == "__main__":
    unittest.main()
