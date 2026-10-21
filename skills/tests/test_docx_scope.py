"""Keep the Word validator scoped to DOCX and DOTX packages."""

import os
import re
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VALIDATE = ROOT / "docx" / "scripts" / "office" / "validate.py"


class DocxScopeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-scope-")
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name)
        self.env = dict(os.environ, PATH="", PYTHONDONTWRITEBYTECODE="1")
        self.original = self.work / "output.docx"
        instructions = (ROOT / "docx" / "SKILL.md").read_text(encoding="utf-8")
        example = re.search(r"```json\n(.*?)```", instructions, re.S).group(1)
        spec = self.work / "document.json"
        spec.write_text(example, encoding="utf-8")
        result = subprocess.run(
            [sys.executable, "-I", "-S", "-B", str(ROOT / "docx/scripts/create_docx.py"), str(spec), str(self.original)], cwd=self.work,
            env=self.env, capture_output=True, text=True,
        )
        if result.returncode != 0:
            raise AssertionError(result.stdout + result.stderr)

    def run_cli(self, *args):
        return subprocess.run(
            [sys.executable, "-I", "-S", "-B", str(VALIDATE), *map(str, args)],
            cwd=self.work, env=self.env, capture_output=True, text=True,
        )

    def test_rejects_non_word_packed_inputs_even_with_docx_original(self):
        for suffix in (".pptx", ".xlsx", ".potx", ".xltx"):
            with self.subTest(suffix=suffix):
                packed = self.work / f"input{suffix}"
                packed.write_bytes(b"not a package")
                result = self.run_cli(packed, "--original", self.original)
                self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
                self.assertIn("Packed input", result.stderr)

    def test_accepts_docx_and_dotx_packed_inputs(self):
        docx = self.work / "input.docx"
        docx.write_bytes(self.original.read_bytes())
        docx_result = self.run_cli(docx)
        self.assertEqual(docx_result.returncode, 0, docx_result.stdout + docx_result.stderr)
        self.assertIn("Basic structure checks PASSED.", docx_result.stdout)

        dotx = self.work / "input.dotx"
        document_content_type = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        )
        template_content_type = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
        )
        with zipfile.ZipFile(self.original) as source, zipfile.ZipFile(dotx, "w") as target:
            for entry in source.infolist():
                data = source.read(entry.filename)
                if entry.filename == "[Content_Types].xml":
                    content_types = data.decode("utf-8")
                    self.assertIn(document_content_type, content_types)
                    data = content_types.replace(
                        document_content_type, template_content_type, 1
                    ).encode("utf-8")
                target.writestr(entry, data)
        dotx_result = self.run_cli(dotx)
        self.assertEqual(dotx_result.returncode, 0, dotx_result.stdout + dotx_result.stderr)
        self.assertIn("Basic structure checks PASSED.", dotx_result.stdout)

    def test_accepts_unpacked_directory_with_docx_original(self):
        unpacked = self.work / "unpacked"
        unpacked.mkdir()
        with zipfile.ZipFile(self.original) as package:
            package.extractall(unpacked)
        result = self.run_cli(unpacked, "--original", self.original)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("Basic structure checks PASSED.", result.stdout)


if __name__ == "__main__":
    unittest.main()
