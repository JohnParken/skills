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

    def make_modified_package(self, *, tracked=False, malformed=False):
        modified = self.work / "modified.docx"
        with zipfile.ZipFile(self.original) as source, zipfile.ZipFile(modified, "w") as target:
            for entry in source.infolist():
                data = source.read(entry.filename)
                if entry.filename == "word/document.xml":
                    if malformed:
                        data = b"<not-xml"
                    elif tracked:
                        data = data.replace(
                            b"</w:p>",
                            b'<w:ins w:id="1" w:author="Reviewer" w:date="2026-01-01T00:00:00Z">'
                            b"<w:r><w:t>Added</w:t></w:r></w:ins></w:p>", 1,
                        )
                    else:
                        data = data.replace(b"The team", b"A team", 1)
                target.writestr(entry, data)
        return modified

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

    def test_original_alone_rejects_untracked_text(self):
        modified = self.make_modified_package()
        result = self.run_cli(modified, "--original", self.original)
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn("Document text doesn't match", result.stdout)

    def test_original_alone_accepts_valid_tracked_change(self):
        modified = self.make_modified_package(tracked=True)
        result = self.run_cli(modified, "--original", self.original)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_author_is_compatibility_option_with_same_result(self):
        modified = self.make_modified_package()
        default = self.run_cli(modified, "--original", self.original)
        compatibility = self.run_cli(
            modified, "--original", self.original, "--author", "ignored-value"
        )
        self.assertEqual(default.returncode, 1, default.stdout + default.stderr)
        self.assertEqual(compatibility.returncode, default.returncode,
                         compatibility.stdout + compatibility.stderr)
        self.assertIn("does not filter authors", compatibility.stdout)

    def test_author_without_original_remains_cli_error(self):
        result = self.run_cli(self.original, "--author", "Reviewer")
        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("--author requires --original", result.stderr)

    def test_structure_failure_still_fails_with_original(self):
        malformed = self.make_modified_package(malformed=True)
        result = self.run_cli(malformed, "--original", self.original)
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        self.assertIn("FAILED", result.stderr + result.stdout)


if __name__ == "__main__":
    unittest.main()
