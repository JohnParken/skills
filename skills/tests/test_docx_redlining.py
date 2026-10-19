"""Focused tests for DOCX redlining validation and text diagnostics."""

import contextlib
import io
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OFFICE = ROOT / "docx" / "scripts" / "office"
sys.path.insert(0, str(OFFICE))
from validators.redlining import RedliningValidator  # noqa: E402


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def document_xml(body):
    return (
        f'<w:document xmlns:w="{W_NS}"><w:body>{body}'
        "</w:body></w:document>"
    )


def paragraph(text):
    return f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"


class DocxRedliningTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.unpacked = self.root / "modified"
        (self.unpacked / "word").mkdir(parents=True)
        self.original = self.root / "original.docx"

    def tearDown(self):
        self.temp_dir.cleanup()

    def write_documents(self, original_body, modified_body):
        with zipfile.ZipFile(self.original, "w") as archive:
            archive.writestr("word/document.xml", document_xml(original_body))
        (self.unpacked / "word/document.xml").write_text(
            document_xml(modified_body), encoding="utf-8"
        )

    def validate_capturing_output(self):
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            result = RedliningValidator(self.unpacked, self.original).validate()
        return result, output.getvalue()

    def test_untracked_modification_fails_with_text_diff(self):
        self.write_documents(paragraph("before"), paragraph("after"))

        result, output = self.validate_capturing_output()

        self.assertFalse(result)
        self.assertIn("Differences:", output)
        self.assertIn("--- original", output)
        self.assertIn("+++ modified", output)
        self.assertIn("-before", output)
        self.assertIn("+after", output)

    def test_matching_and_properly_tracked_text_pass(self):
        self.write_documents(
            paragraph("Hello"),
            paragraph("Hello")
            + '<w:p><w:ins w:author="Editor" w:date="2026-01-01T00:00:00Z">'
            "<w:r><w:t> added</w:t></w:r></w:ins></w:p>",
        )

        result, output = self.validate_capturing_output()

        self.assertTrue(result)
        self.assertEqual("", output)

    def test_diff_handles_empty_unicode_multiline_and_context(self):
        validator = RedliningValidator(self.unpacked, self.original)
        empty_diff = validator._get_text_diff("", "café ☕\n東京\n")
        self.assertIn("--- original", empty_diff)
        self.assertIn("+++ modified", empty_diff)
        self.assertIn("+café ☕", empty_diff)
        self.assertIn("+東京", empty_diff)

        newline_diff = validator._get_text_diff("final line", "final line\n")
        self.assertIn("--- original", newline_diff)
        self.assertIn("- 'final line'", newline_diff)
        self.assertIn("+ 'final line\\n'", newline_diff)

        contextual_diff = validator._get_text_diff(
            "first\nkeep this context\nold value\nlast",
            "first\nkeep this context\nnew value\nlast",
        )
        self.assertIn(" keep this context", contextual_diff)
        self.assertIn("-old value", contextual_diff)
        self.assertIn("+new value", contextual_diff)
        self.assertIn(" last", contextual_diff)


if __name__ == "__main__":
    unittest.main()
