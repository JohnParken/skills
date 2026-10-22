"""Replacement count failures identify common template mismatch causes."""

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "docx" / "scripts"))
from create_docx import create_document  # noqa: E402
from edit_docx import replace_text  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", W)


def q(local):
    return f"{{{W}}}{local}"


class TemplateDiagnosticTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-template-diagnostics-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / "template.docx"
        self.output = self.root / "filled.docx"

    def create_template(self, *paragraphs):
        create_document({
            "blocks": [
                {"type": "paragraph", "text": text}
                for text in paragraphs
            ],
        }, self.source)

    def patch_document(self, update):
        with zipfile.ZipFile(self.source) as package:
            parts = {name: package.read(name) for name in package.namelist()}
        document = ET.fromstring(parts["word/document.xml"])
        update(document)
        parts["word/document.xml"] = ET.tostring(
            document, encoding="UTF-8", xml_declaration=True,
        )
        patched = self.root / "patched-template.docx"
        with zipfile.ZipFile(patched, "w", compression=zipfile.ZIP_DEFLATED) as package:
            for name, content in parts.items():
                package.writestr(name, content)
        self.source = patched

    def assert_rejected_without_replacing_output(self, old, message, *, count=1):
        before = self.source.read_bytes()
        sentinel = b"keep existing output"
        self.output.write_bytes(sentinel)
        with self.assertRaisesRegex(ValueError, r"Expected .* supported occurrence\(s\)") as error:
            replace_text(self.source, self.output, {
                "replacements": [{"old": old, "new": "replacement", "count": count}],
            })
        details = str(error.exception)
        self.assertIn(message, details)
        self.assertIn(
            "cross-run, field, revision and complex-container edits are unsupported",
            details,
        )
        self.assertEqual(self.source.read_bytes(), before)
        self.assertEqual(self.output.read_bytes(), sentinel)

    def test_absent_search_text_is_reported(self):
        self.create_template("ordinary template text")
        self.assert_rejected_without_replacing_output(
            "{{company}}", "The search text is absent from the selected parts.",
        )

    def test_text_split_across_runs_in_one_paragraph_is_reported(self):
        self.create_template("{{company}}")

        def split_run(document):
            paragraph = document.find(f".//{q('p')}")
            first_run = next(child for child in paragraph if child.tag == q("r"))
            first_text = first_run.find(q("t"))
            first_text.text = "{{com"
            second_run = ET.SubElement(paragraph, q("r"))
            ET.SubElement(second_run, q("t")).text = "pany}}"

        self.patch_document(split_run)
        self.assert_rejected_without_replacing_output(
            "{{company}}", "split across text/run nodes within a paragraph",
        )

    def test_occurrence_in_content_control_is_reported_as_protected(self):
        self.create_template("{{company}}")

        def wrap_run(document):
            paragraph = document.find(f".//{q('p')}")
            run = next(child for child in paragraph if child.tag == q("r"))
            paragraph.remove(run)
            control = ET.Element(q("sdt"))
            ET.SubElement(control, q("sdtPr"))
            content = ET.SubElement(control, q("sdtContent"))
            content.append(run)
            paragraph.append(control)

        self.patch_document(wrap_run)
        self.assert_rejected_without_replacing_output(
            "{{company}}", "protected or complex content",
        )

    def test_excess_ordinary_matches_are_reported(self):
        self.create_template("repeat", "repeat")
        self.assert_rejected_without_replacing_output(
            "repeat", "Too many ordinary eligible occurrences were found.",
        )


if __name__ == "__main__":
    unittest.main()
