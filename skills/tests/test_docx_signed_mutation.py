"""Reject signed-package markers before document mutations, without deleting them."""

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "docx/scripts"))
from comment import add_comment
from edit_docx import replace_text
from edit_structure import edit_structure
from office.package import unpack_document
from test_docx_stdlib_core import package_parts, write_package


class SignedMutationTests(unittest.TestCase):
    def test_existing_mutators_preserve_source_and_outputs_when_signature_markers_exist(self):
        for relationship in (False, True):
            with self.subTest(relationship=relationship), tempfile.TemporaryDirectory() as temporary:
                root = Path(temporary)
                source, output = root / "source.docx", root / "edited.docx"
                name = "custom/signature.xml" if relationship else "_xmlsignatures/sig1.xml"
                relation = ('<Relationship Id="sig" Type="http://schemas.openxmlformats.org/package/2006/'
                            'relationships/digital-signature/signature" Target="custom/signature.xml"/>') if relationship else ""
                write_package(source, package_parts(extra={name: "<simulated/>"}, rels_extra=relation))
                original = source.read_bytes()
                output.write_bytes(b"keep")
                with self.assertRaisesRegex(ValueError, "Signed"):
                    replace_text(source, output, {"replacements": [{"old": "Hello", "new": "New"}]})
                with self.assertRaisesRegex(ValueError, "Signed"):
                    edit_structure(source, output, {"operations": [{"type": "insert_paragraph",
                        "paragraph_index": 0, "expected_text": "Hello", "position": "after", "text": "New"}]})
                self.assertEqual(source.read_bytes(), original)
                self.assertEqual(output.read_bytes(), b"keep")
                directory = unpack_document(source, root / "package")
                before = {path.relative_to(directory): path.read_bytes()
                          for path in directory.rglob("*") if path.is_file()}
                with self.assertRaisesRegex(ValueError, "Signed"):
                    add_comment(directory, "Comment")
                after = {path.relative_to(directory): path.read_bytes()
                         for path in directory.rglob("*") if path.is_file()}
                self.assertEqual(after, before)


if __name__ == "__main__":
    unittest.main()
