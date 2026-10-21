"""Relationship semantics on small local fixtures; no network requests."""

import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "docx/scripts"))
from office.validators.structure import CT, R, REL, W, StructureValidator

A = "http://schemas.openxmlformats.org/drawingml/2006/main"


class RelationshipSecurityTests(unittest.TestCase):
    def validate(self, *, attribute="embed", target="media/item.bin", mode="", hyperlink=False):
        with tempfile.TemporaryDirectory(prefix="docx-relationship-security-") as temporary:
            root = Path(temporary)
            (root / "_rels").mkdir()
            (root / "word/_rels").mkdir(parents=True)
            (root / "word/media").mkdir()
            (root / "word/media/item.bin").write_bytes(b"local payload")
            (root / "[Content_Types].xml").write_text(
                f'<Types xmlns="{CT}">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Default Extension="bin" ContentType="application/octet-stream"/>'
                '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                '</Types>', encoding="utf-8")
            (root / "_rels/.rels").write_text(
                f'<Relationships xmlns="{REL}"><Relationship Id="main" Type="{R}/officeDocument" '
                'Target="word/document.xml"/></Relationships>', encoding="utf-8")
            content = ('<w:hyperlink r:id="item"><w:r><w:t>链接</w:t></w:r></w:hyperlink>'
                       if hyperlink else f'<w:r><w:drawing><a:blip r:{attribute}="item"/></w:drawing></w:r>')
            (root / "word/document.xml").write_text(
                f'<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:a="{A}">'
                f'<w:body><w:p>{content}</w:p></w:body></w:document>', encoding="utf-8")
            (root / "word/_rels/document.xml.rels").write_text(
                f'<Relationships xmlns="{REL}"><Relationship Id="item" Type="{R}/'
                f'{"hyperlink" if hyperlink else "image"}" Target="{target}" '
                f'TargetMode="{mode}"/></Relationships>', encoding="utf-8")
            validator = StructureValidator(root, extension=".docx")
            with contextlib.redirect_stdout(io.StringIO()):
                result = validator.validate()
            return result, validator.errors

    def test_external_embedded_relationship_is_rejected(self):
        valid, errors = self.validate(target="https://example.invalid/image.png", mode="External")
        self.assertFalse(valid)
        self.assertIn("must be internal", " ".join(errors))

    def test_existing_internal_embedded_relationship_remains_supported(self):
        valid, errors = self.validate()
        self.assertTrue(valid, errors)

    def test_external_hyperlinks_and_explicit_links_remain_supported(self):
        for hyperlink in (False, True):
            with self.subTest(hyperlink=hyperlink):
                valid, errors = self.validate(attribute="link", hyperlink=hyperlink,
                                              target="https://example.invalid/resource", mode="External")
                self.assertTrue(valid, errors)

    def test_internal_uri_and_package_escape_checks_remain_enforced(self):
        for target, diagnostic in (("https://example.invalid/image.png", "TargetMode=External"),
                                   ("../../../escape.bin", "escapes the package")):
            with self.subTest(target=target):
                valid, errors = self.validate(target=target)
                self.assertFalse(valid)
                self.assertIn(diagnostic, " ".join(errors))


if __name__ == "__main__":
    unittest.main()
