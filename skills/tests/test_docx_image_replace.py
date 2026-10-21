"""Media-part replacement verifies hashes, relationships, usage and preservation."""

import hashlib
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
from replace_image import replace_image
from office import xmlio
from test_docx_generator import png_chunk, tiny_png
from test_docx_stdlib_core import W, R, write_package

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"


class ImageReplaceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-image-replace-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source, self.output = self.root / "source.docx", self.root / "out.docx"
        self.original_image = self.root / "original.png"
        self.replacement = self.root / "replacement.png"
        self.original_image.write_bytes(tiny_png())
        self.replacement.write_bytes(tiny_png()[:-12] + png_chunk(b"tEXt", b"Description\0replacement") + tiny_png()[-12:])
        create_document({"blocks": [{"type": "paragraph", "text": "图片说明"},
                                    {"type": "image", "path": str(self.original_image),
                                     "width_mm": 25, "height_mm": 20, "alt": "像素图片"}]}, self.source)
        with zipfile.ZipFile(self.source) as package:
            self.media_part = next(name for name in package.namelist() if name.startswith("word/media/"))
            self.digest = hashlib.sha256(package.read(self.media_part)).hexdigest()

    def specification(self, **updates):
        value = {"media_part": self.media_part, "expected_sha256": self.digest,
                 "expected_relationships": 1, "expected_references": 1}
        value.update(updates)
        return value

    def run_replace(self, **updates):
        return replace_image(self.source, self.replacement, self.output, self.specification(**updates))

    def rejected(self, **updates):
        before = self.source.read_bytes()
        self.output.write_bytes(b"existing output")
        with self.assertRaises(ValueError):
            self.run_replace(**updates)
        self.assertEqual(self.source.read_bytes(), before)
        self.assertEqual(self.output.read_bytes(), b"existing output")

    def mutate(self, callback):
        with zipfile.ZipFile(self.source) as package:
            parts = {name: package.read(name) for name in package.namelist()}
        callback(parts)
        write_package(self.source, parts)

    def add_shared_drawing(self, separate_relationship=False):
        def shared(parts):
            dom = xmlio.parse_dom(parts["word/document.xml"])
            drawing = dom.getElementsByTagNameNS(W, "drawing")[0]
            run = drawing.parentNode.cloneNode(True)
            run.getElementsByTagNameNS(WP, "docPr")[0].setAttribute("id", "900")
            if separate_relationship:
                blip = run.getElementsByTagNameNS(A, "blip")[0]
                rid = blip.getAttributeNS(R, "embed")
                rels = xmlio.parse_dom(parts["word/_rels/document.xml.rels"])
                relation = next(node for node in rels.getElementsByTagNameNS(REL, "Relationship")
                                if node.getAttribute("Id") == rid).cloneNode(True)
                relation.setAttribute("Id", "rShared")
                rels.documentElement.appendChild(relation)
                parts["word/_rels/document.xml.rels"] = rels.toxml(encoding="UTF-8")
                blip.setAttributeNS(R, "r:embed", "rShared")
            drawing.parentNode.parentNode.appendChild(run)
            parts["word/document.xml"] = dom.toxml(encoding="UTF-8")
        self.mutate(shared)

    def test_single_reference_changes_only_requested_media_part(self):
        self.run_replace()
        with zipfile.ZipFile(self.source) as before, zipfile.ZipFile(self.output) as after:
            self.assertIsNone(after.testzip())
            self.assertEqual(set(before.namelist()), set(after.namelist()))
            for name in before.namelist():
                if name == self.media_part:
                    self.assertEqual(after.read(name), self.replacement.read_bytes())
                else:
                    self.assertEqual(before.read(name), after.read(name), name)

    def test_shared_media_requires_explicit_all_and_both_exact_counts(self):
        for separate in (False, True):
            with self.subTest(separate=separate):
                self.add_shared_drawing(separate)
                relationships = 2 if separate else 1
                self.rejected(expected_relationships=relationships, expected_references=2)
                self.rejected(all_references=True, expected_relationships=relationships, expected_references=1)
                self.run_replace(all_references=True, expected_relationships=relationships, expected_references=2)
                with zipfile.ZipFile(self.output) as package:
                    self.assertEqual(package.read(self.media_part), self.replacement.read_bytes())
                # Restore source for the next independent shared-reference case.
                self._reset_original()

    def _reset_original(self):
        create_document({"blocks": [{"type": "image", "path": str(self.original_image),
                                     "width_mm": 25, "height_mm": 20, "alt": "像素图片"}]}, self.source)
        with zipfile.ZipFile(self.source) as package:
            self.media_part = next(name for name in package.namelist() if name.startswith("word/media/"))

    def test_invalid_specification_preserves_existing_output(self):
        for updates in ({"expected_sha256": "0" * 64}, {"expected_sha256": []},
                        {"expected_references": True}, {"expected_relationships": 2},
                        {"expected_references": 0}, {"all_references": "true"},
                        {"media_part": "../" + self.media_part}, {"media_part": []}, {"unknown": True}):
            with self.subTest(updates=updates):
                self.rejected(**updates)

    def test_replacement_magic_extension_and_content_type_must_agree(self):
        self.replacement.write_bytes(b"not a PNG")
        self.rejected()
        self.replacement.write_bytes(tiny_png())
        original_path = self.replacement
        self.replacement = self.root / "wrong.jpg"
        self.replacement.write_bytes(tiny_png())
        self.rejected()
        self.replacement = original_path
        def types(parts):
            parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(b'image/png', b'image/jpeg')
        self.mutate(types)
        self.rejected()

    def test_truncated_png_and_jpeg_are_rejected_before_writing(self):
        for name, content in (("truncated.png", b"\x89PNG\r\n\x1a\n"),
                              ("truncated.jpg", b"\xff\xd8\xff\xe0missing EOI")):
            with self.subTest(name=name):
                self.replacement = self.root / name
                self.replacement.write_bytes(content)
                self.rejected()

    def test_non_image_and_orphan_relationships_to_target_are_rejected(self):
        for mode in ("non-image", "orphan"):
            self._reset_original()
            def relation(parts):
                rels = xmlio.parse_dom(parts["word/_rels/document.xml.rels"])
                image = next(node for node in rels.getElementsByTagNameNS(REL, "Relationship")
                             if node.getAttribute("Type") == R + "/image")
                if mode == "non-image":
                    image.setAttribute("Type", R + "/hyperlink")
                else:
                    extra = image.cloneNode(True)
                    extra.setAttribute("Id", "rOrphan")
                    rels.documentElement.appendChild(extra)
                parts["word/_rels/document.xml.rels"] = rels.toxml(encoding="UTF-8")
            self.mutate(relation)
            self.rejected(all_references=True, expected_relationships=2 if mode == "orphan" else 1)

    def test_unknown_extension_and_vml_relationship_references_are_rejected(self):
        for mode in ("extension", "vml"):
            self._reset_original()
            def reference(parts):
                dom = xmlio.parse_dom(parts["word/document.xml"])
                blip = dom.getElementsByTagNameNS(A, "blip")[0]
                rid = blip.getAttributeNS(R, "embed")
                if mode == "extension":
                    dom.documentElement.setAttribute("xmlns:x", "urn:vendor")
                    dom.documentElement.setAttributeNS("urn:vendor", "x:reference", rid)
                else:
                    blip.setAttributeNS(R, "r:id", rid)
                    blip.removeAttributeNS(R, "embed")
                parts["word/document.xml"] = dom.toxml(encoding="UTF-8")
            self.mutate(reference)
            self.rejected()

    def test_original_magic_and_missing_part_are_rejected(self):
        def corrupt(parts):
            parts[self.media_part] = b"corrupt"
        self.mutate(corrupt)
        self.rejected(expected_sha256=hashlib.sha256(b"corrupt").hexdigest())
        self.rejected(media_part="word/media/missing.png")

    def test_unknown_picture_container_and_mixed_link_embed_are_rejected(self):
        for mode in ("container", "linked"):
            self._reset_original()
            def mutate(parts):
                dom = xmlio.parse_dom(parts["word/document.xml"])
                blip = dom.getElementsByTagNameNS(A, "blip")[0]
                if mode == "container":
                    dom.documentElement.setAttribute("xmlns:x", "urn:vendor")
                    fill = blip.parentNode
                    wrapper = dom.createElementNS("urn:vendor", "x:wrapper")
                    fill.parentNode.replaceChild(wrapper, fill)
                    wrapper.appendChild(fill)
                else:
                    blip.setAttributeNS(R, "r:link", blip.getAttributeNS(R, "embed"))
                parts["word/document.xml"] = dom.toxml(encoding="UTF-8")
            self.mutate(mutate)
            self.rejected()

    def test_jpeg_family_and_dotx_are_supported_without_decoding(self):
        # Signature-only fixture: verifies routing, not JPEG decoding/rendering.
        jpeg = self.root / "original.jpg"
        jpeg.write_bytes(b"\xff\xd8\xff\xe0original\xff\xd9")
        self.replacement = self.root / "replacement.jpeg"
        self.replacement.write_bytes(b"\xff\xd8\xff\xe0replacement\xff\xd9")
        create_document({"blocks": [{"type": "image", "path": str(jpeg),
                                     "width_mm": 25, "height_mm": 20, "alt": "签名测试"}]}, self.source)
        with zipfile.ZipFile(self.source) as package:
            parts = {name: package.read(name) for name in package.namelist()}
        self.media_part = next(name for name in parts if name.startswith("word/media/"))
        self.digest = hashlib.sha256(parts[self.media_part]).hexdigest()
        parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
            b'wordprocessingml.document.main+xml', b'wordprocessingml.template.main+xml',
        )
        self.source, self.output = self.root / "source.dotx", self.root / "out.dotx"
        write_package(self.source, parts)
        self.run_replace()
        with zipfile.ZipFile(self.output) as package:
            self.assertEqual(package.read(self.media_part), self.replacement.read_bytes())

    def test_signed_packages_are_rejected(self):
        self.mutate(lambda parts: parts.update({"_xmlsignatures/signature.xml": b"<signature/>"}))
        self.rejected()

    def test_same_file_alias_and_extension_conversion_are_rejected(self):
        for destination in (self.source, self.output.with_suffix(".dotx")):
            with self.assertRaises(ValueError):
                replace_image(self.source, self.replacement, destination, self.specification())
        alias = self.root / "alias.docx"
        alias.symlink_to(self.source)
        with self.assertRaises(ValueError):
            replace_image(self.source, self.replacement, alias, self.specification())

    def test_cli_works_without_external_programs(self):
        (self.root / "image.json").write_text(json.dumps(self.specification()), encoding="utf-8")
        result = subprocess.run([sys.executable, "-I", "-S", "-B", str(ROOT / "docx/scripts/replace_image.py"),
                                 "source.docx", "replacement.png", "image.json", "out.docx"], cwd=self.root,
                                env=dict(os.environ, PATH=""), capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()
