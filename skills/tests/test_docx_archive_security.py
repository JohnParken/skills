"""Portable ZIP names and repack limits must fail before publishing output."""

import sys
import struct
import tempfile
import unittest
import zipfile
import zlib
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "docx/scripts"))
from create_docx import create_document
from office import helpers
from office.package import pack_document, unpack_document


class ArchiveSecurityTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-archive-security-")
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name)
        self.source = create_document({"blocks": [{"type": "paragraph", "text": "正文"}]}, self.work / "input.docx")

    def test_case_unicode_and_directory_aliases_are_rejected_before_extract(self):
        for case_index, names in enumerate((("word/document.xml", "word/Document.xml"),
                      ("custom/café.bin", "custom/cafe\u0301.bin"),
                      ("custom/Path/a.bin", "custom/path/b.bin"),
                      ("custom/node", "custom/node/child.bin"))):
            with self.subTest(names=names):
                archive = self.work / "aliases.docx"
                with zipfile.ZipFile(self.source) as original, zipfile.ZipFile(archive, "w") as output:
                    for name in original.namelist():
                        if name not in names:
                            output.writestr(name, original.read(name))
                    for name in names:
                        output.writestr(name, b"small fixture")
                destination = self.work / f"extracted{case_index}"
                with self.assertRaises(ValueError):
                    unpack_document(archive, destination)
                self.assertFalse(destination.exists())

    def test_nul_truncated_original_zip_name_is_rejected(self):
        item = zipfile.ZipInfo("safe.bin\x00hidden")
        archive = mock.Mock()
        archive.infolist.return_value = [item]
        with self.assertRaises(ValueError):
            helpers.check_archive(archive)

    def test_nonportable_sanitized_names_and_extra_directory_slashes_are_rejected(self):
        for name in ('custom/a?.bin', 'custom/a*.bin', 'custom/quote".bin',
                     'custom/trailing.', 'custom/trailing ', 'custom/CON.bin',
                     'custom/com1.txt', 'custom/LPT²', 'word///'):
            with self.subTest(name=name):
                archive = mock.Mock()
                archive.infolist.return_value = [zipfile.ZipInfo(name)]
                with self.assertRaises(ValueError):
                    helpers.check_archive(archive)

    def test_archive_path_depth_and_encoded_length_have_explicit_limits(self):
        for name in ("a/" * 128 + "file.bin", "中文" * 1400 + ".bin"):
            with self.subTest(name_length=len(name)):
                archive = mock.Mock()
                archive.infolist.return_value = [zipfile.ZipInfo(name)]
                with self.assertRaises(ValueError):
                    helpers.check_archive(archive)

    def test_pack_enforces_part_total_and_member_limits_before_replace(self):
        directory = unpack_document(self.source, self.work / "package")
        output = self.work / "output.docx"
        for limit, value in (("MAX_PART_BYTES", 8), ("MAX_PACKAGE_BYTES", 8), ("MAX_MEMBERS", 1)):
            with self.subTest(limit=limit):
                output.write_bytes(b"keep existing output")
                with mock.patch.object(helpers, limit, value):
                    with self.assertRaises(ValueError):
                        pack_document(directory, output)
                self.assertEqual(output.read_bytes(), b"keep existing output")
                self.assertFalse(list(self.work.glob("output.docx.*.tmp")))

    def test_pack_rejects_existing_and_dangling_output_symlinks(self):
        directory = unpack_document(self.source, self.work / "package")
        for packer in (pack_document, helpers.rezip):
            for existing in (True, False):
                with self.subTest(packer=packer.__name__, existing=existing):
                    target = self.work / f"unrelated-{packer.__name__}-{existing}.docx"
                    if existing:
                        target.write_bytes(b"keep target")
                    output = self.work / f"linked-{packer.__name__}-{existing}.docx"
                    output.symlink_to(target)
                    before = set(self.work.iterdir())
                    with self.assertRaisesRegex(ValueError, "[Ss]ymbolic link"):
                        packer(directory, output)
                    self.assertTrue(output.is_symlink())
                    self.assertEqual(set(self.work.iterdir()), before)
                    if existing:
                        self.assertEqual(target.read_bytes(), b"keep target")
                        target.unlink()
                    else:
                        self.assertFalse(target.exists())
                    output.unlink()

    def test_repack_detects_case_aliases_even_on_case_sensitive_hosts(self):
        directory = unpack_document(self.source, self.work / "package")
        # Directory spelling aliases are represented as mocked ZIP metadata so
        # this regression also runs on case-insensitive APFS without overwrites.
        archive = mock.Mock()
        archive.infolist.return_value = [zipfile.ZipInfo("word/a.bin"), zipfile.ZipInfo("Word/b.bin")]
        with self.assertRaises(ValueError):
            helpers.check_archive(archive)

    def test_small_declared_parts_do_not_request_full_global_decompression_budget(self):
        archive = mock.MagicMock()
        info = zipfile.ZipInfo("small.bin")
        info.file_size = 4
        archive.getinfo.return_value = info
        stream = archive.open.return_value.__enter__.return_value
        stream.read.return_value = b"data"
        self.assertEqual(helpers.read_part(archive, "small.bin"), b"data")
        self.assertLessEqual(stream.read.call_args.args[0], info.file_size + 1)

    def test_extract_does_not_expand_hidden_stream_with_a_large_crc_read(self):
        archive = self.work / "hidden.docx"
        with zipfile.ZipFile(self.source) as source, zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as target:
            for name in source.namelist():
                target.writestr(name, source.read(name))
            target.writestr("custom.bin", b"data" + b"x" * 65536)
        data = bytearray(archive.read_bytes())
        with zipfile.ZipFile(archive) as package:
            local = package.getinfo("custom.bin").header_offset
        struct.pack_into("<L", data, local + 14, zlib.crc32(b"data"))
        struct.pack_into("<L", data, local + 22, 4)
        central = data.rfind(b"PK\x01\x02")
        self.assertIn(b"custom.bin", data[central:])
        struct.pack_into("<L", data, central + 16, zlib.crc32(b"data"))
        struct.pack_into("<L", data, central + 24, 4)
        archive.write_bytes(data)
        original = zlib.decompressobj
        expanded = []

        class MeasuredDecompressor:
            def __init__(self, *args, **kwargs):
                self.delegate = original(*args, **kwargs)

            def __getattr__(self, name):
                return getattr(self.delegate, name)

            def decompress(self, *args, **kwargs):
                output = self.delegate.decompress(*args, **kwargs)
                expanded.append(len(output))
                return output

        with mock.patch.object(zlib, "decompressobj", MeasuredDecompressor):
            destination = unpack_document(archive, self.work / "hidden-parts")
        self.assertEqual((destination / "custom.bin").read_bytes(), b"data")
        self.assertLessEqual(max(expanded), 8192)


if __name__ == "__main__":
    unittest.main()
