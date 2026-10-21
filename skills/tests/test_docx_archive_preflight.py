"""Reject oversized or ambiguous ZIP metadata before constructing ZipInfo objects."""

import io
import os
import struct
import sys
import tempfile
import unittest
import zipfile
import zlib
from contextlib import nullcontext
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "docx/scripts"))
from office import archiveio


def archive_bytes(count=1, *, compression=zipfile.ZIP_DEFLATED, comment=b"", extra=b""):
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", compression=compression) as package:
        package.comment = comment
        for index in range(count):
            info = zipfile.ZipInfo(f"word/中文{index}.xml")
            info.compress_type = compression
            info.extra = extra
            package.writestr(info, b"<x>content</x>")
    return stream.getvalue()


def eocd_fields(data):
    position = data.rfind(b"PK\x05\x06")
    return position, list(struct.unpack_from("<4s4H2LH", data, position))


def with_eocd(data, **changes):
    position, values = eocd_fields(data)
    fields = {"disk": 1, "central_disk": 2, "disk_count": 3, "count": 4,
              "size": 5, "offset": 6, "comment": 7}
    for key, value in changes.items():
        values[fields[key]] = value
    return data[:position] + struct.pack("<4s4H2LH", *values) + data[position + 22:]


class ArchivePreflightTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-archive-preflight-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root / "input.docx"

    def rejected_before_zipfile(self, data, *, limits=None):
        self.source.write_bytes(data)
        with mock.patch.multiple(archiveio, **limits) if limits else nullcontext():
            with mock.patch.object(archiveio.zipfile, "ZipFile", side_effect=AssertionError("constructor reached")) as constructor:
                with self.assertRaises(ValueError):
                    with archiveio.open_archive(self.source):
                        self.fail("Archive was unexpectedly accepted")
                constructor.assert_not_called()

    def test_stored_deflated_unicode_and_ordinary_comments_work(self):
        for compression in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
            for comment in (b"", "中文 ZIP 注释".encode(), b"x" * 65535):
                with self.subTest(compression=compression, comment_size=len(comment)):
                    self.source.write_bytes(archive_bytes(3, compression=compression, comment=comment))
                    with archiveio.open_archive(self.source) as package:
                        self.assertEqual(len(package.infolist()), 3)
                        self.assertEqual(package.comment, comment)
                        self.assertIsNone(package.testzip())
                        self.assertEqual(package.read("word/中文1.xml"), b"<x>content</x>")

    def test_declared_member_limit_precedes_constructor(self):
        self.rejected_before_zipfile(archive_bytes(3), limits={"MAX_MEMBERS": 2})

    def test_unknown_well_formed_extra_fields_are_preserved(self):
        extra = struct.pack("<HH", 0xCAFE, 3) + b"abc"
        self.source.write_bytes(archive_bytes(extra=extra))
        with archiveio.open_archive(self.source) as package:
            self.assertEqual(package.infolist()[0].extra, extra)
            self.assertIsNone(package.testzip())

    def test_manual_record_count_rejects_lying_eocd(self):
        data = with_eocd(archive_bytes(3), disk_count=1, count=1)
        self.rejected_before_zipfile(data, limits={"MAX_MEMBERS": 2})

    def test_central_byte_and_compressed_file_limits_precede_constructor(self):
        data = archive_bytes(2)
        self.rejected_before_zipfile(data, limits={"MAX_CENTRAL_BYTES": 32})
        self.rejected_before_zipfile(data, limits={"MAX_ARCHIVE_BYTES": 64})

    def test_multidisk_and_zip64_fields_are_rejected_early(self):
        data = archive_bytes()
        for changes in ({"disk": 1}, {"central_disk": 1}, {"disk_count": 0},
                        {"disk_count": 65535, "count": 65535}, {"size": 0xFFFFFFFF},
                        {"offset": 0xFFFFFFFF}):
            with self.subTest(changes=changes):
                self.rejected_before_zipfile(with_eocd(data, **changes))
        self.rejected_before_zipfile(archive_bytes(extra=struct.pack("<HH", 1, 0)))

    def test_prefixes_trailing_junk_and_concatenated_archives_are_rejected(self):
        data = archive_bytes()
        for invalid in (b"MZprefix" + data, data + b"tail", data + data):
            with self.subTest(size=len(invalid)):
                self.rejected_before_zipfile(invalid)
        position, values = eocd_fields(data)
        prefix = b"MZprefix"
        adjusted = bytearray(with_eocd(data, offset=values[6] + len(prefix)))
        struct.pack_into("<L", adjusted, values[6] + 42, len(prefix))
        self.rejected_before_zipfile(prefix + adjusted)

    def test_directory_offsets_lengths_and_record_signature_are_checked(self):
        data = archive_bytes()
        _, values = eocd_fields(data)
        for changes in ({"offset": values[6] + 1}, {"size": values[5] - 1}, {"comment": 1}):
            self.rejected_before_zipfile(with_eocd(data, **changes))
        corrupt = bytearray(data)
        corrupt[values[6]:values[6] + 4] = b"BAD!"
        self.rejected_before_zipfile(corrupt)
        corrupt = bytearray(data)
        struct.pack_into("<H", corrupt, values[6] + 28, 65535)
        self.rejected_before_zipfile(corrupt)

    def test_central_zip64_and_multidisk_members_are_checked(self):
        data = archive_bytes()
        _, values = eocd_fields(data)
        for relative, fmt, value in ((20, "<L", 0xFFFFFFFF), (24, "<L", 0xFFFFFFFF),
                                     (42, "<L", 0xFFFFFFFF), (34, "<H", 1)):
            corrupt = bytearray(data)
            struct.pack_into(fmt, corrupt, values[6] + relative, value)
            self.rejected_before_zipfile(corrupt)

    def test_zip64_locator_is_rejected_without_legacy_sentinels(self):
        # A locator embedded in the last member comment still makes zipfile
        # consult ZIP64 metadata. It must not bypass the inspected EOCD bounds.
        stream = io.BytesIO()
        with zipfile.ZipFile(stream, "w") as package:
            info = zipfile.ZipInfo("word/document.xml")
            info.comment = struct.pack("<4sLQL", b"PK\x06\x07", 0, 0, 1)
            package.writestr(info, b"<x/>")
        data = stream.getvalue()
        self.rejected_before_zipfile(data)
        self.rejected_before_zipfile(with_eocd(data, comment=65535) + b"x" * 65535)

    def test_empty_zip_is_compatible_but_regular_files_are_required(self):
        self.source.write_bytes(archive_bytes(0))
        with archiveio.open_archive(self.source) as package:
            self.assertEqual(package.namelist(), [])
        with self.assertRaises(ValueError):
            with archiveio.open_archive(self.root):
                self.fail("directory accepted")
        if hasattr(os, "mkfifo"):
            fifo = self.root / "input.fifo"
            os.mkfifo(fifo)
            with self.assertRaises(ValueError):
                with archiveio.open_archive(fifo):
                    self.fail("FIFO accepted")

    def test_deflate_corruption_is_reported_as_value_error(self):
        data = bytearray(archive_bytes())
        name_length, extra_length = struct.unpack_from("<HH", data, 26)
        data[30 + name_length + extra_length] = 0x07  # Reserved DEFLATE block type.
        self.source.write_bytes(data)
        with self.assertRaisesRegex(ValueError, "decompress|DEFLATE|compressed"):
            with archiveio.open_archive(self.source) as package:
                package.read(package.namelist()[0])

    def test_context_wraps_zlib_errors_and_closes_handle(self):
        self.source.write_bytes(archive_bytes())
        with self.assertRaisesRegex(ValueError, "decompress|DEFLATE|compressed"):
            with archiveio.open_archive(self.source) as package:
                raise zlib.error("fixture compressed stream failure")
        self.assertIsNone(package.fp)


if __name__ == "__main__":
    unittest.main()
