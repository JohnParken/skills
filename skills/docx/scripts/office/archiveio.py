"""Bound ZIP metadata before zipfile builds its central-directory object list.

Support regular, single-disk, non-ZIP64 ZIP archives with stored/deflated
members. Reject concatenated prefixes, trailing bytes and ambiguous metadata.
Member paths, types and expanded sizes still pass the existing archive guard.
This is a resource preflight, not a complete ZIP parser or a concurrency lock.
"""

from contextlib import contextmanager
import os
import stat
import struct
import zipfile
import zlib

try:
    from .helpers import MAX_MEMBERS, MAX_PACKAGE_BYTES, check_archive
except ImportError:
    from helpers import MAX_MEMBERS, MAX_PACKAGE_BYTES, check_archive

MAX_CENTRAL_BYTES = 4 * 1024 * 1024
MAX_EOCD_BYTES = 22 + 65535
MAX_ARCHIVE_BYTES = MAX_PACKAGE_BYTES + MAX_CENTRAL_BYTES + MAX_EOCD_BYTES
EOCD = struct.Struct("<4s4H2LH")
CENTRAL = struct.Struct("<4s6H3L5H2L")


def _read_exact(stream, length):
    data = stream.read(length)
    if len(data) != length:
        raise ValueError("Truncated ZIP metadata")
    return data


def _preflight(stream, file_size):
    if file_size < EOCD.size or file_size > MAX_ARCHIVE_BYTES:
        raise ValueError("ZIP compressed file size exceeds limits or is truncated")
    tail_size = min(file_size, MAX_EOCD_BYTES)
    stream.seek(file_size - tail_size)
    tail = _read_exact(stream, tail_size)
    position = tail.rfind(b"PK\x05\x06")
    if position < 0 or position + EOCD.size > len(tail):
        raise ValueError("Missing or ambiguous ZIP end record")
    values = EOCD.unpack_from(tail, position)
    _, disk, central_disk, disk_count, count, size, offset, comment_length = values
    end_offset = file_size - tail_size + position
    if position + EOCD.size + comment_length != tail_size:
        raise ValueError("ZIP trailing bytes or inconsistent comment length are unsupported")
    if disk or central_disk or disk_count != count:
        raise ValueError("Multi-disk or inconsistent ZIP entry counts are unsupported")
    if count == 0xFFFF or size == 0xFFFFFFFF or offset == 0xFFFFFFFF:
        raise ValueError("ZIP64 archives are unsupported")
    # zipfile consults this locator even when the legacy EOCD has no sentinel.
    # Do not let embedded metadata replace the bounded directory we inspected.
    if end_offset >= 20:
        stream.seek(end_offset - 20)
        if _read_exact(stream, 4) == b"PK\x06\x07":
            raise ValueError("ZIP64 locator records are unsupported")
    if count > MAX_MEMBERS:
        raise ValueError("ZIP central directory exceeds the member limit")
    if size > MAX_CENTRAL_BYTES:
        raise ValueError("ZIP central directory exceeds the byte limit")
    if offset + size != end_offset:
        raise ValueError("ZIP central directory offset/size mismatch or concatenated prefix")
    stream.seek(offset)
    central = _read_exact(stream, size)
    cursor = actual_count = 0
    first_local = None
    while cursor < size:
        if size - cursor < CENTRAL.size:
            raise ValueError("Truncated ZIP central directory record")
        fields = CENTRAL.unpack_from(central, cursor)
        if fields[0] != b"PK\x01\x02":
            raise ValueError("Invalid ZIP central directory record signature")
        actual_count += 1
        if actual_count > MAX_MEMBERS or actual_count > count:
            raise ValueError("ZIP central directory member count mismatch or limit exceeded")
        needed, compressed, expanded = fields[2], fields[8], fields[9]
        name_length, extra_length, comment_length, member_disk = fields[10:14]
        local_offset = fields[16]
        if needed >= 45 or compressed == 0xFFFFFFFF or expanded == 0xFFFFFFFF or local_offset == 0xFFFFFFFF:
            raise ValueError("ZIP64 or unsupported ZIP extraction versions are unsupported")
        if member_disk:
            raise ValueError("Multi-disk ZIP members are unsupported")
        record_end = cursor + CENTRAL.size + name_length + extra_length + comment_length
        if record_end > size or local_offset + 30 > offset:
            raise ValueError("ZIP central record lengths or local-header offsets are inconsistent")
        extra_cursor = cursor + CENTRAL.size + name_length
        extra_end = extra_cursor + extra_length
        while extra_cursor < extra_end:
            if extra_end - extra_cursor < 4:
                raise ValueError("Truncated ZIP extra-field header")
            identifier, length = struct.unpack_from("<HH", central, extra_cursor)
            extra_cursor += 4
            if identifier == 1:
                raise ValueError("ZIP64 extra fields are unsupported")
            if length > extra_end - extra_cursor:
                raise ValueError("Truncated ZIP extra-field payload")
            extra_cursor += length
        first_local = local_offset if first_local is None else min(first_local, local_offset)
        cursor = record_end
    if actual_count != count:
        raise ValueError("ZIP central directory entry count does not match the end record")
    if not count:
        if offset:
            raise ValueError("Empty ZIP archives cannot have prefixes")
    else:
        if first_local != 0:
            raise ValueError("ZIP prefixes and concatenated archives are unsupported")
        stream.seek(0)
        if _read_exact(stream, 4) != b"PK\x03\x04":
            raise ValueError("Invalid first ZIP local-header signature")
    stream.seek(0)


@contextmanager
def open_archive(path):
    """Yield a guarded read-only ZipFile after bounded metadata preflight.

    Keep the same file handle for the preflight and ZIP reads. Inputs must remain
    stable while processed; this interface does not isolate hostile concurrent
    writers. zlib failures during caller reads are reported as ValueError.
    """
    flags = os.O_RDONLY | getattr(os, "O_NONBLOCK", 0) | getattr(os, "O_BINARY", 0)
    descriptor = os.open(path, flags)
    try:
        information = os.fstat(descriptor)
        if not stat.S_ISREG(information.st_mode):
            raise ValueError("ZIP input must be a regular file")
        stream = os.fdopen(descriptor, "rb")
        descriptor = None
        with stream:
            _preflight(stream, information.st_size)
            if os.fstat(stream.fileno()).st_size != information.st_size:
                raise ValueError("ZIP input changed during metadata preflight")
            try:
                with zipfile.ZipFile(stream, "r") as archive:
                    check_archive(archive)
                    yield archive
            except zlib.error as exc:
                raise ValueError("Invalid compressed ZIP data: decompression failed") from exc
    finally:
        if descriptor is not None:
            os.close(descriptor)
