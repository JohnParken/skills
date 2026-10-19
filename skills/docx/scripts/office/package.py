"""Unpack and repack DOCX/DOTX files using Python's ZIP support.

Usage:
    python scripts/office/package.py unpack input.docx unpacked/
    python scripts/office/package.py pack unpacked/ output.docx

This checks ZIP integrity and required parts, not XML schemas or page layout.
"""

import argparse
import sys
import tempfile
import zipfile
from pathlib import Path

try:
    from .helpers import rezip, safe_extract
except ImportError:
    from helpers import rezip, safe_extract


REQUIRED_PARTS = ("[Content_Types].xml", "_rels/.rels", "word/document.xml")


def _check_extension(path: Path) -> None:
    if path.suffix.lower() not in {".docx", ".dotx"}:
        raise ValueError("Expected a .docx or .dotx file; format conversion is unsupported")


def _check_parts(directory: Path) -> None:
    missing = [name for name in REQUIRED_PARTS if not (directory / name).is_file()]
    if missing:
        raise ValueError("Missing required package parts: " + ", ".join(missing))


def unpack_document(source: str | Path, destination: str | Path) -> Path:
    """Extract into a new or empty directory, preserving other destinations."""
    source, destination = Path(source), Path(destination)
    _check_extension(source)
    if destination.is_symlink() or (
        destination.exists()
        and (not destination.is_dir() or any(destination.iterdir()))
    ):
        raise ValueError("Unpack destination must be a new or empty directory")
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="docx-unpack-", dir=destination.parent) as tmp:
        staging = Path(tmp) / "package"
        staging.mkdir()
        with zipfile.ZipFile(source) as archive:
            names = archive.namelist()
            if len(names) != len(set(names)):
                raise ValueError("Duplicate ZIP members are unsupported")
            bad_member = archive.testzip()
            if bad_member:
                raise ValueError(f"Corrupt ZIP member: {bad_member}")
            safe_extract(archive, staging)
        _check_parts(staging)
        if destination.exists():
            destination.rmdir()
        staging.replace(destination)
    return destination


def pack_document(source: str | Path, destination: str | Path) -> Path:
    """Repack XML and media without launching an archiver or following links."""
    source, destination = Path(source), Path(destination)
    _check_extension(destination)
    if not source.is_dir() or source.is_symlink():
        raise ValueError("Pack source must be a directory, not a symbolic link")
    if destination.resolve().is_relative_to(source.resolve()):
        raise ValueError("Pack destination must be outside the source directory")
    for path in source.rglob("*"):
        if path.is_symlink():
            raise ValueError(f"Symbolic links are unsupported: {path}")
    _check_parts(source)
    destination.parent.mkdir(parents=True, exist_ok=True)
    rezip(source, destination)
    return destination


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("operation", choices=("unpack", "pack"))
    parser.add_argument("source", type=Path)
    parser.add_argument("destination", type=Path)
    args = parser.parse_args(argv)
    try:
        operation = unpack_document if args.operation == "unpack" else pack_document
        output = operation(args.source, args.destination)
    except (OSError, ValueError, RuntimeError, zipfile.BadZipFile) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    print(output)
    return 0


if __name__ == "__main__":
    sys.exit(main())
