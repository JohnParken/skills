"""Replace one PNG/JPEG media part, with explicit hash and reference guards.

JSON requires media_part, expected_sha256, expected_relationships and
expected_references. Shared parts require all_references=true. Only internal
image relationships used by DrawingML picture blips are supported. No XML,
dimensions, crop or layout is changed. Magic bytes are checked, not decoded.
"""

import argparse
import hashlib
import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from office.helpers import MAX_PART_BYTES
from office.package import pack_document, unpack_document
from office.validators.structure import R, W, StructureValidator

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
IMAGE_RELATION = R + "/image"
IMAGE_TYPES = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg"}


def _image_kind(path, data):
    extension = Path(path).suffix.lower().lstrip(".")
    content_type = IMAGE_TYPES.get(extension)
    valid = (content_type == "image/png" and data.startswith(b"\x89PNG\r\n\x1a\n")
             and data.endswith(b"\x00\x00\x00\x00IEND\xaeB\x60\x82")) or (
        content_type == "image/jpeg" and data.startswith(b"\xff\xd8\xff") and data.endswith(b"\xff\xd9")
    )
    if not valid:
        raise ValueError("PNG/JPEG extension, magic and terminator must agree; conversion is unsupported")
    return content_type


def _specification(value):
    required = {"media_part", "expected_sha256", "expected_relationships", "expected_references"}
    if not isinstance(value, dict) or not required <= set(value) or set(value) - required - {"all_references"}:
        raise ValueError("Expected media part, SHA-256 and both exact relationship/reference counts")
    part = value["media_part"]
    if not isinstance(part, str) or not re.fullmatch(
        r"word/media/[A-Za-z0-9][A-Za-z0-9._-]*\.(?:png|jpe?g)", part, re.IGNORECASE
    ):
        raise ValueError("media_part must be a safe PNG/JPEG member directly under word/media")
    digest = value["expected_sha256"]
    if not isinstance(digest, str) or not re.fullmatch(r"[0-9A-Fa-f]{64}", digest):
        raise ValueError("expected_sha256 must be exactly 64 hexadecimal digits")
    for key in ("expected_relationships", "expected_references"):
        if type(value[key]) is not int or value[key] < 1:
            raise ValueError(f"{key} must be a positive integer")
    if type(value.get("all_references", False)) is not bool:
        raise ValueError("all_references must be a boolean")
    return part


def _picture_reference(node, parents):
    if node.tag != f"{{{A}}}blip" or f"{{{R}}}link" in node.attrib:
        return False
    current = node
    for expected in (f"{{{PIC}}}blipFill", f"{{{PIC}}}pic", f"{{{A}}}graphicData", f"{{{A}}}graphic"):
        current = parents.get(current)
        if current is None or current.tag != expected:
            return False
        if expected == f"{{{A}}}graphicData" and current.get("uri") != PIC:
            return False
    current = parents.get(current)
    if current is None or current.tag not in {f"{{{WP}}}inline", f"{{{WP}}}anchor"}:
        return False
    current = parents.get(current)
    return current is not None and current.tag == f"{{{W}}}drawing"


def _references(validator, part):
    selected = {}
    for owner, relations in validator.relations.items():
        for rid, (kind, target) in relations.items():
            if target != part:
                continue
            if kind != IMAGE_RELATION:
                raise ValueError("Non-image relationship points to the selected media; replacement is unsupported")
            selected.setdefault(owner, set()).add(rid)
    if not selected:
        raise ValueError("The selected media has no internal image relationships")
    count = 0
    used = set()
    for owner, identifiers in selected.items():
        root = validator.roots.get(owner)
        if root is None or not owner.startswith("word/") or owner.endswith(".rels"):
            raise ValueError("An image relationship has an unsupported/opaque owner")
        parents = {child: parent for parent in root.iter() for child in parent}
        for node in root.iter():
            for attribute, value in node.attrib.items():
                if value not in identifiers:
                    continue
                if attribute != f"{{{R}}}embed" or not _picture_reference(node, parents):
                    raise ValueError("Unknown or non-picture reference uses an image relationship ID")
                used.add((owner, value))
                count += 1
    relationships = {(owner, rid) for owner, identifiers in selected.items() for rid in identifiers}
    if used != relationships:
        raise ValueError("Orphan image relationships are unsupported; every selected relationship must be used")
    return len(relationships), count


def replace_image(source, replacement, destination, specification):
    """Replace one same-format media member; all other member bytes are retained."""
    source, replacement, destination = Path(source), Path(replacement), Path(destination)
    if source.suffix.lower() not in {".docx", ".dotx"} or source.suffix.lower() != destination.suffix.lower():
        raise ValueError("Input and output must have the same DOCX/DOTX extension")
    if source.resolve() == destination.resolve() or (destination.exists() and source.samefile(destination)):
        raise ValueError("Output must not refer to the input file")
    part = _specification(specification)
    with replacement.open("rb") as stream:
        replacement_bytes = stream.read(MAX_PART_BYTES + 1)
    if len(replacement_bytes) > MAX_PART_BYTES:
        raise ValueError("Replacement image exceeds the 32 MiB member limit")
    replacement_type = _image_kind(replacement, replacement_bytes)
    with tempfile.TemporaryDirectory(prefix="docx-image-replace-") as temporary:
        unpacked = unpack_document(source, Path(temporary) / "package")
        validator = StructureValidator(unpacked, extension=source.suffix.lower())
        if not validator.validate():
            raise ValueError("Input failed implemented structure checks")
        if any(name.startswith("_xmlsignatures/") for name in validator.files) or any(
            "/digital-signature/" in kind
            for relations in validator.relations.values() for kind, _ in relations.values()
        ):
            raise ValueError("Signed packages are unsupported for media replacement")
        if part not in validator.files:
            raise ValueError("The requested media part does not exist")
        path = unpacked / part
        original = path.read_bytes()
        if hashlib.sha256(original).hexdigest() != specification["expected_sha256"].lower():
            raise ValueError("expected_sha256 does not match the original media bytes")
        content_type = _image_kind(part, original)
        if replacement_type != content_type or validator.content_types.get(part) != content_type:
            raise ValueError("Original, replacement and package content types must use the same PNG/JPEG format")
        relationships, references = _references(validator, part)
        if relationships != specification["expected_relationships"] or references != specification["expected_references"]:
            raise ValueError(f"Image usage mismatch: found {relationships} relationship(s), {references} picture reference(s)")
        if (relationships > 1 or references > 1) and not specification.get("all_references", False):
            raise ValueError("Shared media replacement requires explicit all_references=true")
        path.write_bytes(replacement_bytes)
        if not StructureValidator(unpacked, extension=destination.suffix.lower()).validate():
            raise ValueError("Edited document failed implemented structure checks")
        return pack_document(unpacked, destination)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("replacement", type=Path)
    parser.add_argument("selection", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args(argv)
    try:
        result = replace_image(args.input, args.replacement, args.output,
                               json.loads(args.selection.read_text(encoding="utf-8")))
    except (OSError, ValueError, RuntimeError, zipfile.BadZipFile) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
