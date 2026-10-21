"""Replace exact text in supported plain runs, preserving other package parts.

Usage: python scripts/edit_docx.py input.docx replacements.json output.docx
JSON: {"replacements": [{"old": "{{name}}", "new": "Ada", "count": 1}]}
"""

import argparse
import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from office.package import pack_document, unpack_document
from office.xmlio import parse_dom, read_xml
from office.validators.structure import StructureValidator, W

XML_NS = "http://www.w3.org/XML/1998/namespace"
SAFE_ANCESTORS = {"r", "p", "tc", "tr", "tbl", "body", "document", "hdr", "ftr"}
BLOCKED_PARAGRAPH = {"fldChar", "instrText", "commentRangeStart", "commentRangeEnd",
                     "commentReference", "ins", "del", "moveFrom", "moveTo"}


def _text(node):
    return "".join(c.data for c in node.childNodes
                   if c.nodeType in (c.TEXT_NODE, c.CDATA_SECTION_NODE))


def _eligible(node):
    if node.parentNode.namespaceURI != W or node.parentNode.localName != "r":
        return False
    if any(c.nodeType == c.ELEMENT_NODE for c in node.childNodes):
        return False
    parent = node.parentNode
    while parent.nodeType == parent.ELEMENT_NODE:
        if parent.namespaceURI != W or parent.localName not in SAFE_ANCESTORS:
            return False
        if parent.localName == "p":
            if any(parent.getElementsByTagNameNS(W, name) for name in BLOCKED_PARAGRAPH):
                return False
        parent = parent.parentNode
    return True


def replace_text(source, destination, specification, *, parts=None, require_no_placeholders=False):
    """Apply all replacements simultaneously; require explicit occurrence counts."""
    source, destination = Path(source), Path(destination)
    if source.suffix.lower() not in (".docx", ".dotx") or source.suffix.lower() != destination.suffix.lower():
        raise ValueError("Input and output must have the same DOCX/DOTX extension")
    if source.resolve() == destination.resolve() or (
        destination.exists() and source.samefile(destination)
    ):
        raise ValueError("Output must not refer to the input file")
    if not isinstance(specification, dict) or set(specification) != {"replacements"}:
        raise ValueError("Expected an object containing only replacements")
    operations = specification["replacements"]
    if not isinstance(operations, list) or not operations:
        raise ValueError("replacements must be a nonempty list")
    seen = set()
    for operation in operations:
        if not isinstance(operation, dict) or set(operation) - {"old", "new", "count"}:
            raise ValueError("Each replacement supports old, new, count only")
        old, new, count = operation.get("old"), operation.get("new"), operation.get("count", 1)
        if not isinstance(old, str) or not old or old in seen or not isinstance(new, str):
            raise ValueError("Search text must be unique and nonempty; replacement must be text")
        if type(count) is not int or count < 1:
            raise ValueError("count must be a positive integer")
        if any(ord(c) < 32 or 0xD800 <= ord(c) <= 0xDFFF or c in "\ufffe\uffff" for c in new):
            raise ValueError("Replacement must be single-line XML text without control characters")
        seen.add(old)
    parts = list(parts or ["word/document.xml"])
    if len(parts) != len(set(parts)) or any(
        not re.fullmatch(r"word/(document|header[0-9]+|footer[0-9]+)\.xml", part)
        for part in parts
    ):
        raise ValueError("Only document.xml and numbered header/footer parts are supported")
    with tempfile.TemporaryDirectory(prefix="docx-edit-") as tmp:
        unpacked = unpack_document(source, Path(tmp) / "package")
        if not StructureValidator(unpacked, extension=source.suffix.lower()).validate():
            raise ValueError("Input failed basic structure checks")
        counts = [0] * len(operations)
        changes = []
        documents = {}
        for part in parts:
            path = unpacked / part
            if not path.is_file():
                raise ValueError(f"Requested part is missing: {part}")
            dom = parse_dom(read_xml(path))
            documents[part] = dom
            for node in dom.getElementsByTagNameNS(W, "t"):
                if not _eligible(node):
                    continue
                original = _text(node)
                matches = []
                for index, operation in enumerate(operations):
                    for match in re.finditer(re.escape(operation["old"]), original):
                        matches.append((match.start(), match.end(), operation["new"]))
                        counts[index] += 1
                matches.sort()
                if any(left[1] > right[0] for left, right in zip(matches, matches[1:])):
                    raise ValueError("Overlapping replacements are unsupported")
                updated = original
                for start, end, value in reversed(matches):
                    updated = updated[:start] + value + updated[end:]
                if matches:
                    changes.append((part, node, updated))
        for operation, count in zip(operations, counts):
            expected = operation.get("count", 1)
            if count != expected:
                raise ValueError(
                    f"Expected {expected} supported occurrence(s) of {operation['old']!r}, found {count}; "
                    "cross-run, field, revision and complex-container edits are unsupported"
                )
        modified = set()
        for part, node, value in changes:
            for child in list(node.childNodes):
                node.removeChild(child)
            node.appendChild(node.ownerDocument.createTextNode(value))
            node.setAttributeNS(XML_NS, "xml:space", "preserve")
            modified.add(part)
        for part in modified:
            (unpacked / part).write_bytes(documents[part].toxml(encoding="UTF-8"))
        if not StructureValidator(
            unpacked, extension=destination.suffix.lower(),
            require_no_placeholders=require_no_placeholders,
        ).validate():
            raise ValueError("Edited document failed basic structure checks")
        return pack_document(unpacked, destination)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("replacements", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--parts", nargs="+", default=["word/document.xml"])
    parser.add_argument("--require-no-placeholders", action="store_true")
    args = parser.parse_args(argv)
    try:
        with args.replacements.open(encoding="utf-8") as stream:
            specification = json.load(stream)
        result = replace_text(
            args.input, args.output, specification, parts=args.parts,
            require_no_placeholders=args.require_no_placeholders,
        )
    except (OSError, ValueError, RuntimeError, zipfile.BadZipFile) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
