"""Replace exact text in supported plain runs, preserving other package parts.

Usage: python scripts/edit_docx.py input.docx replacements.json output.docx
JSON: {"replacements": [{"old": "{{name}}", "new": "Ada", "count": 1}]}
"""

import argparse
from bisect import bisect_left, bisect_right
import json
import re
import sys
import tempfile
import zipfile
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from office.package import pack_document, unpack_document
from office.helpers import XML_SPACE, reject_signed_package
from office.xmlio import parse_dom, read_xml
from office.validators.structure import StructureValidator, W

XML_NS = "http://www.w3.org/XML/1998/namespace"
XMLNS_NS = "http://www.w3.org/2000/xmlns/"
SAFE_ANCESTORS = {"r", "p", "tc", "tr", "tbl", "body", "document", "hdr", "ftr"}
BLOCKED_PARAGRAPH = {"fldChar", "instrText", "commentRangeStart", "commentRangeEnd",
                     "commentReference", "ins", "del", "moveFrom", "moveTo",
                     "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd"}
PLAIN_RUN_CHILDREN = {"rPr", "t", "tab", "br", "cr", "noBreakHyphen", "softHyphen"}
MOVE_RANGE_STARTS = {"moveFromRangeStart": "from", "moveToRangeStart": "to"}
MOVE_RANGE_ENDS = {"moveFromRangeEnd": "from", "moveToRangeEnd": "to"}


def _text(node):
    return "".join(c.data for c in node.childNodes
                   if c.nodeType in (c.TEXT_NODE, c.CDATA_SECTION_NODE))


def _eligible(node):
    if node.parentNode.namespaceURI != W or node.parentNode.localName != "r":
        return False
    if any(c.nodeType not in (c.TEXT_NODE, c.CDATA_SECTION_NODE) for c in node.childNodes):
        return False
    if any(attr.namespaceURI != XMLNS_NS and not (
        attr.namespaceURI == XML_NS and attr.localName == "space"
    ) for attr in node.attributes.values()):
        return False
    if any(c.nodeType == c.ELEMENT_NODE and (
        c.namespaceURI != W or c.localName not in PLAIN_RUN_CHILDREN
    ) for c in node.parentNode.childNodes):
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


def _diagnostic_text_sequences(paragraph):
    """Return contiguous plain-text nodes, keeping complex containers separate."""
    sequences = []
    current = []

    def flush():
        nonlocal current
        if current:
            sequences.append(current)
            current = []

    def nested_text_nodes(container):
        found = []
        pending = list(reversed(container.childNodes))
        while pending:
            node = pending.pop()
            if node.nodeType != node.ELEMENT_NODE:
                continue
            if node.namespaceURI == W and node.localName == "p":
                continue
            if node.namespaceURI == W and node.localName == "t":
                found.append(node)
            else:
                pending.extend(reversed(node.childNodes))
        return found

    def visit(node):
        if node.nodeType != node.ELEMENT_NODE:
            if node.nodeType in (node.TEXT_NODE, node.CDATA_SECTION_NODE):
                if node.data.strip():
                    flush()
            else:
                flush()
            return
        if node.namespaceURI == W and node.localName in ("p", "r"):
            for child in node.childNodes:
                visit(child)
            return
        if node.namespaceURI == W and node.localName in ("pPr", "rPr"):
            return
        if node.namespaceURI == W and node.localName == "t":
            current.append(node)
            return
        flush()
        protected_nodes = nested_text_nodes(node)
        if protected_nodes:
            sequences.append(protected_nodes)

    for child in paragraph.childNodes:
        visit(child)
    flush()
    return sequences


def _count_mismatch_evidence(documents, blocked_by_part, operations, indexes):
    """Find split or protected occurrences in the already parsed part DOMs."""
    evidence = {index: {"protected": False, "split": False} for index in indexes}
    if not indexes:
        return evidence
    for part, dom in documents.items():
        blocked_nodes = blocked_by_part[part]
        for paragraph in dom.getElementsByTagNameNS(W, "p"):
            for sequence in _diagnostic_text_sequences(paragraph):
                pieces = []
                starts = []
                ends = []
                offset = 0
                for node in sequence:
                    text = _text(node)
                    if not text:
                        continue
                    starts.append(offset)
                    offset += len(text)
                    ends.append(offset)
                    pieces.append((node, text))
                if len(pieces) < 2:
                    continue
                paragraph_text = "".join(text for _, text in pieces)
                for index in indexes:
                    if evidence[index]["protected"] and evidence[index]["split"]:
                        continue
                    old = operations[index]["old"]
                    for match in re.finditer("(?=" + re.escape(old) + ")", paragraph_text):
                        start, end = match.start(), match.start() + len(old)
                        first = bisect_right(ends, start)
                        last = bisect_left(starts, end) - 1
                        if first >= last:
                            continue
                        participants = [node for node, _ in pieces[first:last + 1]]
                        if all(
                            node not in blocked_nodes and _eligible(node)
                            for node in participants
                        ):
                            evidence[index]["split"] = True
                        else:
                            evidence[index]["protected"] = True
                        if evidence[index]["protected"] and evidence[index]["split"]:
                            break
    return evidence


def _count_mismatch_message(operation, expected, found, evidence):
    if found > expected:
        detail = "Too many ordinary eligible occurrences were found."
    elif evidence["protected"] and evidence["split"]:
        detail = (
            "The search text occurs in protected or complex content and is split across "
            "text/run nodes within a paragraph."
        )
    elif evidence["protected"]:
        detail = "The search text occurs in protected or complex content."
    elif evidence["split"]:
        detail = "The search text is split across text/run nodes within a paragraph."
    elif found == 0:
        detail = "The search text is absent from the selected parts."
    else:
        detail = "The ordinary eligible occurrence count is lower than requested."
    return (
        f"Expected {expected} supported occurrence(s) of {operation['old']!r}, found {found}; "
        f"{detail} cross-run, field, revision and complex-container edits are unsupported"
    )


def _range_text_nodes(dom):
    """Exclude field/comment/movement text across paragraph boundaries."""
    blocked = set()
    fields = []
    comments = set()
    moves = set()
    seen_moves = set()
    for node in dom.getElementsByTagName("*"):
        if node.namespaceURI != W:
            continue
        name = node.localName
        if name == "fldChar":
            kind = node.getAttributeNS(W, "fldCharType")
            if kind == "begin":
                fields.append(False)
            elif kind == "separate" and fields and not fields[-1]:
                fields[-1] = True
            elif kind == "end" and fields:
                fields.pop()
            else:
                raise ValueError("Unbalanced or invalid complex field markers; editing is unsupported")
        elif name == "commentRangeStart":
            comments.add(node.getAttributeNS(W, "id"))
        elif name == "commentRangeEnd":
            comments.discard(node.getAttributeNS(W, "id"))
        elif name in MOVE_RANGE_STARTS or name in MOVE_RANGE_ENDS:
            value = node.getAttributeNS(W, "id")
            key = (MOVE_RANGE_STARTS.get(name, MOVE_RANGE_ENDS.get(name)), value)
            if not value.isascii() or not value.isdigit():
                raise ValueError("Invalid movement range ID; editing is unsupported")
            if name in MOVE_RANGE_STARTS:
                if key in seen_moves:
                    raise ValueError("Duplicate movement range; editing is unsupported")
                seen_moves.add(key)
                moves.add(key)
            elif key not in moves:
                raise ValueError("Unbalanced movement range; editing is unsupported")
            else:
                moves.remove(key)
        elif name == "t" and (fields or comments or moves):
            blocked.add(node)
    if fields:
        raise ValueError("Unclosed complex field; editing is unsupported")
    if moves:
        raise ValueError("Unclosed movement range; editing is unsupported")
    return blocked


def _preserves_space(node):
    """Resolve the nearest explicit xml:space, including ancestor defaults."""
    while node.nodeType == node.ELEMENT_NODE:
        if node.hasAttributeNS(XML_NS, "space"):
            value = node.getAttributeNS(XML_NS, "space")
            if value not in ("default", "preserve"):
                raise ValueError("Invalid xml:space; editing is unsupported")
            return value == "preserve"
        node = node.parentNode
    return False


def _same_xml(left, right):
    """Compare every parsed node and attribute, including opaque extensions."""
    pending = [(left, right)]
    while pending:
        left_node, right_node = pending.pop()
        if (left_node.nodeType, left_node.nodeName, left_node.namespaceURI, left_node.nodeValue) != (
            right_node.nodeType, right_node.nodeName, right_node.namespaceURI, right_node.nodeValue
        ):
            return False
        left_attrs = _attribute_values(left_node)
        right_attrs = _attribute_values(right_node)
        if left_attrs != right_attrs or len(left_node.childNodes) != len(right_node.childNodes):
            return False
        pending.extend(zip(left_node.childNodes, right_node.childNodes))
    return (left.version or "1.0") == (right.version or "1.0") and left.standalone == right.standalone


def _attribute_values(node):
    if node.attributes is None:
        return []
    return sorted((attr.name, attr.namespaceURI, attr.value) for attr in node.attributes.values())


def _serialize_checked(dom):
    data = dom.toxml(encoding="UTF-8", standalone=dom.standalone)
    reopened = parse_dom(data)
    if not _same_xml(dom, reopened):
        raise ValueError("XML serialization cannot preserve all node/attribute values; editing is unsupported")
    return data


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
        validator = StructureValidator(unpacked, extension=source.suffix.lower())
        if not validator.validate():
            raise ValueError("Input failed basic structure checks")
        reject_signed_package(validator.files, (kind for relationships in validator.relations.values()
                                               for kind, _ in relationships.values()))
        counts = [0] * len(operations)
        protected_matches = [False] * len(operations)
        changes = []
        documents = {}
        blocked_by_part = {}
        for part in parts:
            path = unpacked / part
            if not path.is_file():
                raise ValueError(f"Requested part is missing: {part}")
            dom = parse_dom(read_xml(path))
            documents[part] = dom
            blocked_nodes = _range_text_nodes(dom)
            blocked_by_part[part] = blocked_nodes
            for node in dom.getElementsByTagNameNS(W, "t"):
                eligible = node not in blocked_nodes and _eligible(node)
                if not eligible:
                    original = _text(node)
                    for index, operation in enumerate(operations):
                        if operation["old"] in original:
                            protected_matches[index] = True
                    continue
                original = _text(node)
                matches = []
                for index, operation in enumerate(operations):
                    # Look ahead so overlapping occurrences of the same rule
                    # are detected as well as overlaps between different rules.
                    old = operation["old"]
                    for match in re.finditer("(?=" + re.escape(old) + ")", original):
                        matches.append((match.start(), match.start() + len(old), operation["new"]))
                        counts[index] += 1
                matches.sort()
                if any(left[1] > right[0] for left, right in zip(matches, matches[1:])):
                    raise ValueError("Overlapping replacements are unsupported")
                updated = original
                for start, end, value in reversed(matches):
                    updated = updated[:start] + value + updated[end:]
                if matches:
                    if not _preserves_space(node) and original != original.strip(XML_SPACE):
                        raise ValueError("Unpreserved edge XML whitespace is ambiguous; editing is unsupported")
                    changes.append((part, node, updated))
        mismatch = next((
            index for index, (operation, count) in enumerate(zip(operations, counts))
            if count != operation.get("count", 1)
        ), None)
        if mismatch is not None:
            index = mismatch
            operation, count = operations[index], counts[index]
            expected = operation.get("count", 1)
            evidence = {"protected": False, "split": False}
            if count < expected:
                evidence = _count_mismatch_evidence(
                    documents, blocked_by_part, operations, [index],
                )[index]
            evidence["protected"] = evidence["protected"] or protected_matches[index]
            raise ValueError(_count_mismatch_message(operation, expected, count, evidence))
        modified = set()
        for part, node, value in changes:
            for child in list(node.childNodes):
                node.removeChild(child)
            if value:
                node.appendChild(node.ownerDocument.createTextNode(value))
            node.setAttributeNS(XML_NS, "xml:space", "preserve")
            modified.add(part)
        for part in modified:
            (unpacked / part).write_bytes(_serialize_checked(documents[part]))
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
