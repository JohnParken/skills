"""Manage a strict subset of range comments transactionally using only stdlib.

One JSON action per call: list, add, update, delete or reply. Mutations require
a separate DOCX/DOTX output. This is independent of the low-level comment.py.
"""

import argparse
import contextlib
import io
import json
import re
import secrets
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from edit_docx import XML_NS, XMLNS_NS, _preserves_space, _serialize_checked, _text
from edit_structure import _blocked_anchors, _elements, _plain_paragraph, _prepare_run, _single_line
from office.helpers import XML_SPACE, opc_target
from office.package import pack_document, unpack_document
from office.validators.structure import StructureValidator, W
from office.xmlio import parse_dom, read_xml

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
MAX_ID = 2 ** 31 - 1
COMMENT_PARTS = {
    "word/comments.xml": (W, "w", "comments", "comment",
                          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                          "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"),
    "word/commentsExtended.xml": (W15, "w15", "commentsEx", "commentEx",
                                  "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
                                  "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"),
    "word/commentsIds.xml": (W16CID, "w16cid", "commentsIds", "commentId",
                             "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
                             "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"),
    "word/commentsExtensible.xml": (W16CEX, "w16cex", "commentsExtensible", "commentExtensible",
                                    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
                                    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"),
}
META_ATTRS = {
    "word/commentsExtended.xml": {"paraId", "paraIdParent", "done"},
    "word/commentsIds.xml": {"paraId", "durableId"},
    "word/commentsExtensible.xml": {"durableId", "dateUtc"},
}
ON_OFF_PROPERTIES = {
    "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
    "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
    "rtl", "cs", "specVanish", "oMath",
}
BORDER_ATTRIBUTES = {"val", "sz", "space", "color", "themeColor", "themeTint", "themeShade", "shadow", "frame"}
SHADING_ATTRIBUTES = {"val", "color", "fill", "themeColor", "themeTint", "themeShade", "themeFill", "themeFillTint", "themeFillShade"}
RUN_PROPERTIES = {name: {"val"} for name in ON_OFF_PROPERTIES | {
    "rStyle", "sz", "szCs", "highlight", "vertAlign", "spacing", "w", "kern", "position", "em", "effect",
}}
RUN_PROPERTIES.update({
    "rFonts": {"ascii", "hAnsi", "eastAsia", "cs", "hint", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"},
    "color": {"val", "themeColor", "themeTint", "themeShade"},
    "u": {"val", "color", "themeColor", "themeTint", "themeShade"},
    "lang": {"val", "eastAsia", "bidi"}, "bdr": BORDER_ATTRIBUTES, "shd": SHADING_ATTRIBUTES,
    "fitText": {"val", "id"}, "eastAsianLayout": {"id", "combine", "combineBrackets", "vert", "vertCompress"},
})
PARAGRAPH_PROPERTIES = {name: {"val"} for name in {
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "widowControl", "contextualSpacing",
    "suppressLineNumbers", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
    "jc", "textAlignment", "outlineLvl",
}}
PARAGRAPH_PROPERTIES.update({
    "spacing": {"before", "after", "line", "lineRule", "beforeAutospacing", "afterAutospacing", "beforeLines", "afterLines"},
    "ind": {"left", "right", "start", "end", "hanging", "firstLine", "leftChars", "rightChars", "startChars", "endChars", "hangingChars", "firstLineChars"},
    "shd": SHADING_ATTRIBUTES,
})


def _id(value, label="id", *, xml=False):
    if xml:
        if not isinstance(value, str) or not re.fullmatch(r"[0-9]{1,10}", value):
            raise ValueError(f"Invalid {label}; unsupported comment metadata")
        value = int(value)
    if type(value) is not int or not 0 <= value <= MAX_ID:
        raise ValueError(f"{label} must be an integer from 0 to {MAX_ID}")
    return value


def _hex(value, label):
    if not isinstance(value, str) or not re.fullmatch(r"[0-9A-Fa-f]{8}", value):
        raise ValueError(f"Invalid {label}; unsupported comment metadata")
    return value.upper()


def _children(node, namespace, name):
    return [child for child in node.childNodes if child.nodeType == child.ELEMENT_NODE
            and child.namespaceURI == namespace and child.localName == name]


def _known_attrs(node, allowed, *, root=False):
    for attr in node.attributes.values():
        if attr.namespaceURI == XMLNS_NS:
            continue
        if root and (attr.namespaceURI, attr.localName) == (MC, "Ignorable"):
            continue
        if (attr.namespaceURI, attr.localName) not in allowed:
            raise ValueError(f"Unknown attribute on {node.nodeName}; comment mutation is unsupported")


def _only_children(node, allowed):
    for child in node.childNodes:
        if child.nodeType == child.ELEMENT_NODE:
            if (child.namespaceURI, child.localName) not in allowed:
                raise ValueError(f"Unknown content in {node.nodeName}; comment mutation is unsupported")
        elif child.nodeType != child.TEXT_NODE or child.data.strip(XML_SPACE):
            raise ValueError(f"Complex XML in {node.nodeName}; comment mutation is unsupported")


def _new(dom, namespace, prefix, name):
    node = dom.createElementNS(namespace, prefix + ":" + name)
    node.setAttributeNS(XMLNS_NS, "xmlns:" + prefix, namespace)
    return node


def _attr(node, namespace, prefix, name, value):
    node.setAttributeNS(XMLNS_NS, "xmlns:" + prefix, namespace)
    node.setAttributeNS(namespace, prefix + ":" + name, str(value))


def _format_leaf(node, attributes):
    _known_attrs(node, {(W, name) for name in attributes})
    _only_children(node, set())


def _format_tree(node, model, *, extra=(), repeated=False):
    """Validate every descendant that a whole-comment deletion would remove.

    This is a deliberately explicit formatting subset, not general OOXML/XSD
    validation. Unknown attributes on otherwise known formatting nodes matter
    just as much as unknown child elements.
    """
    _known_attrs(node, set())
    _only_children(node, {(W, name) for name in set(model) | set(extra)})
    names = [child.localName for child in _elements(node)]
    if not repeated and len(names) != len(set(names)):
        raise ValueError("Duplicate comment formatting properties are unsupported")
    for child in _elements(node):
        name = child.localName
        if name in model:
            _format_leaf(child, model[name])
        elif name == "rPr":
            _format_tree(child, RUN_PROPERTIES)
        elif name == "numPr":
            _format_tree(child, {"ilvl": {"val"}, "numId": {"val"}})
        elif name == "tabs":
            _format_tree(child, {"tab": {"val", "leader", "pos"}}, repeated=True)
        elif name == "pBdr":
            _format_tree(child, {key: BORDER_ATTRIBUTES for key in ("top", "left", "bottom", "right", "between", "bar")})


def _run_properties(run):
    _prepare_run(run)
    for properties in _children(run, W, "rPr"):
        _format_tree(properties, RUN_PROPERTIES)


def _comment_content(comment):
    _known_attrs(comment, {(W, key) for key in ("id", "author", "date", "initials")})
    if not comment.getAttributeNS(W, "author"):
        raise ValueError("Comment author is missing; mutation is unsupported")
    _only_children(comment, {(W, "p")})
    paragraphs = _children(comment, W, "p")
    if len(paragraphs) != 1:
        raise ValueError("Only single-paragraph comments can be mutated")
    paragraph = paragraphs[0]
    _known_attrs(paragraph, {(W14, "paraId"), (W14, "textId"), (XML_NS, "space")})
    _only_children(paragraph, {(W, "pPr"), (W, "r")})
    if len(_children(paragraph, W, "pPr")) > 1:
        raise ValueError("Duplicate comment paragraph properties")
    for properties in _children(paragraph, W, "pPr"):
        _format_tree(properties, PARAGRAPH_PROPERTIES, extra={"rPr", "numPr", "tabs", "pBdr"})
    texts = []
    for run in _children(paragraph, W, "r"):
        _known_attrs(run, {(W, key) for key in ("rsidR", "rsidRPr", "rsidDel")})
        _run_properties(run)
        _only_children(run, {(W, "rPr"), (W, "t"), (W, "annotationRef")})
        for reference in _children(run, W, "annotationRef"):
            _format_leaf(reference, set())
        for text in _children(run, W, "t"):
            _known_attrs(text, {(XML_NS, "space")})
            if any(child.nodeType not in (child.TEXT_NODE, child.CDATA_SECTION_NODE)
                   for child in text.childNodes):
                raise ValueError("Complex comment text is unsupported")
            texts.append(text)
    if len(texts) != 1:
        raise ValueError("Only one plain comment text node is supported")
    return paragraph, texts[0]


def _validate_basic(directory, extension):
    validator = StructureValidator(directory, extension=extension)
    with contextlib.redirect_stdout(io.StringIO()):
        valid = validator.validate()
    if not valid:
        raise ValueError("Basic structure checks failed: " + "; ".join(validator.errors[:5]))


def _state(directory):
    documents = {"word/document.xml": parse_dom(read_xml(directory / "word/document.xml"))}
    files = {path.relative_to(directory).as_posix() for path in directory.rglob("*") if path.is_file()}
    errors = []
    signed = any(name.lower().startswith("_xmlsignatures/") for name in files)
    known_types = {values[4]: part for part, values in COMMENT_PARTS.items()}
    for path in directory.rglob("*.rels"):
        dom = parse_dom(read_xml(path))
        for node in dom.getElementsByTagNameNS(REL, "Relationship"):
            kind = node.getAttribute("Type")
            if "digital-signature" in kind.lower():
                signed = True
            if any(token in kind.lower() for token in ("/comments", "/people", "/person")):
                if kind not in known_types or path.relative_to(directory).as_posix() != "word/_rels/document.xml.rels":
                    errors.append("Unknown comment/person relationship dependencies")
                elif (node.getAttribute("TargetMode") not in ("", "Internal")
                      or opc_target(node.getAttribute("Target"), "word/document.xml") != known_types[kind]):
                    errors.append("Unsupported comment metadata relationship target")
    if signed:
        errors.append("Digitally signed packages cannot be mutated")
    unexpected = [name for name in files if (
        name.lower().startswith("word/comments") and name not in COMMENT_PARTS
    ) or name.lower() in {"word/people.xml", "word/persons.xml"}
        or re.fullmatch(r"word/_rels/comments[^/]*\.rels", name)]
    if unexpected:
        errors.append("Unknown comment/person dependencies: " + ", ".join(sorted(unexpected)))
    present = set(COMMENT_PARTS) & files
    mode = "none" if not present else "classic" if present == {"word/comments.xml"} else "modern"
    if present and "word/comments.xml" not in present:
        errors.append("Modern metadata exists without comments.xml")
    if mode == "modern" and present != set(COMMENT_PARTS):
        errors.append("Modern comment metadata must contain all four supported parts")
    records = {}
    for part in present:
        documents[part] = parse_dom(read_xml(directory / part))
    state = {"documents": documents, "records": records, "mode": mode, "errors": errors,
             "files": files, "metadata": {}, "anchors": {}}
    if "word/comments.xml" not in present:
        return state
    root = documents["word/comments.xml"].documentElement
    try:
        if (root.namespaceURI, root.localName) != (W, "comments"):
            raise ValueError("Wrong comments root")
        _known_attrs(root, set(), root=True)
        _only_children(root, {(W, "comment")})
        for comment in _children(root, W, "comment"):
            value = _id(comment.getAttributeNS(W, "id"), xml=True)
            if value in records:
                raise ValueError("Duplicate comment IDs")
            record = {"id": value, "node": comment, "text": "".join(
                _text(node) for node in comment.getElementsByTagNameNS(W, "t")),
                "author": comment.getAttributeNS(W, "author"), "parent_id": None}
            records[value] = record
            try:
                paragraph, text = _comment_content(comment)
                record.update(paragraph=paragraph, text_node=text)
            except ValueError as exc:
                errors.append(f"Comment {value}: {exc}")
        if mode == "modern" and present == set(COMMENT_PARTS):
            _modern_state(state)
        _anchor_state(state)
        _metadata_relationships(directory, state)
    except ValueError as exc:
        errors.append(str(exc))
    return state


def _modern_state(state):
    records, documents, metadata = state["records"], state["documents"], state["metadata"]
    paragraphs = {}
    for value, record in records.items():
        if "paragraph" not in record:
            raise ValueError("Complex modern comments cannot be mutated")
        para_id = _hex(record["paragraph"].getAttributeNS(W14, "paraId"), "paraId")
        if para_id in paragraphs:
            raise ValueError("Duplicate comment paraId")
        paragraphs[para_id] = value
        record["para_id"] = para_id
    for part in META_ATTRS:
        namespace, _, root_name, child_name, _, _ = COMMENT_PARTS[part]
        root = documents[part].documentElement
        if (root.namespaceURI, root.localName) != (namespace, root_name):
            raise ValueError(f"Wrong metadata root in {part}")
        _known_attrs(root, set(), root=True)
        _only_children(root, {(namespace, child_name)})
        mapping = {}
        for node in _children(root, namespace, child_name):
            _known_attrs(node, {(namespace, key) for key in META_ATTRS[part]})
            _only_children(node, set())
            key_name = "durableId" if part.endswith("commentsExtensible.xml") else "paraId"
            key = _hex(node.getAttributeNS(namespace, key_name), key_name)
            if key in mapping:
                raise ValueError(f"Duplicate {key_name} in {part}")
            mapping[key] = node
        metadata[part] = mapping
    extensions = metadata["word/commentsExtended.xml"]
    identities = metadata["word/commentsIds.xml"]
    extensible = metadata["word/commentsExtensible.xml"]
    if set(extensions) != set(paragraphs) or set(identities) != set(paragraphs):
        raise ValueError("Comment paraId metadata is not one-to-one")
    durable_ids = set()
    for para_id, value in paragraphs.items():
        record = records[value]
        extension = extensions[para_id]
        done = extension.getAttributeNS(W15, "done")
        if done not in ("", "0", "1", "true", "false"):
            raise ValueError("Unknown resolved-state value")
        record["resolved"] = done in {"1", "true"}
        parent = extension.getAttributeNS(W15, "paraIdParent")
        if parent:
            parent = _hex(parent, "paraIdParent")
            if parent not in paragraphs or parent == para_id:
                raise ValueError("Missing or self-referential comment parent")
            record["parent_id"] = paragraphs[parent]
        durable = _hex(identities[para_id].getAttributeNS(W16CID, "durableId"), "durableId")
        if durable in durable_ids:
            raise ValueError("Duplicate comment durableId")
        durable_ids.add(durable)
        record["durable_id"] = durable
    if durable_ids != set(extensible):
        raise ValueError("Comment durableId metadata is not one-to-one")
    for value, record in records.items():
        visited = {value}
        parent = record["parent_id"]
        depth = 0
        while parent is not None:
            if parent in visited:
                raise ValueError("Cyclic comment parents")
            visited.add(parent)
            depth += 1
            parent = records[parent]["parent_id"]
        if depth > 1:
            raise ValueError("Only one reply level is supported")


def _anchor_state(state):
    dom = state["documents"]["word/document.xml"]
    roots = _children(dom.documentElement, W, "body")
    if (dom.documentElement.namespaceURI, dom.documentElement.localName) != (W, "document") or len(roots) != 1:
        raise ValueError("Comment anchors require one document body")
    body = roots[0]
    paragraphs = _children(body, W, "p")
    collected = {}
    for kind in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        for node in dom.getElementsByTagNameNS(W, kind):
            value = _id(node.getAttributeNS(W, "id"), xml=True)
            if value not in state["records"]:
                raise ValueError("Unknown comment anchor ID")
            group = collected.setdefault(value, {})
            if kind in group:
                raise ValueError("Duplicate comment anchor")
            group[kind] = node
    if set(collected) != set(state["records"]):
        raise ValueError("Each supported comment needs one body range and reference")
    for value, group in collected.items():
        if len(group) != 3:
            raise ValueError("Unpaired comment anchors")
        start, end, reference = (group[key] for key in
                                 ("commentRangeStart", "commentRangeEnd", "commentReference"))
        for marker in (start, end, reference):
            _known_attrs(marker, {(W, "id")})
            _only_children(marker, set())
        paragraph = start.parentNode
        if paragraph not in paragraphs or end.parentNode is not paragraph:
            raise ValueError("Only same-paragraph direct body ranges are supported")
        reference_run = reference.parentNode
        if (reference_run.namespaceURI, reference_run.localName) != (W, "r") or reference_run.parentNode is not paragraph:
            raise ValueError("Unsupported comment reference placement")
        _only_children(reference_run, {(W, "rPr"), (W, "commentReference")})
        _known_attrs(reference_run, {(W, key) for key in ("rsidR", "rsidRPr", "rsidDel")})
        if len(_children(reference_run, W, "commentReference")) != 1:
            raise ValueError("Mixed comment reference runs are unsupported")
        _run_properties(reference_run)
        children = list(paragraph.childNodes)
        positions = tuple(children.index(node) for node in (start, end, reference_run))
        if not positions[0] < positions[1] < positions[2]:
            raise ValueError("Comment anchors are not ordered")
        group.update(paragraph=paragraph, paragraph_index=paragraphs.index(paragraph),
                     reference_run=reference_run, positions=positions)
    for value, group in collected.items():
        parent_id = state["records"][value]["parent_id"]
        if parent_id is not None:
            parent = collected[parent_id]
            if group["paragraph"] is not parent["paragraph"] or not (
                parent["positions"][0] < group["positions"][0]
                and group["positions"][1] < parent["positions"][1]
            ):
                raise ValueError("Reply anchors must be nested inside their parent")
    state["anchors"] = collected


def _metadata_relationships(directory, state):
    path = directory / "word/_rels/document.xml.rels"
    if not path.exists():
        raise ValueError("Missing comments relationships")
    dom = parse_dom(read_xml(path))
    relationships = dom.getElementsByTagNameNS(REL, "Relationship")
    known_types = {item[4]: part for part, item in COMMENT_PARTS.items()}
    seen = set()
    for node in relationships:
        kind = node.getAttribute("Type")
        if any(token in kind.lower() for token in ("/comments", "/people", "/person")):
            if kind not in known_types:
                raise ValueError("Unknown comment relationship type")
            part = known_types[kind]
            target = opc_target(node.getAttribute("Target"), "word/document.xml", node.getAttribute("TargetMode"))
            if target != part or part not in state["files"] or part in seen:
                raise ValueError("Unsupported comment relationship target or duplicate")
            seen.add(part)
    if seen != set(COMMENT_PARTS) & state["files"]:
        raise ValueError("Missing metadata relationship")
    types = parse_dom(read_xml(directory / "[Content_Types].xml"))
    for part in seen:
        overrides = [node for node in types.getElementsByTagNameNS(CT, "Override")
                     if node.getAttribute("PartName") == "/" + part]
        if len(overrides) != 1 or overrides[0].getAttribute("ContentType") != COMMENT_PARTS[part][5]:
            raise ValueError("Incorrect comment content type")


def _fresh_hex(used):
    for _ in range(128):
        value = f"{secrets.randbelow(0x7FFFFFFE) + 1:08X}"
        if value not in used:
            return value
    raise ValueError("Could not allocate a unique comment metadata ID")


def _ensure_parts(directory, state):
    documents = state["documents"]
    for part, (namespace, prefix, root_name, _, _, _) in COMMENT_PARTS.items():
        if part not in documents:
            documents[part] = parse_dom(f'<{prefix}:{root_name} xmlns:{prefix}="{namespace}"/>')
    rels_part = "word/_rels/document.xml.rels"
    rels_path = directory / rels_part
    rels = parse_dom(read_xml(rels_path)) if rels_path.exists() else parse_dom(f'<Relationships xmlns="{REL}"/>')
    types = parse_dom(read_xml(directory / "[Content_Types].xml"))
    ids = {node.getAttribute("Id") for node in rels.getElementsByTagNameNS(REL, "Relationship")}
    for part, values in COMMENT_PARTS.items():
        if part in state["files"]:
            continue
        index = 1
        while "rId" + str(index) in ids:
            index += 1
        rid = "rId" + str(index)
        ids.add(rid)
        relationship = _new(rels, REL, "rel", "Relationship")
        for key, value in (("Id", rid), ("Type", values[4]), ("Target", part.removeprefix("word/"))):
            relationship.setAttribute(key, value)
        rels.documentElement.appendChild(relationship)
        override = _new(types, CT, "ct", "Override")
        override.setAttribute("PartName", "/" + part)
        override.setAttribute("ContentType", values[5])
        types.documentElement.appendChild(override)
    documents[rels_part] = rels
    documents["[Content_Types].xml"] = types


def _add_definition(state, text, author, initials, parent_id=None):
    records, documents = state["records"], state["documents"]
    value = max(records, default=-1) + 1
    _id(value)
    used_para = {record["para_id"] for record in records.values()}
    used_durable = {record["durable_id"] for record in records.values()}
    para_id, durable = _fresh_hex(used_para), _fresh_hex(used_durable)
    stamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    dom = documents["word/comments.xml"]
    comment = _new(dom, W, "w", "comment")
    for key, field in (("id", value), ("author", author), ("initials", initials), ("date", stamp)):
        _attr(comment, W, "w", key, field)
    paragraph = _new(dom, W, "w", "p")
    _attr(paragraph, W14, "w14", "paraId", para_id)
    run = _new(dom, W, "w", "r")
    node = _new(dom, W, "w", "t")
    node.setAttributeNS(XML_NS, "xml:space", "preserve")
    node.appendChild(dom.createTextNode(text))
    run.appendChild(node)
    paragraph.appendChild(run)
    comment.appendChild(paragraph)
    dom.documentElement.appendChild(comment)
    for part, fields in (
        ("word/commentsExtended.xml", {"paraId": para_id, "done": "0", **(
            {"paraIdParent": records[parent_id]["para_id"]} if parent_id is not None else {})}),
        ("word/commentsIds.xml", {"paraId": para_id, "durableId": durable}),
        ("word/commentsExtensible.xml", {"durableId": durable, "dateUtc": stamp}),
    ):
        namespace, prefix, _, child_name, _, _ = COMMENT_PARTS[part]
        dom = documents[part]
        child = _new(dom, namespace, prefix, child_name)
        for key, field in fields.items():
            _attr(child, namespace, prefix, key, field)
        dom.documentElement.appendChild(child)
    return value


def _anchor_for_add(dom, specification):
    body = _children(dom.documentElement, W, "body")[0]
    paragraphs = _children(body, W, "p")
    index = specification["paragraph_index"]
    if type(index) is not int or not 0 <= index < len(paragraphs):
        raise ValueError("paragraph_index must identify a direct body paragraph using a zero-based integer")
    paragraph = paragraphs[index]
    runs, text = _plain_paragraph(paragraph, _blocked_anchors(dom))
    if text != specification["expected_text"]:
        raise ValueError("Anchor paragraph text mismatch")
    for run in runs:
        _prepare_run(run)
        for node in _elements(run, "t"):
            if not _preserves_space(node) and _text(node) != _text(node).strip(XML_SPACE):
                raise ValueError("Ambiguous anchor whitespace; comment mutation is unsupported")
    needle = specification["anchor_text"]
    starts, cursor = [], 0
    while (cursor := text.find(needle, cursor)) != -1:
        starts.append(cursor)
        cursor += 1
    if len(starts) != 1:
        raise ValueError("Anchor text must match exactly once in the selected paragraph")
    boundaries, offset = {0: 0}, 0
    for index, run in enumerate(runs):
        value = "".join(_text(node) for node in _elements(run, "t"))
        if not value:
            raise ValueError("Empty anchor runs are unsupported")
        offset += len(value)
        boundaries[offset] = index + 1
    start, end = starts[0], starts[0] + len(needle)
    if start not in boundaries or end not in boundaries:
        raise ValueError("Anchor must cover complete consecutive runs; run splitting is unsupported")
    return paragraph, runs[boundaries[start]], runs[boundaries[end] - 1]


def _insert_markers(dom, paragraph, first, last, value, *, before_end=None):
    start, end = (_new(dom, W, "w", key) for key in ("commentRangeStart", "commentRangeEnd"))
    for node in (start, end):
        _attr(node, W, "w", "id", value)
    paragraph.insertBefore(start, first)
    paragraph.insertBefore(end, last.nextSibling)
    reference_run = _new(dom, W, "w", "r")
    reference = _new(dom, W, "w", "commentReference")
    _attr(reference, W, "w", "id", value)
    reference_run.appendChild(reference)
    paragraph.insertBefore(reference_run, before_end if before_end is not None else end.nextSibling)


def _reply_anchor(state, parent_id, value):
    group = state["anchors"][parent_id]
    paragraph = group["paragraph"]
    dom = state["documents"]["word/document.xml"]
    clean = dom.cloneNode(True)
    for kind in ("commentRangeStart", "commentRangeEnd"):
        for node in list(clean.getElementsByTagNameNS(W, kind)):
            node.parentNode.removeChild(node)
    for node in list(clean.getElementsByTagNameNS(W, "commentReference")):
        node.parentNode.parentNode.removeChild(node.parentNode)
    clean_body = _children(clean.documentElement, W, "body")[0]
    clean_paragraph = _children(clean_body, W, "p")[group["paragraph_index"]]
    _plain_paragraph(clean_paragraph, _blocked_anchors(clean))
    children = list(paragraph.childNodes)
    inside = children[group["positions"][0] + 1:group["positions"][1]]
    runs = [node for node in inside if node.nodeType == node.ELEMENT_NODE and node.namespaceURI == W
            and node.localName == "r" and not _children(node, W, "commentReference")]
    if not runs:
        raise ValueError("Empty reply anchors are unsupported")
    # Cover the same original runs. The reply's reference stays outside the
    # parent range so references never become part of another comment's text.
    _insert_markers(dom, paragraph, runs[0], runs[-1], value,
                    before_end=group["reference_run"].nextSibling)


def _specification(value):
    if (not isinstance(value, dict) or not isinstance(value.get("action"), str)
            or value["action"] not in {"list", "add", "update", "delete", "reply"}):
        raise ValueError("Expected one comment action: list, add, update, delete or reply")
    action = value["action"]
    fields = {"action"} | ({"paragraph_index", "expected_text", "anchor_text", "text", "author", "initials"}
                            if action == "add" else {"id", "expected_text", "text", "author", "initials"}
                            if action == "reply" else {"id", "expected_text", "text"}
                            if action == "update" else {"id", "expected_text"}
                            if action == "delete" else set())
    if set(value) - fields:
        raise ValueError("Unsupported comment action fields")
    if action != "list":
        _single_line(value.get("expected_text"), "expected_text", empty=action != "add")
    if action in {"update", "delete", "reply"}:
        _id(value.get("id"))
    if action in {"add", "reply", "update"}:
        _single_line(value.get("text"), "text")
    if action in {"add", "reply"}:
        _single_line(value.get("author"), "author")
        _single_line(value.get("initials", ""), "initials", empty=True)
    if action == "add":
        _single_line(value.get("anchor_text"), "anchor_text")
        if "paragraph_index" not in value:
            raise ValueError("paragraph_index is required")
    return action


def manage_comments(source, specification, destination=None):
    """Inspect or atomically mutate one supported comment action."""
    action = _specification(specification)
    source = Path(source)
    if source.suffix.lower() not in {".docx", ".dotx"}:
        raise ValueError("Input must be DOCX or DOTX")
    if action == "list":
        if destination is not None:
            raise ValueError("list is read-only and takes no document output")
    else:
        if destination is None:
            raise ValueError("Mutations require a separate output document")
        destination = Path(destination)
        if destination.suffix.lower() != source.suffix.lower():
            raise ValueError("Input and output must have the same DOCX/DOTX extension")
        if source.resolve() == destination.resolve() or (destination.exists() and source.samefile(destination)):
            raise ValueError("Output must not refer to the input")
        if destination.is_symlink():
            raise ValueError("Symbolic link outputs are unsupported")
    with tempfile.TemporaryDirectory(prefix="docx-comment-manage-") as temporary:
        directory = unpack_document(source, Path(temporary) / "package")
        state = _state(directory)
        if action == "list":
            try:
                _validate_basic(directory, source.suffix.lower())
            except ValueError as exc:
                state["errors"].append(str(exc))
            return {"action": "list", "mode": state["mode"], "mutable": not state["errors"],
                    "diagnostics": state["errors"], "comments": [
                        {"id": value, "text": record["text"], "author": record["author"],
                         "parent_id": record["parent_id"], "paragraph_index": state["anchors"].get(value, {}).get("paragraph_index")}
                        for value, record in sorted(state["records"].items())]}
        _validate_basic(directory, source.suffix.lower())
        if state["errors"]:
            raise ValueError("Comment mutation is unsupported: " + "; ".join(state["errors"]))
        documents, records = state["documents"], state["records"]
        if action == "add":
            if state["mode"] == "classic" and records:
                raise ValueError("Adding to classic comments requires an unsupported metadata upgrade")
            paragraph, first, last = _anchor_for_add(documents["word/document.xml"], specification)
            _ensure_parts(directory, state)
            value = _add_definition(state, specification["text"], specification["author"], specification.get("initials", ""))
            _insert_markers(documents["word/document.xml"], paragraph, first, last, value)
        else:
            value = specification["id"]
            if value not in records or records[value]["text"] != specification["expected_text"]:
                raise ValueError("Comment ID or expected_text guard mismatch")
            record = records[value]
            if action == "reply":
                if state["mode"] != "modern" or record["parent_id"] is not None:
                    raise ValueError("Replies require supported modern root comments; nested replies are unsupported")
                if record.get("resolved"):
                    raise ValueError("Replies to resolved comments are unsupported")
                reply_id = _add_definition(state, specification["text"], specification["author"],
                                           specification.get("initials", ""), parent_id=value)
                _reply_anchor(state, value, reply_id)
                value = reply_id
            elif action == "update":
                text = record["text_node"]
                for child in list(text.childNodes):
                    text.removeChild(child)
                text.appendChild(text.ownerDocument.createTextNode(specification["text"]))
                text.setAttributeNS(XML_NS, "xml:space", "preserve")
            else:
                if any(item["parent_id"] == value for item in records.values()):
                    raise ValueError("Delete replies first; cascading comment deletion is unsupported")
                record["node"].parentNode.removeChild(record["node"])
                group = state["anchors"][value]
                for node in (group["commentRangeStart"], group["commentRangeEnd"], group["reference_run"]):
                    node.parentNode.removeChild(node)
                if state["mode"] == "modern":
                    for part, key in (("word/commentsExtended.xml", record["para_id"]),
                                      ("word/commentsIds.xml", record["para_id"]),
                                      ("word/commentsExtensible.xml", record["durable_id"])):
                        node = state["metadata"][part][key]
                        node.parentNode.removeChild(node)
        for part, dom in documents.items():
            path = directory / part
            if path.exists() and parse_dom(read_xml(path)).toxml() == dom.toxml():
                continue
            data = _serialize_checked(dom)
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(data)
        _validate_basic(directory, destination.suffix.lower())
        checked = _state(directory)
        if checked["errors"]:
            raise ValueError("Edited comment metadata failed checks: " + "; ".join(checked["errors"]))
        pack_document(directory, destination)
    return {"action": action, "comment_id": value}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("action", type=Path, help="JSON containing exactly one action")
    parser.add_argument("output", type=Path, nargs="?")
    args = parser.parse_args(argv)
    try:
        with args.action.open(encoding="utf-8") as stream:
            specification = json.load(stream)
        result = manage_comments(args.input, specification, args.output)
    except (OSError, ValueError, RuntimeError, zipfile.BadZipFile) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
