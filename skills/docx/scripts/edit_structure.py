"""Insert plain body paragraphs or format plain runs with verified anchors.

This deliberately excludes table, field, revision and complex-container editing.
All operations address the original document and write a separate output file.
"""

import argparse
import json
import math
import re
import sys
import tempfile
import zipfile
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from edit_docx import XML_NS, XMLNS_NS, _eligible, _range_text_nodes, _serialize_checked, _text
from office.package import pack_document, unpack_document
from office.helpers import reject_signed_package
from office.validators.structure import StructureValidator, W
from office.xmlio import parse_dom, read_xml

RUN_PROPERTY_ORDER = (
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
    "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
    "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
    "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "oMath",
)
FORMAT_FIELDS = {"bold", "italic", "size_pt", "color", "font_east_asia", "font_latin"}
ANCHOR_RANGE_STARTS = {
    "bookmarkStart": "bookmark", "permStart": "permission",
    "customXmlInsRangeStart": "customXmlIns", "customXmlDelRangeStart": "customXmlDel",
    "customXmlMoveFromRangeStart": "customXmlMoveFrom", "customXmlMoveToRangeStart": "customXmlMoveTo",
}
ANCHOR_RANGE_ENDS = {name.removesuffix("Start") + "End": kind for name, kind in ANCHOR_RANGE_STARTS.items()}


def _elements(node, name=None):
    return [child for child in node.childNodes if child.nodeType == child.ELEMENT_NODE
            and (name is None or (child.namespaceURI == W and child.localName == name))]


def _single_line(value, label, *, empty=False):
    if not isinstance(value, str) or (not empty and not value):
        raise ValueError(f"{label} must be {'possibly empty ' if empty else 'nonempty '}text")
    if any(ord(char) < 32 or 0xD800 <= ord(char) <= 0xDFFF or char in "\ufffe\uffff" for char in value):
        raise ValueError(f"{label} must be single-line XML text")
    return value


def _index(value, length, label):
    if type(value) is not int or not 0 <= value < length:
        raise ValueError(f"{label} must identify an existing item using a zero-based integer")
    return value


def _blocked_anchors(dom):
    blocked = _range_text_nodes(dom)
    active, seen = set(), set()
    for node in dom.getElementsByTagName("*"):
        if node.namespaceURI != W:
            continue
        name = node.localName
        if name in ANCHOR_RANGE_STARTS or name in ANCHOR_RANGE_ENDS:
            value = node.getAttributeNS(W, "id")
            key = (ANCHOR_RANGE_STARTS.get(name, ANCHOR_RANGE_ENDS.get(name)), value)
            if not value:
                raise ValueError("Missing anchor range ID; editing is unsupported")
            if name in ANCHOR_RANGE_STARTS:
                if key in seen:
                    raise ValueError("Duplicate anchor range; editing is unsupported")
                seen.add(key)
                active.add(key)
            elif key not in active:
                raise ValueError("Unbalanced anchor range; editing is unsupported")
            else:
                active.remove(key)
        elif name == "t" and active:
            blocked.add(node)
    if active:
        raise ValueError("Unclosed anchor range; editing is unsupported")
    return blocked


def _plain_paragraph(paragraph, blocked):
    children = _elements(paragraph)
    if any(child.namespaceURI != W or child.localName not in {"pPr", "r"} for child in children):
        raise ValueError("Anchor paragraph has a complex container or range; editing is unsupported")
    if len(_elements(paragraph, "pPr")) > 1:
        raise ValueError("Duplicate paragraph properties; editing is unsupported")
    for name in ("pPrChange", "rPrChange", "sectPr"):
        if paragraph.getElementsByTagNameNS(W, name):
            raise ValueError("Anchor has revision/section properties; editing is unsupported")
    runs = _elements(paragraph, "r")
    nodes = []
    for run in runs:
        if any(child.namespaceURI != W or child.localName not in {"rPr", "t"}
               for child in _elements(run)):
            raise ValueError("Anchor run is not plain text; editing is unsupported")
        texts = _elements(run, "t")
        if not texts or any(node in blocked or not _eligible(node) for node in texts):
            raise ValueError("Anchor includes a field, revision, range or unknown text; editing is unsupported")
        nodes.extend(texts)
    text = "".join(_text(node) for node in nodes)
    if not text:
        raise ValueError("Empty anchor paragraphs are unsupported")
    return runs, text


def _format_spec(value):
    if not isinstance(value, dict) or not value or set(value) - FORMAT_FIELDS:
        raise ValueError("format must contain supported nonempty run formatting fields")
    result = value.copy()
    for key, field in value.items():
        if key in {"bold", "italic"}:
            if type(field) is not bool:
                raise ValueError(f"format.{key} must be a boolean")
        elif key == "size_pt":
            if (type(field) not in (int, float) or not 1 <= field <= 200
                    or not math.isfinite(field) or not float(field * 2).is_integer()):
                raise ValueError("format.size_pt must be 1..200 in half-point increments")
        elif key == "color":
            if not isinstance(field, str) or not re.fullmatch(r"[0-9A-Fa-f]{6}", field):
                raise ValueError("format.color must be six hexadecimal RGB digits")
            result[key] = field.upper()
        else:
            _single_line(field, f"format.{key}")
    return result


def _prepare_run(run):
    properties = _elements(run, "rPr")
    if len(properties) > 1:
        raise ValueError("Duplicate run properties; editing is unsupported")
    if properties:
        names = []
        for child in _elements(properties[0]):
            if child.namespaceURI != W or child.localName not in RUN_PROPERTY_ORDER:
                raise ValueError("Unknown or revised run properties; editing is unsupported")
            names.append(child.localName)
        if len(names) != len(set(names)):
            raise ValueError("Duplicate run property; editing is unsupported")


def _new_element(dom, name):
    node = dom.createElementNS(W, "w:" + name)
    node.setAttributeNS(XMLNS_NS, "xmlns:w", W)
    return node


def _set_property(properties, name, attributes):
    existing = _elements(properties, name)
    if existing:
        node = existing[0]
    else:
        node = _new_element(properties.ownerDocument, name)
        next_node = next((child for child in _elements(properties)
                          if RUN_PROPERTY_ORDER.index(child.localName) > RUN_PROPERTY_ORDER.index(name)), None)
        properties.insertBefore(node, next_node)
    prefix = node.prefix or "w"
    if not node.prefix:
        node.setAttributeNS(XMLNS_NS, "xmlns:w", W)
    for key, value in attributes.items():
        attribute = node.getAttributeNodeNS(W, key)
        if attribute is not None:
            attribute.value = str(value)
        else:
            node.setAttributeNS(W, prefix + ":" + key, str(value))
    return node


def _apply_format(run, specification):
    properties = _elements(run, "rPr")
    if properties:
        properties = properties[0]
    else:
        properties = _new_element(run.ownerDocument, "rPr")
        run.insertBefore(properties, run.firstChild)
    for key, names in (("bold", ("b", "bCs")), ("italic", ("i", "iCs"))):
        if key in specification:
            for name in names:
                _set_property(properties, name, {"val": "1" if specification[key] else "0"})
    if "size_pt" in specification:
        for name in ("sz", "szCs"):
            _set_property(properties, name, {"val": int(specification["size_pt"] * 2)})
    if "color" in specification:
        color = _set_property(properties, "color", {"val": specification["color"]})
        for name in ("themeColor", "themeTint", "themeShade"):
            if color.hasAttributeNS(W, name):
                color.removeAttributeNS(W, name)
    fonts = {}
    if "font_east_asia" in specification:
        fonts["eastAsia"] = specification["font_east_asia"]
    if "font_latin" in specification:
        fonts.update(ascii=specification["font_latin"], hAnsi=specification["font_latin"])
    if fonts:
        node = _set_property(properties, "rFonts", fonts)
        for name in fonts:
            theme_name = name + "Theme"
            if node.hasAttributeNS(W, theme_name):
                node.removeAttributeNS(W, theme_name)


def edit_structure(source, destination, specification, *, require_no_placeholders=False):
    """Apply operations to direct body paragraphs; indices refer to original XML."""
    source, destination = Path(source), Path(destination)
    if source.suffix.lower() not in {".docx", ".dotx"} or source.suffix.lower() != destination.suffix.lower():
        raise ValueError("Input and output must have the same DOCX/DOTX extension")
    if source.resolve() == destination.resolve() or (destination.exists() and source.samefile(destination)):
        raise ValueError("Output must not refer to the input file")
    if not isinstance(specification, dict) or set(specification) != {"operations"}:
        raise ValueError("Expected an object containing only operations")
    operations = specification["operations"]
    if not isinstance(operations, list) or not operations:
        raise ValueError("operations must be a nonempty list")
    with tempfile.TemporaryDirectory(prefix="docx-structure-edit-") as temporary:
        unpacked = unpack_document(source, Path(temporary) / "package")
        validator = StructureValidator(unpacked, extension=source.suffix.lower())
        if not validator.validate():
            raise ValueError("Input failed implemented structure checks")
        reject_signed_package(validator.files, (kind for relationships in validator.relations.values()
                                               for kind, _ in relationships.values()))
        dom = parse_dom(read_xml(unpacked / "word/document.xml"))
        body = dom.documentElement.getElementsByTagNameNS(W, "body")[0]
        paragraphs = _elements(body, "p")
        blocked = _blocked_anchors(dom)
        paragraph_styles = set()
        styles_path = unpacked / "word/styles.xml"
        if styles_path.is_file():
            styles = parse_dom(read_xml(styles_path))
            paragraph_styles = {node.getAttributeNS(W, "styleId")
                                for node in styles.getElementsByTagNameNS(W, "style")
                                if node.getAttributeNS(W, "type") == "paragraph"}
        plan, seen = [], set()
        for operation in operations:
            if not isinstance(operation, dict):
                raise ValueError("Each operation must be an object")
            kind = operation.get("type")
            common = {"type", "paragraph_index", "expected_text"}
            allowed = common | ({"position", "text", "style"} if kind == "insert_paragraph"
                                else {"run_index", "expected_run_text", "format"})
            if not isinstance(kind, str) or kind not in {"insert_paragraph", "format_run"} or set(operation) - allowed:
                raise ValueError("Unknown operation or unsupported operation fields")
            index = _index(operation.get("paragraph_index"), len(paragraphs), "paragraph_index")
            paragraph = paragraphs[index]
            runs, actual = _plain_paragraph(paragraph, blocked)
            expected = _single_line(operation.get("expected_text"), "expected_text")
            if actual != expected:
                raise ValueError(f"Anchor text mismatch at body paragraph {index}")
            if kind == "insert_paragraph":
                position = operation.get("position")
                if not isinstance(position, str) or position not in {"before", "after"}:
                    raise ValueError("position must be before or after")
                _single_line(operation.get("text"), "text")
                if "style" in operation:
                    _single_line(operation["style"], "style")
                    if operation["style"] not in paragraph_styles:
                        raise ValueError("style must identify an existing paragraph style")
                key = (kind, index, position)
                plan.append((kind, paragraph, operation))
            else:
                run_index = _index(operation.get("run_index"), len(runs), "run_index")
                run = runs[run_index]
                expected_run = _single_line(operation.get("expected_run_text"), "expected_run_text")
                if "".join(_text(node) for node in _elements(run, "t")) != expected_run:
                    raise ValueError(f"Run text mismatch at paragraph {index}, run {run_index}")
                _prepare_run(run)
                formatting = _format_spec(operation.get("format"))
                key = (kind, index, run_index)
                plan.append((kind, run, formatting))
            if key in seen:
                raise ValueError("Multiple operations on the same target/position are unsupported")
            seen.add(key)
        for kind, target, options in plan:
            if kind == "format_run":
                _apply_format(target, options)
                continue
            paragraph = _new_element(dom, "p")
            if "style" in options:
                properties = _new_element(dom, "pPr")
                style = _new_element(dom, "pStyle")
                style.setAttributeNS(W, "w:val", options["style"])
                properties.appendChild(style)
                paragraph.appendChild(properties)
            run = _new_element(dom, "r")
            text = _new_element(dom, "t")
            text.setAttributeNS(XML_NS, "xml:space", "preserve")
            text.appendChild(dom.createTextNode(options["text"]))
            run.appendChild(text)
            paragraph.appendChild(run)
            body.insertBefore(paragraph, target if options["position"] == "before" else target.nextSibling)
        (unpacked / "word/document.xml").write_bytes(_serialize_checked(dom))
        if not StructureValidator(unpacked, extension=destination.suffix.lower(),
                                  require_no_placeholders=require_no_placeholders).validate():
            raise ValueError("Edited document failed implemented structure checks")
        return pack_document(unpacked, destination)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("operations", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--require-no-placeholders", action="store_true")
    args = parser.parse_args(argv)
    try:
        specification = json.loads(args.operations.read_text(encoding="utf-8"))
        result = edit_structure(args.input, args.output, specification,
                                require_no_placeholders=args.require_no_placeholders)
    except (OSError, ValueError, RuntimeError, zipfile.BadZipFile) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
