"""Insert/delete one data row in a direct-body rectangular plain-text table.

JSON: {"operation": {"type": "insert_row", "table_index": 0, "row_index": 1,
"expected_rows": 3, "expected_text": ["A", "B"], "position": "after",
"cells": ["new A", "new B"]}}. Delete uses type=delete_row and omits cells
and position. Indices are zero-based. Header rows and the final data row are
protected. Insertion clones the selected data row's plain cell/run formatting.
"""

import argparse
import json
import sys
import tempfile
import zipfile
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from edit_docx import XML_NS, XMLNS_NS, _eligible, _serialize_checked, _text
from edit_structure import RUN_PROPERTY_ORDER, _blocked_anchors, _elements, _index, _single_line
from office.package import pack_document, unpack_document
from office.validators.structure import StructureValidator, W
from office.xmlio import parse_dom, read_xml

TABLE_PROPERTIES = {
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
    "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
}
ROW_PROPERTIES = {"cantSplit", "tblHeader", "trHeight", "tblCellSpacing", "jc"}
CELL_PROPERTIES = {"tcW", "tcBorders", "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark"}
PARAGRAPH_PROPERTIES = {
    "pStyle", "keepNext", "keepLines", "widowControl", "snapToGrid", "spacing", "ind",
    "contextualSpacing", "bidi", "textAlignment", "jc",
}
PROPERTY_CHILDREN = {"top", "bottom", "left", "right", "start", "end", "insideH", "insideV", "tl2br", "tr2bl"}


def _safe_attributes(node, *, structural=False):
    for attr in node.attributes.values():
        if attr.namespaceURI == XMLNS_NS:
            continue
        if attr.namespaceURI == XML_NS and attr.localName == "space":
            if attr.value not in {"default", "preserve"}:
                raise ValueError("Invalid xml:space in table; editing is unsupported")
            continue
        if attr.namespaceURI != W or attr.localName.lower().endswith("id"):
            raise ValueError("Unknown/reference-bearing table attributes; editing is unsupported")
        if structural and attr.localName not in {"rsidR", "rsidRPr", "rsidDel", "rsidTr", "rsidP", "rsidRDefault"}:
            raise ValueError("Unknown structural table attributes; editing is unsupported")


def _children(node, allowed):
    _safe_attributes(node, structural=True)
    result = []
    for child in node.childNodes:
        if child.nodeType == child.TEXT_NODE and not child.data.strip(" \t\r\n"):
            continue
        if child.nodeType != child.ELEMENT_NODE or child.namespaceURI != W or child.localName not in allowed:
            raise ValueError("Complex/unknown table content; editing is unsupported")
        result.append(child)
    return result


def _properties(node, allowed):
    names = set()
    _safe_attributes(node)
    for child in node.childNodes:
        if child.nodeType == child.TEXT_NODE and not child.data.strip(" \t\r\n"):
            continue
        if child.nodeType != child.ELEMENT_NODE or child.namespaceURI != W or child.localName not in allowed:
            raise ValueError("Complex/unknown table properties; editing is unsupported")
        if child.localName in names:
            raise ValueError("Duplicate table properties; editing is unsupported")
        names.add(child.localName)
        _safe_attributes(child)
        if _elements(child):
            _properties(child, PROPERTY_CHILDREN)
        elif any(c.nodeType != c.TEXT_NODE or c.data.strip(" \t\r\n") for c in child.childNodes):
            raise ValueError("Complex table property values; editing is unsupported")


def _header(row):
    nodes = row.getElementsByTagNameNS(W, "tblHeader")
    if not nodes:
        return False
    value = nodes[0].getAttributeNS(W, "val")
    if value not in {"", "1", "0", "true", "false", "on", "off"}:
        raise ValueError("Invalid repeated-header flag; editing is unsupported")
    return value not in {"0", "false", "off"}


def _plain_table(table, blocked):
    _children(table, {"tblPr", "tblGrid", "tr"})
    properties = _elements(table, "tblPr")
    grids = _elements(table, "tblGrid")
    rows = _elements(table, "tr")
    if len(properties) > 1 or len(grids) != 1 or not rows:
        raise ValueError("A rectangular table requires one explicit grid and data rows")
    if properties:
        _properties(properties[0], TABLE_PROPERTIES)
    _children(grids[0], {"gridCol"})
    columns = len(_elements(grids[0], "gridCol"))
    if columns < 1:
        raise ValueError("An explicit nonempty table grid is required")
    for column in _elements(grids[0], "gridCol"):
        _safe_attributes(column)
        if column.childNodes:
            raise ValueError("Complex grid columns are unsupported")
    records = []
    seen_data = False
    for row in rows:
        _children(row, {"trPr", "tc"})
        row_properties = _elements(row, "trPr")
        if len(row_properties) > 1:
            raise ValueError("Duplicate row properties are unsupported")
        if row_properties:
            _properties(row_properties[0], ROW_PROPERTIES)
        is_header = _header(row)
        if is_header and seen_data:
            raise ValueError("Repeated header rows must be leading and contiguous")
        seen_data = seen_data or not is_header
        cells = _elements(row, "tc")
        if len(cells) != columns:
            raise ValueError("Only rectangular, unmerged tables are supported")
        texts = []
        for cell in cells:
            _children(cell, {"tcPr", "p"})
            cell_properties = _elements(cell, "tcPr")
            paragraphs = _elements(cell, "p")
            if len(cell_properties) > 1 or len(paragraphs) != 1:
                raise ValueError("Each cell must have one plain paragraph")
            if cell_properties:
                _properties(cell_properties[0], CELL_PROPERTIES)
            paragraph = paragraphs[0]
            _children(paragraph, {"pPr", "r"})
            paragraph_properties = _elements(paragraph, "pPr")
            runs = _elements(paragraph, "r")
            if len(paragraph_properties) > 1 or len(runs) != 1:
                raise ValueError("Each cell paragraph must have one plain run")
            if paragraph_properties:
                _properties(paragraph_properties[0], PARAGRAPH_PROPERTIES)
            _children(runs[0], {"rPr", "t"})
            run_properties = _elements(runs[0], "rPr")
            nodes = _elements(runs[0], "t")
            if len(run_properties) > 1 or len(nodes) != 1:
                raise ValueError("Each cell run must have one text node")
            if run_properties:
                _properties(run_properties[0], set(RUN_PROPERTY_ORDER))
            node = nodes[0]
            if node in blocked or not _eligible(node):
                raise ValueError("Active range or complex cell text; editing is unsupported")
            texts.append(_text(node))
        records.append((row, is_header, texts))
    return records


def edit_tables(source, destination, specification, *, require_no_placeholders=False):
    """Apply exactly one guarded insert_row/delete_row operation to a new file."""
    source, destination = Path(source), Path(destination)
    if source.suffix.lower() not in {".docx", ".dotx"} or source.suffix.lower() != destination.suffix.lower():
        raise ValueError("Input and output must have the same DOCX/DOTX extension")
    if source.resolve() == destination.resolve() or (destination.exists() and source.samefile(destination)):
        raise ValueError("Output must not refer to the input file")
    if not isinstance(specification, dict) or set(specification) != {"operation"}:
        raise ValueError("Expected only one operation object")
    operation = specification["operation"]
    if not isinstance(operation, dict):
        raise ValueError("operation must be an object")
    kind = operation.get("type")
    common = {"type", "table_index", "row_index", "expected_rows", "expected_text"}
    if not isinstance(kind, str) or kind not in {"insert_row", "delete_row"}:
        raise ValueError("Supported operation types are insert_row and delete_row")
    required = common | ({"position", "cells"} if kind == "insert_row" else set())
    if set(operation) != required or type(operation["expected_rows"]) is not int:
        raise ValueError("Missing, unknown or invalid operation fields")
    with tempfile.TemporaryDirectory(prefix="docx-table-edit-") as temporary:
        unpacked = unpack_document(source, Path(temporary) / "package")
        validator = StructureValidator(unpacked, extension=source.suffix.lower())
        if not validator.validate():
            raise ValueError("Input failed implemented structure checks")
        if any(name.startswith("_xmlsignatures/") for name in validator.files) or any(
            "/digital-signature/" in kind
            for relations in validator.relations.values() for kind, _ in relations.values()
        ):
            raise ValueError("Signed packages are unsupported for structural editing")
        dom = parse_dom(read_xml(unpacked / "word/document.xml"))
        body = dom.documentElement.getElementsByTagNameNS(W, "body")[0]
        tables = _elements(body, "tbl")
        table_index = _index(operation["table_index"], len(tables), "table_index")
        table = tables[table_index]
        rows = _plain_table(table, _blocked_anchors(dom))
        if operation["expected_rows"] != len(rows):
            raise ValueError("expected_rows does not match the original table")
        row_index = _index(operation["row_index"], len(rows), "row_index")
        anchor, is_header, original_text = rows[row_index]
        expected = operation["expected_text"]
        if not isinstance(expected, list) or expected != original_text:
            raise ValueError("expected_text must match every original anchor-row cell")
        if is_header:
            raise ValueError("Header rows are protected; choose a data-row anchor")
        if kind == "delete_row":
            if sum(not header for _, header, _ in rows) <= 1:
                raise ValueError("The final remaining data row is protected")
            table.removeChild(anchor)
        else:
            position, values = operation["position"], operation["cells"]
            if not isinstance(position, str) or position not in {"before", "after"}:
                raise ValueError("position must be before or after")
            if not isinstance(values, list) or len(values) != len(original_text):
                raise ValueError("cells must provide one text value for every table column")
            for value in values:
                _single_line(value, "cell text", empty=True)
            new_row = anchor.cloneNode(True)
            for cell, value in zip(_elements(new_row, "tc"), values):
                node = cell.getElementsByTagNameNS(W, "t")[0]
                for child in list(node.childNodes):
                    node.removeChild(child)
                if value:
                    node.appendChild(dom.createTextNode(value))
                node.setAttributeNS(XML_NS, "xml:space", "preserve")
            table.insertBefore(new_row, anchor if position == "before" else anchor.nextSibling)
        (unpacked / "word/document.xml").write_bytes(_serialize_checked(dom))
        if not StructureValidator(unpacked, extension=destination.suffix.lower(),
                                  require_no_placeholders=require_no_placeholders).validate():
            raise ValueError("Edited document failed implemented structure checks")
        return pack_document(unpacked, destination)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("operation", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--require-no-placeholders", action="store_true")
    args = parser.parse_args(argv)
    try:
        result = edit_tables(args.input, args.output, json.loads(args.operation.read_text(encoding="utf-8")),
                             require_no_placeholders=args.require_no_placeholders)
    except (OSError, ValueError, RuntimeError, zipfile.BadZipFile) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
