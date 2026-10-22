"""Merge a conservative subset of adjacent plain WordprocessingML runs.

Only ordinary paragraphs in word/document.xml are processed. Fields, revisions,
ranges, complex runs and unknown formatting are left unchanged. Supported text
uses inherited xml:space; no renderer or XSD validation is claimed.

Usage:
    python merge_runs.py unpacked/
    python merge_runs.py input.docx -o merged.docx

The directory entry point writes document.xml atomically. Packaged input uses
atomic repacking; signed packages and symbolic links are rejected.
"""

import argparse
import contextlib
import io
import os
import sys
import tempfile
import zipfile
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from edit_docx import XML_NS, XMLNS_NS, _eligible, _preserves_space, _serialize_checked, _text
from edit_structure import _blocked_anchors
from office.archiveio import open_archive
from office.helpers import (
    MAX_MEMBERS, MAX_PACKAGE_BYTES, MAX_PART_BYTES, atomic_write_bytes, rendered_text,
    reject_signed_package, rezip, safe_extract,
)
from office.validators.structure import StructureValidator
from office.xmlio import parse_dom, read_xml

WORDML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
RSID_ATTRIBUTES = {"rsidR", "rsidRPr", "rsidDel"}
PROOF_TYPES = {"spellStart", "spellEnd", "gramStart", "gramEnd"}
PROPERTY_ATTRIBUTES = {name: {"val"} for name in (
    "rStyle", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
    "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish",
    "webHidden", "spacing", "w", "kern", "position", "sz", "szCs", "highlight",
    "effect", "vertAlign", "rtl", "cs", "em", "specVanish", "oMath",
)}
PROPERTY_ATTRIBUTES.update({
    "rFonts": {"ascii", "hAnsi", "eastAsia", "cs", "hint", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"},
    "color": {"val", "themeColor", "themeTint", "themeShade"},
    "u": {"val", "color", "themeColor", "themeTint", "themeShade"},
    "lang": {"val", "eastAsia", "bidi"},
})


def _elements(node, name=None):
    return [child for child in node.childNodes if child.nodeType == child.ELEMENT_NODE
            and (name is None or (child.namespaceURI == WORDML_NS and child.localName == name))]


def _attributes_allowed(node, allowed):
    return all(attr.namespaceURI == XMLNS_NS or (attr.namespaceURI, attr.localName) in allowed
               for attr in node.attributes.values())


def _children_allowed(node, allowed):
    for child in node.childNodes:
        if child.nodeType == child.ELEMENT_NODE:
            if (child.namespaceURI, child.localName) not in allowed:
                return False
        elif child.nodeType != child.TEXT_NODE or child.data.strip(" \t\r\n"):
            return False
    return True


def _plain_properties(properties):
    if not _attributes_allowed(properties, set()):
        return False
    if not _children_allowed(properties, {(WORDML_NS, name) for name in PROPERTY_ATTRIBUTES}):
        return False
    names = [child.localName for child in _elements(properties)]
    if len(names) != len(set(names)):
        return False
    for child in _elements(properties):
        if not _attributes_allowed(child, {(WORDML_NS, name) for name in PROPERTY_ATTRIBUTES[child.localName]}):
            return False
        if not _children_allowed(child, set()):
            return False
    return True


def _plain_run(run, blocked):
    allowed = {(WORDML_NS, name) for name in RSID_ATTRIBUTES} | {(XML_NS, "space")}
    if not _attributes_allowed(run, allowed):
        return False
    if not _children_allowed(run, {(WORDML_NS, "rPr"), (WORDML_NS, "t")}):
        return False
    properties = _elements(run, "rPr")
    if len(properties) > 1 or (properties and not _plain_properties(properties[0])):
        return False
    texts = _elements(run, "t")
    if not texts:
        return False
    for text in texts:
        if text in blocked or not _eligible(text):
            return False
        _preserves_space(text)  # Reject invalid xml:space rather than guess.
    return True


def _proof_error(node):
    return (node.namespaceURI == WORDML_NS and node.localName == "proofErr"
            and _attributes_allowed(node, {(WORDML_NS, "type")})
            and node.getAttributeNS(WORDML_NS, "type") in PROOF_TYPES
            and _children_allowed(node, set()))


def _paragraph_runs(paragraph, blocked):
    paragraph_attrs = {(WORDML_NS, name) for name in (
        "rsidR", "rsidRPr", "rsidDel", "rsidP", "rsidRDefault",
    )} | {(XML_NS, "space"), (W14_NS, "paraId"), (W14_NS, "textId")}
    if not _attributes_allowed(paragraph, paragraph_attrs):
        return []
    if not _children_allowed(paragraph, {(WORDML_NS, name) for name in ("pPr", "r", "proofErr")}):
        return []
    if len(_elements(paragraph, "pPr")) > 1:
        return []
    for properties in _elements(paragraph, "pPr"):
        for node in [properties, *properties.getElementsByTagName("*")]:
            if node.namespaceURI != WORDML_NS or any(
                attr.namespaceURI not in {WORDML_NS, XMLNS_NS, XML_NS}
                for attr in node.attributes.values()
            ):
                return []
    for name in ("pPrChange", "rPrChange", "sectPr"):
        if paragraph.getElementsByTagNameNS(WORDML_NS, name):
            return []
    if any(not _proof_error(node) for node in _elements(paragraph, "proofErr")):
        return []
    runs = _elements(paragraph, "r")
    return runs if runs and all(_plain_run(run, blocked) for run in runs) else []


def _format_key(run):
    properties = _elements(run, "rPr")
    return properties[0].toxml() if properties else None


def _combine(group):
    texts = [text for run in group for text in _elements(run, "t")]
    value = "".join(rendered_text(_text(text), _preserves_space(text)) for text in texts)
    target, first = group[0], texts[0]
    for child in list(first.childNodes):
        first.removeChild(child)
    if value:
        first.appendChild(first.ownerDocument.createTextNode(value))
    first.setAttributeNS(XML_NS, "xml:space", "preserve")
    for text in list(_elements(target, "t"))[1:]:
        target.removeChild(text)
    for run in group[1:]:
        run.parentNode.removeChild(run)
    if len(group) > 1:
        for name in RSID_ATTRIBUTES:
            if target.hasAttributeNS(WORDML_NS, name):
                target.removeAttributeNS(WORDML_NS, name)


def _check_directory(directory, extension):
    if not directory.is_dir() or directory.is_symlink():
        raise ValueError("Input must be an unpacked directory without symbolic links")
    paths = []
    total = 0
    for path in directory.rglob("*"):
        if path.is_symlink():
            raise ValueError("Symbolic links in unpacked packages are unsupported")
        paths.append(path)
        if len(paths) > MAX_MEMBERS:
            raise ValueError("Unpacked package exceeds the member limit")
        if path.is_file():
            size = path.stat().st_size
            total += size
            if size > MAX_PART_BYTES or total > MAX_PACKAGE_BYTES:
                raise ValueError("Unpacked package exceeds the size limit")
    files = {path.relative_to(directory).as_posix() for path in paths if path.is_file()}
    relation_types = []
    for path in paths:
        if path.is_file() and path.name.lower().endswith(".rels"):
            dom = parse_dom(read_xml(path))
            relation_types.extend(node.getAttribute("Type")
                                  for node in dom.getElementsByTagNameNS(REL_NS, "Relationship"))
    reject_signed_package(files, relation_types)
    validator = StructureValidator(directory, extension=extension)
    with contextlib.redirect_stdout(io.StringIO()):
        valid = validator.validate()
    if not valid:
        raise ValueError("Input failed basic structure checks: " + "; ".join(validator.errors[:5]))


def merge_runs(input_dir: str, *, extension=None) -> tuple[int, str]:
    """Merge supported runs; return controlled errors without modifying the input."""
    try:
        directory = Path(input_dir)
        if extension not in (None, ".docx", ".dotx"):
            raise ValueError("Unsupported document extension")
        _check_directory(directory, extension)
        document = directory / "word/document.xml"
        original = read_xml(document)
        dom = parse_dom(original)
        blocked = _blocked_anchors(dom)
        merged = 0
        changed = False
        for paragraph in dom.getElementsByTagNameNS(WORDML_NS, "p"):
            runs = _paragraph_runs(paragraph, blocked)
            if not runs:
                continue
            for marker in _elements(paragraph, "proofErr"):
                paragraph.removeChild(marker)
                changed = True
            index = 0
            while index < len(runs):
                end = index + 1
                while end < len(runs) and _format_key(runs[index]) == _format_key(runs[end]):
                    end += 1
                group = runs[index:end]
                if len(group) > 1 or len(_elements(group[0], "t")) > 1:
                    _combine(group)
                    merged += len(group) - 1
                    changed = True
                index = end
        if changed:
            atomic_write_bytes(document, _serialize_checked(dom))
        return merged, f"Merged {merged} runs"
    except (OSError, ValueError, RuntimeError, zipfile.BadZipFile) as exc:
        return 0, f"Error: {exc}"


def _merge_or_die(path, *, extension=None):
    _, message = merge_runs(str(path), extension=extension)
    if message.startswith("Error"):
        raise ValueError(message)
    return message


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Unpacked directory or DOCX/DOTX file")
    parser.add_argument("-o", "--output", type=Path,
                        help="Required output document for packaged input")
    args = parser.parse_args(argv)
    source = args.input
    try:
        if source.is_symlink():
            raise ValueError("Symbolic link inputs are unsupported")
        if source.is_dir():
            if args.output is not None:
                raise ValueError("Directory input is modified in place and takes no --output")
            print(_merge_or_die(source))
        elif source.is_file() and source.suffix.lower() in {".docx", ".dotx"}:
            if args.output is None:
                raise ValueError("Packaged input requires an explicit --output path")
            output = args.output
            same_path = output.resolve(strict=False) == source.resolve(strict=True)
            same_file = output.exists() and os.path.samefile(source, output)
            if same_path or same_file:
                raise ValueError("Output must be distinct from input")
            if output.suffix.lower() != source.suffix.lower():
                raise ValueError("Input and output must have the same DOCX/DOTX extension")
            if output.is_symlink():
                raise ValueError("Symbolic link outputs are unsupported")
            with tempfile.TemporaryDirectory(prefix="docx-merge-") as temporary:
                directory = Path(temporary)
                with open_archive(source) as package:
                    safe_extract(package, directory)
                message = _merge_or_die(directory, extension=source.suffix.lower())
                output.parent.mkdir(parents=True, exist_ok=True)
                rezip(directory, output)
            print(f"{message}; wrote {output}")
        else:
            raise ValueError("Input must be an unpacked directory or DOCX/DOTX file")
    except (OSError, ValueError, RuntimeError, zipfile.BadZipFile) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
