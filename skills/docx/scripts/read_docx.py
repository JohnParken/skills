"""Extract readable text from a DOCX/DOTX package without office binaries."""

from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path
from xml.etree.ElementTree import Element

from defusedxml import ElementTree as SafeET
from defusedxml.common import DefusedXmlException

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_REVISION_TAGS = {
    "accepted": {"del", "moveFrom"},
    "original": {"ins", "moveTo"},
    "all": set(),
}
_BLOCK_WRAPPERS = {"sdt", "sdtContent", "customXml", "ins", "del", "moveFrom", "moveTo", "smartTag"}


def _tag(local: str) -> str:
    return f"{{{W}}}{local}"


def _paragraph_text(p: Element, revisions: str) -> str:
    out: list[str] = []

    def visit(node: Element, blocked: bool = False) -> None:
        name = node.tag.rsplit("}", 1)[-1]
        blocked = blocked or name in _REVISION_TAGS[revisions]
        if blocked or name == "txbxContent":
            return
        if name in ("t", "delText"):
            if node.text:
                out.append(node.text)
            return
        if name == "tab":
            out.append("\t")
            return
        if name in ("br", "cr"):
            out.append("\n")
            return
        for child in node:
            # Nested paragraphs belong to separate blocks and are handled by block traversal.
            if child.tag == _tag("p"):
                continue
            visit(child, blocked)

    visit(p)
    return "".join(out)


def _blocks(node: Element, revisions: str) -> list[str]:
    result: list[str] = []
    for child in node:
        name = child.tag.rsplit("}", 1)[-1]
        if child.tag == _tag("p"):
            result.append(_paragraph_text(child, revisions))
        elif child.tag == _tag("tbl"):
            result.extend(_table(child, revisions))
        elif name in _BLOCK_WRAPPERS:
            if name not in _REVISION_TAGS[revisions]:
                result.extend(_blocks(child, revisions))
    return result


def _table(tbl: Element, revisions: str) -> list[str]:
    rows: list[str] = []
    for tr in tbl.findall(_tag("tr")):
        cells = []
        for tc in tr.findall(_tag("tc")):
            cells.append("\n".join(_blocks(tc, revisions)))
        rows.append("\t".join(cells))
    return rows


def _part_text(root: Element, part: str, revisions: str) -> str:
    if part.endswith("footnotes.xml") or part.endswith("endnotes.xml"):
        kind = "footnote" if part.endswith("footnotes.xml") else "endnote"
        elements = [e for e in root if e.tag == _tag(kind) and e.get(_tag("type")) not in ("separator", "continuationSeparator")]
        chunks = []
        for e in elements:
            blocks = _blocks(e, revisions)
            if blocks:
                chunks.append("\n".join(blocks))
        return "\n".join(chunks)
    if part.endswith("comments.xml"):
        chunks = []
        for e in root.findall(_tag("comment")):
            blocks = _blocks(e, revisions)
            if blocks:
                chunks.append("\n".join(blocks))
        return "\n".join(chunks)
    return "\n".join(_blocks(root, revisions))


def read_docx(path: str | Path, *, revisions: str = "accepted", include_related: bool = False) -> str:
    """Return document text in reading order; this does not accept revisions or reproduce layout."""
    if revisions not in _REVISION_TAGS:
        raise ValueError("revisions must be 'accepted', 'original', or 'all'")
    try:
        with zipfile.ZipFile(path) as package:
            try:
                document = SafeET.fromstring(package.read("word/document.xml"))
            except KeyError as exc:
                raise ValueError("DOCX package is missing word/document.xml") from exc
            if document.tag != _tag("document"):
                raise ValueError("word/document.xml does not contain a w:document root")
            body = document.find(_tag("body"))
            if body is None:
                raise ValueError("word/document.xml is missing w:body")
            pieces = _blocks(body, revisions)
            if include_related:
                names = package.namelist()
                categories = (
                    ("Headers", "word/header"),
                    ("Footers", "word/footer"),
                    ("Footnotes", "word/footnotes.xml"),
                    ("Endnotes", "word/endnotes.xml"),
                    ("Comments", "word/comments.xml"),
                )
                for label, prefix in categories:
                    for part in sorted(n for n in names if n.startswith(prefix) and n.endswith(".xml")):
                        root = SafeET.fromstring(package.read(part))
                        text = _part_text(root, part, revisions)
                        if text.strip():
                            pieces.extend((f"[{label}: {part}]", text))
    except (SafeET.ParseError, DefusedXmlException) as exc:
        raise ValueError(f"invalid or unsafe XML in DOCX package: {exc}") from exc
    except (zipfile.BadZipFile, OSError) as exc:
        raise ValueError(f"cannot read DOCX package: {exc}") from exc
    return "\n".join(pieces)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="input .docx or .dotx file")
    parser.add_argument("-o", "--output", type=Path)
    parser.add_argument("--revisions", choices=tuple(_REVISION_TAGS), default="accepted")
    parser.add_argument("--include-related", action="store_true")
    args = parser.parse_args(argv)
    if args.input.suffix.lower() not in (".docx", ".dotx"):
        parser.error("input must have a .docx or .dotx extension")
    try:
        if args.output and args.output.exists():
            try:
                same_file = args.input.samefile(args.output)
            except OSError:
                same_file = args.input.resolve() == args.output.resolve()
            if same_file:
                parser.error("output must not refer to the input file")
        text = read_docx(args.input, revisions=args.revisions, include_related=args.include_related)
        if args.output:
            args.output.write_text(text, encoding="utf-8")
        else:
            sys.stdout.write(text)
    except (ValueError, OSError, zipfile.BadZipFile) as exc:
        parser.error(str(exc))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
