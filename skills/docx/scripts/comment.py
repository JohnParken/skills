"""Add comments to a DOCX document.

Accepts either an unpacked directory OR a .docx/.dotx file directly.
Directory edits are staged and published using a recoverable two-rename swap,
not one atomic multi-file operation; use a stable trusted task directory.

Usage:
    # Against an unpacked directory (writes satellite files in place)
    python comment.py unpacked/ "Comment text"
    python comment.py unpacked/ "Reply text" --parent 0

    # Against a .docx directly (extracts, writes satellite files, rezips)
    python comment.py contract.docx "This cap is too low" -o annotated.docx
    python comment.py contract.docx "Comment" --id 5      # explicit ID

The comment ID is auto-assigned (max existing + 1) unless --id is given.
Plain text is XML-escaped automatically; if you pass already-escaped text
(e.g. &amp;, &#x2019;) use --raw for escaped text only, never element markup.

After running, add markers to word/document.xml so the comment is visible:
  <w:commentRangeStart w:id="N"/>
  ... commented content ...
  <w:commentRangeEnd w:id="N"/>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="N"/></w:r>
"""

import argparse
import os
import random
import shutil
import stat
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from office.xmlio import parse_dom, read_xml
from office.archiveio import open_archive
from edit_docx import XMLNS_NS, _serialize_checked
from xml.parsers.expat import ExpatError
from xml.sax.saxutils import escape as xml_escape

from office.helpers import (
    MAX_MEMBERS, MAX_PACKAGE_BYTES, MAX_PART_BYTES, opc_target, reject_signed_package,
    rezip as _rezip, safe_extract as _safe_extract,
)

TEMPLATE_DIR = Path(__file__).resolve().parents[1] / "assets" / "comments"
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
}
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

COMMENT_XML = """\
<w:comment w:id="{id}" w:author="{author}" w:date="{date}" w:initials="{initials}">
  <w:p w14:paraId="{para_id}" w14:textId="77777777">
    <w:r>
      <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
      <w:annotationRef/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:color w:val="000000"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
      </w:rPr>
      <w:t xml:space="preserve">{text}</w:t>
    </w:r>
  </w:p>
</w:comment>"""

COMMENT_MARKER_TEMPLATE = """
Add to word/document.xml (markers must be direct children of w:p, never inside w:r):
  <w:commentRangeStart w:id="{cid}"/>
  <w:r>...</w:r>
  <w:commentRangeEnd w:id="{cid}"/>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{cid}"/></w:r>"""

REPLY_MARKER_TEMPLATE = """
Nest markers inside parent {pid}'s markers (direct children of w:p, never inside w:r):
  <w:commentRangeStart w:id="{pid}"/><w:commentRangeStart w:id="{cid}"/>
  <w:r>...</w:r>
  <w:commentRangeEnd w:id="{cid}"/><w:commentRangeEnd w:id="{pid}"/>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{pid}"/></w:r>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{cid}"/></w:r>"""

def _generate_hex_id() -> str:
    return f"{random.randint(0, 0x7FFFFFFE):08X}"


def _append_xml(xml_path: Path, root_tag: str, content: str) -> None:
    dom = parse_dom(read_xml(xml_path))
    prefix, local = root_tag.split(":", 1)
    roots = dom.getElementsByTagNameNS(NS[prefix], local)
    if not roots:
        raise ValueError(f"Missing expected comment root: {root_tag}")
    root = roots[0]
    for key, uri in NS.items():
        existing = root.getAttribute("xmlns:" + key)
        if existing and existing != uri:
            raise ValueError(f"Conflicting namespace prefix: {key}")
        root.setAttributeNS(XMLNS_NS, "xmlns:" + key, uri)
    ns_attrs = " ".join(f'xmlns:{k}="{v}"' for k, v in NS.items())
    wrapper_dom = parse_dom(f"<root {ns_attrs}>{content}</root>")
    for child in wrapper_dom.documentElement.childNodes:  
        if child.nodeType == child.ELEMENT_NODE:
            root.appendChild(dom.importNode(child, True))
    xml_path.write_bytes(_serialize_checked(dom))


def _find_para_id(comments_path: Path, comment_id: int) -> str | None:
    dom = parse_dom(read_xml(comments_path))
    for c in dom.getElementsByTagNameNS(NS["w"], "comment"):
        if c.getAttributeNS(NS["w"], "id") == str(comment_id):
            for p in c.getElementsByTagNameNS(NS["w"], "p"):
                if pid := p.getAttributeNS(NS["w14"], "paraId"):
                    return pid
    return None


def _next_comment_id(comments_path: Path) -> int:
    if not comments_path.exists():
        return 0
    dom = parse_dom(read_xml(comments_path))
    ids = []
    for c in dom.getElementsByTagNameNS(NS["w"], "comment"):
        value = c.getAttributeNS(NS["w"], "id")
        if not value.isascii() or not value.isdigit() or len(value) > 10 or int(value) > 2 ** 31 - 1:
            raise ValueError("Invalid existing comment ID")
        ids.append(int(value))
    if len(ids) != len(set(ids)):
        raise ValueError("Duplicate existing comment IDs")
    return (max(ids) + 1) if ids else 0


def _get_next_rid(rels_path: Path) -> int:
    dom = parse_dom(read_xml(rels_path))
    max_rid = 0
    for rel in dom.getElementsByTagNameNS(REL_NS, "Relationship"):
        rid = rel.getAttribute("Id")
        if rid and rid.startswith("rId"):
            try:
                max_rid = max(max_rid, int(rid[3:]))
            except ValueError:
                pass
    return max_rid + 1


def _has_relationship(rels_path: Path, target: str) -> bool:
    dom = parse_dom(read_xml(rels_path))
    return any(
        rel.getAttribute("Target") == target
        for rel in dom.getElementsByTagNameNS(REL_NS, "Relationship")
    )


def _has_content_type(ct_path: Path, part_name: str) -> bool:
    dom = parse_dom(read_xml(ct_path))
    return any(
        o.getAttribute("PartName") == part_name
        for o in dom.getElementsByTagNameNS(CT_NS, "Override")
    )


_COMMENT_RELS = [
    ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", "comments.xml"),
    ("http://schemas.microsoft.com/office/2011/relationships/commentsExtended", "commentsExtended.xml"),
    ("http://schemas.microsoft.com/office/2016/09/relationships/commentsIds", "commentsIds.xml"),
    ("http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible", "commentsExtensible.xml"),
]
_COMMENT_OVERRIDES = [
    ("/word/comments.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"),
    ("/word/commentsExtended.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"),
    ("/word/commentsIds.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"),
    ("/word/commentsExtensible.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"),
]


def _ensure_comment_relationships(unpacked_dir: Path) -> None:
    rels_path = unpacked_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        rels_path.parent.mkdir(parents=True, exist_ok=True)
        rels_path.write_text(f'<Relationships xmlns="{REL_NS}"/>', encoding="utf-8")
    dom = parse_dom(read_xml(rels_path))
    root = dom.documentElement
    comment_types = {rel_type for rel_type, _ in _COMMENT_RELS}
    existing = set()
    for rel in dom.getElementsByTagNameNS(REL_NS, "Relationship"):
        if rel.getAttribute("Type") not in comment_types:
            continue
        part = opc_target(
            rel.getAttribute("Target"),
            "word/document.xml",
            rel.getAttribute("TargetMode"),
        )
        if part is not None:
            existing.add(part)
    next_rid = _get_next_rid(rels_path)
    changed = False
    for rel_type, target in _COMMENT_RELS:
        if opc_target(target, "word/document.xml") in existing:
            continue
        tag = (root.prefix + ":" if root.prefix else "") + "Relationship"
        rel = dom.createElementNS(REL_NS, tag)
        rel.setAttribute("Id", f"rId{next_rid}")
        rel.setAttribute("Type", rel_type)
        rel.setAttribute("Target", target)
        root.appendChild(rel)  
        next_rid += 1
        changed = True
    if changed:
        rels_path.write_bytes(_serialize_checked(dom))


def _ensure_comment_content_types(unpacked_dir: Path) -> None:
    ct_path = unpacked_dir / "[Content_Types].xml"
    if not ct_path.exists():
        return
    dom = parse_dom(read_xml(ct_path))
    root = dom.documentElement
    existing = {
        o.getAttribute("PartName")
        for o in dom.getElementsByTagNameNS(CT_NS, "Override")
    }
    changed = False
    for part_name, content_type in _COMMENT_OVERRIDES:
        if part_name in existing:
            continue
        tag = (root.prefix + ":" if root.prefix else "") + "Override"
        override = dom.createElementNS(CT_NS, tag)
        override.setAttribute("PartName", part_name)
        override.setAttribute("ContentType", content_type)
        root.appendChild(override)  
        changed = True
    if changed:
        ct_path.write_bytes(_serialize_checked(dom))


def _unique_hex_id(used):
    used = {value.upper() for value in used}
    for _ in range(128):
        value = _generate_hex_id()
        if value.upper() not in used:
            return value
    raise ValueError("Unable to allocate a unique comment metadata ID")


def _preflight(unpacked_dir, text, author, initials, comment_id, parent_id, raw):
    if unpacked_dir.is_symlink() or not unpacked_dir.is_dir():
        raise ValueError("Input must be a directory without symbolic links")
    files = []
    total = count = 0
    for path in unpacked_dir.rglob("*"):
        mode = path.lstat().st_mode
        count += 1
        if stat.S_ISLNK(mode):
            raise ValueError("Symbolic links in the package are unsupported")
        if not (stat.S_ISREG(mode) or stat.S_ISDIR(mode)):
            raise ValueError("Only regular package files and directories are supported")
        if count > MAX_MEMBERS:
            raise ValueError("Package exceeds the directory entry limit")
        if stat.S_ISREG(mode):
            size = path.stat().st_size
            total += size
            if size > MAX_PART_BYTES or total > MAX_PACKAGE_BYTES:
                raise ValueError("Package exceeds size limits")
            files.append(path)
    relationship_types = []
    for path in files:
        if path.name.lower().endswith(".rels"):
            dom = parse_dom(read_xml(path))
            relationship_types.extend(node.getAttribute("Type")
                                      for node in dom.getElementsByTagNameNS(REL_NS, "Relationship"))
    reject_signed_package((path.relative_to(unpacked_dir).as_posix() for path in files), relationship_types)
    for label, value in (("text", text), ("author", author), ("initials", initials)):
        if not isinstance(value, str) or any(
            ord(char) < 32 or 0xD800 <= ord(char) <= 0xDFFF or char in "\ufffe\uffff" for char in value
        ):
            raise ValueError(f"{label} must be single-line XML text")
        if label == "author" and not value:
            raise ValueError("author must be nonempty")
    for label, value in (("comment_id", comment_id), ("parent_id", parent_id)):
        if value is not None and (type(value) is not int or not 0 <= value <= 2 ** 31 - 1):
            raise ValueError(f"{label} must be a nonnegative 32-bit integer")
    if raw:
        wrapper = parse_dom("<text>" + text + "</text>").documentElement
        if any(child.nodeType != child.TEXT_NODE for child in wrapper.childNodes):
            raise ValueError("raw accepts escaped text only, not markup")
        if any(ord(char) < 32 for child in wrapper.childNodes for char in child.data):
            raise ValueError("raw text must be single-line")
    para_ids, durable_ids = set(), set()
    comment_roots = {"word/comments.xml": (NS["w"], "comments"),
                     "word/commentsExtended.xml": (NS["w15"], "commentsEx"),
                     "word/commentsIds.xml": (NS["w16cid"], "commentsIds"),
                     "word/commentsExtensible.xml": (NS["w16cex"], "commentsExtensible")}
    expected_roots = {**comment_roots, "[Content_Types].xml": (CT_NS, "Types"),
                      "word/_rels/document.xml.rels": (REL_NS, "Relationships"),
                      "word/document.xml": (NS["w"], "document")}
    for relative in ("[Content_Types].xml", "word/document.xml"):
        if not (unpacked_dir / relative).is_file():
            raise ValueError(f"Missing required package XML: {relative}")
    for relative in ("[Content_Types].xml", "word/_rels/document.xml.rels",
                     "word/comments.xml", "word/commentsExtended.xml", "word/commentsIds.xml",
                     "word/commentsExtensible.xml", "word/document.xml"):
        path = unpacked_dir / relative
        if not path.is_file():
            continue
        dom = parse_dom(read_xml(path))
        _serialize_checked(dom)
        if (dom.documentElement.namespaceURI, dom.documentElement.localName) != expected_roots[relative]:
            raise ValueError(f"Invalid package XML root: {relative}")
        if relative in comment_roots:
            for prefix, uri in NS.items():
                existing = dom.documentElement.getAttribute("xmlns:" + prefix)
                if existing and existing != uri:
                    raise ValueError(f"Conflicting comment namespace prefix: {prefix}")
        for node in dom.getElementsByTagName("*"):
            for attr in node.attributes.values():
                if attr.localName in {"paraId", "paraIdParent"} and attr.namespaceURI in NS.values():
                    para_ids.add(attr.value)
                if attr.localName == "durableId" and attr.namespaceURI in NS.values():
                    durable_ids.add(attr.value)
    return para_ids, durable_ids


def _add_comment_inplace(
    unpacked_dir: Path | str,
    text: str,
    comment_id: int | None = None,
    author: str = "Claude",
    initials: str = "C",
    parent_id: int | None = None,
    raw: bool = False,
) -> tuple[int, str, str]:
    unpacked_dir = Path(unpacked_dir)
    para_ids, durable_ids = _preflight(unpacked_dir, text, author, initials, comment_id, parent_id, raw)
    if not raw:
        text = xml_escape(text)
    author = xml_escape(author, {'"': "&quot;"})
    initials = xml_escape(initials, {'"': "&quot;"})
    word = unpacked_dir / "word"
    if not word.exists():
        raise FileNotFoundError(f"{word} not found (not an unpacked .docx?)")

    comments = word / "comments.xml"
    next_id = _next_comment_id(comments)
    if comment_id is None:
        comment_id = next_id
    if comment_id > 2 ** 31 - 1:
        raise ValueError("Comment IDs exhausted")
    if comments.exists():
        dom = parse_dom(read_xml(comments))
        if any(int(node.getAttributeNS(NS["w"], "id")) == comment_id
               for node in dom.getElementsByTagNameNS(NS["w"], "comment")):
            raise ValueError("Comment ID already exists")

    parent_para = None
    if parent_id is not None:
        parent_para = _find_para_id(comments, parent_id) if comments.exists() else None
        if not parent_para:
            raise ValueError(f"parent comment {parent_id} not found")

    para_id, durable_id = _unique_hex_id(para_ids), _unique_hex_id(durable_ids)
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    if not comments.exists():
        shutil.copy(TEMPLATE_DIR / "comments.xml", comments)
    _ensure_comment_relationships(unpacked_dir)
    _ensure_comment_content_types(unpacked_dir)
    _append_xml(
        comments,
        "w:comments",
        COMMENT_XML.format(
            id=comment_id, author=author, date=ts, initials=initials,
            para_id=para_id, text=text,
        ),
    )

    ext = word / "commentsExtended.xml"
    if not ext.exists():
        shutil.copy(TEMPLATE_DIR / "commentsExtended.xml", ext)
    if parent_para is not None:
        _append_xml(
            ext, "w15:commentsEx",
            f'<w15:commentEx w15:paraId="{para_id}" w15:paraIdParent="{parent_para}" w15:done="0"/>',
        )
    else:
        _append_xml(
            ext, "w15:commentsEx",
            f'<w15:commentEx w15:paraId="{para_id}" w15:done="0"/>',
        )

    ids = word / "commentsIds.xml"
    if not ids.exists():
        shutil.copy(TEMPLATE_DIR / "commentsIds.xml", ids)
    _append_xml(
        ids, "w16cid:commentsIds",
        f'<w16cid:commentId w16cid:paraId="{para_id}" w16cid:durableId="{durable_id}"/>',
    )

    extensible = word / "commentsExtensible.xml"
    if not extensible.exists():
        shutil.copy(TEMPLATE_DIR / "commentsExtensible.xml", extensible)
    _append_xml(
        extensible, "w16cex:commentsExtensible",
        f'<w16cex:commentExtensible w16cex:durableId="{durable_id}" w16cex:dateUtc="{ts}"/>',
    )

    action = "reply" if parent_id is not None else "comment"
    return comment_id, para_id, f"Added {action} id={comment_id} (paraId={para_id})"


def add_comment(
    unpacked_dir: Path | str,
    text: str,
    comment_id: int | None = None,
    author: str = "Claude",
    initials: str = "C",
    parent_id: int | None = None,
    raw: bool = False,
) -> tuple[int, str, str]:
    """Stage all parts before publishing a recoverable directory replacement.

    This is two directory renames, not a single atomic multi-file transaction.
    Catchable install failures restore the backup; an interrupted process or a
    failed restoration may leave the original under the reported backup path.
    Use a trusted, stable task directory without concurrent writers/readers.
    """
    source = Path(unpacked_dir)
    _preflight(source, text, author, initials, comment_id, parent_id, raw)
    source = source.resolve()
    if Path.cwd().resolve().is_relative_to(source):
        raise ValueError("Comment input must not contain the current working directory; use a separate unpacked directory")
    transaction = Path(tempfile.mkdtemp(prefix=f".{source.name}.comment-", dir=source.parent))
    planned, backup = transaction / "planned", transaction / "backup"
    installed = keep_transaction = False
    retained = None
    try:
        shutil.copytree(source, planned)
        result = _add_comment_inplace(planned, text, comment_id, author, initials, parent_id, raw)
        _preflight(planned, text, author, initials, comment_id, parent_id, raw)
        try:
            os.replace(source, backup)
            os.replace(planned, source)
            installed = True
        except BaseException as exc:
            if backup.exists():
                if source.exists():
                    raise RuntimeError(f"Comment install outcome requires inspection; original backup retained at {backup}") from exc
                try:
                    os.replace(backup, source)
                except BaseException as restore_error:
                    if backup.exists():
                        raise RuntimeError(f"Comment rollback failed; restore the original directory from {backup}") from restore_error
                    if source.is_dir():
                        # The restore completed before the late error. Do not
                        # invent a missing recovery path or hide the install failure.
                        raise exc from restore_error
                    keep_transaction = True
                    raise RuntimeError(f"Comment rollback outcome requires inspection at {source} and {transaction}") from restore_error
            raise
        try:
            shutil.rmtree(backup)
        except OSError:
            if backup.exists():
                retained = backup
    finally:
        # Never automatically delete the original after an uncertain outcome or
        # failed rollback. Its location is included in the error/result above.
        if not keep_transaction and not backup.exists():
            try:
                shutil.rmtree(transaction)
            except OSError:
                if transaction.exists():
                    retained = transaction
                    if not installed:
                        print(f"Warning: comment cleanup retained at {transaction}", file=sys.stderr)
    if retained is not None:
        kind = "backup " if retained == backup else ""
        return result[0], result[1], result[2] + f"; installed, cleanup {kind}retained at {retained}"
    return result


def main() -> None:
    p = argparse.ArgumentParser(description="Add a comment to a DOCX (directory or .docx file).")
    p.add_argument("input", help="Unpacked DOCX directory OR a .docx/.dotx file")
    p.add_argument("text", help="Comment text (plain text; XML-escaped automatically)")
    p.add_argument("--raw", action="store_true",
                   help="Treat text as pre-escaped XML (skip automatic escaping)")
    p.add_argument("--id", type=int, dest="comment_id",
                   help="Comment ID (default: auto-assign as max existing + 1)")
    p.add_argument("--author", default="Claude", help="Author name")
    p.add_argument("--initials", default="C", help="Author initials")
    p.add_argument("--parent", type=int, help="Parent comment ID (makes this a reply)")
    p.add_argument("-o", "--output",
                   help="Output .docx path (only used when input is a .docx; default: overwrite input)")
    args = p.parse_args()

    src = Path(args.input)

    try:
        if src.is_dir():
            if args.output:
                print("Warning: --output ignored for directory input", file=sys.stderr)
            cid, _, msg = add_comment(
                src, args.text, comment_id=args.comment_id,
                author=args.author, initials=args.initials,
                parent_id=args.parent, raw=args.raw,
            )
            print(msg)
        elif src.is_file() and src.suffix.lower() in (".docx", ".dotx"):
            out = Path(args.output) if args.output else src
            if out.is_symlink() or out.suffix.lower() != src.suffix.lower():
                raise ValueError("Comment output must have the same DOCX/DOTX extension and not be a symbolic link")
            with tempfile.TemporaryDirectory() as tmp:
                tmp_path = Path(tmp)
                with open_archive(src) as zf:
                    _safe_extract(zf, tmp_path)
                cid, _, msg = add_comment(
                    tmp_path, args.text, comment_id=args.comment_id,
                    author=args.author, initials=args.initials,
                    parent_id=args.parent, raw=args.raw,
                )
                _rezip(tmp_path, out)
            print(msg)
            print(f"Wrote {out} (comment defined; add markers to word/document.xml to make it visible)")
        else:
            print(f"Error: {src} is neither a directory nor a .docx/.dotx file", file=sys.stderr)
            sys.exit(1)
    except (OSError, ValueError, RuntimeError, zipfile.BadZipFile, ExpatError) as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    if args.parent is not None:
        print(REPLY_MARKER_TEMPLATE.format(pid=args.parent, cid=cid))
    else:
        print(COMMENT_MARKER_TEMPLATE.format(cid=cid))


if __name__ == "__main__":
    main()
