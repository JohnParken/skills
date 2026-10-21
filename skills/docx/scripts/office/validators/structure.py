"""Basic OPC and WordprocessingML checks; deliberately not a schema validator."""

from collections import Counter
from pathlib import Path

try:
    from .. import xmlio
    from ..helpers import MAX_MEMBERS, MAX_PACKAGE_BYTES, MAX_PART_BYTES, opc_target
except ImportError:
    import xmlio
    from helpers import MAX_MEMBERS, MAX_PACKAGE_BYTES, MAX_PART_BYTES, opc_target

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
MAIN_TYPES = {
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    ".dotx": "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
}
REQUIRED = ("[Content_Types].xml", "_rels/.rels", "word/document.xml")


def w(name):
    return f"{{{W}}}{name}"


class StructureValidator:
    def __init__(self, directory, *, extension=None, require_no_placeholders=False):
        self.directory = Path(directory)
        self.extension = extension
        self.require_no_placeholders = require_no_placeholders
        self.errors = []
        self.roots = {}
        self.files = set()
        self.relations = {}
        self.content_types = {}

    def error(self, message):
        self.errors.append(message)

    def validate(self):
        self.errors = []
        self.roots = {}
        self.files = set()
        self.relations = {}
        self.content_types = {}
        try:
            self._load()
            if not self.errors:
                self._types()
                self._relationships()
                self._word()
        except (OSError, ValueError) as exc:
            self.error(str(exc))
        for message in self.errors:
            print(f"FAILED - {message}")
        return not self.errors

    def _load(self):
        if not self.directory.is_dir() or self.directory.is_symlink():
            raise ValueError("Input must be a directory without symbolic links")
        total = count = 0
        for path in self.directory.rglob("*"):
            if path.is_symlink():
                raise ValueError(f"Symbolic link in package: {path.name}")
            if not path.is_file():
                continue
            count += 1
            size = path.stat().st_size
            total += size
            if count > MAX_MEMBERS or size > MAX_PART_BYTES or total > MAX_PACKAGE_BYTES:
                raise ValueError("Package exceeds size or member limits")
            name = path.relative_to(self.directory).as_posix()
            self.files.add(name)
            if name.lower().endswith((".xml", ".rels")):
                try:
                    self.roots[name] = xmlio.parse(path).getroot()
                except ValueError as exc:
                    self.error(f"{name}: {exc}")
        for name in REQUIRED:
            if name not in self.roots:
                self.error(f"Missing or invalid required XML part: {name}")

    def _types(self):
        root = self.roots["[Content_Types].xml"]
        if root.tag != f"{{{CT}}}Types":
            self.error("Invalid content types root")
            return
        defaults, overrides = {}, {}
        for node in root:
            kind = node.tag
            content_type = node.get("ContentType", "")
            if kind == f"{{{CT}}}Default":
                key = node.get("Extension", "").lower()
                if not key or key in defaults or not content_type:
                    self.error("Empty or duplicate content type default")
                defaults[key] = content_type
            elif kind == f"{{{CT}}}Override":
                part = node.get("PartName", "")
                if not part.startswith("/") or not content_type:
                    self.error("Invalid content type override")
                    continue
                key = opc_target(part, "")
                if key in overrides or key not in self.files:
                    self.error(f"Duplicate or missing content type part: {part}")
                overrides[key] = content_type
            else:
                self.error("Unexpected content type element")
        for name in self.files - {"[Content_Types].xml"}:
            value = overrides.get(name) or defaults.get(name.rsplit(".", 1)[-1].lower())
            if not value:
                self.error(f"Missing content type for {name}")
            self.content_types[name] = value
            if name.endswith(".rels") and value != "application/vnd.openxmlformats-package.relationships+xml":
                self.error(f"Invalid relationships content type: {name}")
        main_type = self.content_types.get("word/document.xml")
        if self.extension:
            if main_type != MAIN_TYPES[self.extension]:
                self.error("Main content type does not match DOCX/DOTX extension")
        elif main_type not in MAIN_TYPES.values():
            self.error("Unsupported main document content type")

    def _relationships(self):
        for name, root in self.roots.items():
            if not name.endswith(".rels"):
                continue
            if root.tag != f"{{{REL}}}Relationships":
                self.error(f"Invalid relationships root: {name}")
                continue
            path = Path(name)
            if name == "_rels/.rels":
                owner = ""
            elif path.parent.name == "_rels":
                owner = (path.parent.parent / path.name[:-5]).as_posix()
                if owner not in self.files:
                    self.error(f"Relationship source is missing: {owner}")
            else:
                self.error(f"Invalid relationships part location: {name}")
                continue
            rels = {}
            for node in root:
                rid, kind, target = node.get("Id"), node.get("Type"), node.get("Target")
                mode = node.get("TargetMode", "")
                if (
                    node.tag != f"{{{REL}}}Relationship" or not rid or rid in rels
                    or not kind or not target or mode not in ("", "Internal", "External")
                ):
                    self.error(f"Invalid or duplicate relationship: {name}: {rid}")
                    continue
                resolved = opc_target(target, owner, mode)
                if resolved is not None and resolved not in self.files:
                    self.error(f"Missing relationship target: {name}: {resolved}")
                rels[rid] = (kind, resolved)
            self.relations[owner] = rels
        mains = [target for kind, target in self.relations.get("", {}).values()
                 if kind == R + "/officeDocument"]
        if mains != ["word/document.xml"]:
            self.error("Package must have one internal officeDocument relationship to word/document.xml")
        for name, root in self.roots.items():
            ids = self.relations.get(name, {})
            for node in root.iter():
                for attr, value in node.attrib.items():
                    if attr in (f"{{{R}}}id", f"{{{R}}}embed", f"{{{R}}}link") and value not in ids:
                        self.error(f"Unresolved relationship reference {value} in {name}")

    def _word(self):
        root = self.roots["word/document.xml"]
        body = root.find(w("body"))
        if root.tag != w("document") or body is None or len(root.findall(w("body"))) != 1:
            self.error("Main document must contain one w:body in w:document")
            return
        sections = body.findall(w("sectPr"))
        if len(sections) > 1 or (sections and body[-1] is not sections[0]):
            self.error("Body-level sectPr must be unique and last")
        if any(child.tag in (w("r"), w("t"), w("tr"), w("tc")) for child in body):
            self.error("Invalid body child: place runs in paragraphs and cells in tables")
        styles = self.roots.get("word/styles.xml")
        style_ids = set()
        if styles is not None:
            if styles.tag != w("styles"):
                self.error("Invalid styles root")
            for style in styles.findall(w("style")):
                sid = style.get(w("styleId"))
                if not sid or sid in style_ids:
                    self.error("Empty or duplicate style ID")
                style_ids.add(sid)
        numbering = self.roots.get("word/numbering.xml")
        num_ids, abstract_ids = set(), set()
        if numbering is not None:
            if numbering.tag != w("numbering"):
                self.error("Invalid numbering root")
            for tag, key, seen in (("abstractNum", "abstractNumId", abstract_ids), ("num", "numId", num_ids)):
                for node in numbering.findall(w(tag)):
                    value = node.get(w(key))
                    if not value or not value.isdigit() or value in seen:
                        self.error(f"Invalid or duplicate {key}")
                    seen.add(value)
            for num in numbering.findall(w("num")):
                abstract = num.find(w("abstractNumId"))
                if abstract is None or abstract.get(w("val")) not in abstract_ids:
                    self.error("Numbering references missing abstractNum")
        doc_rels = set(self.relations.get("word/document.xml", {}).values())
        for part, kind in (("word/styles.xml", "styles"), ("word/numbering.xml", "numbering"),
                           ("word/comments.xml", "comments")):
            if part in self.files and (R + "/" + kind, part) not in doc_rels:
                self.error(f"Missing document relationship for {part}")
        for name, part in self.roots.items():
            if not name.startswith("word/") or name.endswith(".rels"):
                continue
            revision_ids = set()
            parents = {child: parent for parent in part.iter() for child in parent}
            for node in part.iter():
                if node.tag in (w("bold"), w("italic")):
                    self.error(f"Invalid run property name; use w:b/w:i in {name}")
                if node.tag in (w("t"), w("delText")):
                    parent = parents.get(node)
                    if parent is None or parent.tag != w("r"):
                        self.error(f"Text must be inside w:r in {name}")
                if node.tag in {w("pStyle"), w("rStyle"), w("tblStyle"), w("basedOn"), w("next"), w("link")}:
                    if node.get(w("val")) not in style_ids:
                        self.error(f"Unknown style {node.get(w('val'))} in {name}")
                if node.tag == w("numId") and node.get(w("val")) not in num_ids | {"0"}:
                    self.error(f"Unknown numbering ID in {name}")
                if node.tag in (w("ins"), w("del")):
                    value = node.get(w("id"), "")
                    if not value.isdigit() or value in revision_ids or not node.get(w("author")):
                        self.error(f"Invalid/duplicate revision ID or missing author in {name}")
                    revision_ids.add(value)
            if self.require_no_placeholders:
                for paragraph in part.iter(w("p")):
                    text = "".join(n.text or "" for n in paragraph.iter(w("t")))
                    if "{{" in text or "}}" in text:
                        self.error(f"Unfilled placeholder in {name}")
                        break
        self._comments()

    def _comments(self):
        part = self.roots.get("word/comments.xml")
        definitions = []
        if part is not None:
            if part.tag != w("comments"):
                self.error("Invalid comments root")
            definitions = [node.get(w("id")) for node in part.findall(w("comment"))]
            if any(not value or not value.isdigit() for value in definitions) or len(set(definitions)) != len(definitions):
                self.error("Invalid or duplicate comment IDs")
        ids = set(definitions)
        referenced = set()
        owners = {}
        for name, root in self.roots.items():
            if not name.startswith("word/") or name == "word/comments.xml":
                continue
            counters = {kind: Counter() for kind in ("commentRangeStart", "commentRangeEnd", "commentReference")}
            open_ranges = set()
            for node in root.iter():
                kind = node.tag.removeprefix("{" + W + "}")
                if kind not in counters:
                    continue
                value = node.get(w("id"))
                counters[kind][value] += 1
                referenced.add(value)
                if value not in ids:
                    self.error(f"Unknown comment {value} in {name}")
                if kind == "commentRangeStart":
                    open_ranges.add(value)
                elif kind == "commentRangeEnd":
                    if value not in open_ranges:
                        self.error(f"Comment range ends before start: {value} in {name}")
                    open_ranges.discard(value)
            union = set().union(*(c.keys() for c in counters.values()))
            for value in union:
                if value in owners and owners[value] != name:
                    self.error(f"Comment {value} is anchored in multiple parts")
                owners[value] = name
                # This subset checks range comments, not point comments.
                if any(c[value] != 1 for c in counters.values()):
                    self.error(f"Unpaired or duplicate comment anchors: {value} in {name}")
        for value in ids - referenced:
            self.error(f"Comment {value} has no document anchor")
