#!/usr/bin/env python3
"""Read-only, standard-library checks for a skill directory."""
from __future__ import annotations

import argparse
import ast
import json
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import unquote, urlsplit


@dataclass
class Issue:
    level: str
    rule: str
    path: str
    message: str


ALLOWED_ROOT = {"SKILL.md", "requirements.txt", "scripts", "references", "assets"}
SUSPICIOUS = {"readme", "readme.md", "cache", "__pycache__", "log", "logs", "tmp", "temp"}
DIRECTORY_NAME = re.compile(r"[A-Za-z_](?:[A-Za-z_-]*[A-Za-z_])?")
TEMP_SUFFIXES = {".pyc", ".pyo", ".log", ".tmp", ".bak"}


def _without_comment(value: str) -> str:
    """Remove an inline YAML comment without changing quoted text."""
    quote = None
    escaped = False
    for index, char in enumerate(value):
        if escaped:
            escaped = False
        elif quote == '"' and char == "\\":
            escaped = True
        elif char == quote:
            quote = None
        elif quote is None and char in "'\"" and index == 0:
            quote = char
        elif quote is None and char == "#" and (index == 0 or value[index - 1].isspace()):
            return value[:index].rstrip()
    return value


def _unquote(value: str) -> tuple[str, bool]:
    """Decode a limited quoted scalar; return (value, supported)."""
    if not value.startswith(("'", '"')):
        if (
            value.lower() in {"", "null", "~", "true", "false", "yes", "no", "on", "off", ".inf", ".nan"}
            or re.fullmatch(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[-+]?\d+)?", value, re.I)
            or value.startswith(("[", "{", "&", "*", "!", "|", ">", "@", "`"))
            or ": " in value
        ):
            return value, False
        return value, True
    quote = value[0]
    if len(value) < 2 or value[-1] != quote:
        return value, False
    inner = value[1:-1]
    if quote == "'":
        if not re.fullmatch(r"(?:[^']|'')*", inner):
            return inner, False
        return inner.replace("''", "'"), True
    # JSON quoted strings form a useful, deliberately limited YAML subset.
    try:
        decoded = json.loads(value)
    except (ValueError, TypeError):
        return inner, False
    return decoded, True


def _frontmatter(text: str, rel: str, issues: list[Issue]) -> tuple[dict[str, object], int | None]:
    lines = text.splitlines()
    if not lines or lines[0] != "---":
        issues.append(Issue("error", "frontmatter", rel, "Expected exact --- delimiter lines around YAML frontmatter."))
        return {}, None
    try:
        end = lines.index("---", 1)
    except ValueError:
        issues.append(Issue("error", "frontmatter", rel, "Frontmatter has no exact closing --- delimiter."))
        return {}, None
    fields: dict[str, object] = {}
    seen: set[str] = set()
    invalid_fields: set[str] = set()
    supported = True
    i = 1
    while i < end:
        line = lines[i]
        if not line.strip() or line.lstrip().startswith("#"):
            i += 1
            continue
        m = re.fullmatch(r"([A-Za-z_][\w.-]*):(?:[ \t]*(.*))?", line)
        if not m:
            supported = False
            i += 1
            continue
        key, raw = m.group(1), _without_comment((m.group(2) or "").strip())
        if key in seen:
            issues.append(Issue("error", "duplicate-frontmatter", rel, f"Frontmatter key '{key}' is repeated."))
        seen.add(key)
        if key == "dependency":
            if raw:
                issues.append(Issue("error", "python-dependencies", rel, "dependency must be a mapping containing python as a list."))
                i += 1
                continue
            i += 1
            if i >= end or not re.fullmatch(r"  python:\s*", lines[i]):
                issues.append(Issue("error", "python-dependencies", rel, "dependency must contain a python list."))
                continue
            i += 1
            values: list[str] = []
            while i < end and re.match(r"^    -(?:\s|$)", lines[i]):
                item_raw = _without_comment(re.sub(r"^    -\s*", "", lines[i]).strip())
                item, ok = _unquote(item_raw)
                if not ok:
                    supported = False
                    if not item_raw.startswith(("[", "{", "&", "*", "!")):
                        issues.append(Issue("error", "python-dependencies", rel, "dependency.python entries must be supported nonempty strings."))
                if not item.strip():
                    issues.append(Issue("error", "python-dependencies", rel, "dependency.python entries must be nonempty strings."))
                values.append(item)
                i += 1
            if not values:
                issues.append(Issue("error", "python-dependencies", rel, "dependency.python must be a nonempty list."))
            fields["dependency.python"] = values
            continue
        if raw in {"|", ">", "|-", ">-"} or raw.startswith(("[", "{", "&", "*", "!")):
            supported = False
            i += 1
            continue
        value, ok = _unquote(raw)
        supported &= ok
        if not ok:
            invalid_fields.add(key)
        if raw.lower() in {"null", "~", "true", "false"} or re.fullmatch(r"[-+]?\d+(?:\.\d+)?", raw):
            value = None
        fields[key] = value
        i += 1
    if not supported:
        issues.append(Issue("warning", "yaml-manual-review", rel, "Frontmatter uses YAML syntax this checker does not parse; review manually."))
    for key in ("name", "description"):
        if key in invalid_fields or not isinstance(fields.get(key), str) or not fields[key].strip():
            issues.append(Issue("error", "required-frontmatter", rel, f"Required nonempty string field '{key}' is missing or not a supported string."))
    desc = fields.get("description")
    if isinstance(desc, str) and not 100 <= len(desc) <= 150:
        issues.append(Issue("warning", "description-length", rel, f"Description is {len(desc)} characters; recommended length is 100–150."))
    return fields, end


def _outside_fences(text: str) -> str:
    lines, fence = [], None
    for line in text.splitlines():
        match = re.match(r"^\s*(`{3,}|~{3,})(.*)$", line)
        if match:
            marker = match.group(1)
            if fence is None:
                fence = marker
            elif marker[0] == fence[0] and len(marker) >= len(fence) and not match.group(2).strip():
                fence = None
            lines.append("")
        else:
            lines.append("" if fence else line)
    return "\n".join(lines)


def _walk(skill: Path, issues: list[Issue]):
    """Collect regular files once, without traversing nested symlinks."""
    collected = []

    def on_error(exc):
        issues.append(Issue("error", "unreadable", str(exc.filename), f"Cannot read directory: {exc}"))

    for current, dirs, files in os.walk(skill, followlinks=False, onerror=on_error):
        here = Path(current)
        if here != skill and not dirs and not files:
            issues.append(Issue("warning", "empty-directory", here.relative_to(skill).as_posix(), "Directory is empty."))
        for entry in list(dirs):
            path = here / entry
            rel = path.relative_to(skill).as_posix()
            if path.is_symlink():
                issues.append(Issue("error", "symlink", rel, "Nested symlink directory is not traversed."))
                dirs.remove(entry)
            elif entry.lower() in SUSPICIOUS:
                issues.append(Issue("warning", "suspicious-entry", rel, "Empty/cache/log/tmp/README directories are discouraged."))
            elif not DIRECTORY_NAME.fullmatch(entry):
                issues.append(Issue("error", "directory-name", rel, "Use letters, underscores, or internal hyphens in directory names."))
        for filename in files:
            path = here / filename
            rel = path.relative_to(skill).as_posix()
            if path.is_symlink():
                issues.append(Issue("error", "symlink", rel, "Nested symlink file is not read."))
                continue
            collected.append(path)
            try:
                if path.stat().st_size == 0:
                    issues.append(Issue("warning", "empty-file", rel, "File is empty."))
            except OSError as exc:
                issues.append(Issue("error", "unreadable", rel, f"Cannot inspect file: {exc}"))
            if filename.lower() in SUSPICIOUS or path.suffix.lower() in TEMP_SUFFIXES:
                issues.append(Issue("warning", "suspicious-entry", rel, "README/cache/log/tmp files are discouraged."))
    return collected


def _markdown_links(markdown: Path, skill: Path, issues: list[Issue]) -> None:
    rel = markdown.relative_to(skill).as_posix()
    try:
        text = markdown.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as exc:
        issues.append(Issue("error", "unreadable", rel, f"Cannot read Markdown file: {exc}"))
        return
    text = _outside_fences(text)
    for match in re.finditer(r"!??\[[^\]]*\]\(([^)]*)\)", text):
        raw = match.group(1).strip()
        if not raw:
            issues.append(Issue("error", "broken-link", rel, "Markdown link has an empty target."))
            continue
        angled = re.match(r"^<([^>]+)>", raw)
        target = angled.group(1) if angled else raw.split()[0]
        try:
            parts = urlsplit(target)
        except ValueError:
            issues.append(Issue("error", "broken-link", rel, f"Malformed link target '{target}'."))
            continue
        if parts.scheme or parts.netloc or target.startswith("/") or re.match(r"^[A-Za-z]:[/\\]", target):
            issues.append(Issue("warning", "resource-link", rel, f"Resource link '{target}' is external or absolute."))
            continue
        dest = unquote(parts.path)
        resolved = markdown
        if dest:
            try:
                resolved = (markdown.parent / dest).resolve()
            except (OSError, RuntimeError, ValueError):
                issues.append(Issue("error", "broken-link", rel, f"Cannot resolve link '{target}'."))
                continue
            try:
                resolved.relative_to(skill.resolve())
            except ValueError:
                issues.append(Issue("error", "link-escape", rel, f"Local Markdown link '{target}' escapes the skill directory."))
                continue
            if not resolved.exists():
                issues.append(Issue("error", "broken-link", rel, f"Local Markdown link '{target}' does not exist."))
                continue
        if parts.fragment:
            anchor_file = resolved
            if anchor_file.is_file() and anchor_file.suffix.lower() == ".md":
                try:
                    anchor_text = anchor_file.read_text(encoding="utf-8")
                except (OSError, UnicodeError) as exc:
                    issues.append(Issue("error", "unreadable", anchor_file.relative_to(skill).as_posix(), f"Cannot read link target: {exc}"))
                    continue
                anchors = {re.sub(r"[^\w -]", "", h.group(1).lower()).strip().replace(" ", "-") for h in re.finditer(r"(?m)^#{1,6}\s+(.+?)\s*#*\s*$", _outside_fences(anchor_text))}
                if unquote(parts.fragment).lower() not in anchors:
                    issues.append(Issue("error", "broken-anchor", rel, f"Markdown anchor '#{parts.fragment}' does not exist."))


def check_skill(skill: Path) -> list[Issue]:
    skill = Path(skill)
    issues: list[Issue] = []
    try:
        skill = skill.resolve()
    except (OSError, RuntimeError, ValueError):
        return [Issue("error", "directory", str(skill), "Cannot resolve skill directory.")]
    if not skill.is_dir():
        return [Issue("error", "directory", str(skill), "Skill path is not a directory.")]
    if not DIRECTORY_NAME.fullmatch(skill.name):
        issues.append(Issue("error", "directory-name", ".", "Use letters, underscores, or internal hyphens in directory names."))
    files = _walk(skill, issues)
    try:
        entries = list(skill.iterdir())
    except OSError as exc:
        return issues + [Issue("error", "unreadable", ".", f"Cannot inspect skill directory: {exc}")]
    for entry in entries:
        if entry.name not in ALLOWED_ROOT:
            level = "warning" if entry.name.lower() in SUSPICIOUS or entry.name.startswith(".") else "error"
            issues.append(Issue(level, "root-entry", entry.name, "Root entry is not in the allowed skill layout."))
        elif entry.is_symlink():
            continue
        elif entry.is_dir() and entry.name not in {"scripts", "references", "assets"}:
            issues.append(Issue("error", "root-entry", entry.name, "Allowed root file name is not a directory."))
        elif entry.name in {"scripts", "references", "assets"} and not entry.is_dir():
            issues.append(Issue("error", "root-entry", entry.name, "Resource root entry must be a directory."))
    skill_file = skill / "SKILL.md"
    text = ""
    if skill_file not in files:
        issues.append(Issue("error", "skill-file", "SKILL.md", "Expected uppercase SKILL.md at the skill root."))
    else:
        try:
            text = skill_file.read_text(encoding="utf-8")
        except (OSError, UnicodeError) as exc:
            issues.append(Issue("error", "unreadable", "SKILL.md", f"Cannot read SKILL.md: {exc}"))
        else:
            _, end = _frontmatter(text, "SKILL.md", issues)
            body = "\n".join(text.splitlines()[end + 1:]).strip() if end is not None else ""
            if not body:
                issues.append(Issue("error", "empty-body", "SKILL.md", "Skill body must be nonempty."))
            if len(body.splitlines()) > 500:
                issues.append(Issue("warning", "body-length", "SKILL.md", f"Skill body has {len(body.splitlines())} lines; recommended maximum is 500."))
    pyfiles = [p for p in files if p.suffix == ".py"]
    if pyfiles and skill / "requirements.txt" not in files:
        issues.append(Issue("error", "requirements", "requirements.txt", "Python scripts require requirements.txt."))
    for pyfile in pyfiles:
        try:
            ast.parse(pyfile.read_text(encoding="utf-8"), filename=str(pyfile))
        except (SyntaxError, UnicodeError, OSError) as exc:
            issues.append(Issue("error", "python-syntax", pyfile.relative_to(skill).as_posix(), f"Cannot parse Python file: {exc}"))
    refs = skill / "references"
    reference_files = [p for p in files if p.is_relative_to(refs) and p.suffix.lower() == ".md"]
    if reference_files:
        for md in reference_files:
            try:
                content = md.read_text(encoding="utf-8")
            except (OSError, UnicodeError) as exc:
                issues.append(Issue("error", "unreadable", md.relative_to(skill).as_posix(), f"Cannot read reference: {exc}"))
                continue
            if len(content.splitlines()) > 100 and not re.search(r"(?im)^\s*(#{1,6}\s+)?(table of contents|contents|目次|目录)\s*$", content):
                issues.append(Issue("warning", "reference-toc", md.relative_to(skill).as_posix(), "Reference exceeds 100 lines and has no recognizable table of contents."))
    markdowns = ([skill_file] if skill_file in files else []) + reference_files
    for md in markdowns:
        if md.is_file() and not md.is_symlink():
            _markdown_links(md, skill, issues)
    return issues


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("paths", nargs="+", type=Path)
    parser.add_argument("--strict", action="store_true", help="Treat warnings as failures.")
    args = parser.parse_args(argv)
    failed = False
    for path in args.paths:
        issues = check_skill(path)
        if not issues:
            print(f"OK {path}")
        for issue in issues:
            print(f"{issue.level.upper()} {path}:{issue.path}: [{issue.rule}] {issue.message}")
            failed |= issue.level == "error" or (args.strict and issue.level == "warning")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
