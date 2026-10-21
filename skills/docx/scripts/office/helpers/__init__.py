import os
import posixpath
import re
import stat
import tempfile
import urllib.parse
import unicodedata
import zipfile
from pathlib import Path

OOXML_FAMILY = {
    ".docx": "docx",
    ".dotx": "docx",
}

_SCHEME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9+.\-]*:")

MAX_PART_BYTES = 32 * 1024 * 1024
MAX_PACKAGE_BYTES = 128 * 1024 * 1024
MAX_MEMBERS = 4096
MAX_PATH_BYTES = 4096
MAX_PATH_COMPONENTS = 128


def reject_signed_package(files, relation_types=()):
    """Refuse mutations that would leave a package's signatures invalid."""
    if any(str(name).lower().startswith("_xmlsignatures/") for name in files) or any(
        "digital-signature" in kind.lower() for kind in relation_types
    ):
        raise ValueError("Signed packages are unsupported for editing")


def check_archive(zf: zipfile.ZipFile) -> None:
    """Check names, types and declared sizes before decompressing any members."""
    _check_entries(zf.infolist())


def _check_entries(entries) -> None:
    if len(entries) > MAX_MEMBERS:
        raise ValueError("Archive exceeds the 4096 member limit")
    seen = set()
    namespace = {}
    total = 0
    for item in entries:
        name = item.filename
        path = name[:-1] if name.endswith("/") else name
        if (
            name != item.orig_filename or not path or "\\" in path or "\x00" in path or ":" in path
            or path.startswith("/") or any(p in ("", ".", "..") for p in path.split("/"))
        ):
            raise ValueError(f"unsafe archive entry: {name!r}")
        for component in path.split("/"):
            stem = component.split(".", 1)[0].upper()
            if (any(character in component for character in '<>"|?*')
                    or component.endswith((" ", "."))
                    or stem in {"CON", "PRN", "AUX", "NUL"}
                    or re.fullmatch(r"(?:COM|LPT)[1-9¹²³]", stem)):
                raise ValueError(f"Nonportable archive entry: {name!r}")
        if path in seen:
            raise ValueError(f"Duplicate ZIP members: {name!r}")
        seen.add(path)
        segments = path.split("/")
        if len(segments) > MAX_PATH_COMPONENTS or len(path.encode("utf-8")) > MAX_PATH_BYTES:
            raise ValueError(f"ZIP path depth or length limit exceeded: {name!r}")
        children = namespace
        for index, component in enumerate(segments):
            canonical = unicodedata.normalize("NFC", component).casefold()
            node = children.setdefault(canonical, {"spelling": component, "children": {}, "kind": None})
            if node["spelling"] != component:
                raise ValueError(f"Case or Unicode alias in ZIP names: {name!r}")
            is_directory = index < len(segments) - 1 or item.is_dir()
            if is_directory:
                if node["kind"] == "file":
                    raise ValueError(f"File/directory ZIP name conflict: {name!r}")
                node["kind"] = "directory"
            else:
                if node["kind"] == "directory":
                    raise ValueError(f"File/directory ZIP name conflict: {name!r}")
                node["kind"] = "file"
            children = node["children"]
        if stat.S_ISLNK(item.external_attr >> 16):
            raise ValueError(f"symlink archive entry not allowed: {name!r}")
        if item.flag_bits & 1:
            raise ValueError("Encrypted ZIP members are unsupported")
        if item.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
            raise ValueError("Only stored and deflated ZIP members are supported")
        if item.is_dir() and item.file_size:
            raise ValueError("ZIP directory entries must not contain data")
        total += item.file_size
        if item.file_size > MAX_PART_BYTES or total > MAX_PACKAGE_BYTES:
            raise ValueError("Archive exceeds part (32 MiB) or total (128 MiB) size limit")


def read_part(zf: zipfile.ZipFile, name: str) -> bytes:
    info = zf.getinfo(name)
    if info.file_size > MAX_PART_BYTES:
        raise ValueError(f"Part exceeds 32 MiB: {name}")
    with zf.open(info) as stream:
        data = stream.read(min(info.file_size + 1, MAX_PART_BYTES + 1))
    if len(data) > MAX_PART_BYTES:
        raise ValueError(f"Part exceeds 32 MiB: {name}")
    if len(data) != info.file_size:
        raise ValueError(f"ZIP member size mismatch: {name}")
    return data


def opc_target(target: str, source_part: str, target_mode: str = "") -> str | None:
    if not target:
        return None
    if target_mode.lower() == "external":
        return None
    target = urllib.parse.unquote(target)
    if _SCHEME_RE.match(target) or target.startswith("//"):
        raise ValueError("URI relationships must declare TargetMode=External")
    target = target.split("#", 1)[0]

    if "\\" in target:
        raise ValueError(f"relationship target is not a POSIX part name: {target!r}")

    if target.startswith("/"):
        joined = target.lstrip("/")
    else:
        joined = posixpath.join(posixpath.dirname(source_part), target)

    parts: list[str] = []
    for segment in posixpath.normpath(joined).split("/"):
        if segment in ("", "."):
            continue
        if segment == "..":
            if not parts:
                raise ValueError(f"relationship target escapes the package: {target!r}")
            parts.pop()
        else:
            parts.append(segment)

    if not parts:
        raise ValueError(f"relationship target resolves to nothing: {target!r}")
    return "/".join(parts)


def rels_source_part(rels_file: Path, unpacked_dir: Path) -> str:
    owner_dir = rels_file.parent.parent.relative_to(unpacked_dir)
    return posixpath.join(owner_dir.as_posix(), rels_file.name[: -len(".rels")]).lstrip("./")


def part_text(data: bytes) -> str:
    return data.decode("utf-8", "surrogateescape")


XML_SPACE = " \t\r\n"


def rendered_text(text: str, preserve: bool) -> str:
    return text if preserve else text.strip(XML_SPACE)


def atomic_write_bytes(destination: Path, data: bytes) -> None:
    """Replace a regular output only after its complete data is safely staged."""
    destination = Path(destination)
    if destination.is_dir():
        raise IsADirectoryError(f"Is a directory: {destination}")
    if destination.is_symlink() or destination.exists() and not destination.is_file():
        raise ValueError("Output must be a regular file, not a symbolic link")
    descriptor, temporary_name = tempfile.mkstemp(
        prefix="." + destination.name + ".", suffix=".tmp", dir=destination.parent,
    )
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        if destination.exists():
            os.chmod(temporary, stat.S_IMODE(destination.stat().st_mode))
        os.replace(temporary, destination)
    finally:
        if temporary.exists():
            temporary.unlink()


def safe_extract(zf: zipfile.ZipFile, dest: Path) -> None:
    check_archive(zf)
    if dest.is_symlink() or dest.exists() and (not dest.is_dir() or any(dest.iterdir())):
        raise ValueError("Extraction directory must be new or empty and not a symbolic link")
    dest.mkdir(parents=True, exist_ok=True)
    dest = dest.resolve()
    for m in zf.infolist():
        if stat.S_ISLNK(m.external_attr >> 16):
            raise ValueError(f"symlink archive entry not allowed: {m.filename!r}")
        target = (dest / m.filename).resolve()
        if not target.is_relative_to(dest):
            raise ValueError(f"unsafe archive entry: {m.filename!r}")
        if m.is_dir():
            target.mkdir(parents=True, exist_ok=True)
        else:
            data = read_part(zf, m.filename)
            target.parent.mkdir(parents=True, exist_ok=True)
            with target.open("xb") as stream:
                stream.write(data)


def rezip(src_dir: Path, out_path: Path) -> None:
    if src_dir.is_symlink() or not src_dir.is_dir():
        raise ValueError("Pack source must be a directory without symbolic links")
    files, entries = [], []
    for path in sorted(src_dir.rglob("*")):
        mode = path.lstat().st_mode
        if stat.S_ISLNK(mode) or not (stat.S_ISREG(mode) or stat.S_ISDIR(mode)):
            raise ValueError(f"Only regular package files and directories are supported: {path.name}")
        if stat.S_ISDIR(mode):
            continue
        item = zipfile.ZipInfo(path.relative_to(src_dir).as_posix())
        item.file_size = path.stat().st_size
        files.append(path)
        entries.append(item)
    _check_entries(entries)
    ct = src_dir / "[Content_Types].xml"
    fd, tmp_name = tempfile.mkstemp(
        prefix=out_path.name + ".", suffix=".tmp", dir=out_path.parent
    )
    tmp_out = Path(tmp_name)
    try:
        with os.fdopen(fd, "wb") as fh:
            with zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as zf:
                if ct.exists():
                    zf.write(ct, ct.relative_to(src_dir), compress_type=zipfile.ZIP_STORED)
                for f in files:
                    if f == ct:
                        continue
                    zf.write(f, f.relative_to(src_dir))
        with zipfile.ZipFile(tmp_out) as checked:
            check_archive(checked)
            if checked.testzip() is not None:
                raise ValueError("Repacked ZIP failed its CRC check")
        if out_path.exists():
            mode = out_path.stat().st_mode & 0o777
        else:
            umask = os.umask(0)
            os.umask(umask)
            mode = 0o666 & ~umask
        os.chmod(tmp_out, mode)
        os.replace(tmp_out, out_path)
    finally:
        if tmp_out.exists():
            tmp_out.unlink()
