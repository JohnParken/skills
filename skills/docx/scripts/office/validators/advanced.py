"""Explicit table-grid and numbering invariants, not OOXML schema validation.

Implicit/augmented grids and ignored restart/offset values are legal. Do not
invent fixed-width-table constraints or simulate the reader's numbering engine.
Legacy horizontal merges, wrapped rows/cells and style-linked numbering are
reported as unverified rather than rejected on guessed layout semantics.
"""

import re

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
INTEGER = re.compile(r"[+-]?[0-9]+\Z")


def w(name):
    return f"{{{W}}}{name}"


def _integer(raw, location, error, *, minimum=None, maximum=None):
    value = raw.strip(" \t\r\n") if isinstance(raw, str) else ""
    if not INTEGER.fullmatch(value):
        error(f"{location}: expected an integer, got {raw!r}")
        return None
    # DecimalNumberType is a signed 32-bit value; bound before int conversion.
    significant = value.lstrip("+-").lstrip("0") or "0"
    if len(significant) > 10:
        error(f"{location}: integer is outside the supported 32-bit range")
        return None
    result = int(significant)
    if value.startswith("-"):
        result = -result
    lower = -(2 ** 31) if minimum is None else minimum
    upper = 2 ** 31 - 1 if maximum is None else maximum
    if result < lower or result > upper:
        error(f"{location}: integer {result} must be in {lower}..{upper}")
        return None
    return result


def _single(parent, name, location, error):
    nodes = [] if parent is None else parent.findall(w(name))
    if len(nodes) > 1:
        error(f"{location}/{name}: duplicate elements")
    return nodes[0] if nodes else None


def _value(parent, name, location, error, *, default=None, minimum=None, maximum=None):
    node = _single(parent, name, location, error)
    if node is None:
        return default
    return _integer(node.get(w("val")), f"{location}/{name}", error,
                    minimum=minimum, maximum=maximum)


def validate_tables(roots, error, warning):
    for part_name, root in roots.items():
        if not part_name.startswith("word/") or part_name.endswith(".rels"):
            continue
        for table_index, table in enumerate(root.iter(w("tbl")), 1):
            location = f"{part_name}:tbl[{table_index}]"
            _validate_table(table, location, error, warning)


def _validate_table(table, location, error, warning):
    grids = table.findall(w("tblGrid"))
    if len(grids) > 1:
        error(f"{location}/tblGrid: duplicate table grids")
    columns = len(grids[0].findall(w("gridCol"))) if grids else 0
    rows = table.findall(w("tr"))
    # These wrappers require revision/content-control interpretation. Check
    # nested tables separately, but do not make continuity claims here.
    wrappers = {w("sdt"), w("customXml"), w("ins"), w("del"), w("moveFrom"), w("moveTo")}
    wrapped = any(child.tag in wrappers or not child.tag.startswith("{" + W + "}")
                  for child in table)
    wrapped = wrapped or any(
        any(child.tag in wrappers or not child.tag.startswith("{" + W + "}") for child in row)
        or row.find(w("trPr") + "/" + w("ins")) is not None
        or row.find(w("trPr") + "/" + w("del")) is not None
        for row in rows
    )
    wrapped = wrapped or any(cell.find(w("tcPr") + "/" + w("cellMerge")) is not None
                             for row in rows for cell in row.findall(w("tc")))
    legacy = any(cell.find(w("tcPr") + "/" + w("hMerge")) is not None
                 for row in rows for cell in row.findall(w("tc")))
    continuity = not wrapped and not legacy
    if not continuity:
        reason = "legacy hMerge" if legacy else "wrapped/revised rows or cells"
        warning(f"{location}: {reason}; vertical merge continuity is unverified")
    previous = set()
    for row_index, row in enumerate(rows, 1):
        row_location = f"{location}/tr[{row_index}]"
        props = _single(row, "trPr", row_location, error)
        before = _value(props, "gridBefore", row_location, error, default=0, minimum=0)
        after = _value(props, "gridAfter", row_location, error, default=0, minimum=0)
        if before is None or after is None:
            previous = set()
            continue
        if columns and before > columns:
            warning(f"{row_location}/gridBefore: {before} exceeds declared grid; ignored by the format")
            before = 0
        cursor = before
        current = set()
        for cell_index, cell in enumerate(row.findall(w("tc")), 1):
            cell_location = f"{row_location}/tc[{cell_index}]"
            cell_props = _single(cell, "tcPr", cell_location, error)
            span = _value(cell_props, "gridSpan", cell_location, error, default=1, minimum=1)
            merge = _single(cell_props, "vMerge", cell_location, error)
            if span is None:
                # Later column positions cannot be established after bad span.
                continuity = False
                continue
            interval = (cursor, cursor + span)
            cursor += span
            if merge is None:
                continue
            mode = merge.get(w("val"), "continue")
            if mode not in ("restart", "continue"):
                error(f"{cell_location}/vMerge: expected restart or continue, got {mode!r}")
                continue
            if continuity and mode == "continue" and interval not in previous:
                error(f"{cell_location}/vMerge: continuation at grid columns "
                      f"[{interval[0]},{interval[1]}) has no matching merge in the preceding row")
            current.add(interval)
        if columns and after and cursor + after != columns:
            warning(f"{row_location}/gridAfter: conflicts with remaining declared grid; ignored by the format")
        # gridSpan may legally augment the grid for following rows.
        columns = max(columns, cursor)
        previous = current


def _check_level(level, location, error, warning):
    index = _integer(level.get(w("ilvl")), f"{location}/@ilvl", error, minimum=0, maximum=8)
    _value(level, "start", location, error)
    restart = _value(level, "lvlRestart", location, error, minimum=0)
    if index is not None and restart is not None and restart > index:
        warning(f"{location}/lvlRestart: {restart} does not name an earlier level; ignored by the format")
    return index


def validate_numbering(roots, error, warning):
    numbering = roots.get("word/numbering.xml")
    if numbering is None:
        _paragraph_levels(roots, {}, error, warning)
        return
    abstract_levels = {}
    linked = set()
    for position, abstract in enumerate(numbering.findall(w("abstractNum")), 1):
        identifier = abstract.get(w("abstractNumId"))
        location = f"word/numbering.xml:abstractNum[{position}](@abstractNumId={identifier})"
        levels = set()
        for level_position, level in enumerate(abstract.findall(w("lvl")), 1):
            level_location = f"{location}/lvl[{level_position}]"
            index = _check_level(level, level_location, error, warning)
            if index in levels:
                error(f"{level_location}/@ilvl: duplicate level {index}")
            if index is not None:
                levels.add(index)
        abstract_levels[identifier] = levels
        if abstract.find(w("numStyleLink")) is not None:
            linked.add(identifier)
            warning(f"{location}/numStyleLink: linked style numbering levels are unverified")
    effective = {}
    for position, num in enumerate(numbering.findall(w("num")), 1):
        identifier = num.get(w("numId"))
        location = f"word/numbering.xml:num[{position}](@numId={identifier})"
        reference = _single(num, "abstractNumId", location, error)
        abstract_id = reference.get(w("val")) if reference is not None else None
        levels = set(abstract_levels.get(abstract_id, set()))
        overrides = set()
        for override_position, override in enumerate(num.findall(w("lvlOverride")), 1):
            override_location = f"{location}/lvlOverride[{override_position}]"
            index = _integer(override.get(w("ilvl")), f"{override_location}/@ilvl", error,
                             minimum=0, maximum=8)
            if index in overrides:
                error(f"{override_location}/@ilvl: duplicate override level {index}")
            if index is not None:
                overrides.add(index)
            _value(override, "startOverride", override_location, error)
            level = _single(override, "lvl", override_location, error)
            if level is not None:
                nested_index = _check_level(level, f"{override_location}/lvl", error, warning)
                if nested_index is not None and index is not None and nested_index != index:
                    warning(f"{override_location}/lvl/@ilvl: differs from override index; effective level is unverified")
                if index is not None:
                    levels.add(index)
        # None explicitly means resolution requires a linked style, not that
        # all requested levels exist. Preserve compatibility and report scope.
        effective[identifier] = None if abstract_id in linked else levels
    _paragraph_levels(roots, effective, error, warning)


def _paragraph_levels(roots, effective, error, warning):
    for part_name, root in roots.items():
        if not part_name.startswith("word/") or part_name.endswith(".rels"):
            continue
        for position, paragraph in enumerate(root.iter(w("p")), 1):
            location = f"{part_name}:p[{position}]/pPr/numPr"
            props = paragraph.find(w("pPr"))
            numbering = _single(props, "numPr", location.rsplit("/", 1)[0], error)
            if numbering is None:
                continue
            level = _value(numbering, "ilvl", location, error, minimum=0, maximum=8)
            reference = _single(numbering, "numId", location, error)
            # Missing numId/ilvl can be inherited from styles; only verify an
            # explicit pair. numId=0 removes numbering and needs no level.
            identifier = reference.get(w("val")) if reference is not None else None
            if identifier == "0":
                continue
            if identifier is None or level is None:
                warning(f"{location}: inherited numbering ID/level is unverified")
                continue
            if identifier not in effective:
                continue
            levels = effective[identifier]
            if levels is not None and level not in levels:
                error(f"{location}: numId={identifier} ilvl={level} references an undefined numbering level")
