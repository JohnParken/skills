"""Check generated layout intent against normalized input, without rendering.

This generator-specific contract is not an OOXML/XSD validator. Expectations
come from input values, without calling the XML serializers a second time.
Run after StructureValidator has bounded and validated the staged archive.
"""

import hashlib

from ..helpers import read_part
from .structure import R, w

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _expect(actual, expected, location):
    if actual != expected:
        raise ValueError(f"Generated layout mismatch at {location}")


def _attrs(parent, path):
    node = parent.find(path) if parent is not None else None
    return node.attrib if node is not None else None


def _properties(parent, expected, location, skip=()):
    children = [child for child in parent if child.tag not in skip] if parent is not None else []
    _expect(len(children), len(expected), location)
    _expect({child.tag: child.attrib for child in children}, expected, location)
    _expect(any(len(child) for child in children), False, location)


def _value(value):
    return {w("val"): str(value)}


def _format_maps(formatting):
    """Translate supported input units to expected paragraph/run properties."""
    paragraph, run = {}, {}
    fonts = {}
    if "font_east_asia" in formatting:
        fonts[w("eastAsia")] = formatting["font_east_asia"]
    if "font_latin" in formatting:
        fonts.update({w(key): formatting["font_latin"] for key in ("ascii", "hAnsi", "cs")})
    if fonts:
        run[w("rFonts")] = fonts
    for key, tag in (("bold", "b"), ("italic", "i")):
        if key in formatting:
            run[w(tag)] = _value(int(formatting[key]))
    if "underline" in formatting:
        run[w("u")] = _value("single" if formatting["underline"] else "none")
    if "size_pt" in formatting:
        for tag in ("sz", "szCs"):
            run[w(tag)] = _value(int(formatting["size_pt"] * 2 + 0.5))
    if "color" in formatting:
        run[w("color")] = _value(formatting["color"])
    spacing = {}
    for key, attr in (("space_before_pt", "before"), ("space_after_pt", "after")):
        if key in formatting:
            spacing[w(attr)] = str(int(formatting[key] * 20 + 0.5))
    if "line_spacing_pt" in formatting:
        spacing.update({w("line"): str(int(formatting["line_spacing_pt"] * 20 + 0.5)),
                        w("lineRule"): "exact"})
    elif "line_spacing_multiple" in formatting:
        spacing.update({w("line"): str(int(formatting["line_spacing_multiple"] * 240 + 0.5)),
                        w("lineRule"): "auto"})
    if spacing:
        paragraph[w("spacing")] = spacing
    if "first_line_chars" in formatting:
        paragraph[w("ind")] = {w("firstLineChars"): str(int(formatting["first_line_chars"] * 100 + 0.5))}
    if "align" in formatting:
        paragraph[w("jc")] = _value("both" if formatting["align"] == "justify" else formatting["align"])
    return paragraph, run


def _styles(spec, root):
    styles = {node.get(w("styleId")): node for node in root.findall(w("style"))}
    normal = {"font_east_asia": "仿宋", "font_latin": "Times New Roman", "size_pt": 16,
              "bold": False, "space_before_pt": 0, "space_after_pt": 0,
              "line_spacing_multiple": 1.15, "align": "left"}
    expected = {"Normal": normal, "BodyText": spec["styles"]["body"],
                "TableBody": spec["styles"]["table_body"], "TableHeader": spec["styles"]["table_header"],
                "Title": spec["styles"]["title"], "Subtitle": spec["styles"]["subtitle"]}
    expected.update({f"Heading{level}": spec["styles"][f"heading{level}"] for level in range(1, 5)})
    if spec["related_format"] is not None:
        expected.update({"Header": spec["related_format"], "Footer": spec["related_format"]})
    if spec["list_format"] is not None:
        expected["ListParagraph"] = spec["list_format"]
    for name, formatting in expected.items():
        style = styles.get(name)
        _expect(style is not None, True, f"styles.{name}")
        _expect(style.get(w("type")), "paragraph", f"styles.{name}.type")
        if name != "Normal":
            _expect(_attrs(style, w("basedOn")), _value("Normal"), f"styles.{name}.basedOn")
        else:
            _expect(style.get(w("default")), "1", "styles.Normal.default")
        paragraph, run = _format_maps(formatting)
        if name.startswith("Heading"):
            paragraph[w("outlineLvl")] = _value(int(name[-1]) - 1)
        if name.startswith("Heading") or name in {"Title", "Subtitle"}:
            paragraph[w("keepNext")] = {}
        _properties(style.find(w("pPr")), paragraph, f"styles.{name}.paragraph")
        _properties(style.find(w("rPr")), run, f"styles.{name}.run")
    paragraph, run = _format_maps(normal)
    paragraph.pop(w("jc"))
    _properties(root.find(f"{w('docDefaults')}/{w('pPrDefault')}/{w('pPr')}"), paragraph, "defaults.paragraph")
    _properties(root.find(f"{w('docDefaults')}/{w('rPrDefault')}/{w('rPr')}"), run, "defaults.run")


def _paragraph(node, text_runs, style, formatting, location, numbering=None, image=None):
    _expect(node.tag, w("p"), location)
    props, _ = _format_maps(formatting)
    if style is not None:
        props[w("pStyle")] = _value(style)
    ppr = node.find(w("pPr"))
    _properties(ppr, props, location + ".paragraph", skip=(w("numPr"),))
    num_props = ppr.find(w("numPr")) if ppr is not None else None
    expected_num = {} if numbering is None else {
        w("numId"): _value(numbering[0]), w("ilvl"): _value(numbering[1]),
    }
    _properties(num_props, expected_num, location + ".numbering")
    runs = node.findall(w("r"))
    _expect(len(runs), len(text_runs) + int(image is not None), location + ".runs")
    _expect(len(node), len(runs) + int(ppr is not None), location + ".children")
    for index, ((text, run_format), run) in enumerate(zip(text_runs, runs)):
        _, expected_run = _format_maps(run_format)
        _properties(run.find(w("rPr")), expected_run, f"{location}.runs[{index}].format")
        _expect([child.tag for child in run if child.tag != w("rPr")], [w("t")], location + ".text")
        _expect(run.find(w("t")).text or "", text, location + ".text")
        _expect(run.find(w("t")).attrib,
                {XML_SPACE: "preserve"} if text[:1].isspace() or text[-1:].isspace() else {},
                location + ".whitespace")
    if image is not None:
        _expect([child.tag for child in runs[-1]], [w("drawing")], location + ".image")
        return runs[-1].find(w("drawing"))
    return None


def _border(fields):
    if fields["style"] == "none":
        return _value("nil")
    return {w("val"): "single", w("sz"): str(int(fields["width_pt"] * 8 + 0.5)),
            w("space"): "0", w("color"): fields["color"]}


def _table(node, block, location, check_image):
    _expect(node.tag, w("tbl"), location)
    layout = block["format"]
    widths = block["widths_twips"]
    props = node.find(w("tblPr"))
    expected = {w("tblW"): {w("w"): str(sum(widths)), w("type"): "dxa"},
                w("jc"): _value(layout["align"]), w("tblLayout"): {w("type"): "fixed"}}
    if layout["cell_fill"] != "none":
        expected[w("shd")] = {w("val"): "clear", w("fill"): layout["cell_fill"]}
    _properties(props, expected, location + ".layout", skip=(w("tblBorders"), w("tblCellMar")))
    expected_borders = {}
    for edge in ("top", "bottom", "left", "right", "insideH", "insideV"):
        fields = {"style": layout["border_style"], "width_pt": layout["border_width_pt"],
                  "color": layout["border_color"], **layout.get("borders", {}).get(edge, {})}
        expected_borders[w(edge)] = _border(fields)
    _properties(props.find(w("tblBorders")), expected_borders, location + ".borders")
    margins = {w(side): {w("w"): str(int(layout[f"cell_margin_{side}_mm"] * 1440 / 25.4 + 0.5)),
                        w("type"): "dxa"} for side in ("top", "left", "bottom", "right")}
    _properties(props.find(w("tblCellMar")), margins, location + ".margins")
    grid = node.find(w("tblGrid"))
    _expect([child.attrib for child in grid] if grid is not None else None,
            [{w("w"): str(width)} for width in widths], location + ".widths")
    rows = node.findall(w("tr"))
    _expect(len(rows), len(block["rows"]), location + ".rows")
    _expect(len(node), len(rows) + 2, location + ".children")
    for row_index, (row, values) in enumerate(zip(rows, block["rows"])):
        row_location = f"{location}.rows[{row_index}]"
        is_header = row_index < block["header_rows"]
        expected_row = {}
        if not layout["allow_row_split"]:
            expected_row[w("cantSplit")] = _value("true")
        if is_header and layout["repeat_header"]:
            expected_row[w("tblHeader")] = _value("true")
        _properties(row.find(w("trPr")), expected_row, row_location + ".layout")
        cells = row.findall(w("tc"))
        _expect(len(cells), len(values), row_location + ".columns")
        for column, (cell, value, width) in enumerate(zip(cells, values, widths)):
            cell_location = f"{row_location}[{column}]"
            cell_props = cell.find(w("tcPr"))
            expected_cell = {w("tcW"): {w("w"): str(width), w("type"): "dxa"},
                             w("vAlign"): _value(layout["vertical_align"])}
            if not is_header and (row_index - block["header_rows"]) % 2 == 1 and layout["band_fill"] != "none":
                expected_cell[w("shd")] = {w("val"): "clear", w("fill"): layout["band_fill"]}
            _properties(cell_props, expected_cell, cell_location + ".layout", skip=(w("tcBorders"),))
            borders = value.get("borders", {}) if isinstance(value, dict) else {}
            _properties(cell_props.find(w("tcBorders")), {w(edge): _border(fields) for edge, fields in borders.items()},
                        cell_location + ".borders")
            text = value if isinstance(value, str) else value["text"]
            picture = value.get("image") if isinstance(value, dict) else None
            paragraphs = cell.findall(w("p"))
            _expect(len(paragraphs), 1, cell_location + ".paragraphs")
            formatting = block["header_format"] if is_header else block["body_format"]
            drawing = _paragraph(paragraphs[0], [(text, formatting)], "TableHeader" if is_header else "TableBody",
                                 formatting, cell_location, image=picture)
            if picture is not None:
                check_image(drawing, picture, cell_location)


def verify_generation(spec, validator):
    """Reject loss or corruption of supported generation intent before publish."""
    roots = validator.roots
    for part in ("word/document.xml", "word/styles.xml", "word/numbering.xml"):
        _expect(part in roots, True, part)
    document = roots["word/document.xml"]
    body = document.find(w("body"))
    _styles(spec, roots["word/styles.xml"])
    page = spec["page_twips"]
    section = body.find(w("sectPr"))
    size = {w("w"): str(page["width"]), w("h"): str(page["height"])}
    if page["width"] > page["height"]:
        size[w("orient")] = "landscape"
    _expect(_attrs(section, w("pgSz")), size, "page.size")
    margins = {w(side): str(page[side]) for side in ("top", "right", "bottom", "left")}
    margins.update({w("header"): str(min(720, page["top"])), w("footer"): str(min(720, page["bottom"])),
                    w("gutter"): "0"})
    _expect(_attrs(section, w("pgMar")), margins, "page.margins")
    for key, part, style in (("header", "word/header1.xml", "Header"), ("footer", "word/footer1.xml", "Footer")):
        present = spec[key] is not None or key == "footer" and spec["page_numbers"]
        _expect(part in roots, present, key)
        if present:
            expected_style = _value(style) if spec["related_format"] is not None else None
            _expect(_attrs(roots[part], f"{w('p')}/{w('pPr')}/{w('pStyle')}"), expected_style, key + ".style")
            references = section.findall(w(key + "Reference")) if section is not None else []
            _expect(len(references), 1, key + ".reference")
            relation = validator.relations["word/document.xml"].get(references[0].get("{" + R + "}id"))
            _expect(relation, (R + "/" + key, part), key + ".reference")
            paragraphs = roots[part].findall(w("p"))
            _expect(len(paragraphs), 1, key + ".paragraphs")
            _properties(paragraphs[0].find(w("pPr")),
                        {w("pStyle"): expected_style} if expected_style is not None else {}, key + ".format")
            for run in paragraphs[0].findall(w("r")):
                _properties(run.find(w("rPr")), {}, key + ".run")
    if "word/footer1.xml" in roots:
        footer = roots["word/footer1.xml"]
        _expect([node.text for node in footer.iter(w("instrText"))],
                [" PAGE "] if spec["page_numbers"] else [], "footer.page_field")
        _expect([node.get(w("fldCharType")) for node in footer.iter(w("fldChar"))],
                ["begin", "separate", "end"] if spec["page_numbers"] else [], "footer.page_field")

    media_hashes, drawing_ids = {}, set()

    def check_image(drawing, picture, location):
        _expect(drawing is not None, True, location + ".drawing")
        extent = {"cx": str(picture["cx"]), "cy": str(picture["cy"])}
        _expect(_attrs(drawing, f"{WP}inline/{WP}extent"), extent, location + ".size")
        _expect(_attrs(drawing, f".//{A}xfrm/{A}ext"), extent, location + ".transform")
        properties = drawing.find(f"{WP}inline/{WP}docPr")
        _expect(properties is not None, True, location + ".description")
        _expect(properties.get("descr"), picture["alt"], location + ".description")
        identifier = properties.get("id")
        _expect(identifier is not None and identifier not in drawing_ids, True, location + ".id")
        drawing_ids.add(identifier)
        nonvisual = drawing.find(f".//{PIC}cNvPr")
        _expect(nonvisual is not None, True, location + ".picture_description")
        _expect(nonvisual.get("descr"), picture["alt"], location + ".picture_description")
        blip = drawing.find(f".//{A}blip")
        _expect(blip is not None, True, location + ".media")
        relation = validator.relations["word/document.xml"].get(blip.get("{" + R + "}embed"))
        _expect(relation is not None and relation[0] == R + "/image" and relation[1] is not None,
                True, location + ".media")
        target = relation[1]
        if target not in media_hashes:
            media_hashes[target] = hashlib.sha256(read_part(validator.archive, target)).digest()
        _expect(media_hashes[target], picture["image_digest"], location + ".media")

    children = [child for child in body if child.tag != w("sectPr")]
    cursor = 0
    shared_lists, instances = {}, []
    numbering = roots["word/numbering.xml"]
    numbers = {node.get(w("numId")): node for node in numbering.findall(w("num"))}
    abstracts = {node.get(w("abstractNumId")): node for node in numbering.findall(w("abstractNum"))}
    for index, block in enumerate(spec["blocks"]):
        location = f"blocks[{index}]"
        kind = block["type"]
        count = len(block["items"]) if kind == "list" else 1
        _expect(cursor + count <= len(children), True, location + ".missing")
        nodes = children[cursor:cursor + count]
        cursor += count
        if kind == "list":
            identifier = block["list_id"]
            if identifier is None or identifier not in shared_lists:
                number_id = len(instances) + 1
                instances.append(number_id)
                if identifier is not None:
                    shared_lists[identifier] = number_id
            else:
                number_id = shared_lists[identifier]
            abstract_id = int(block["ordered"]) + (2 if identifier is not None else 0)
            number = numbers.get(str(number_id))
            _expect(_attrs(number, w("abstractNumId")), _value(abstract_id), location + ".list_instance")
            abstract = abstracts.get(str(abstract_id))
            levels = {node.get(w("ilvl")): node for node in abstract.findall(w("lvl"))} if abstract is not None else {}
            level = levels.get(str(block["level"]))
            _expect(_attrs(level, w("numFmt")), _value("decimal" if block["ordered"] else "bullet"), location + ".list_format")
            _expect(_attrs(level, w("start")), _value(1), location + ".list_start")
            label = ".".join(f"%{position + 1}" for position in range(block["level"] + 1)) + "." if block["ordered"] else "•"
            _expect(_attrs(level, w("lvlText")), _value(label), location + ".list_label")
            indent = str(720 * (block["level"] + 1))
            _expect(_attrs(level, f"{w('pPr')}/{w('ind')}"),
                    {w("left"): indent, w("hanging"): "360"}, location + ".list_indent")
            _expect(_attrs(level, f"{w('pPr')}/{w('tabs')}/{w('tab')}"),
                    {w("val"): "num", w("pos"): indent}, location + ".list_tab")
            _expect(_attrs(level, w("lvlRestart")),
                    _value(block["level"]) if block["level"] else None, location + ".list_restart")
            for node, text in zip(nodes, block["items"]):
                _paragraph(node, [(text, {})], "ListParagraph" if spec["list_format"] is not None else None,
                           {}, location, numbering=(number_id, block["level"]))
        elif kind == "table":
            _table(nodes[0], block, location, check_image)
        elif kind == "page_break":
            node = nodes[0]
            _expect(node.tag, w("p"), location)
            _expect([child.tag for child in node], [w("r")], location)
            run = node[0]
            _expect([child.tag for child in run], [w("br")], location)
            _expect(run[0].attrib, {w("type"): "page"}, location)
        elif kind == "image":
            drawing = _paragraph(nodes[0], [], None, {}, location, image=block)
            check_image(drawing, block, location)
        else:
            formatting = block["format"].copy()
            if block.get("align") is not None:
                formatting["align"] = block["align"]
            style = {"title": "Title", "subtitle": "Subtitle", "paragraph": "BodyText"}.get(kind)
            if kind == "heading":
                style = f"Heading{block['level']}"
            content = block["content"] if kind == "paragraph" else block["text"]
            if isinstance(content, str):
                text_runs = [(content, formatting)]
            else:
                text_runs = []
                for run in content:
                    run_format = {**formatting, **run}
                    if "font" in run:
                        run_format.update({"font_east_asia": run.get("font_east_asia", run["font"]),
                                           "font_latin": run.get("font_latin", run["font"])})
                    text_runs.append((run["text"], run_format))
            _paragraph(nodes[0], text_runs, style, formatting, location)
    _expect(cursor, len(children), "body.extra_blocks")
    _expect(len(numbers), len(instances), "numbering.instances")
