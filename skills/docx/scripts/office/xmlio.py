"""Bounded XML parsing using only the Python standard library.

Reject DTDs and entity declarations through parser callbacks (including UTF-16),
then build an ElementTree or a namespace-preserving DOM. Never resolve resources.
"""

from pathlib import Path
from xml.etree import ElementTree as ET
from xml.dom import minidom
from xml.parsers import expat

MAX_XML_BYTES = 16 * 1024 * 1024
MAX_XML_DEPTH = 128
MAX_XML_ELEMENTS = 200_000
# Bound allocations not represented by the element count or input byte size.
# Namespace URIs in normal OOXML are short; this ceiling is checked in a
# namespace-unaware pass before Expat can multiply a URI into expanded names.
MAX_XML_NAMESPACE_URI_CHARS = 2048
MAX_XML_EXPANDED_NAME_CHARS = 32 * 1024 * 1024
MAX_XML_ATTRIBUTES_PER_ELEMENT = 1024
MAX_XML_ATTRIBUTES = 200_000
MAX_XML_NODES = 400_000
ParseError = ValueError


class XmlBudget:
    """Cumulative allocation counters for one package/read/validation operation.

    Use a fresh instance for each independent operation and pass it to all XML
    parts retained by that operation. Each parse consumes budget, including
    repeated parses and failed inputs; this is intentionally fail-closed.
    The same constants also remain enforced independently for each XML part.
    """

    def __init__(self):
        self.elements = 0
        self.attributes = 0
        self.nodes = 0
        self.expanded_name_chars = 0

    def _consume(self, kind, amount, limit):
        value = getattr(self, kind) + amount
        setattr(self, kind, value)
        if value > limit:
            raise ValueError(f"XML shared {kind.replace('_', ' ')} limit exceeded")


def _checked(data, *, budget=None):
    if budget is not None and not isinstance(budget, XmlBudget):
        raise ValueError("budget must be an XmlBudget instance or None")
    if isinstance(data, str):
        data = data.encode("utf-8")
    if not isinstance(data, bytes):
        raise ValueError("XML input must be bytes or text")
    if len(data) > MAX_XML_BYTES:
        raise ValueError("XML exceeds the 16 MiB limit")
    if expat.version_info < (2, 6, 0):
        raise ValueError("XML input requires Python with Expat >= 2.6.0")
    parser = expat.ParserCreate()
    depth = elements = attributes = nodes = namespace_chars = 0
    text_open = in_cdata = False

    def reject(*args):
        raise ValueError("DTD and entity declarations are forbidden")

    def add_nodes(count):
        nonlocal nodes
        nodes += count
        if nodes > MAX_XML_NODES:
            raise ValueError("XML node limit exceeded")
        if budget is not None:
            budget._consume("nodes", count, MAX_XML_NODES)

    def start(name, attrs):
        nonlocal depth, elements, attributes, namespace_chars, text_open
        depth += 1
        elements += 1
        if depth > MAX_XML_DEPTH or elements > MAX_XML_ELEMENTS:
            raise ValueError("XML nesting or element limit exceeded")
        attributes += len(attrs)
        if len(attrs) > MAX_XML_ATTRIBUTES_PER_ELEMENT or attributes > MAX_XML_ATTRIBUTES:
            raise ValueError("XML attribute limit exceeded")
        if budget is not None:
            budget._consume("elements", 1, MAX_XML_ELEMENTS)
            budget._consume("attributes", len(attrs), MAX_XML_ATTRIBUTES)
        add_nodes(1 + len(attrs))
        text_open = False
        for key, value in attrs.items():
            if key == "xmlns" or key.startswith("xmlns:"):
                if len(value) > MAX_XML_NAMESPACE_URI_CHARS:
                    raise ValueError("XML namespace URI limit exceeded")
                namespace_chars += len(key) + len(value)
                if namespace_chars > MAX_XML_EXPANDED_NAME_CHARS:
                    raise ValueError("XML expanded name limit exceeded")
                if budget is not None:
                    budget._consume("expanded_name_chars", len(key) + len(value),
                                    MAX_XML_EXPANDED_NAME_CHARS)

    def end(name):
        nonlocal depth, text_open
        depth -= 1
        text_open = False

    def characters(value):
        nonlocal text_open
        # Expat may split one DOM text node into callbacks at entity references
        # or internal buffer boundaries. Count contiguous character data once.
        if value and not in_cdata and not text_open:
            add_nodes(1)
            text_open = True

    def miscellaneous(*args):
        nonlocal text_open
        add_nodes(1)
        text_open = False

    def start_cdata():
        nonlocal in_cdata, text_open
        add_nodes(1)
        in_cdata = True
        text_open = False

    def end_cdata():
        nonlocal in_cdata, text_open
        in_cdata = False
        text_open = False

    parser.StartDoctypeDeclHandler = reject
    parser.EntityDeclHandler = reject
    parser.ExternalEntityRefHandler = reject
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = characters
    parser.CommentHandler = miscellaneous
    parser.ProcessingInstructionHandler = miscellaneous
    parser.StartCdataSectionHandler = start_cdata
    parser.EndCdataSectionHandler = end_cdata
    try:
        parser.Parse(data, True)
        # The first pass bounds raw attributes and namespace URI lengths. Only
        # then enable expansion, which otherwise allocates URI * local-name
        # combinations before the application receives a start-element event.
        expanded_chars = namespace_chars

        def expanded_start(name, attrs):
            nonlocal expanded_chars
            amount = len(name) + sum(len(key) for key in attrs)
            expanded_chars += amount
            if expanded_chars > MAX_XML_EXPANDED_NAME_CHARS:
                raise ValueError("XML expanded name limit exceeded")
            if budget is not None:
                budget._consume("expanded_name_chars", amount, MAX_XML_EXPANDED_NAME_CHARS)

        names_parser = expat.ParserCreate(namespace_separator="}")
        names_parser.StartDoctypeDeclHandler = reject
        names_parser.EntityDeclHandler = reject
        names_parser.ExternalEntityRefHandler = reject
        names_parser.StartElementHandler = expanded_start
        names_parser.Parse(data, True)
    except expat.ExpatError as exc:
        raise ValueError(f"Invalid XML: {exc}") from exc
    return data


def read_xml(path):
    with Path(path).open("rb") as stream:
        data = stream.read(MAX_XML_BYTES + 1)
    if len(data) > MAX_XML_BYTES:
        raise ValueError("XML exceeds the 16 MiB limit")
    return data


def fromstring(data, *, budget=None):
    try:
        return ET.fromstring(_checked(data, budget=budget))
    except ET.ParseError as exc:
        raise ValueError(f"Invalid XML: {exc}") from exc


def parse(path, *, budget=None):
    return ET.ElementTree(fromstring(read_xml(path), budget=budget))


def parse_dom(data, *, budget=None):
    try:
        return minidom.parseString(_checked(data, budget=budget))
    except expat.ExpatError as exc:
        raise ValueError(f"Invalid XML: {exc}") from exc
