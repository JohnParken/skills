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
ParseError = ValueError


def _checked(data):
    if isinstance(data, str):
        data = data.encode("utf-8")
    if not isinstance(data, bytes):
        raise ValueError("XML input must be bytes or text")
    if len(data) > MAX_XML_BYTES:
        raise ValueError("XML exceeds the 16 MiB limit")
    if expat.version_info < (2, 6, 0):
        raise ValueError("XML input requires Python with Expat >= 2.6.0")
    parser = expat.ParserCreate()
    depth = elements = 0

    def reject(*args):
        raise ValueError("DTD and entity declarations are forbidden")

    def start(name, attrs):
        nonlocal depth, elements
        depth += 1
        elements += 1
        if depth > MAX_XML_DEPTH or elements > MAX_XML_ELEMENTS:
            raise ValueError("XML nesting or element limit exceeded")

    def end(name):
        nonlocal depth
        depth -= 1

    parser.StartDoctypeDeclHandler = reject
    parser.EntityDeclHandler = reject
    parser.ExternalEntityRefHandler = reject
    parser.StartElementHandler = start
    parser.EndElementHandler = end
    try:
        parser.Parse(data, True)
    except expat.ExpatError as exc:
        raise ValueError(f"Invalid XML: {exc}") from exc
    return data


def read_xml(path):
    with Path(path).open("rb") as stream:
        data = stream.read(MAX_XML_BYTES + 1)
    if len(data) > MAX_XML_BYTES:
        raise ValueError("XML exceeds the 16 MiB limit")
    return data


def fromstring(data):
    try:
        return ET.fromstring(_checked(data))
    except ET.ParseError as exc:
        raise ValueError(f"Invalid XML: {exc}") from exc


def parse(path):
    return ET.ElementTree(fromstring(read_xml(path)))


def parse_dom(data):
    return minidom.parseString(_checked(data))
