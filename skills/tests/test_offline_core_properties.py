"""Regression coverage for bundled, offline OPC core property schemas."""

import importlib.util
import types
from pathlib import Path
import sys
import unittest
from urllib.parse import urlparse

from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
SKILLS = ("xlsx", "pptx")
XSD_NS = "http://www.w3.org/2001/XMLSchema"
CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"


def schema_root(skill):
    return ROOT / skill / "scripts/office/schemas"


class RejectRemoteResolver(etree.Resolver):
    """Fail schema compilation if any dependency is fetched by URL."""

    def resolve(self, url, public_id, context):
        if urlparse(url).scheme not in ("", "file"):
            raise OSError("remote schema resolution blocked: " + url)
        return None


class OfflineCorePropertiesTests(unittest.TestCase):
    def test_actual_schema_dependencies_are_present_and_local(self):
        for skill in SKILLS:
            with self.subTest(skill=skill):
                root = schema_root(skill)
                pending = [root / "ecma/fouth-edition/opc-coreProperties.xsd"]
                visited = set()
                while pending:
                    current = pending.pop().resolve()
                    self.assertTrue(current.is_file(), str(current))
                    if current in visited:
                        continue
                    visited.add(current)
                    tree = etree.parse(str(current))
                    for node in tree.xpath(
                        "//xs:import[@schemaLocation] | //xs:include[@schemaLocation] | //xs:redefine[@schemaLocation]",
                        namespaces={"xs": XSD_NS},
                    ):
                        location = node.get("schemaLocation")
                        dependency = urlparse(location)
                        self.assertFalse(dependency.scheme, location)
                        self.assertFalse(dependency.netloc, location)
                        resolved = (current.parent / location).resolve()
                        self.assertTrue(resolved.is_relative_to(root.resolve()), location)
                        self.assertTrue(resolved.is_file(), f"{current}: {location}")
                        pending.append(resolved)

    def test_compiles_and_validates_core_properties_offline(self):
        for skill in SKILLS:
            with self.subTest(skill=skill):
                xsd = schema_root(skill) / "ecma/fouth-edition/opc-coreProperties.xsd"
                parser = etree.XMLParser(no_network=True, resolve_entities=False, load_dtd=False)
                parser.resolvers.add(RejectRemoteResolver())
                schema = etree.XMLSchema(etree.parse(str(xsd.resolve()), parser))
                legal = etree.fromstring(
                    f'''<cp:coreProperties xmlns:cp="{CORE_NS}" xmlns:dc="{DC_NS}"
                        xmlns:dcterms="{DCTERMS_NS}" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                      <dc:title>Report</dc:title><dc:creator>Ada</dc:creator>
                      <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-02T03:04:05Z</dcterms:created>
                    </cp:coreProperties>'''
                )
                self.assertTrue(schema.validate(legal), str(schema.error_log))
                for invalid in (
                    f'<cp:coreProperties xmlns:cp="{CORE_NS}"><cp:lastPrinted>yesterday</cp:lastPrinted></cp:coreProperties>',
                    f'<cp:coreProperties xmlns:cp="{CORE_NS}"><cp:unknown>value</cp:unknown></cp:coreProperties>',
                ):
                    with self.subTest(skill=skill, xml=invalid):
                        self.assertFalse(schema.validate(etree.fromstring(invalid)))

    def test_validator_loader_uses_bundled_schema(self):
        for skill in SKILLS:
            with self.subTest(skill=skill):
                office = ROOT / skill / "scripts/office"
                module_path = office / "validators/base.py"
                module_name = f"{skill}_offline_base_test"
                added = str(office) not in sys.path
                if added:
                    sys.path.insert(0, str(office))
                stubs = {}
                try:
                    # The loader only needs these names at import time; keep this
                    # regression independent of each skill's runtime packages.
                    for name, value in (
                        ("defusedxml", types.ModuleType("defusedxml")),
                        ("defusedxml.minidom", types.ModuleType("defusedxml.minidom")),
                        ("helpers", types.ModuleType("helpers")),
                    ):
                        stubs[name] = sys.modules.get(name)
                        sys.modules[name] = value
                    sys.modules["helpers"].safe_extract = lambda *args, **kwargs: None
                    spec = importlib.util.spec_from_file_location(module_name, module_path)
                    module = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(module)
                    schema = module._load_schema(str(schema_root(skill) / "ecma/fouth-edition/opc-coreProperties.xsd"))
                    self.assertIsInstance(schema, etree.XMLSchema)
                finally:
                    sys.modules.pop(module_name, None)
                    for name, previous in stubs.items():
                        if previous is None:
                            sys.modules.pop(name, None)
                        else:
                            sys.modules[name] = previous
                    if added:
                        sys.path.remove(str(office))


if __name__ == "__main__":
    unittest.main()
