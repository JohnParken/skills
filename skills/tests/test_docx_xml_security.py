"""Small-limit regressions for XML allocation budgets; never stress the host."""

import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "docx/scripts"))
from office import xmlio


class XmlSecurityTests(unittest.TestCase):
    def rejected(self, data, setting, limit, diagnostic):
        for parser in (xmlio.fromstring, xmlio.parse_dom):
            with self.subTest(parser=parser.__name__, setting=setting):
                with mock.patch.object(xmlio, setting, limit, create=True):
                    with self.assertRaisesRegex(ValueError, diagnostic):
                        parser(data)

    def test_namespace_uri_is_bounded_before_expansion(self):
        data = '<r xmlns:p="urn:' + 'x' * 40 + '"><p:a/></r>'
        self.rejected(data, "MAX_XML_NAMESPACE_URI_CHARS", 32, "namespace URI")

    def test_expanded_element_names_have_a_cumulative_budget(self):
        data = '<r xmlns:p="urn:' + 'x' * 24 + '">'
        data += ''.join(f'<p:item{index}/>' for index in range(8)) + '</r>'
        self.rejected(data, "MAX_XML_EXPANDED_NAME_CHARS", 180, "expanded name")

    def test_expanded_attribute_names_have_a_cumulative_budget(self):
        data = '<r xmlns:p="urn:' + 'x' * 24 + '" '
        data += ' '.join(f'p:a{index}="v"' for index in range(8)) + '/>'
        self.rejected(data, "MAX_XML_EXPANDED_NAME_CHARS", 180, "expanded name")

    def test_namespace_declarations_are_included_in_the_name_budget(self):
        data = '<r ' + ' '.join(f'xmlns:p{index}="urn:namespace{index}"'
                                  for index in range(6)) + '/>'
        self.rejected(data, "MAX_XML_EXPANDED_NAME_CHARS", 100, "expanded name")

    def test_attribute_limits_include_namespace_declarations(self):
        self.rejected('<r xmlns:p="urn:p" a="1" b="2" c="3"/>',
                      "MAX_XML_ATTRIBUTES_PER_ELEMENT", 3, "attribute")
        self.rejected('<r a="1"><c b="2" c="3" d="4"/></r>',
                      "MAX_XML_ATTRIBUTES", 3, "attribute")

    def test_dom_node_budget_includes_attributes(self):
        self.rejected('<r a="1" b="2" c="3"/>', "MAX_XML_NODES", 3, "node")

    def test_dom_node_budget_includes_non_element_content(self):
        cases = ('<r>' + '<!--keep-->' * 5 + '</r>',
                 '<r>' + '<?keep data?>' * 5 + '</r>',
                 '<r>' + '<![CDATA[text]]>' * 5 + '</r>',
                 '<r>A<c/>B<c/>C</r>')
        for data in cases:
            with self.subTest(data=data):
                self.rejected(data, "MAX_XML_NODES", 5, "node")

    def test_contiguous_text_callbacks_do_not_create_false_node_failures(self):
        data = '<r>' + '&#65;' * 20 + '</r>'
        with mock.patch.object(xmlio, "MAX_XML_NODES", 2, create=True):
            self.assertEqual(xmlio.fromstring(data).text, 'A' * 20)
            self.assertEqual(xmlio.parse_dom(data).documentElement.firstChild.data, 'A' * 20)

    def test_normal_namespaces_encodings_and_miscellaneous_nodes_remain_supported(self):
        data = ('<?xml version="1.0"?><document xmlns="urn:doc" xmlns:q="urn:doc">'
                '<!--keep--><?keep data?><p q:align="left"><![CDATA[中文 & text]]></p>'
                '</document>')
        for encoding in ("utf-8", "utf-16"):
            with self.subTest(encoding=encoding):
                payload = data.replace('<?xml version="1.0"?>',
                                       f'<?xml version="1.0" encoding="{encoding}"?>').encode(encoding)
                self.assertEqual(xmlio.fromstring(payload)[0].text, '中文 & text')
                self.assertEqual(xmlio.parse_dom(payload).documentElement.namespaceURI, 'urn:doc')

    def test_shared_budgets_stop_multiple_individually_valid_xml_parts(self):
        cases = (
            ("MAX_XML_ELEMENTS", 3, '<r><a/></r>', "element"),
            ("MAX_XML_ATTRIBUTES", 3, '<r a="1" b="2"/>', "attribute"),
            ("MAX_XML_NODES", 3, '<r>A</r>', "node"),
            ("MAX_XML_EXPANDED_NAME_CHARS", 32, '<p:r xmlns:p="urn:x"/>', "expanded"),
        )
        for setting, limit, data, diagnostic in cases:
            for parser in (xmlio._checked, xmlio.fromstring, xmlio.parse_dom):
                with self.subTest(parser=parser.__name__, setting=setting):
                    with mock.patch.object(xmlio, setting, limit):
                        parser(data)
                        parser(data)
                        budget = xmlio.XmlBudget()
                        parser(data, budget=budget)
                        with self.assertRaisesRegex(ValueError, "shared.*" + diagnostic):
                            parser(data, budget=budget)

    def test_file_parser_forwards_shared_budget(self):
        with tempfile.TemporaryDirectory(prefix="docx-xml-budget-") as temporary:
            first, second = Path(temporary) / "first.xml", Path(temporary) / "second.xml"
            first.write_text('<r><a/></r>', encoding="utf-8")
            second.write_text('<r><b/></r>', encoding="utf-8")
            with mock.patch.object(xmlio, "MAX_XML_ELEMENTS", 3):
                xmlio.parse(first)
                xmlio.parse(second)
                budget = xmlio.XmlBudget()
                xmlio.parse(first, budget=budget)
                with self.assertRaisesRegex(ValueError, "shared.*element"):
                    xmlio.parse(second, budget=budget)

    def test_shared_budget_never_loosens_per_xml_limits(self):
        budget = xmlio.XmlBudget()
        with mock.patch.object(xmlio, "MAX_XML_ATTRIBUTES_PER_ELEMENT", 1):
            with self.assertRaisesRegex(ValueError, "attribute"):
                xmlio.fromstring('<r a="1" b="2"/>', budget=budget)
        with self.assertRaisesRegex(ValueError, "XmlBudget"):
            xmlio.fromstring('<r/>', budget={})


if __name__ == "__main__":
    unittest.main()
