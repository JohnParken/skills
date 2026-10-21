"""Constructed OOXML fixtures for explicit table and numbering invariants."""

import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "docx/scripts"))
from office.package import unpack_document
from office.validators.structure import StructureValidator
from test_docx_stdlib_core import R, REL, W, package_parts, write_package


def cell(properties="", content="<w:p/>"):
    return f"<w:tc><w:tcPr>{properties}</w:tcPr>{content}</w:tc>"


def row(cells, properties=""):
    return f"<w:tr><w:trPr>{properties}</w:trPr>{cells}</w:tr>"


def table(rows, columns=3):
    grid = "" if columns is None else "<w:tblGrid>" + '<w:gridCol w:w="1000"/>' * columns + "</w:tblGrid>"
    return f"<w:tbl>{grid}{rows}</w:tbl>"


def level(index, properties=""):
    return f'<w:lvl w:ilvl="{index}"><w:start w:val="1"/>{properties}</w:lvl>'


class AdvancedStructureTests(unittest.TestCase):
    def validate(self, body="<w:p/>", numbering=None):
        with tempfile.TemporaryDirectory(prefix="docx-advanced-structure-") as temporary:
            directory = Path(temporary)
            extra = {}
            if numbering is not None:
                extra = {
                    "word/numbering.xml": f'<w:numbering xmlns:w="{W}">{numbering}</w:numbering>',
                    "word/_rels/document.xml.rels": (
                        f'<Relationships xmlns="{REL}"><Relationship Id="n" '
                        f'Type="{R}/numbering" Target="numbering.xml"/></Relationships>'
                    ),
                }
            document = f'<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>'
            archive = directory / "input.docx"
            write_package(archive, package_parts(document, extra=extra))
            unpacked = unpack_document(archive, directory / "unpacked")
            validator = StructureValidator(unpacked, extension=".docx")
            with contextlib.redirect_stdout(io.StringIO()):
                success = validator.validate()
            return success, validator

    def assert_invalid(self, body="<w:p/>", numbering=None, contains=()):
        success, validator = self.validate(body, numbering)
        self.assertFalse(success, "Fixture should fail an explicit structural invariant")
        report = "\n".join(validator.errors)
        for expected in contains:
            self.assertIn(expected, report)

    def test_valid_spanned_vertical_merges_and_omitted_cells(self):
        first = row(cell('<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>') + cell())
        second = row(cell('<w:gridSpan w:val="2"/><w:vMerge/>') + cell())
        third = row(cell(), '<w:gridBefore w:val="1"/><w:gridAfter w:val="1"/>')
        self.assertTrue(self.validate(table(first + second + third))[0])

    def test_orphan_and_changed_span_vertical_merges_report_cell_location(self):
        for rows in (
            row(cell('<w:vMerge/>')),
            row(cell('<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>'))
            + row(cell('<w:vMerge w:val="continue"/>')),
        ):
            with self.subTest(rows=rows):
                self.assert_invalid(table(rows), contains=("word/document.xml", "tbl[1]", "tr[", "tc[1]", "vMerge"))

    def test_missing_vmerge_ends_chain_before_later_continue(self):
        rows = row(cell('<w:vMerge w:val="restart"/>')) + row(cell()) + row(cell('<w:vMerge/>'))
        self.assert_invalid(table(rows), contains=("tr[3]", "vMerge"))

    def test_grid_before_changes_column_position_and_prevents_false_merge(self):
        rows = row(cell('<w:vMerge w:val="restart"/>'), '<w:gridBefore w:val="1"/>')
        rows += row(cell('<w:vMerge/>'))
        self.assert_invalid(table(rows), contains=("tr[2]", "vMerge"))

    def test_invalid_grid_values_and_duplicate_properties_fail(self):
        for cell_props, row_props in (
            ('<w:gridSpan w:val="0"/>', ""),
            ('<w:gridSpan w:val="x"/>', ""),
            ('<w:gridSpan w:val="1"/><w:gridSpan w:val="2"/>', ""),
            ("", '<w:gridBefore w:val="-1"/>'),
            ("", '<w:gridAfter w:val="NaN"/>'),
            ('<w:vMerge w:val="invalid"/>', ""),
        ):
            with self.subTest(cell_props=cell_props, row_props=row_props):
                self.assert_invalid(table(row(cell(cell_props), row_props)), contains=("tbl[1]", "tr[1]"))

    def test_non_xml_whitespace_in_numeric_values_is_rejected(self):
        for padding in ("\u3000", "\u00a0"):
            for value in (padding + "1", "1" + padding):
                with self.subTest(value=value, field="gridSpan"):
                    self.assert_invalid(table(row(cell(f'<w:gridSpan w:val="{value}"/>'))),
                                        contains=("gridSpan", "expected an integer"))
                with self.subTest(value=value, field="ilvl"):
                    self.assert_invalid(
                        numbering='<w:abstractNum w:abstractNumId="0">'
                        + level(value) + '</w:abstractNum>',
                        contains=("ilvl", "expected an integer"),
                    )
                with self.subTest(value=value, field="lvlRestart"):
                    self.assert_invalid(
                        numbering='<w:abstractNum w:abstractNumId="0">'
                        + level(1, f'<w:lvlRestart w:val="{value}"/>') + '</w:abstractNum>',
                        contains=("lvlRestart", "expected an integer"),
                    )

    def test_xml_numeric_whitespace_and_leading_zeros_remain_supported(self):
        for value in (" 001 ", "\t001\r\n", "&#9;001&#13;&#10;"):
            with self.subTest(value=value):
                self.assertTrue(self.validate(table(row(cell(f'<w:gridSpan w:val="{value}"/>'))))[0])
                numbering = '<w:abstractNum w:abstractNumId="0">' + level(
                    value, f'<w:lvlRestart w:val="{value}"/>',
                ) + '</w:abstractNum>'
                self.assertTrue(self.validate(numbering=numbering)[0])

    def test_nested_table_merge_state_is_separate(self):
        inner = table(row(cell('<w:vMerge/>')), 1)
        outer = table(row(cell('<w:vMerge w:val="restart"/>', inner + '<w:p/>')), 1)
        self.assert_invalid(outer, contains=("tbl[2]", "tr[1]", "tc[1]", "vMerge"))

    def test_implicit_or_augmented_grid_is_not_rejected(self):
        rows = row(cell('<w:gridSpan w:val="4"/><w:vMerge w:val="restart"/>'))
        rows += row(cell('<w:gridSpan w:val="4"/><w:vMerge/>'))
        for columns in (None, 1):
            with self.subTest(columns=columns):
                self.assertTrue(self.validate(table(rows, columns))[0])

    def test_ignored_grid_offsets_are_not_rejected(self):
        rows = row(cell(), '<w:gridBefore w:val="7"/><w:gridAfter w:val="99"/>')
        self.assertTrue(self.validate(table(rows, 1))[0])

    def test_grid_before_uses_grid_augmented_by_preceding_row(self):
        rows = row(cell('<w:gridSpan w:val="2"/>')
                   + cell('<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>'))
        rows += row(cell('<w:gridSpan w:val="2"/><w:vMerge/>'), '<w:gridBefore w:val="2"/>')
        self.assertTrue(self.validate(table(rows, 1))[0])

    def test_large_numeric_values_do_not_allocate_per_column(self):
        body = table(row(cell('<w:gridSpan w:val="2147483647"/>')), 1)
        self.assertTrue(self.validate(body)[0])
        self.assert_invalid(table(row(cell('<w:gridSpan w:val="99999999999"/>'))),
                            contains=("gridSpan", "32-bit"))
        zeros = "0" * 5000 + "1"
        self.assertTrue(self.validate(table(row(cell(f'<w:gridSpan w:val="{zeros}"/>'))))[0])

    def test_legacy_hmerge_is_reported_as_unverified_without_false_error(self):
        body = table(row(cell('<w:hMerge w:val="restart"/><w:vMerge/>')))
        success, validator = self.validate(body)
        self.assertTrue(success)
        self.assertIn("hMerge", " ".join(validator.warnings))

    def test_wrapped_table_rows_have_explicit_unverified_diagnostic(self):
        body = table('<w:sdt><w:sdtContent>' + row(cell('<w:vMerge/>')) + '</w:sdtContent></w:sdt>')
        success, validator = self.validate(body)
        self.assertTrue(success)
        self.assertIn("unverified", " ".join(validator.warnings))

    def test_duplicate_or_out_of_range_numbering_levels_fail(self):
        for levels in (level(0) + level(0), level(9), level(-1), level("中文")):
            with self.subTest(levels=levels):
                self.assert_invalid(numbering=f'<w:abstractNum w:abstractNumId="3">{levels}</w:abstractNum>',
                                    contains=("word/numbering.xml", "abstractNum", "ilvl"))

    def test_override_indices_and_start_override_values_fail(self):
        for override in (
            '<w:lvlOverride w:ilvl="9"/>',
            '<w:lvlOverride w:ilvl="0"/><w:lvlOverride w:ilvl="0"/>',
            '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="x"/></w:lvlOverride>',
            '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/><w:startOverride w:val="2"/></w:lvlOverride>',
        ):
            with self.subTest(override=override):
                numbering = f'<w:abstractNum w:abstractNumId="3">{level(0)}</w:abstractNum>'
                numbering += f'<w:num w:numId="5"><w:abstractNumId w:val="3"/>{override}</w:num>'
                self.assert_invalid(numbering=numbering, contains=("word/numbering.xml", "num[", "lvlOverride"))

    def test_invalid_restart_value_fails_and_ignored_later_level_does_not(self):
        bad = '<w:abstractNum w:abstractNumId="0">' + level(1, '<w:lvlRestart w:val="x"/>') + '</w:abstractNum>'
        self.assert_invalid(numbering=bad, contains=("lvlRestart",))
        for restart in (0, 1, 9):
            good = '<w:abstractNum w:abstractNumId="0">' + level(1, f'<w:lvlRestart w:val="{restart}"/>') + '</w:abstractNum>'
            self.assertTrue(self.validate(numbering=good)[0])

    def test_direct_numbering_reference_requires_defined_level(self):
        numbering = f'<w:abstractNum w:abstractNumId="0">{level(0)}</w:abstractNum>'
        numbering += '<w:num w:numId="5"><w:abstractNumId w:val="0"/></w:num>'
        body = '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="5"/></w:numPr></w:pPr></w:p>'
        self.assert_invalid(body, numbering, contains=("word/document.xml", "p[1]", "numId=5", "ilvl=1"))

    def test_paragraph_level_range_and_duplicate_ilvl_fail(self):
        numbering = f'<w:abstractNum w:abstractNumId="0">{level(0)}</w:abstractNum>'
        numbering += '<w:num w:numId="5"><w:abstractNumId w:val="0"/></w:num>'
        for values in ('<w:ilvl w:val="9"/>', '<w:ilvl w:val="0"/><w:ilvl w:val="1"/>'):
            body = f'<w:p><w:pPr><w:numPr>{values}<w:numId w:val="5"/></w:numPr></w:pPr></w:p>'
            self.assert_invalid(body, numbering, contains=("word/document.xml", "ilvl"))

    def test_override_can_supply_level_and_signed_start(self):
        numbering = f'<w:abstractNum w:abstractNumId="0">{level(0)}</w:abstractNum>'
        numbering += ('<w:num w:numId="5"><w:abstractNumId w:val="0"/>'
                      '<w:lvlOverride w:ilvl="1"><w:startOverride w:val="-2"/>'
                      + level(1) + '</w:lvlOverride></w:num>')
        body = '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="5"/></w:numPr></w:pPr></w:p>'
        self.assertTrue(self.validate(body, numbering)[0])

    def test_style_linked_numbering_is_explicitly_unverified_not_missing_level(self):
        numbering = ('<w:abstractNum w:abstractNumId="0"><w:numStyleLink w:val="List"/></w:abstractNum>'
                     '<w:num w:numId="5"><w:abstractNumId w:val="0"/></w:num>')
        body = '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="5"/></w:numPr></w:pPr></w:p>'
        success, validator = self.validate(body, numbering)
        self.assertTrue(success)
        self.assertIn("numStyleLink", " ".join(validator.warnings))


if __name__ == "__main__":
    unittest.main()
