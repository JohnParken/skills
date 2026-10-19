import sys
import tempfile
import unittest
import zipfile
from contextlib import redirect_stderr
from io import StringIO
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "docx" / "scripts"))
from read_docx import main, read_docx


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def package(path, document, **parts):
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("word/document.xml", document)
        for name, xml in parts.items():
            zf.writestr("word/" + name, xml)


def doc(body):
    return f'<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>'


class ReadDocxTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / "sample.docx"

    def tearDown(self):
        self.tmp.cleanup()

    def test_paragraph_table_order_unicode_hyperlink_tabs_and_breaks(self):
        body = ('<w:p><w:r><w:t>café</w:t></w:r><w:hyperlink><w:r><w:t> link</w:t></w:r></w:hyperlink>'
                '<w:r><w:tab/></w:r><w:r><w:br/><w:t>終</w:t></w:r></w:p>'
                '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc>'
                '<w:tc><w:p><w:r><w:t>two</w:t></w:r></w:p><w:tbl><w:tr><w:tc>'
                '<w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc>'
                '</w:tr></w:tbl><w:p><w:r><w:t>end</w:t></w:r></w:p>')
        package(self.path, doc(body))
        self.assertEqual(read_docx(self.path), "café link\t\n終\ncell\ttwo\nnested\nend")

    def test_revision_modes(self):
        body = ('<w:p><w:r><w:t>A</w:t></w:r><w:ins><w:r><w:t>I</w:t></w:r></w:ins>'
                '<w:del><w:delText>D</w:delText></w:del><w:moveFrom><w:r><w:t>F</w:t></w:r></w:moveFrom>'
                '<w:moveTo><w:r><w:t>T</w:t></w:r></w:moveTo></w:p>')
        package(self.path, doc(body))
        self.assertEqual(read_docx(self.path), "AIT")
        self.assertEqual(read_docx(self.path, revisions="original"), "ADF")
        self.assertEqual(read_docx(self.path, revisions="all"), "AIDFT")

    def test_wrapped_blocks_and_block_level_revision_views(self):
        body = ('<w:sdt><w:sdtContent><w:p><w:r><w:t>content</w:t></w:r></w:p></w:sdtContent></w:sdt>'
                '<w:customXml><w:tbl><w:tr><w:tc><w:p><w:r><w:t>custom</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:customXml>'
                '<w:ins><w:p><w:r><w:t>inserted</w:t></w:r></w:p></w:ins>'
                '<w:del><w:p><w:r><w:t>deleted</w:t></w:r></w:p></w:del>')
        package(self.path, doc(body))
        self.assertEqual(read_docx(self.path), "content\ncustom\ninserted")
        self.assertEqual(read_docx(self.path, revisions="original"), "content\ncustom\ndeleted")
        self.assertEqual(read_docx(self.path, revisions="all"), "content\ncustom\ninserted\ndeleted")

    def test_related_parts_and_note_separators(self):
        body = '<w:p><w:r><w:t>main</w:t></w:r></w:p>'
        package(self.path, doc(body), **{
            "header1.xml": f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>head</w:t></w:r></w:p></w:hdr>',
            "footnotes.xml": f'<w:footnotes xmlns:w="{W}"><w:footnote w:type="separator"><w:p><w:r><w:t>skip</w:t></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:footnote></w:footnotes>',
            "comments.xml": f'<w:comments xmlns:w="{W}"><w:comment w:id="1"><w:p><w:r><w:t>comment</w:t></w:r></w:p></w:comment></w:comments>',
        })
        result = read_docx(self.path, include_related=True)
        self.assertEqual(result, "main\n[Headers: word/header1.xml]\nhead\n[Footnotes: word/footnotes.xml]\nnote\n[Comments: word/comments.xml]\ncomment")

    def test_malformed_zip_xml_and_missing_document_fail_clearly(self):
        self.path.write_bytes(b"not a zip")
        with self.assertRaisesRegex(ValueError, "cannot read DOCX package"):
            read_docx(self.path)
        with zipfile.ZipFile(self.path, "w") as zf:
            zf.writestr("word/document.xml", "<broken>")
        with self.assertRaisesRegex(ValueError, "invalid or unsafe XML"):
            read_docx(self.path)
        with zipfile.ZipFile(self.path, "w") as zf:
            zf.writestr("word/document.xml", '<!DOCTYPE x [<!ENTITY e "boom">]><w:document xmlns:w="' + W + '"><w:body><w:p><w:r><w:t>&e;</w:t></w:r></w:p></w:body></w:document>')
        with self.assertRaisesRegex(ValueError, "invalid or unsafe XML"):
            read_docx(self.path)
        with zipfile.ZipFile(self.path, "w") as zf:
            zf.writestr("word/document.xml", f'<w:document xmlns:w="{W}"/>')
        with self.assertRaisesRegex(ValueError, "missing w:body"):
            read_docx(self.path)
        with zipfile.ZipFile(self.path, "w") as zf:
            zf.writestr("other.xml", "<root/>")
        with self.assertRaisesRegex(ValueError, "missing word/document.xml"):
            read_docx(self.path)

    def test_cli_refuses_same_input_output_and_preserves_bytes(self):
        package(self.path, doc('<w:p><w:r><w:t>keep</w:t></w:r></w:p>'))
        before = self.path.read_bytes()
        with redirect_stderr(StringIO()) as err, self.assertRaises(SystemExit) as raised:
            main([str(self.path), "-o", str(self.path)])
        self.assertEqual(raised.exception.code, 2)
        self.assertIn("output must not refer to the input file", err.getvalue())
        self.assertEqual(self.path.read_bytes(), before)

        alias = self.path.with_name("alias.docx")
        alias.symlink_to(self.path)
        with redirect_stderr(StringIO()), self.assertRaises(SystemExit):
            main([str(self.path), "-o", str(alias)])
        self.assertEqual(self.path.read_bytes(), before)

    def test_cli_reports_output_write_errors(self):
        package(self.path, doc('<w:p><w:r><w:t>keep</w:t></w:r></w:p>'))
        output_dir = Path(self.tmp.name) / "folder.docx"
        output_dir.mkdir()
        with redirect_stderr(StringIO()) as err, self.assertRaises(SystemExit) as raised:
            main([str(self.path), "-o", str(output_dir)])
        self.assertEqual(raised.exception.code, 2)
        self.assertIn("Is a directory", err.getvalue())


if __name__ == "__main__":
    unittest.main()
