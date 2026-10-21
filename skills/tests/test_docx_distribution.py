"""Keep the local DOCX distribution equal to source and usable after extraction."""

import json
import hashlib
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ARCHIVE = ROOT / "docx.zip"


class DocxDistributionTests(unittest.TestCase):
    def test_distribution_matches_source_without_metadata_or_caches(self):
        expected = {path.relative_to(ROOT).as_posix(): path.read_bytes()
                    for path in (ROOT / "docx").rglob("*") if path.is_file()}
        with zipfile.ZipFile(ARCHIVE) as package:
            self.assertIsNone(package.testzip())
            names = package.namelist()
            self.assertEqual(len(names), len(set(names)))
            self.assertEqual(set(names), set(expected))
            for name in names:
                with self.subTest(member=name):
                    self.assertNotIn(".DS_Store", Path(name).parts)
                    self.assertNotIn("__MACOSX", Path(name).parts)
                    self.assertNotIn("__pycache__", Path(name).parts)
                    self.assertFalse(name.endswith(".pyc"))
                    self.assertEqual(package.read(name), expected[name])

    def test_extracted_skill_chinese_generation_editing_and_reading(self):
        with tempfile.TemporaryDirectory(prefix="docx-distribution-test-") as temp:
            work = Path(temp)
            with zipfile.ZipFile(ARCHIVE) as package:
                package.extractall(work)
            scripts = work / "docx/scripts"
            from test_docx_generator import png_chunk, tiny_png
            image = tiny_png()
            (work / "pixel.png").write_bytes(image)
            replacement = image[:-12] + png_chunk(b"tEXt", b"test\x00replacement") + image[-12:]
            (work / "replacement.png").write_bytes(replacement)

            def cli(script, *args):
                result = subprocess.run(
                    [sys.executable, "-I", "-S", "-B", str(scripts / script), *args],
                    cwd=work, env=dict(os.environ, PATH=""), capture_output=True, text=True,
                )
                self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
                return result.stdout

            self.assertIn("check-ok", cli("create_docx.py", "--check"))
            spec = {"blocks": [
                {"type": "heading", "level": 1, "text": "中文公文"},
                {"type": "heading", "level": 4, "text": "四级标题"},
                {"type": "paragraph", "text": "{{ 公司名称 }} “中英混排”"},
                {"type": "table", "header_rows": 2,
                 "format": {"borders": {"top": {"width_pt": 1, "color": "123456"}}},
                 "rows": [["项目", "状态"], ["子项目", "子状态"], ["测试", "通过"]]},
                {"type": "image", "path": "pixel.png", "width_mm": 10, "height_mm": 10, "alt": "测试图"},
                {"type": "list", "list_id": "实例", "level": 0, "ordered": True, "items": ["父项"]},
                {"type": "list", "list_id": "实例", "level": 1, "ordered": True, "items": ["子项甲"]},
                {"type": "paragraph", "text": "列表说明"},
                {"type": "list", "list_id": "实例", "level": 1, "ordered": True, "items": ["子项乙"]},
                {"type": "table", "rows": [[{"text": "单元格图片", "image": {
                    "path": "pixel.png", "width_mm": 10, "height_mm": 10, "alt": "单元格图"},
                    "borders": {"top": {"color": "AA0000", "width_pt": 1.5}}}]]},
            ]}
            (work / "spec.json").write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
            rules = {"replacements": [{"old": "{{ 公司名称 }}", "new": "中文 & English"}]}
            (work / "rules.json").write_text(json.dumps(rules, ensure_ascii=False), encoding="utf-8")
            cli("create_docx.py", "spec.json", "output.docx")
            cli("edit_docx.py", "output.docx", "rules.json", "filled.docx", "--require-no-placeholders")
            self.assertIn("PASSED", cli("office/validate.py", "filled.docx", "--require-no-placeholders"))
            text = cli("read_docx.py", "filled.docx")
            for expected in ("中文公文", "四级标题", "中文 & English “中英混排”", "项目\t状态", "测试\t通过"):
                self.assertIn(expected, text)
            operations = {"operations": [
                {"type": "insert_paragraph", "paragraph_index": 2,
                 "expected_text": "中文 & English “中英混排”", "position": "after",
                 "text": "追加说明", "style": "BodyText"},
                {"type": "format_run", "paragraph_index": 2,
                 "expected_text": "中文 & English “中英混排”", "run_index": 0,
                 "expected_run_text": "中文 & English “中英混排”", "format": {"bold": True}},
            ]}
            (work / "operations.json").write_text(json.dumps(operations, ensure_ascii=False), encoding="utf-8")
            cli("edit_structure.py", "filled.docx", "operations.json", "structured.docx")
            self.assertIn("PASSED", cli("office/validate.py", "structured.docx"))
            self.assertIn("追加说明", cli("read_docx.py", "structured.docx"))
            structure = json.loads(cli("read_docx.py", "structured.docx", "--format", "json"))
            self.assertEqual(structure["schema_version"], 1)
            self.assertEqual(structure["parts"][0]["name"], "word/document.xml")
            row_operation = {"operation": {"type": "insert_row", "table_index": 0,
                "row_index": 2, "expected_rows": 3, "expected_text": ["测试", "通过"],
                "position": "after", "cells": ["新增", "待复核"]}}
            (work / "row.json").write_text(json.dumps(row_operation, ensure_ascii=False), encoding="utf-8")
            cli("edit_tables.py", "structured.docx", "row.json", "rows.docx")
            self.assertIn("新增\t待复核", cli("read_docx.py", "rows.docx"))
            with zipfile.ZipFile(work / "rows.docx") as package:
                media = next(name for name in package.namelist() if name.startswith('word/media/'))
                digest = hashlib.sha256(package.read(media)).hexdigest()
            selection = {"media_part": media, "expected_sha256": digest,
                         "expected_relationships": 1, "expected_references": 1}
            (work / "selection.json").write_text(json.dumps(selection), encoding="utf-8")
            cli("replace_image.py", "rows.docx", "replacement.png", "selection.json", "media.docx")
            self.assertIn("PASSED", cli("office/validate.py", "media.docx"))
            with zipfile.ZipFile(work / "rows.docx") as before_media, zipfile.ZipFile(work / "media.docx") as after_media:
                self.assertEqual(after_media.read(media), replacement)
                for name in before_media.namelist():
                    if name != media:
                        self.assertEqual(before_media.read(name), after_media.read(name))
            def comments(source, target, action):
                (work / "comment-action.json").write_text(json.dumps(action, ensure_ascii=False), encoding="utf-8")
                args = (source, "comment-action.json") + ((target,) if target is not None else ())
                return json.loads(cli("manage_comments.py", *args))

            added = comments("media.docx", "commented.docx", {"action": "add", "paragraph_index": 3,
                "expected_text": "追加说明", "anchor_text": "追加说明", "text": "请复核", "author": "审阅人"})
            parent = added["comment_id"]
            listed = comments("commented.docx", None, {"action": "list"})
            self.assertTrue(listed["mutable"], listed["diagnostics"])
            self.assertEqual(len(listed["comments"]), 1)
            replied = comments("commented.docx", "replied.docx", {"action": "reply", "id": parent,
                "expected_text": "请复核", "text": "已核验", "author": "经办人"})
            comments("replied.docx", "updated.docx", {"action": "update", "id": parent,
                "expected_text": "请复核", "text": "请补证据"})
            comments("updated.docx", "reply-deleted.docx", {"action": "delete", "id": replied["comment_id"],
                "expected_text": "已核验"})
            comments("reply-deleted.docx", "comments-deleted.docx", {"action": "delete", "id": parent,
                "expected_text": "请补证据"})
            self.assertEqual(comments("comments-deleted.docx", None, {"action": "list"})["comments"], [])
            self.assertIn("PASSED", cli("office/validate.py", "comments-deleted.docx"))
            with zipfile.ZipFile(work / "output.docx") as before, zipfile.ZipFile(work / "filled.docx") as after:
                self.assertIsNone(after.testzip())
                self.assertEqual(set(before.namelist()), set(after.namelist()))
                for name in before.namelist():
                    if name != "word/document.xml":
                        self.assertEqual(before.read(name), after.read(name))


if __name__ == "__main__":
    unittest.main()
