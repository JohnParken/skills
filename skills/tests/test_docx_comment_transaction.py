"""Failure injection for low-level comment directory publication and recovery."""

import sys
import io
import contextlib
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "docx/scripts"))
import comment
from create_docx import create_document
from office.package import unpack_document


class CommentTransactionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-comment-transaction-test-")
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name)
        source = create_document({"blocks": [{"type": "paragraph", "text": "正文"}]}, self.work / "source.docx")
        self.directory = unpack_document(source, self.work / "package")
        (self.directory / "opaque.bin").write_bytes(b"unknown bytes\x00unchanged")

    def snapshot(self, directory=None):
        return {path.relative_to(directory or self.directory).as_posix(): path.read_bytes()
                for path in (directory or self.directory).rglob("*") if path.is_file()}

    def test_late_comment_append_failure_does_not_publish_any_partial_part(self):
        before = self.snapshot()
        append = comment._append_xml
        calls = 0

        def failing_append(*args, **kwargs):
            nonlocal calls
            calls += 1
            if calls == 3:
                raise OSError("injected third metadata write failure")
            return append(*args, **kwargs)

        with mock.patch.object(comment, "_append_xml", failing_append):
            with self.assertRaisesRegex(OSError, "third metadata"):
                comment.add_comment(self.directory, "新批注")
        self.assertEqual(self.snapshot(), before)
        self.assertEqual(sorted(path.name for path in self.work.iterdir()), ["package", "source.docx"])

    def test_success_preserves_non_target_members_and_existing_comments(self):
        before = self.snapshot()
        first = comment.add_comment(self.directory, "父批注", author="甲")
        second = comment.add_comment(self.directory, "回复", parent_id=first[0], author="乙")
        self.assertEqual((first[0], second[0]), (0, 1))
        changed = {"[Content_Types].xml", "word/_rels/document.xml.rels", "word/comments.xml",
                   "word/commentsExtended.xml", "word/commentsIds.xml", "word/commentsExtensible.xml"}
        for name, data in before.items():
            if name not in changed:
                self.assertEqual((self.directory / name).read_bytes(), data)
        self.assertEqual(sorted(path.name for path in self.work.iterdir()), ["package", "source.docx"])

    def test_copy_failure_leaves_source_and_no_transaction_artifacts(self):
        before = self.snapshot()
        with mock.patch.object(comment.shutil, "copytree", side_effect=OSError("copy failed")):
            with self.assertRaisesRegex(OSError, "copy failed"):
                comment.add_comment(self.directory, "批注")
        self.assertEqual(self.snapshot(), before)
        self.assertFalse(list(self.work.glob(".package.comment-*")))

    def test_failed_backup_or_install_restores_original_directory(self):
        replace = comment.os.replace
        for failing_move in (1, 2):
            with self.subTest(failing_move=failing_move):
                before = self.snapshot()
                calls = 0

                def failing_replace(source, destination):
                    nonlocal calls
                    calls += 1
                    if calls == failing_move:
                        raise OSError("injected directory move failure")
                    return replace(source, destination)

                with mock.patch.object(comment.os, "replace", failing_replace):
                    with self.assertRaisesRegex(OSError, "directory move"):
                        comment.add_comment(self.directory, "批注")
                self.assertEqual(self.snapshot(), before)
                self.assertFalse(list(self.work.glob(".package.comment-*")))

    def test_failed_restore_keeps_original_backup_and_reports_exact_location(self):
        before = self.snapshot()
        replace = comment.os.replace
        calls = 0

        def failing_replace(source, destination):
            nonlocal calls
            calls += 1
            if calls >= 2:
                raise OSError("install and restore unavailable")
            return replace(source, destination)

        with mock.patch.object(comment.os, "replace", failing_replace):
            with self.assertRaisesRegex(RuntimeError, "rollback failed") as failure:
                comment.add_comment(self.directory, "批注")
        self.assertFalse(self.directory.exists())
        transactions = list(self.work.glob(".package.comment-*"))
        self.assertEqual(len(transactions), 1)
        backup = transactions[0] / "backup"
        self.assertIn(str(backup), str(failure.exception))
        self.assertEqual(self.snapshot(backup), before)
        replace(backup, self.directory)

    def test_directory_containing_working_directory_is_rejected(self):
        before = self.snapshot()
        with mock.patch.object(comment.Path, "cwd", return_value=self.directory):
            with self.assertRaisesRegex(ValueError, "current working directory"):
                comment.add_comment(self.directory, "批注")
        self.assertEqual(self.snapshot(), before)

    def test_package_cli_planning_failure_preserves_input_and_existing_output(self):
        source = self.work / "source.docx"
        original = source.read_bytes()
        output = self.work / "output.docx"
        output.write_bytes(b"keep existing output")
        with mock.patch.object(comment, "_append_xml", side_effect=OSError("metadata unavailable")):
            with mock.patch.object(sys, "argv", ["comment.py", str(source), "批注", "-o", str(output)]):
                with contextlib.redirect_stderr(io.StringIO()) as errors:
                    with self.assertRaises(SystemExit) as exit_status:
                        comment.main()
        self.assertEqual(exit_status.exception.code, 1)
        self.assertIn("metadata unavailable", errors.getvalue())
        self.assertEqual(source.read_bytes(), original)
        self.assertEqual(output.read_bytes(), b"keep existing output")

    def test_success_with_backup_cleanup_failure_reports_completed_install(self):
        before = self.snapshot()
        remove = comment.shutil.rmtree

        def failing_cleanup(path, *args, **kwargs):
            if Path(path).name == "backup":
                raise OSError("backup cleanup unavailable")
            return remove(path, *args, **kwargs)

        with mock.patch.object(comment.shutil, "rmtree", failing_cleanup):
            result = comment.add_comment(self.directory, "已安装")
        self.assertIn("installed, cleanup backup retained", result[2])
        backup = next(self.work.glob(".package.comment-*/backup"))
        self.assertIn(str(backup), result[2])
        self.assertEqual(self.snapshot(backup), before)
        self.assertTrue((self.directory / "word/comments.xml").is_file())

    def test_new_metadata_crossing_package_limit_is_rejected_before_install(self):
        before = self.snapshot()
        total = sum(len(data) for data in before.values())
        with mock.patch.object(comment, "MAX_PACKAGE_BYTES", total + 1):
            with self.assertRaisesRegex(ValueError, "size limits"):
                comment.add_comment(self.directory, "批注")
        self.assertEqual(self.snapshot(), before)
        self.assertFalse(list(self.work.glob(".package.comment-*")))

    def test_package_cli_rejects_output_links_and_cross_format_names(self):
        source = self.work / "source.docx"
        original = source.read_bytes()
        target = self.work / "target.docx"
        target.write_bytes(b"keep target")
        alias = self.work / "alias.docx"
        alias.symlink_to(target)
        for output in (alias, self.work / "other.dotx"):
            with self.subTest(output=output):
                with mock.patch.object(sys, "argv", ["comment.py", str(source), "批注", "-o", str(output)]):
                    with contextlib.redirect_stderr(io.StringIO()):
                        with self.assertRaises(SystemExit) as failure:
                            comment.main()
                self.assertEqual(failure.exception.code, 1)
                self.assertEqual(source.read_bytes(), original)
                self.assertEqual(target.read_bytes(), b"keep target")

    def test_restore_that_completes_before_error_reports_original_install_failure(self):
        before = self.snapshot()
        replace = comment.os.replace
        calls = 0

        def late_failure(source, destination):
            nonlocal calls
            calls += 1
            if calls == 2:
                raise OSError("original install failure")
            result = replace(source, destination)
            if calls == 3:
                raise OSError("late restore error after completed rename")
            return result

        with mock.patch.object(comment.os, "replace", late_failure):
            with self.assertRaisesRegex(OSError, "original install failure"):
                comment.add_comment(self.directory, "批注")
        self.assertEqual(self.snapshot(), before)
        self.assertFalse(list(self.work.glob(".package.comment-*")))

    def test_transaction_cleanup_failure_does_not_hide_successful_install(self):
        remove = comment.shutil.rmtree

        def failing_cleanup(path, *args, **kwargs):
            if Path(path).name.startswith(".package.comment-"):
                raise OSError("transaction cleanup unavailable")
            return remove(path, *args, **kwargs)

        with mock.patch.object(comment.shutil, "rmtree", failing_cleanup):
            result = comment.add_comment(self.directory, "已安装")
        transaction = next(self.work.glob(".package.comment-*"))
        self.assertIn("installed", result[2])
        self.assertIn(str(transaction), result[2])
        self.assertTrue((self.directory / "word/comments.xml").is_file())


if __name__ == "__main__":
    unittest.main()
