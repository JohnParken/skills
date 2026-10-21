"""Tests for the DOCX generator's local runtime diagnostic."""

import contextlib
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "docx" / "scripts" / "create_docx.py"
sys.path.insert(0, str(SCRIPT.parent))
import create_docx  # noqa: E402


class RuntimeCheckTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="docx-runtime-check-")
        self.addCleanup(self.temp.cleanup)
        self.work = Path(self.temp.name)
        self.env = dict(os.environ, PATH="", PYTHONDONTWRITEBYTECODE="1")

    def run_cli(self, *args):
        return subprocess.run(
            [sys.executable, "-I", "-S", "-B", str(SCRIPT), *map(str, args)],
            cwd=self.work, env=self.env, capture_output=True, text=True,
        )

    def test_check_succeeds_without_external_programs_or_leftovers(self):
        result = self.run_cli("--check")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("output_dir=. writable=True", result.stdout)
        self.assertIn("check-ok", result.stdout)
        self.assertEqual(list(self.work.iterdir()), [])

    def test_check_rejects_positionals_and_generation_requires_both(self):
        for args in ((), ("input.json",), ("--check", "input.json"),
                     ("--check", "input.json", "output.docx")):
            with self.subTest(args=args):
                self.assertEqual(self.run_cli(*args).returncode, 2)

    def test_probe_denial_is_sanitized_and_fails_without_check_ok(self):
        output = io.StringIO()
        errors = io.StringIO()
        with mock.patch.object(create_docx.tempfile, "TemporaryFile", side_effect=OSError("secret/path")):
            with contextlib.redirect_stdout(output), contextlib.redirect_stderr(errors):
                result = create_docx.main(["--check"])
        self.assertEqual(result, 1)
        self.assertIn("output_dir=. writable=False", output.getvalue())
        self.assertNotIn("check-ok", output.getvalue())
        self.assertNotIn("secret/path", errors.getvalue())

    def test_old_expat_is_reported_as_environment_failure(self):
        output = io.StringIO()
        errors = io.StringIO()
        with mock.patch.object(create_docx.expat, "EXPAT_VERSION", "2.5.0"):
            with contextlib.redirect_stdout(output), contextlib.redirect_stderr(errors):
                result = create_docx.main(["--check"])
        self.assertEqual(result, 1)
        self.assertNotIn("check-ok", output.getvalue())
        self.assertIn("runtime environment check failed", errors.getvalue())

    def test_relative_input_and_output_still_generate_valid_zip(self):
        (self.work / "spec.json").write_text(json.dumps({"blocks": []}), encoding="utf-8")
        result = self.run_cli("spec.json", "result.docx")
        self.assertEqual(result.returncode, 0, result.stderr)
        with zipfile.ZipFile(self.work / "result.docx") as package:
            self.assertIn("word/document.xml", package.namelist())
            self.assertTrue(package.testzip() is None)


if __name__ == "__main__":
    unittest.main()
