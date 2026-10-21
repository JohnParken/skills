"""Enforce the skill's dependency boundary in an isolated interpreter."""

import ast
import os
import subprocess
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "docx/scripts"


class StdlibEnvironmentTests(unittest.TestCase):
    def test_compressed_input_entrypoints_use_the_bounded_archive_opener(self):
        # Direct constructors are reserved for the common bounded opener and
        # the two modules that write/check their own staged output archives.
        allowed = {"create_docx.py", "office/helpers/__init__.py", "office/archiveio.py"}
        for path in SCRIPTS.rglob("*.py"):
            relative = path.relative_to(SCRIPTS).as_posix()
            if relative in allowed:
                continue
            with self.subTest(script=relative):
                tree = ast.parse(path.read_text())
                for node in ast.walk(tree):
                    if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute):
                        self.assertNotEqual(node.func.attr, "ZipFile",
                                            "Use office.archiveio.open_archive for untrusted input")

    def test_isolated_interpreter_cannot_import_third_party_packages(self):
        result = subprocess.run(
            [sys.executable, "-I", "-S", "-B", "-c",
             "import sys,importlib.util;"
             "assert sys.flags.isolated and sys.flags.no_site;"
             "assert not any('site-packages' in p for p in sys.path);"
             "assert all(importlib.util.find_spec(n) is None for n in "
             "('docx','lxml','defusedxml','docxtpl','jinja2','yaml'))"],
            env=dict(os.environ, PATH=""), capture_output=True, text=True,
        )
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_runtime_imports_are_standard_library_or_skill_local(self):
        local = {"office", "helpers", "validators", "xmlio", "create_docx", "read_docx", "edit_docx",
                 "edit_structure", "archiveio"}
        for path in SCRIPTS.rglob("*.py"):
            with self.subTest(script=path.relative_to(SCRIPTS)):
                tree = ast.parse(path.read_text())
                for node in ast.walk(tree):
                    if isinstance(node, ast.Import):
                        modules = [alias.name.split(".", 1)[0] for alias in node.names]
                    elif isinstance(node, ast.ImportFrom) and not node.level:
                        modules = [(node.module or "").split(".", 1)[0]]
                    else:
                        continue
                    for module in modules:
                        self.assertIn(module, sys.stdlib_module_names | local)
                        self.assertNotEqual(module, "subprocess", "Skill must not launch external programs")


if __name__ == "__main__":
    unittest.main()
