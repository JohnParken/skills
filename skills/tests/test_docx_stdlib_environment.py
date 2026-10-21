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
                 "edit_structure"}
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
