"""Check syntax and local Python identifier naming for the DOCX UX change.

This standard-library check does not replace a general-purpose linter. It skips
import/API attribute names and XML strings, and never writes bytecode caches.
"""

import ast
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
FILES = (
    "docx/scripts/create_docx.py",
    "docx/scripts/edit_docx.py",
    "docx/scripts/office/validate.py",
    "docx/scripts/office/validators/structure.py",
    "docx/scripts/office/validators/generation.py",
    "tests/test_docx_optimization.py",
    "tests/test_docx_ux.py",
    "tests/test_docx_generator.py",
    "tests/test_docx_distribution.py",
    "tests/test_docx_template_diagnostics.py",
    "tests/benchmark_docx_pipeline.py",
    "tests/check_docx_ux_static.py",
)
SNAKE = re.compile(r"_?[a-z][a-z0-9_]*\Z|__[a-z][a-z0-9_]*__\Z")
CONSTANT = re.compile(r"_?[A-Z][A-Z0-9_]*\Z")
CLASS = re.compile(r"[A-Z][A-Za-z0-9]*\Z")


class NamingCheck(ast.NodeVisitor):
    def __init__(self, path):
        self.path = path
        self.function_depth = 0
        self.errors = []

    def check(self, name, node, pattern):
        if name.startswith("visit_") and hasattr(ast, name.removeprefix("visit_")):
            return  # ast.NodeVisitor callback names are an external API.
        if name in {"setUp", "tearDown", "setUpClass", "tearDownClass"}:
            return  # unittest.TestCase lifecycle names are an external API.
        if not pattern.fullmatch(name) and name != "_":
            self.errors.append(f"{self.path}:{node.lineno}: invalid local identifier {name!r}")

    def visit_FunctionDef(self, node):
        self.check(node.name, node, SNAKE)
        self.function_depth += 1
        for argument in (*node.args.posonlyargs, *node.args.args, *node.args.kwonlyargs):
            self.check(argument.arg, argument, SNAKE)
        for argument in (node.args.vararg, node.args.kwarg):
            if argument:
                self.check(argument.arg, argument, SNAKE)
        self.generic_visit(node)
        self.function_depth -= 1

    visit_AsyncFunctionDef = visit_FunctionDef

    def visit_ClassDef(self, node):
        self.check(node.name, node, CLASS)
        self.generic_visit(node)

    def visit_Name(self, node):
        if isinstance(node.ctx, ast.Store):
            if self.function_depth:
                self.check(node.id, node, SNAKE)
            else:
                # Module-level comprehension variables are ordinary variables.
                if not CONSTANT.fullmatch(node.id):
                    self.check(node.id, node, SNAKE)


def main():
    errors = []
    for relative in FILES:
        path = ROOT / relative
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path), feature_version=(3, 10))
        check = NamingCheck(relative)
        check.visit(tree)
        errors.extend(check.errors)
    for error in errors:
        print(error)
    print(f"DOCX UX syntax and identifier checks: {len(FILES)} files, {len(errors)} errors")
    return bool(errors)


if __name__ == "__main__":
    raise SystemExit(main())
