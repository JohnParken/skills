"""Check this skill's deliberately simple metadata and packaging using stdlib."""

import ast
import re
import unittest
from pathlib import Path

SKILL = Path(__file__).resolve().parents[1] / "docx"


class SkillContractTests(unittest.TestCase):
    def test_metadata_instructions_and_dependency_manifest(self):
        text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
        self.assertTrue(text.startswith("---\n"))
        _, metadata, body = text.split("---", 2)
        fields = dict(line.split(": ", 1) for line in metadata.strip().splitlines())
        self.assertEqual(fields["name"], "docx")
        self.assertTrue(100 <= len(fields["description"]) <= 150)
        self.assertLessEqual(len(body.splitlines()), 500)
        self.assertTrue(body.strip())
        requirements = (SKILL / "requirements.txt").read_text()
        self.assertTrue(requirements.strip())
        self.assertFalse([line for line in requirements.splitlines()
                          if line.strip() and not line.lstrip().startswith("#")])

    def test_resource_layout_links_and_script_syntax(self):
        self.assertLessEqual(
            {p.name for p in SKILL.iterdir()},
            {"SKILL.md", "scripts", "assets", "references", "requirements.txt", "LICENSE.txt"},
        )
        for path in SKILL.rglob("*"):
            with self.subTest(path=path.relative_to(SKILL)):
                if path.is_dir():
                    self.assertRegex(path.name, r"^[A-Za-z_][A-Za-z_-]*[A-Za-z_]$|^[A-Za-z_]$")
                    self.assertTrue(any(path.iterdir()))
                    continue
                self.assertTrue(path.read_bytes())
                self.assertNotEqual(path.name.lower(), "readme.md")
                self.assertNotIn(path.suffix, (".pyc", ".log", ".tmp", ".xsd"))
                if path.suffix == ".py":
                    ast.parse(path.read_text(), filename=str(path), feature_version=(3, 10))
                if path.suffix == ".md":
                    text = path.read_text()
                    self.assertNotRegex(text, r"/Users/|C:\\Users\\|https?://")
                    for target in re.findall(r"\[[^\]]+\]\(([^)]+)\)", text):
                        resource = target.partition("#")[0]
                        self.assertFalse(Path(resource).is_absolute())
                        linked_file = path.parent / resource if resource else path
                        self.assertTrue(linked_file.is_file(), target)
                        self.assertTrue(linked_file.resolve().is_relative_to(SKILL.resolve()))
                    if path.parent.name == "references" and len(text.splitlines()) > 100:
                        self.assertIn("目录", text)


if __name__ == "__main__":
    unittest.main()
