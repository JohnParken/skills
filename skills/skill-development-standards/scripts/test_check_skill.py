"""Focused tests for the read-only skill checker."""
import contextlib
import io
import json
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
import check_skill


DESCRIPTION = "A useful description that explains when and how this skill should be used, with enough detail for discovery and selection by the assistant."


def make_skill(root: Path, body: str = "Instructions.\n", front: str | None = None) -> Path:
    path = root / "sample-skill"
    path.mkdir()
    if front is None:
        front = f"---\nname: sample-skill\nversion: '1.0.0'\ndescription: {DESCRIPTION}\n---\n"
    (path / "SKILL.md").write_text(front + body, encoding="utf-8")
    return path


class CheckSkillTests(unittest.TestCase):
    def test_version_required_type_length_and_characters(self):
        cases = [
            (None, "required-frontmatter"),
            ("null", "required-frontmatter"),
            ("42", "required-frontmatter"),
            ("true", "required-frontmatter"),
            ('""', "required-frontmatter"),
            ('"   "', "required-frontmatter"),
            ('[one]', "required-frontmatter"),
            (json.dumps("a" * 101), "version-length"),
            (json.dumps("版本1"), "version-characters"),
            (json.dumps("v1\nnext"), "version-characters"),
            (json.dumps("v1\x7f"), "version-characters"),
        ]
        for scalar, rule in cases:
            with self.subTest(scalar=scalar), tempfile.TemporaryDirectory() as tmp:
                version_line = "" if scalar is None else f"version: {scalar}\n"
                front = f"---\nname: sample-skill\ndescription: {DESCRIPTION}\n{version_line}---\n"
                skill = make_skill(Path(tmp), front=front)
                issues = check_skill.check_skill(skill)
                self.assertTrue(any(issue.rule == rule and issue.level == "error" for issue in issues))
                for options in ([], ["--strict"]):
                    with contextlib.redirect_stdout(io.StringIO()):
                        self.assertEqual(1, check_skill.main([str(skill), *options]))

    def test_valid_version_boundaries_and_decoded_length(self):
        for value in ("v", "1.0.0-rc.1+build", "release 1", "a" * 100):
            with self.subTest(value=value), tempfile.TemporaryDirectory() as tmp:
                # Escapes count as decoded characters, not source bytes.
                scalar = '"' + "".join(f"\\u{ord(char):04x}" for char in value) + '"'
                front = f"---\nname: sample-skill\ndescription: {DESCRIPTION}\nversion: {scalar}\n---\n"
                skill = make_skill(Path(tmp), front=front)
                self.assertEqual([], check_skill.check_skill(skill))

    def test_clean_skill_and_python_dependency_metadata(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_skill(Path(tmp), "See [guide](references/guide.md).\n", "---\nname: 'sample-skill'\nversion: '1.0.0'\ndescription: \"" + DESCRIPTION + "\"\ndependency:\n  python:\n    - requests\n---\n")
            (skill / "requirements.txt").write_text("requests\n", encoding="utf-8")
            (skill / "scripts").mkdir()
            (skill / "scripts" / "ok.py").write_text("x = 1\n", encoding="utf-8")
            (skill / "references").mkdir()
            (skill / "references" / "guide.md").write_text("Guide\n", encoding="utf-8")
            self.assertEqual([], check_skill.check_skill(skill))

    def test_detects_invalid_layout_links_and_python_without_executing(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_skill(Path(tmp), "[missing](absent.md)\n[escape](../outside.md)\n")
            (skill / "requirements.txt").touch()
            (skill / "scripts").mkdir()
            marker = Path(tmp) / "executed"
            (skill / "scripts" / "bad.py").write_text(f"open({str(marker)!r}, 'w').write('yes')\nif\n", encoding="utf-8")
            (skill / "extra.txt").touch()
            rules = {issue.rule for issue in check_skill.check_skill(skill)}
            self.assertTrue({"python-syntax", "broken-link", "link-escape", "root-entry"} <= rules)
            self.assertFalse(marker.exists())

    def test_warnings_and_strict_cli_exit_codes_and_readonly(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_skill(Path(tmp), "Short body.\n", f"---\nname: sample-skill\nversion: '1.0.0'\ndescription: short\n---\n")
            before = {p.relative_to(skill): (p.stat().st_mtime_ns, p.stat().st_size) for p in skill.rglob("*")}
            out = io.StringIO()
            with contextlib.redirect_stdout(out):
                normal = check_skill.main([str(skill)])
            self.assertEqual(0, normal)
            self.assertIn("WARNING", out.getvalue())
            out = io.StringIO()
            with contextlib.redirect_stdout(out):
                strict = check_skill.main([str(skill), "--strict"])
            self.assertEqual(1, strict)
            after = {p.relative_to(skill): (p.stat().st_mtime_ns, p.stat().st_size) for p in skill.rglob("*")}
            self.assertEqual(before, after)

    def test_manual_review_warning_for_complex_yaml(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_skill(Path(tmp), front=f"---\nname: sample-skill\nversion: '1.0.0'\ndescription: {DESCRIPTION}\ntags: [one, two]\n---\n")
            self.assertIn("yaml-manual-review", {i.rule for i in check_skill.check_skill(skill)})

    def test_required_scalar_types_duplicates_and_bad_quotes_fail(self):
        invalid_fronts = [
            f"---\nname: 42\nversion: '1.0.0'\ndescription: {DESCRIPTION}\n---\n",
            f"---\nname: sample-skill\nversion: '1.0.0'\nname: other\ndescription: {DESCRIPTION}\n---\n",
            f"---\nname: 'sample-skill\nversion: '1.0.0'\ndescription: {DESCRIPTION}\n---\n",
            f"---\nname: sample-skill\nversion: '1.0.0'\ndescription: {DESCRIPTION}\ndependency: requests\n---\n",
            f"---\nname: sample-skill\nversion: '1.0.0'\ndescription: {DESCRIPTION}\ndependency:\n  python:\n    - ''\n---\n",
        ]
        for front in invalid_fronts:
            with self.subTest(front=front), tempfile.TemporaryDirectory() as tmp:
                skill = make_skill(Path(tmp), front=front)
                self.assertTrue(any(i.level == "error" for i in check_skill.check_skill(skill)))

    def test_empty_body_files_nested_names_and_reference_links_and_anchors(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_skill(Path(tmp), body="")
            refs = skill / "references"
            refs.mkdir()
            guide = refs / "guide with spaces.md"
            guide.write_text("# Valid Anchor\n\n[bad](missing.md)\n[bad anchor](#nope)\n[]()\n", encoding="utf-8")
            (skill / "SKILL.md").write_text(f"---\nname: sample-skill\nversion: '1.0.0'\ndescription: {DESCRIPTION}\n---\n", encoding="utf-8")
            (refs / "empty.md").touch()
            (refs / "dir1").mkdir()
            (skill / "scripts2").mkdir()
            rules = {i.rule for i in check_skill.check_skill(skill)}
            self.assertTrue({"empty-body", "empty-file", "directory-name", "broken-link", "broken-anchor"} <= rules)

    def test_links_with_spaces_angles_valid_anchors_and_fences(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_skill(Path(tmp), "[guide](<references/guide with spaces.md#valid-anchor>)\n\n```\n[x](missing.md)\n```\n")
            refs = skill / "references"
            refs.mkdir()
            (refs / "guide with spaces.md").write_text("# Valid Anchor\n", encoding="utf-8")
            rules = {i.rule for i in check_skill.check_skill(skill)}
            self.assertNotIn("broken-link", rules)
            self.assertNotIn("broken-anchor", rules)

    def test_valid_python_is_parsed_but_never_executed_and_requires_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_skill(Path(tmp))
            marker = Path(tmp) / "executed"
            (skill / "scripts").mkdir()
            (skill / "scripts" / "valid.py").write_text(f"open({str(marker)!r}, 'w').write('ran')\n", encoding="utf-8")
            rules = {i.rule for i in check_skill.check_skill(skill)}
            self.assertIn("requirements", rules)
            self.assertFalse(marker.exists())

    def test_read_errors_and_symlinks_are_reported_without_following(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_skill(Path(tmp))
            refs = skill / "references"
            refs.mkdir()
            (refs / "bad.md").write_bytes(b"\xff")
            outside = Path(tmp) / "outside"
            outside.mkdir()
            try:
                (skill / "escape").symlink_to(outside, target_is_directory=True)
            except OSError:
                self.skipTest("symlink creation unavailable")
            rules = {i.rule for i in check_skill.check_skill(skill)}
            self.assertIn("unreadable", rules)
            self.assertIn("symlink", rules)

    def test_current_directory_and_discovery_symlink_resolve_to_same_skill(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_skill(Path(tmp))
            before = Path.cwd()
            try:
                os.chdir(skill)
                self.assertEqual([], check_skill.check_skill(Path(".")))
            finally:
                os.chdir(before)
            alias = Path(tmp) / "discovery-link"
            try:
                alias.symlink_to(skill, target_is_directory=True)
            except OSError:
                self.skipTest("symlink creation unavailable")
            self.assertEqual([], check_skill.check_skill(alias))

    def test_hyphen_boundaries_empty_nested_folders_and_temporary_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_skill(Path(tmp))
            scripts = skill / "scripts"
            scripts.mkdir()
            (scripts / "-invalid").mkdir()
            (scripts / "also-invalid-").mkdir()
            (scripts / "valid--name").mkdir()
            (scripts / "valid--name" / "data.txt").write_text("data")
            (scripts / "run.log").write_text("log")
            (scripts / "run.pyc").write_bytes(b"cache")
            issues = check_skill.check_skill(skill)
            invalid = {i.path for i in issues if i.rule == "directory-name"}
            self.assertEqual({"scripts/-invalid", "scripts/also-invalid-"}, invalid)
            self.assertEqual(2, sum(i.rule == "empty-directory" for i in issues))
            self.assertEqual(2, sum(i.rule == "suspicious-entry" for i in issues))
            renamed = skill.with_name("-sample-skill")
            skill.rename(renamed)
            self.assertIn(".", {i.path for i in check_skill.check_skill(renamed) if i.rule == "directory-name"})

    def test_scalar_comments_escaping_and_dependency_duplicates(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_skill(Path(tmp), front=(
                "---\nname: 'sample-skill' # comment\nversion: '1.0.0'\n"
                f"description: \"{DESCRIPTION}\" # comment\n"
                "label: 'a ''quoted'' label'\n"
                'note: "a: b # literal"\n---\n'
            ))
            self.assertEqual([], check_skill.check_skill(skill))
            base = f"---\nname: sample-skill\nversion: '1.0.0'\ndescription: {DESCRIPTION}\n"
            invalid = [
                'note: "unescaped "quote""\n',
                "dependency:\n  python:\n    - [nested]\n",
                "dependency:\n  python:\n    - pkg\n"
                "dependency:\n  python:\n    - other\n",
            ]
            for suffix in invalid:
                with self.subTest(suffix=suffix):
                    (skill / "SKILL.md").write_text(base + suffix + "---\nInstructions.\n")
                    self.assertTrue(check_skill.check_skill(skill))

    def test_empty_manifest_wrong_resource_type_and_exact_delimiters(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_skill(Path(tmp), front=(
                f"---\nname: sample-skill\nversion: '1.0.0'\ndescription: {DESCRIPTION}\n"
                'note: "--- is part of a scalar"\n---\n'
            ))
            self.assertEqual([], check_skill.check_skill(skill))
            (skill / "scripts").write_text("not a directory")
            self.assertIn("root-entry", {i.rule for i in check_skill.check_skill(skill)})
            (skill / "SKILL.md").write_text("")
            issues = check_skill.check_skill(skill)
            self.assertTrue(any(i.level == "error" for i in issues))

    def test_chinese_reference_toc_and_nested_fence_examples(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_skill(Path(tmp), body="Instructions.\n\n````markdown\n```\n[x](missing.md)\n```\n````\n")
            refs = skill / "references"
            refs.mkdir()
            guide = refs / "guide.md"
            guide.write_text("# Guide\n\n## 目录\n\n- 内容\n" + "内容\n" * 101)
            self.assertEqual([], check_skill.check_skill(skill))
            guide.write_text("# Guide\n" + "内容\n" * 101)
            self.assertIn("reference-toc", {i.rule for i in check_skill.check_skill(skill)})

    def test_invalid_and_encoded_link_paths(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_skill(Path(tmp), body='[valid](<references/a%23b.md> "title")\n')
            refs = skill / "references"
            refs.mkdir()
            (refs / "a#b.md").write_text("# Title\n")
            self.assertEqual([], check_skill.check_skill(skill))
            (skill / "SKILL.md").write_text(
                f"---\nname: sample-skill\nversion: '1.0.0'\ndescription: {DESCRIPTION}\n---\n"
                "[bad](http://[broken)\n"
            )
            self.assertIn("broken-link", {i.rule for i in check_skill.check_skill(skill)})

    def test_unreadable_skill_file_is_reported(self):
        with tempfile.TemporaryDirectory() as tmp:
            skill = make_skill(Path(tmp))
            with mock.patch.object(Path, "read_text", side_effect=PermissionError("denied")):
                self.assertIn("unreadable", {i.rule for i in check_skill.check_skill(skill)})


if __name__ == "__main__":
    unittest.main()
