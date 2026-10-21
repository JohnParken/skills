"""Check DOCX/DOTX packages with explicit structural rules, without XSD or rendering."""

import argparse
import sys
import tempfile
import zipfile
from pathlib import Path

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent))

try:
    from .helpers import OOXML_FAMILY, safe_extract
    from .validators import StructureValidator, RedliningValidator
except ImportError:
    from helpers import OOXML_FAMILY, safe_extract
    from validators import StructureValidator, RedliningValidator


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", type=Path, help="DOCX/DOTX file or unpacked directory")
    parser.add_argument("--original", type=Path, help="Original DOCX/DOTX for optional text comparison")
    parser.add_argument("--author", help="Enable body text revision comparison; requires --original")
    parser.add_argument("--require-no-placeholders", action="store_true",
                        help="Reject remaining {{...}} placeholders")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args(argv)
    if args.author is not None and not args.original:
        parser.error("--author requires --original")
    if not args.path.exists():
        parser.error(f"{args.path} does not exist")
    if args.path.is_file() and args.path.suffix.lower() not in OOXML_FAMILY:
        parser.error("Packed input must be a .docx or .dotx file")
    if args.original and (
        not args.original.is_file() or args.original.suffix.lower() not in OOXML_FAMILY
    ):
        parser.error("--original must be an existing .docx or .dotx file")
    try:
        with tempfile.TemporaryDirectory(prefix="docx-check-") as tmp:
            unpacked = args.path
            extension = None
            if args.path.is_file():
                extension = args.path.suffix.lower()
                unpacked = Path(tmp)
                with zipfile.ZipFile(args.path) as archive:
                    safe_extract(archive, unpacked)
            validator = StructureValidator(
                unpacked, extension=extension,
                require_no_placeholders=args.require_no_placeholders,
            )
            success = validator.validate()
            if success and args.author is not None:
                success = RedliningValidator(
                    unpacked, args.original, verbose=args.verbose
                ).validate()
            if not success:
                return 1
    except (OSError, ValueError, RuntimeError, zipfile.BadZipFile) as exc:
        print(f"FAILED - {exc}", file=sys.stderr)
        return 1
    print("Basic structure checks PASSED. XSD and visual validation were not performed.")
    if args.original and args.author is None:
        print("Original supplied; revision comparison requires --author.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
