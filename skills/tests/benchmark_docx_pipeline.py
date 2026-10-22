"""Measure local DOCX generation and equivalent disk/memory structure checks.

Run with the supported Python and -I -S -B. No external programs or packages.
Numbers cover local processing only, not model generation or tool latency.
"""

import argparse
import json
import random
import statistics
import struct
import sys
import tempfile
import time
import tracemalloc
import zlib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "docx/scripts"))
from create_docx import create_document, _validate_spec
from office.archiveio import open_archive
from office.helpers import safe_extract
from office.validators.structure import StructureValidator


def measure(operation, repeats):
    values = []
    for _ in range(repeats):
        start = time.perf_counter()
        operation()
        values.append(time.perf_counter() - start)
    return round(statistics.median(values), 6)


def png_sample():
    def chunk(kind, payload):
        data = kind + payload
        return struct.pack(">I", len(payload)) + data + struct.pack(">I", zlib.crc32(data) & 0xFFFFFFFF)

    random_source = random.Random(21)
    rows = b"".join(b"\x00" + random_source.randbytes(512 * 3) for _ in range(512))
    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", 512, 512, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


def extended_workloads(work, repeats):
    import zipfile

    output = work / "extended.docx"
    tables = []
    for row_count in (100, 500, 1000):
        spec = {"blocks": [{"type": "table", "header": True,
                            "rows": [[f"第{row}行第{column}列" for column in range(6)]
                                     for row in range(row_count)]}]}
        elapsed = measure(lambda: create_document(spec, output, verify=True), repeats)
        with zipfile.ZipFile(output) as archive:
            xml_bytes = archive.getinfo("word/document.xml").file_size
        tables.append({"rows": row_count, "columns": 6, "verified_generation_seconds": elapsed,
                       "document_xml_bytes": xml_bytes})
    data = png_sample()
    (work / "asset.png").write_bytes(data)
    image_spec = {"deduplicate_images": True, "blocks": [
        {"type": "image", "path": "asset.png", "width_mm": 30, "alt": str(index)} for index in range(12)
    ]}
    tracemalloc.start()
    try:
        normalized = _validate_spec(image_spec, work)
        _, peak = tracemalloc.get_traced_memory()
    finally:
        tracemalloc.stop()
    payloads = {id(block["image_data"]): block["image_data"] for block in normalized["blocks"]}
    create_document(image_spec, output, work, verify=True)
    with zipfile.ZipFile(output) as archive:
        media_count = sum(name.startswith("word/media/") for name in archive.namelist())
    return {"tables": tables, "repeated_image": {
        "references": 12, "source_bytes": len(data), "retained_unique_payloads": len(payloads),
        "retained_payload_bytes": sum(map(len, payloads.values())),
        "preparation_peak_bytes": peak, "packaged_media_count": media_count,
    }}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repeats", type=int, default=7)
    args = parser.parse_args()
    if not 1 <= args.repeats <= 100:
        parser.error("--repeats must be between 1 and 100")
    spec = {"blocks": [{"type": "paragraph", "text": "测试正文 mixed text 123。" * 5}] * 100}
    with tempfile.TemporaryDirectory(prefix="docx-benchmark-") as folder:
        work = Path(folder)
        output = work / "benchmark.docx"
        create_document(spec, output)

        def memory_check():
            with open_archive(output) as archive:
                validator = StructureValidator.from_archive(archive, extension=".docx")
                if not validator.validate(quiet=True):
                    raise RuntimeError(validator.errors)

        def disk_check():
            with tempfile.TemporaryDirectory(dir=work) as temporary:
                directory = Path(temporary)
                with open_archive(output) as archive:
                    safe_extract(archive, directory)
                validator = StructureValidator(directory, extension=".docx")
                if not validator.validate(quiet=True):
                    raise RuntimeError(validator.errors)

        results = {
            "repeats": args.repeats, "blocks": len(spec["blocks"]),
            "generation_seconds": measure(lambda: create_document(spec, output), args.repeats),
            "verified_generation_seconds": measure(lambda: create_document(spec, output, verify=True), args.repeats),
            "memory_validation_seconds": measure(memory_check, args.repeats),
            "disk_validation_seconds": measure(disk_check, args.repeats),
            "extended_workloads": extended_workloads(work, args.repeats),
        }
        print(json.dumps(results, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
