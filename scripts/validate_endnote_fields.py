#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path


def read_document_xml(docx_path: Path) -> str:
    if not docx_path.exists():
        raise FileNotFoundError(docx_path)
    with zipfile.ZipFile(docx_path) as zf:
        return zf.read("word/document.xml").decode("utf-8", "replace")


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate real EndNote fields inside a DOCX.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--min-cites", type=int, default=1)
    parser.add_argument("--require-reflist", action="store_true")
    args = parser.parse_args()

    xml = read_document_xml(args.docx)
    cite_count = xml.count("ADDIN EN.CITE")
    reflist_count = xml.count("ADDIN EN.REFLIST")
    record_count = xml.count("&lt;record&gt;") + xml.count("<record>")
    visible_numeric = len(re.findall(r"\[[0-9]+(?:\s*(?:,|-)\s*[0-9]+)*\]", xml))

    print(f"docx: {args.docx}")
    print(f"ADDIN_EN_CITE: {cite_count}")
    print(f"ADDIN_EN_REFLIST: {reflist_count}")
    print(f"embedded_records: {record_count}")
    print(f"visible_numeric_tokens_in_xml: {visible_numeric}")

    failures: list[str] = []
    if cite_count < args.min_cites:
        failures.append(f"expected at least {args.min_cites} ADDIN EN.CITE fields")
    if args.require_reflist and reflist_count < 1:
        failures.append("expected at least one ADDIN EN.REFLIST field")
    if cite_count and record_count < cite_count:
        failures.append("expected embedded EndNote <record> payloads for citation fields")

    if failures:
        for failure in failures:
            print(f"FAIL: {failure}", file=sys.stderr)
        return 1

    print("OK: EndNote field structure detected")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
