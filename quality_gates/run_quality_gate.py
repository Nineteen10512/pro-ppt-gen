from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path


def inspect_file(path: Path) -> dict:
    errors = []
    warnings = []
    if not path.exists():
        errors.append("file not found")
    elif path.suffix.lower() != ".pptx":
        errors.append("expected .pptx")
    else:
        try:
            with zipfile.ZipFile(path) as zf:
                names = set(zf.namelist())
            if "ppt/presentation.xml" not in names:
                errors.append("missing ppt/presentation.xml")
            if not any(name.startswith("ppt/slides/slide") for name in names):
                errors.append("no slides found")
        except zipfile.BadZipFile:
            errors.append("invalid zip container")
    return {"path": str(path), "errors": errors, "warnings": warnings, "passed": not errors}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--json-report", type=Path)
    args = parser.parse_args()
    report = inspect_file(args.pptx)
    if args.json_report:
        args.json_report.parent.mkdir(parents=True, exist_ok=True)
        args.json_report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False))
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
