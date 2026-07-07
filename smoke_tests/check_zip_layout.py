from __future__ import annotations

import argparse
import zipfile
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("zip_path", type=Path)
    args = parser.parse_args()

    with zipfile.ZipFile(args.zip_path) as zf:
        names = [name.replace("\\", "/") for name in zf.namelist() if not name.endswith("/")]

    errors = []
    if {name.split("/", 1)[0] for name in names} != {"pro_ppt_gen"}:
        errors.append("zip must have single root pro_ppt_gen")
    if any("/.git/" in name or name.startswith("pro_ppt_gen/.git/") for name in names):
        errors.append("zip must not include .git")
    for required in (
        "pro_ppt_gen/SKILL.md",
        "pro_ppt_gen/__init__.py",
        "pro_ppt_gen/ppt_jsx.py",
        "pro_ppt_gen/visual_redesign.py",
        "pro_ppt_gen/references/visual-redesign.md",
        "pro_ppt_gen/shared/import_helper.py",
        "pro_ppt_gen/shared/remote_assets.py",
        "pro_ppt_gen/shared/taste/core.py",
        "pro_ppt_gen/smoke_tests/run_smoke_tests.py",
    ):
        if required not in names:
            errors.append(f"missing {required}")

    if errors:
        for error in errors:
            print(f"FAIL {error}")
        return 1
    print(f"PASS zip layout {args.zip_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
