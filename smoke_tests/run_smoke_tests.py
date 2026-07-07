from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path


BUNDLE_ROOT = Path(__file__).resolve().parents[1]
PACKAGE_PARENT = BUNDLE_ROOT.parent
if str(PACKAGE_PARENT) not in sys.path:
    sys.path.insert(0, str(PACKAGE_PARENT))


def _assert(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def _sample_content() -> dict:
    return {
        "slides": [
            {"layout": "cover", "title": "Growth Plan", "subtitle": "Revenue improves through focused execution"},
            {
                "layout": "timeline",
                "title": "Three-stage execution path",
                "events": [
                    {"date": "Q1", "title": "Diagnose", "desc": "Map funnel bottlenecks"},
                    {"date": "Q2", "title": "Launch", "desc": "Ship priority campaigns"},
                    {"date": "Q3", "title": "Scale", "desc": "Expand proven motions"},
                ],
            },
            {
                "layout": "kpi",
                "title": "Performance indicators show acceleration",
                "items": [
                    {"value": "32%", "label": "Pipeline growth"},
                    {"value": "18%", "label": "Win-rate lift"},
                ],
            },
            {"layout": "thanks", "title": "Q&A"},
        ]
    }


def test_imports() -> None:
    import pro_ppt_gen

    _assert(pro_ppt_gen.__version__ == "1.6.4", "version mismatch")
    for name in ("generate", "taste_check", "visual_redesign_check", "build_visual_redesign_prompt", "apply_visual_redesign_guidance"):
        _assert(hasattr(pro_ppt_gen, name), f"missing API {name}")


def test_visual_redesign() -> None:
    from pro_ppt_gen import visual_redesign_check, build_visual_redesign_prompt, apply_visual_redesign_guidance

    content = _sample_content()
    report = visual_redesign_check(content, strict=True)
    _assert(report["passed"], f"visual redesign should pass: {report['issues']}")
    prompt = build_visual_redesign_prompt(content)
    _assert("structure elements" in prompt, "prompt missing structure guidance")
    guided = apply_visual_redesign_guidance(content)
    _assert("visual_plan" in guided["slides"][1], "visual_plan not applied")


def test_flat_detection() -> None:
    from pro_ppt_gen import visual_redesign_check

    flat = {"slides": [{"layout": "content", "title": "Plan", "bullets": ["One", "Two", "Three"]}]}
    report = visual_redesign_check(flat, strict=True)
    codes = {issue["code"] for issue in report["issues"]}
    _assert("word_doc_on_slide" in codes, "flat text slide not detected")


def test_taste_check() -> None:
    from pro_ppt_gen import taste_check

    report = taste_check(_sample_content(), theme="business", lang="cn", strict=False)
    _assert("visual_redesign" in report, "taste report missing visual redesign section")
    _assert(isinstance(report["issues"], list), "taste issues missing")


def test_generate(out_dir: Path) -> None:
    from pro_ppt_gen import generate

    out = out_dir / "smoke.pptx"
    generate(_sample_content(), str(out), theme="business", lang="en")
    _assert(out.exists() and out.stat().st_size > 5000, "pptx not generated")
    with zipfile.ZipFile(out) as zf:
        names = set(zf.namelist())
    _assert("ppt/presentation.xml" in names, "invalid pptx package")


def run(out_dir: Path) -> int:
    tests = [
        ("imports", test_imports),
        ("visual redesign", test_visual_redesign),
        ("flat detection", test_flat_detection),
        ("taste check", test_taste_check),
        ("generate", lambda: test_generate(out_dir)),
    ]
    failures = 0
    for name, test in tests:
        try:
            test()
            print(f"PASS {name}")
        except Exception as exc:
            failures += 1
            print(f"FAIL {name}: {exc}")
    return failures


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path)
    args = parser.parse_args()
    if args.output_dir:
        out_dir = args.output_dir.resolve()
        if out_dir.exists():
            shutil.rmtree(out_dir)
        out_dir.mkdir(parents=True)
        return run(out_dir)
    with tempfile.TemporaryDirectory(prefix="pro_ppt_smoke_") as tmp:
        return run(Path(tmp))


if __name__ == "__main__":
    raise SystemExit(main())
