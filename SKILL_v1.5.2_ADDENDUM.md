# PRO-PPTX v1.6.3 Addendum

This addendum documents behavior that is already implemented in the archive but
was not clearly described in the earlier `SKILL.md`.

## Shared taste framework

PRO-PPTX taste logic now shares a reusable framework with PRO-DOCX:

- `shared/taste/core.py`
- `shared/taste/rules.py`
- `shared/taste/adapters.py`

This shared layer provides:

- design-read inference
- copy-density checks
- layout-rhythm checks
- preflight report assembly

## Actual `taste_check()` report shape

`ppt_jsx.taste_check(content, theme="academic", lang="cn", strict=False)`
returns:

- `version`
- `score`
- `passed`
- `threshold`
- `design_read`
- `preflight`
- `issues`
- `base_quality_score`
- `story_compliant`

`design_read` includes:

- a one-line reading of the deck
- `design_variance`
- `motion_intensity`
- `visual_density`

## Hard-fail taste rules

The following issue codes are treated as hard failures in v1.6.3:

- `contrast_iron_law`
- `chart_color_variety`
- `placeholder_text`

Practical meaning:

- dark background -> light text
- light background -> dark text
- multi-series or multi-category charts cannot collapse to a single chart color
- placeholder copy such as `TODO`, `TBD`, `[image]`, `placeholder` fails the
  taste gate directly

## Exported helper

The taste module also exports:

```python
from pro_ppt_gen.taste import infer_design_read
```

Use it when you only want the audience/vibe/dials read without running the full
preflight.

## Slide deletion link safety

`ppt_jsx.delete_slide()` now re-links internal slide-jump relationships after a
slide is removed. This prevents surviving slide hyperlinks from still pointing
at stale slide-order targets.

## Local taste-skill reference

The archive contains a local reference copy of the public `taste-skill`
repository under:

- `docs/references/taste-skill-main/`

It is included for reference and maintenance only. It is not imported at
runtime by PRO-PPTX.
