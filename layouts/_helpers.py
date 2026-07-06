"""Internal helpers shared across layout handlers."""
from __future__ import annotations


def pt(v) -> float:
    """Convert a python-pptx Pt value (or raw EMU int) to a float point size."""
    if hasattr(v, "pt"):
        return float(v.pt)
    # Assume EMU
    return float(v) / 12700.0


def emu(v) -> int:
    """Normalize a value to int EMU (handles Emu / Inches / Pt objects)."""
    return int(v)
