"""PPT element animation engine for PRO-PPTX v1.6.3.

Builds the <p:timing> XML subtree that OOXML uses to describe slide-level
element animations (entrance / emphasis / exit).  All timings follow the
PowerPoint 2016+ preset definitions; we expose a small semantic API so
LLMs never specify per-element durations or delays.

Key public API:
    ANIMATION_PRESETS     : dict mapping semantic preset name -> PresetInfo
    TRANSITION_EXT_MAP    : dict for 6 new slide transitions (cube/reveal/zoom/ferris/gallery/conveyor)
    NodeAnim              : dataclass for per-shape animation spec
    build_slide_timing(slide, animations) : inject <p:timing> into slide XML
    apply_ext_transition(slide, kind)     : apply an extended transition
    register_ext_transitions(renderer_map): register ext transitions with Renderer._TRANSITION_MAP

@since v1.6.3
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any, Optional

from lxml import etree

# ----------------------------------------------------------------------
# Namespaces
# ----------------------------------------------------------------------
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"


def _p(tag: str) -> str:
    """Clark-notation helper for p: namespace."""
    return f"{{{P_NS}}}{tag}"


def _a(tag: str) -> str:
    return f"{{{A_NS}}}{tag}"


def _p14(tag: str) -> str:
    return f"{{{P14_NS}}}{tag}"


# Default OOXML durations (ms) — internal only
_ENTRANCE_DUR = 500
_EMPHASIS_DUR = 300
_EXIT_DUR = 400


# ----------------------------------------------------------------------
# Preset metadata
# ----------------------------------------------------------------------

@dataclass
class PresetInfo:
    preset_class: str        # 'entr' | 'emph' | 'exit'
    preset_id: int           # OOXML preset ID (ECMA-376 ST_TLAnimateBehaviorPresetID)
    preset_subtype: int = 0
    duration: int = _ENTRANCE_DUR


ANIMATION_PRESETS: dict[str, PresetInfo] = {
    # ---- Entrance (10) ----
    "fade_in":           PresetInfo("entr", 10, 0,  _ENTRANCE_DUR),
    "appear":            PresetInfo("entr", 1,  0,  1),
    "fly_in_left":       PresetInfo("entr", 2,  4,  _ENTRANCE_DUR),
    "fly_in_right":      PresetInfo("entr", 2,  8,  _ENTRANCE_DUR),
    "fly_in_bottom":     PresetInfo("entr", 2,  2,  _ENTRANCE_DUR),
    "wipe_right":        PresetInfo("entr", 22, 8,  _ENTRANCE_DUR),
    "wipe_up":           PresetInfo("entr", 22, 1,  _ENTRANCE_DUR),
    "zoom_in":           PresetInfo("entr", 23, 0,  _ENTRANCE_DUR),
    "float_in":          PresetInfo("entr", 42, 0,  _ENTRANCE_DUR),
    "split_horizontal":  PresetInfo("entr", 35, 4,  _ENTRANCE_DUR),
    # ---- Emphasis (4) ----
    "pulse":             PresetInfo("emph", 12, 0,  _EMPHASIS_DUR),
    "grow_shrink":       PresetInfo("emph", 9,  0,  _EMPHASIS_DUR),
    "color_shift":       PresetInfo("emph", 6,  0,  _EMPHASIS_DUR),
    "underline_reveal":  PresetInfo("emph", 28, 0,  _EMPHASIS_DUR),
    # ---- Exit (4) ----
    "fade_out":          PresetInfo("exit", 10, 0,  _EXIT_DUR),
    "fly_out_right":     PresetInfo("exit", 2,  8,  _EXIT_DUR),
    "zoom_out":          PresetInfo("exit", 23, 0,  _EXIT_DUR),
    "dissolve_out":      PresetInfo("exit", 26, 0,  _EXIT_DUR),
}


TRANSITION_EXT_MAP: dict[str, dict] = {
    "cube":     {"tag": "cube",     "dir": "l",  "spd": "med"},
    "reveal":   {"tag": "reveal",   "dir": "l",  "spd": "med", "throughBlk": "1"},
    "zoom":     {"tag": "zoom",     "dir": "in", "spd": "med"},
    "ferris":   {"tag": "ferris",   "dir": "l",  "spd": "med"},
    "gallery":  {"tag": "gallery",  "dir": "l",  "spd": "med"},
    "conveyor": {"tag": "conveyor", "dir": "l",  "spd": "med"},
}


def register_ext_transitions(renderer_transition_map: dict) -> None:
    """Register the 6 new extended transitions into a Renderer._TRANSITION_MAP dict."""
    for name in TRANSITION_EXT_MAP:
        if name not in renderer_transition_map:
            renderer_transition_map[name] = ("__ext__", {"ext_tag": name})


# ----------------------------------------------------------------------
# NodeAnim
# ----------------------------------------------------------------------

@dataclass
class NodeAnim:
    shape_id: int
    preset_name: str = "fade_in"
    sequence: str = "one_by_one"
    trigger: str = "on_click"
    node_level: Optional[int] = None


# ----------------------------------------------------------------------
# build_slide_timing
# ----------------------------------------------------------------------

def build_slide_timing(slide, animations: list) -> None:
    """Build and inject <p:timing> subtree for a slide."""
    sld = slide._element

    for old in sld.findall(_p("timing")):
        sld.remove(old)
    if not animations:
        return

    timing = etree.SubElement(sld, _p("timing"))
    tnLst = etree.SubElement(timing, _p("tnLst"))

    par_root = etree.SubElement(tnLst, _p("par"))
    cTn_root = etree.SubElement(par_root, _p("cTn"),
                                 id="1", dur="indefinite",
                                 restart="never", nodeType="tmRoot")
    childTnLst_root = etree.SubElement(cTn_root, _p("childTnLst"))

    seq = etree.SubElement(childTnLst_root, _p("seq"),
                           concurrent="1", nextAc="seek")
    cTn_seq = etree.SubElement(seq, _p("cTn"), id="2",
                                dur="indefinite", nodeType="mainSeq")
    childTnLst_seq = etree.SubElement(cTn_seq, _p("childTnLst"))

    # prev/next condLst
    prev = etree.SubElement(seq, _p("prevCondLst"))
    nxt = etree.SubElement(seq, _p("nextCondLst"))
    nxt_cond = etree.SubElement(nxt, _p("cond"), evt="onClick", delay="0")
    tgtEl = etree.SubElement(nxt_cond, _p("tgtEl"))
    etree.SubElement(tgtEl, _p("sldTgt"))

    next_id = 3
    click_groups = _group_by_click(animations)

    for ci, group in enumerate(click_groups):
        click_par = etree.SubElement(childTnLst_seq, _p("par"))
        cTn_click = etree.SubElement(click_par, _p("cTn"),
                                     id=str(next_id), fill="hold")
        next_id += 1
        stCondLst = etree.SubElement(cTn_click, _p("stCondLst"))
        if ci == 0:
            etree.SubElement(stCondLst, _p("cond"), delay="0")
        else:
            cond = etree.SubElement(stCondLst, _p("cond"), evt="onClick", delay="0")
            t = etree.SubElement(cond, _p("tgtEl"))
            etree.SubElement(t, _p("sldTgt"))
        childTnLst_click = etree.SubElement(cTn_click, _p("childTnLst"))

        for anim in group:
            next_id = _add_shape_anim(childTnLst_click, anim, next_id)

    # bldLst
    bldLst = etree.SubElement(timing, _p("bldLst"))
    for anim in animations:
        spid = str(anim.shape_id)
        if anim.sequence == "by_paragraph":
            etree.SubElement(bldLst, _p("bldP"), spid=spid, grpId="0", build="p")
        elif anim.sequence == "by_series":
            etree.SubElement(bldLst, _p("bldChart"), spid=spid, grpId="0", build="series")
        else:
            etree.SubElement(bldLst, _p("bldP"), spid=spid, grpId="0")


def _group_by_click(animations: list) -> list:
    groups: list = []
    current: list = []
    for a in animations:
        if a.trigger == "on_click" and current:
            groups.append(current)
            current = [a]
        else:
            current.append(a)
    if current:
        groups.append(current)
    if not groups and animations:
        groups = [list(animations)]
    return groups


def _add_shape_anim(parent, anim, next_id: int) -> int:
    preset = ANIMATION_PRESETS.get(anim.preset_name)
    if preset is None:
        return next_id

    spid = str(anim.shape_id)
    pid = str(preset.preset_id)
    psub = str(preset.preset_subtype)
    dur = str(preset.duration)

    par = etree.SubElement(parent, _p("par"))
    cTn = etree.SubElement(par, _p("cTn"), id=str(next_id), fill="hold")
    next_id += 1
    stCondLst = etree.SubElement(cTn, _p("stCondLst"))
    etree.SubElement(stCondLst, _p("cond"), delay="0")
    childTnLst = etree.SubElement(cTn, _p("childTnLst"))

    inner_par = etree.SubElement(childTnLst, _p("par"))
    cTn_inner_attrs = {
        "id": str(next_id),
        "presetID": pid,
        "presetClass": preset.preset_class,
        "presetSubtype": psub,
        "fill": "hold",
        "dur": dur,
    }
    cTn_inner = etree.SubElement(inner_par, _p("cTn"), **cTn_inner_attrs)
    next_id += 1
    stCondLst2 = etree.SubElement(cTn_inner, _p("stCondLst"))
    etree.SubElement(stCondLst2, _p("cond"), delay="0")
    childTnLst2 = etree.SubElement(cTn_inner, _p("childTnLst"))

    if preset.preset_class == "entr":
        set_el = etree.SubElement(childTnLst2, _p("set"))
        cBhvr = etree.SubElement(set_el, _p("cBhvr"))
        cTn_set = etree.SubElement(cBhvr, _p("cTn"),
                                    id=str(next_id), dur="1", fill="hold")
        next_id += 1
        stCL = etree.SubElement(cTn_set, _p("stCondLst"))
        etree.SubElement(stCL, _p("cond"), delay="0")
        tgtEl = etree.SubElement(cBhvr, _p("tgtEl"))
        etree.SubElement(tgtEl, _p("spTgt"), spid=spid)
        attrNameLst = etree.SubElement(cBhvr, _p("attrNameLst"))
        an = etree.SubElement(attrNameLst, _p("attrName"))
        an.text = "style.visibility"
        to = etree.SubElement(set_el, _p("to"))
        etree.SubElement(to, _p("strVal"), val="visible")

    filter_str = _ooxml_filter(anim.preset_name)
    animEffect_attrs = {
        "transition": "in" if preset.preset_class in ("entr", "emph") else "out",
        "filter": filter_str,
    }
    animEffect = etree.SubElement(childTnLst2, _p("animEffect"), **animEffect_attrs)
    cBhvr2 = etree.SubElement(animEffect, _p("cBhvr"))
    etree.SubElement(cBhvr2, _p("cTn"), id=str(next_id), dur=dur)
    next_id += 1
    tgtEl2 = etree.SubElement(cBhvr2, _p("tgtEl"))
    etree.SubElement(tgtEl2, _p("spTgt"), spid=spid)

    return next_id


def _ooxml_filter(preset_name: str) -> str:
    _FILTERS = {
        "fade_in": "fade",
        "fade_out": "fade",
        "appear": "fade",
        "fly_in_left": "fly(fromLeft)",
        "fly_in_right": "fly(fromRight)",
        "fly_in_bottom": "fly(fromBottom)",
        "fly_out_right": "fly(toRight)",
        "wipe_right": "wipe(right)",
        "wipe_up": "wipe(up)",
        "zoom_in": "zoom(in)",
        "zoom_out": "zoom(out)",
        "float_in": "float(bottom)",
        "split_horizontal": "split(horizontal,in)",
        "pulse": "pulse",
        "grow_shrink": "growShrink",
        "color_shift": "colorBlend",
        "underline_reveal": "underline",
        "dissolve_out": "dissolve",
    }
    return _FILTERS.get(preset_name, "fade")


# ----------------------------------------------------------------------
# apply_ext_transition
# ----------------------------------------------------------------------

def apply_ext_transition(slide, transition_type: str) -> bool:
    kind = (transition_type or "").strip().lower()
    if kind not in TRANSITION_EXT_MAP:
        return False
    spec = TRANSITION_EXT_MAP[kind]
    sld = slide._element
    for old in sld.findall(_p("transition")):
        sld.remove(old)
    nsmap = {"p": P_NS, "p14": P14_NS}
    trans = etree.SubElement(sld, _p("transition"), nsmap=nsmap)
    trans.set("spd", spec.get("spd", "med"))
    extLst = etree.SubElement(trans, _p("extLst"))
    ext = etree.SubElement(extLst, _p("ext"), uri=f"{{{P14_NS}}}")
    ext_el = etree.SubElement(ext, _p14(spec["tag"]))
    if "dir" in spec:
        ext_el.set("dir", spec["dir"])
    if "throughBlk" in spec:
        ext_el.set("throughBlk", spec["throughBlk"])
    return True
