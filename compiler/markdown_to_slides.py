"""markdown_to_slides — 简易 Markdown → PPT slides 语义 dict 转换器（P1-2, v1.4）。

支持的 Markdown 语法（简化版，表格/图片复杂语法留给 LLM 直接写 JSON）：

- ``# 大标题``            → cover 页（title=标题，subtitle 取后续到 ``##`` 之前的段落）
- ``## 章节标题``         → section 页（number 自动编号），后续 ``- item``/``1. item``/纯文本归入后续 content 页
- ``- item`` / ``* item`` → content 页 bullets（默认无序列表）
- ``1. item`` / ``2. item`` → content 页 bullets（有序列表，序号由 PPT 排版自动生成）
- 纯文本段落              → 作为 content 页的一段 bullet（或 speaker_notes）
- ``---``                 → 手动分页符
- ``### 标题``            → 作为 content 页的 title（content 页之间的分隔）
- 连续超过 6 个 bullet    → 自动拆页

返回的 dict 可直接传给 ``ppt_jsx.generate()``：

    {
        "meta": {"title": ..., ...},
        "slides": [ {"layout": "cover", ...}, {"layout": "section", ...}, {"layout": "content", ...} ]
    }
"""
from __future__ import annotations

import re


_BULLET_UL = re.compile(r"^\s*[-*]\s+(.+)$")
_BULLET_OL = re.compile(r"^\s*\d+\.\s+(.+)$")
_H1 = re.compile(r"^#\s+(.+)$")
_H2 = re.compile(r"^##\s+(.+)$")
_H3 = re.compile(r"^###\s+(.+)$")
_HR = re.compile(r"^\s*(---|\*\*\*|___)\s*$")
_EMPTY = re.compile(r"^\s*$")

# PPT 单页 bullets 上限（超过则拆页）
MAX_BULLETS_PER_PAGE = 6


def markdown_to_slides(md_text: str, theme: str = "academic", **meta_kwargs) -> dict:
    """将 Markdown 字符串解析为 PPT 语义 dict（可直接传给 ppt_jsx.generate）。

    Args:
        md_text: Markdown 源文本。
        theme: 主题名（仅作为元信息，实际主题由 generate(theme=...) 决定）。
        **meta_kwargs: 透传到 meta，如 author/subtitle/date 等。

    Returns:
        {"meta": {...}, "slides": [...]}。
    """
    lines = md_text.split("\n")
    slides: list[dict] = []

    # 状态
    doc_title: str | None = None
    cover_subtitle_parts: list[str] = []
    in_preamble = True  # 在第一个 # 之前 / # 后到 ## 之间，收集 subtitle

    section_counter = 0
    current_content_title: str | None = None
    current_bullets: list[str] = []

    def _flush_content():
        nonlocal current_bullets, current_content_title
        if not current_bullets:
            current_content_title = None
            return
        # 按 MAX_BULLETS_PER_PAGE 拆页
        chunks = [
            current_bullets[i:i + MAX_BULLETS_PER_PAGE]
            for i in range(0, len(current_bullets), MAX_BULLETS_PER_PAGE)
        ]
        for idx, chunk in enumerate(chunks):
            slide = {
                "layout": "content",
                "title": current_content_title or ("内容" if idx > 0 else ""),
                "bullets": chunk,
            }
            if idx > 0:
                # 续页标题加"(续)"
                if current_content_title:
                    slide["title"] = f"{current_content_title}（续）"
            slides.append(slide)
        current_bullets = []
        current_content_title = None

    def _new_section(title: str):
        nonlocal section_counter
        _flush_content()
        section_counter += 1
        slides.append({
            "layout": "section",
            "number": f"{section_counter:02d}",
            "title": title,
        })

    def _start_content(title: str | None):
        nonlocal current_content_title
        _flush_content()
        if title:
            current_content_title = title

    i = 0
    while i < len(lines):
        line = lines[i].rstrip("\n")
        stripped = line.strip()

        # 空行
        if _EMPTY.match(stripped):
            i += 1
            continue

        # 分页符
        if _HR.match(stripped):
            _flush_content()
            i += 1
            continue

        # # 大标题 → cover
        m = _H1.match(stripped)
        if m:
            _flush_content()
            doc_title = m.group(1).strip()
            in_preamble = True  # 开始收集 subtitle（直到遇到 ## 或 ###）
            cover_subtitle_parts = []
            i += 1
            continue

        # ## 章节 → section
        m = _H2.match(stripped)
        if m:
            # 如果还在收集 cover subtitle，则先收尾 cover
            if in_preamble and doc_title is not None:
                cover_slide = {
                    "layout": "cover",
                    "title": doc_title,
                }
                sub = " ".join(p for p in cover_subtitle_parts if p).strip()
                if sub:
                    cover_slide["subtitle"] = sub
                if meta_kwargs.get("author"):
                    cover_slide["author"] = meta_kwargs["author"]
                if meta_kwargs.get("date"):
                    cover_slide["date"] = meta_kwargs["date"]
                slides.append(cover_slide)
                in_preamble = False
                cover_subtitle_parts = []
            _flush_content()
            _new_section(m.group(1).strip())
            i += 1
            continue

        # ### 小标题 → content title
        m = _H3.match(stripped)
        if m:
            if in_preamble and doc_title is not None:
                # 先结束 cover
                cover_slide = {"layout": "cover", "title": doc_title}
                sub = " ".join(p for p in cover_subtitle_parts if p).strip()
                if sub:
                    cover_slide["subtitle"] = sub
                slides.append(cover_slide)
                in_preamble = False
                cover_subtitle_parts = []
            _start_content(m.group(1).strip())
            i += 1
            continue

        # bullet
        m = _BULLET_UL.match(stripped) or _BULLET_OL.match(stripped)
        if m:
            if in_preamble and doc_title is not None:
                # bullet 出现在 cover 之后正文区，先收尾 cover
                cover_slide = {"layout": "cover", "title": doc_title}
                sub = " ".join(p for p in cover_subtitle_parts if p).strip()
                if sub:
                    cover_slide["subtitle"] = sub
                slides.append(cover_slide)
                in_preamble = False
                cover_subtitle_parts = []
            current_bullets.append(m.group(1).strip())
            i += 1
            continue

        # 普通段落
        if in_preamble and doc_title is not None:
            cover_subtitle_parts.append(stripped)
        else:
            # 正文中的纯文本段落：作为一段 bullet
            # 如果当前已有 bullets，视为新的一段
            # 连续的纯文本行合并为一段
            para_lines = [stripped]
            j = i + 1
            while j < len(lines):
                nxt = lines[j].rstrip("\n")
                nxt_s = nxt.strip()
                if _EMPTY.match(nxt_s) or _HR.match(nxt_s) or _H1.match(nxt_s) or _H2.match(nxt_s) \
                        or _H3.match(nxt_s) or _BULLET_UL.match(nxt_s) or _BULLET_OL.match(nxt_s):
                    break
                para_lines.append(nxt_s)
                j += 1
            para_text = " ".join(para_lines).strip()
            current_bullets.append(para_text)
            i = j
            continue
        i += 1

    # 尾部处理
    if in_preamble and doc_title is not None:
        cover_slide = {"layout": "cover", "title": doc_title}
        sub = " ".join(p for p in cover_subtitle_parts if p).strip()
        if sub:
            cover_slide["subtitle"] = sub
        if meta_kwargs.get("author"):
            cover_slide["author"] = meta_kwargs["author"]
        if meta_kwargs.get("date"):
            cover_slide["date"] = meta_kwargs["date"]
        slides.append(cover_slide)
    _flush_content()

    # 如果完全没有解析出任何 slide，放一个空 cover 防止渲染错误
    if not slides:
        first_nonempty = next((ln.strip() for ln in lines if ln.strip()), "Untitled")
        title_m = _H1.match(first_nonempty)
        title_text = title_m.group(1).strip() if title_m else first_nonempty[:40]
        slides.append({"layout": "cover", "title": title_text})

    meta = {"title": doc_title or (slides[0].get("title") if slides else "Untitled")}
    meta.update(meta_kwargs)
    meta.setdefault("theme", theme)

    return {"meta": meta, "slides": slides}
