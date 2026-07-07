---
name: pro-ppt-gen
description: Use when users need to create or edit PPT/PPTX decks, slides, presentations, courseware, reports, pitches, timelines, charts, KPI pages, image/text layouts, references, speaker notes, page numbers, logo/watermark, Markdown-to-PPTX output, or taste/quality preflight with PRO-PPTX semantic JSON.
---

# PRO-PPTX v1.6.4 — 专业级 PPT 生成（PaperJSX 语义编译架构）

> **核心理念：LLM 只写语义，不碰坐标和样式数值。** 所有位置、字号、颜色、间距由确定性布局引擎基于网格系统和设计令牌计算，输出稳定、可对齐、美观的 16:9 宽屏 PPT。

---

## 版本亮点 v1.6.4（2026-07）

- **可视化重设计内核**：新增 `visual_redesign_check()` / `build_visual_redesign_prompt()` / `apply_visual_redesign_guidance()`，把“内容解构 → 关系识别 → 分页规划 → 信息分层 → 布局设计 → 配色 → prompt”下沉为生成前规划能力。
- **防平铺质量门**：`taste_check()` 默认接入防平铺检查；纯“矩形色块 + 标题 + bullets”触发 `word_doc_on_slide`，核心页结构要素不足触发 `flat_layout_risk`。
- **信息图结构要素要求**：核心叙事页默认至少包含 2 类结构要素：连接线/箭头、节点、空间对比、环形/循环、进度/状态、非对称焦点、层级结构、数据可视化。
- **发布包修复**：补齐 `shared.import_helper` 与 `shared.taste.*`，修复独立导入失败；新版 zip 不包含 `.git`。

## 版本亮点 v1.6.3（2026-07）

- **🆕 12 套场景化模板**（`template_name` 参数）：高考复习 `gaokao_review` / 答辩 `thesis_defense` / 路演 `startup_pitch` / 年会 `year_end_review` / 培训 `training_workshop` / 产品发布 `product_launch` / 学术会议 `academic_conference` / 项目汇报 `project_report` / 极简 `minimal_clean`，覆盖原 academic/business/teaching。
- **🆕 本地 WPS/Office 模板主题自动提取**：`template_path` 参数直接传 .pptx/.potx/.dpt 文件，自动提取色板/字体/背景；`scan_local_templates()` 扫描默认 WPS 模板目录；`check_template_match()` 在 `auto_taste_match=True` 时对 tokens 做 WCAG AA 对比度校验。
- **🆕 18 种元素动画**（`animation` 字段）：入场 10（fade_in/appear/fly_in_left/fly_in_bottom/wipe/zoom/float/split…）+ 强调 4（pulse/grow_shrink/color_shift/underline_reveal）+ 退出 4（fade_out/fly_out_right/zoom_out/dissolve_out）；支持裸字符串预设名 shortcut。
- **🆕 13 种切换**：基础 6（fade/dissolve/push/wipe/cover/split）+ 方向别名 7（wipe_right/push_down/cover_up/split_vertical/…）+ 扩展 6（cube/reveal/zoom/ferris/gallery/conveyor，p14 命名空间，PowerPoint 2010+）。
- **trigger 枚举扩展**：on_click / auto_after / with_prev / after_prev（with_prev/after_prev 与前一动画同步/接续）。
- **向后兼容 100%**：不声明 animation/transition 时零 OOXML 新节点；所有 v1.0–v1.5 JSON 不改字段即可继续使用。

## 版本亮点 v1.5

- **`taste_check()` 审美/手感预检**：吸收 Impeccable 与 Taste skill 的可迁移规则，检查设计读法、版式节奏、AI 味文案、占位文本、视觉素材意图、图表洞察和交付前清单。
- **质量门增强**：`quality_gates/run_quality_gate.py` 对已生成 PPTX 增加轻量 AI 味/重复文本/破折号扫描；semantic JSON 阶段优先使用 `ppt_jsx.taste_check(content)`。
- **随包 smoke tests 覆盖 v1.5**：新增 `PPT taste check`，确保新接口随包可用。

### v1.4 保留亮点

- **演讲者备注（speaker notes）**：所有 14 种版式支持 `notes` 字段，写入 PPT 备注页（演示者视图可见）
- **Markdown 直入**：`generate_from_markdown(md_str)` 支持 # 标题→封面、##→章节、- 列表→bullets 直接生成 PPT
- **references 参考文献版式**：新增 `references` layout，自动分页渲染学术参考文献列表（与 PRO-DOCX 共享 citation.py）
- **PPT SVG 迷你图**：`mini_bar` / `mini_line` / `mini_pie` 三种 SVG 小图表，200-400px 尺寸嵌入 KPI/内容页
- **幻灯片切换效果**：`transition` 字段支持 fade/push/wipe/split/cover/dissolve/none
- **自定义 Logo 水印**：`meta.logo` 支持本地路径或 http(s) URL（自动 5s 超时下载），5 方位可选（tl/tr/bl/br/center）
- **页码 4 种样式**：`meta.page_number_style=plain|slash|chinese|roman`（"5" / "5/12" / "第5页" / "V"），自动跳过封面/目录
- **致谢页增强**：`thanks` 版式新增 `signature`（签名图）、`qr_code`（二维码）、`contact`（联系信息块）字段
- **统一主题字典**：主题色板下沉到 `shared/themes.py`，与 PRO-DOCX 完全对齐
- **输出文件命名**：默认按 `{title}_v{version}.pptx` 自动命名
- **中英双语 docstring**：关键 API 与版式类均提供中英双语文档
- **中文错误提示 + 修复建议**：validators 校验失败时给出明确原因与修改建议
- **References 专业排版**：悬挂缩进 + 方括号编号 + DOI 蓝色下划线；超长自动分页

---

## 1. 六阶段工作流（渐进式披露）

生成一份正式 PPT 推荐按以下六步走，避免一次把大段 JSON 丢给引擎后返工：

### 阶段 1 · 调研（Research，按需）
当主题不熟悉、需要事实数据或案例支撑时，先用 `search_web` 做两层搜索：
- **广度搜索**（≤10 个关键词）：快速扫一遍主题全貌、关键概念、最新数据。
- **深度搜索**（≤5 个关键词）：针对已经确定的章节要点查权威来源、具体数字、引用文献。
调研结果沉淀为后续大纲和 bullets 的事实来源。**这一步是 Agent 工作流，引擎不强制搜索——但对正式汇报/答辩强烈推荐。**

### 阶段 2 · 大纲（Outline）
先写**只含 layout + 标识字段**的精简 JSON（每页只需 layout 和 title/number 等最小字段），调用 `ppt_jsx.outline(content)` 得到页数和标题列表，让用户确认页数、顺序、章节划分。如果 LLM 不想手动指定 layout，可以跳过 layout 字段，调用 `ppt_jsx.auto_layout(content)` 自动推断每页版式，再进入大纲确认。

### 阶段 2.5 · 可视化重设计规划（Visual Redesign）
正式 PPT 渲染前先做视觉规划：
- 调用或参考 `build_visual_redesign_prompt(content)`，先解构内容块、识别关系、判断单页/多页、分配页面角色。
- 核心叙事页禁止平铺：不得只用等大矩形色块承载“标题 + bullets”。
- 每个核心叙事页必须有至少 2 类结构要素：连接线/箭头、节点图形、空间对比、环形/循环、进度/状态、非对称焦点、层级结构、数据可视化。
- 多页核心叙事页之间不能全用同一种布局结构；按关系分别选流程、对比、层级、循环、看板、编辑式大焦点等结构。
- 复杂多页 PPT 生成前，先给用户展示“防平铺总览”：每页页面角色 + 布局结构 + 结构要素，再进入生成。

```python
plan_prompt = ppt_jsx.build_visual_redesign_prompt(content)
guided = ppt_jsx.apply_visual_redesign_guidance(content)
visual = ppt_jsx.visual_redesign_check(guided, strict=True)
```

### 阶段 3 · 配图（Image Hunt）
对 `full_image` 和 `image_text` 页面（以及需要插图的 `content` 页）：
1. 调用 `search_images` 工具按页面主题搜索真实图片（优先真实图片，不用 AI 生成图，避免风格不一致）。
2. 将图片 URL（或下载到本地的路径）填入 `image_path` 字段。
3. `image_path` 直接支持 `http(s)://` URL，引擎会自动 6 秒超时下载到临时文件嵌入；下载失败时显示占位框不影响整体渲染。
配图目录建议放在 `./codeact/output/ppt_assets/` 下。

### 阶段 4 · 内容填充
根据确认的大纲，逐页填充 `bullets` / `table` / `chart` / `image_path` / KPI 数字 / `events` 等正文内容。严格遵守字数建议（见 §4）。每页可选填 `notes: "讲稿..."` 字段，写入 PPT 演讲者备注（演示者视图可见）。

### 阶段 5 · 字数与密度预检
调用 `ppt_jsx.estimate_length(content)`，返回三类问题页面：
- **overflow_slides**：bullets/图表洞察可能溢出页面边界；
- **dense_slides**：信息密度 high，建议拆页或精简；
- **sparse_slides**：信息密度 low，建议合并或补充内容。
LLM 需在渲染前精简/调整这些页面。

### 阶段 6 · 渲染输出
调用 `ppt_jsx.generate(content, output_path, theme, lang)` 生成 `.pptx` 文件。

---

## 2. 14 种版式语义（Semantic JSON 规范）

每页是一个对象，`layout` 字段指定版式，其余字段按版式要求填写。**v1.2 新增 timeline 版式和 notes 通用字段。**

| layout | 用途 | 必填字段 | 可选字段 |
|---|---|---|---|
| `cover` | 封面 | `title` | `subtitle`, `author`, `date`, `institution`, `notes` |
| `toc` | 目录 | `items` (≤8 项) | `title`（默认"目录"）, `notes` |
| `section` | 章节分隔 | `number`, `title` | `notes` |
| `content` | 标题+要点/图表 | `title`（bullets/chart 至少一个） | `bullets`, `chart`, `notes` |
| `two_column` | 双栏（任意图文/图表组合） | `title`, `left`, `right` | `notes` |
| `image_text` | 图文混排 | `title`, `image_path`, `bullets` | `side`（"left"/"right"，默认 left）, `notes` |
| `full_image` | 全图页 | `image_path` | `title`, `overlay`（bool）, `notes` |
| `table` | 表格 | `title`, `headers`, `rows` | `caption`, `notes` |
| `chart` | 整页原生图表 | `title`, `chart` | `caption`, `bullets`（≤3 条洞察）, `notes` |
| `kpi` | 数字卡片（2-4 张） | `title`, `items` | `notes` |
| `quote` | 金句/引言 | `text` | `attribution`, `notes` |
| `summary` | 总结页 | — | `title`（默认"总结"）, `bullets`, `chart`, `notes` |
| `thanks` | 致谢/Q&A | — | `title`（默认"感谢聆听"）, `subtitle`（默认"Q & A"）, `notes` |
| `timeline` | **[v1.2]** 时间轴/里程碑 | `title`, `events` (≤6 项) | `notes` |

**字段细节：**

- `bullets`：字符串数组，或 `{"text": str, "level"?: 1|2|3}` 对象数组，支持三级缩进（●/○/▪）。
- `items` (toc)：`[{"num": "01", "text": "..."}]` 或直接字符串数组。
- `items` (kpi)：`[{"value": "98%", "label": "准确率", "subtext": "同比+5%"}]`。
- `left` / `right` (two_column)：`{"title"?: str, "bullets"?: [...], "chart"?: {...}, "notes"?: [...]}`；每栏可放 bullets 或 chart（二选一），chart 下可选 `notes` 作为简短注释。
- `image_path`：本地文件路径**或 http(s):// URL**（v1.2 新增，引擎自动下载嵌入，6s 超时容错）。
- `events` (timeline)：`[{"date": "2024 Q1", "title": "里程碑", "desc": "简要描述"}]`，≤6 项，横向均匀分布。
- `notes`：每页可选的演讲者备注字符串，写入 PPT 备注页（演示者视图可见）。

---

### 2.5 图表使用（`chart` 字段语义）

所有图表都是 **原生可编辑 OOXML Chart**（PowerPoint/WPS 中双击可改数据），不是图片。

**`chart` 对象字段：**

| 字段 | 类型 | 说明 |
|---|---|---|
| `type` | string | 图表类型，见下表（必填） |
| `title` | string | 图表区内标题（可选） |
| `categories` | string[] | 分类/X 轴标签（必填，scatter 除外） |
| `series` | array | `[{"name": "系列名", "values": [数字,…]}]`；scatter 时 values 为 `[{"x":..,"y":..},…]` |
| `show_legend` | bool | 默认 true；**单系列柱/条/折线/面积/雷达自动隐藏** |
| `legend_position` | string | `"bottom"`/`"right"`/`"top"`/`"left"`，默认 bottom |
| `show_data_labels` | bool | 默认 false；**饼图/环形图默认 true 且显示百分比+类别名** |
| `number_format` | string | 默认 `"0"`；`"$#,##0"` / `"0%"` 等 |
| `y_axis` | object | `{"min"?: number, "max"?: number, "title"?: string}`（可选） |

**支持的 `type` 值：** `column`、`bar`、`stacked_column`、`stacked_bar`、`line`、`line_markers`、`pie`、`doughnut`、`area`、`scatter`、`radar`。

**配色**：引擎从主题色板自动取色循环，LLM 无需指定颜色。

---

## 3. 主题系统（Themes v1.2 升级）

### 3.1 内置主题（14 套）

| theme | 主色 | 风格定位 | 适用场景 |
|---|---|---|---|
| `academic` | 深蓝 #1F3864 + 红棕强调 | 严谨正式（v1.1 默认） | 学位答辩、学术报告 |
| `business` | 蓝灰 #2C3E50 + 橙色强调 | 现代商务（v1.1 默认） | 企业汇报、融资路演 |
| `teaching` | 绿色 #2E7D32，字号更大 | 清晰明亮（v1.1 默认） | 教学课件、培训讲座 |
| `tech` | 深海军蓝 #0A192F + 青蓝霓虹 | 未来科技感 | AI/互联网/产品发布 |
| `dark` | 近黑 #0F172A + 蓝色/琥珀强调 | 暗色沉浸 | 产品发布会、深色模式场景 |
| `ocean` | 深海蓝 #0C4A6E + 青色 | 清凉通透 | 数据、金融、旅行 |
| `forest` | 深绿 #064E3B + 翡翠绿 | 生机自然 | 环保、农业、健康 |
| `nature` | 军绿 #365314 + 土黄 | 大地质感 | 自然、户外、食品 |
| `warm` | 酒红 #7F1D1D + 琥珀 | 温暖柔和 | 生活方式、文创 |
| `sunset` | 棕红 #7C2D12 + 橙粉紫 | 活力时尚 | 营销、潮流 |
| `premium` | 黑 #1C1917 + 金色 #D4AF37 | 高端奢华 | 奢侈品、品牌发布、投资 |
| `chinese_red` | 中国红 #8B0000 + 金色 | 国风喜庆 | 节日、政务、传统文化 |
| `light` / `minimal` | 纯黑白 + 一抹蓝 | 极简苹果风 | 设计作品集、极简报告 |
| `high_contrast` | 黑白高对比 | 高可读/打印友好 | 黑白打印、无障碍场景 |

### 3.2 自然语言主题查询

`theme` 参数直接传中文描述即可，引擎自动关键词匹配：

```python
ppt_jsx.generate(content, "out.pptx", theme="科技感深色背景")  # → tech
ppt_jsx.generate(content, "out.pptx", theme="中国风红色庆典")  # → chinese_red
ppt_jsx.generate(content, "out.pptx", theme="极简白苹果风")    # → minimal/light
ppt_jsx.generate(content, "out.pptx", theme="高端黑金质感")    # → premium
```

### 3.3 自定义主题

传 dict 或 `{"name": "<基础主题>", "overrides": {...}}` 即可局部覆写：

```python
from pptx.dml.color import RGBColor

my_theme = {
    "color": {
        "primary": RGBColor(0x4A, 0x14, 0x8C),   # 紫色主色
        "accent":  RGBColor(0xF5, 0x9E, 0x0B),   # 琥珀强调
        "title_bar_bg": RGBColor(0x4A, 0x14, 0x8C),
    }
}
ppt_jsx.generate(content, "out.pptx", theme=my_theme)

# 或在 business 基础上改一个颜色:
ppt_jsx.generate(content, "out.pptx", theme={"name": "business", "overrides": {"color": {"accent": RGBColor(0xFF, 0x00, 0x80)}}})
```

可覆写字段包括 `color.primary/secondary/accent/bg/text/text_light/title_bar_bg/table_header_bg/section_bg/kpi_bg/quote_border/divider/chart_palette[...]`，以及 `font.size.*`、`font.family.*`。

---

## 4. 字数建议（避免溢出）

| 版式 | 建议上限 |
|---|---|
| `content` / `summary`（纯 bullets） | ≤ 6 条 / 页 |
| `content`（bullets + chart 共存） | bullets ≤ 3 条 |
| `two_column`（双栏 bullets） | 每栏 ≤ 5 条 |
| `image_text` | ≤ 5 条 |
| `chart`（整页图表） | 洞察 bullets ≤ 3 条，每条 ≤ 30 字 |
| `toc` | ≤ 8 个章节项 |
| `kpi` | 2-4 张卡片 |
| `table` | ≤ 10 行 |
| `timeline` | ≤ 6 个事件 |
| 单条 bullet | ≤ 30 个汉字（或 40 英文字符） |

`estimate_length()` 会自动给出 overflow / dense / sparse 三类预警，渲染前务必检查。

### 4.1 防平铺铁律

“平铺”= 页面只由若干矩形色块上下/左右堆叠，每块都是标题 + 文字列表，缺少连接线、节点、关系结构、节奏变化。这种页面等同 Word 文档上色，必须重做。

合格核心页至少满足 2 项：连接线/箭头、节点图形、空间对比、环形/循环、进度/状态、非对称焦点、层级/树形/矩阵、数据可视化。

运行：
```python
report = ppt_jsx.visual_redesign_check(content, strict=True)
```
`word_doc_on_slide` 和 strict 模式下的 `flat_layout_risk` 必须修复后再渲染。

---

## 5. v1.2 新增 API

### 5.1 `ppt_jsx.auto_layout(content, in_place=False) -> content`

根据每页已有语义字段自动推断 `layout`，免去 LLM 手动指定版式的负担。推断规则：
- `events` → `timeline`
- `image_path` 无 bullets → `full_image`；有 bullets → `image_text`
- `chart` 无 bullets → `chart`
- `headers+rows` → `table`
- `left+right` → `two_column`
- `items` 含 value/label 字典 → `kpi`；否则 → `toc`
- `number+title` → `section`
- `title+subtitle/author/...` → `cover`
- 显式 `title` 为"感谢聆听"等 → `thanks`
- `text` 字段 → `quote`
- 其余 → `content`

已显式设置的 layout 字段不会被覆盖。

### 5.2 `ppt_jsx.estimate_length(content) -> dict`

v1.2 新增三个返回字段：
- `dense_slides`: 信息密度过高的页码列表（建议拆页/精简）
- `sparse_slides`: 信息密度过低的页码列表（建议合并/补充）
- `pages[i].density`: `"low" / "medium" / "high"` 三级评分

### 5.2.1 `ppt_jsx.taste_check(content, theme="academic", lang="cn", strict=False) -> dict`

v1.5 新增审美/手感预检。它不会渲染文件，只检查 semantic JSON：

```python
report = ppt_jsx.taste_check(content, theme="business", lang="cn")
print(report["score"], report["passed"])
for issue in report["issues"]:
    print(issue["level"], issue["code"], issue.get("page"), issue["message"])
```

返回内容包括：
- `design_read`: 根据主题、版式和内容推断的一句话设计读法，以及 `design_variance` / `motion_intensity` / `visual_density` 三个拨盘值。
- `preflight`: 交付前检查清单，覆盖文案自检、内容密度、视觉意图、版式节奏、基础质量和叙事结构。
- `issues`: 具体问题列表，例如 `placeholder_text`、`generic_ai_copy`、`missing_visual_asset`、`missing_chart_insight`、`low_layout_variety`。
- `score` / `passed`: 默认 75 分通过；`strict=True` 时阈值提高到 85。
- `visual_redesign`: v1.6.4 新增，返回防平铺和信息图结构检查报告。

### 5.2.2 可视化重设计 API

```python
ppt_jsx.visual_redesign_check(content, strict=False)
ppt_jsx.build_visual_redesign_prompt(content, page_count=None)
ppt_jsx.apply_visual_redesign_guidance(content)
```

- `visual_redesign_check`: 质量门，检查防平铺、结构要素、核心页重复。
- `build_visual_redesign_prompt`: 给 agent 生成 PPT 前的融合式规划 prompt。
- `apply_visual_redesign_guidance`: 给 semantic JSON 增加 `visual_plan` 建议，不破坏原字段。

### 5.3 单页修改 API

```python
# 替换第 N 页（1-based）
ppt_jsx.update_slide("output.pptx", 3, {"layout":"content", "title":"新标题", "bullets":[...]})

# 在指定位置插入新页（position=None 追加；position=-1 在最后一页前插入）
ppt_jsx.add_slide("output.pptx", {"layout":"content", "title":"新页", "bullets":[...]}, position=3)

# 删除第 N 页
ppt_jsx.delete_slide("output.pptx", 5)
```

所有三个 API 都支持 `output_path=` 参数写入新文件，否则原地修改；支持 `theme=` / `lang=` 参数。

### 5.4 远程图片

`image_path` 字段直接填写 URL 即可：

```json
{"layout": "image_text", "title": "示例", "image_path": "https://example.com/photo.jpg", "bullets": [...]}
```

引擎自动以 6 秒超时（连接+读取）下载到临时目录后嵌入；下载失败或超时会显示灰色占位框，整体渲染不会中断。

### 5.5 演讲者备注

每页任意加 `"notes": "演讲稿..."` 字段，内容写入 PPT 备注页，演示者视图可见：

```json
{"layout": "content", "title": "关键发现", "bullets": [...], "notes": "这一页讲 2 分钟，重点强调第二个数据..."}
```

### 5.6 页码

正文页页脚右侧自动显示 `X / Y` 当前页/总页；封面/目录/章节/致谢页不显示页码，保持视觉干净。

---

## 6. 中文字体自动适配

- 英文使用 `Arial`，中文使用 `微软雅黑`（通过 OOXML `eastAsia` 属性设置，确保中文在 PowerPoint/WPS 中正常显示；图表内文字、图例、坐标轴、数据标签、备注也统一设置）。
- 金句/引言页使用 `Georgia` 斜体增强文学感（中文仍回退到微软雅黑）。
- premium 主题的标题字体默认使用 `Georgia`，更具排版质感。
- 无需在 JSON 中指定字体，引擎按主题和语言自动处理。

---

## 7. 最小可运行示例

```python
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from skills.pro_ppt_gen import ppt_jsx

content = {
    "slides": [
        {
            "layout": "cover",
            "title": "2026 H1 经营复盘",
            "subtitle": "数据驱动 · 持续增长",
            "author": "战略部",
            "date": "2026年7月",
            "notes": "开场欢迎，介绍议程",
        },
        {
            "layout": "section",
            "number": "01",
            "title": "核心业绩",
        },
        {
            "layout": "timeline",
            "title": "季度里程碑",
            "events": [
                {"date": "Q1", "title": "产品发布", "desc": "用户破百万"},
                {"date": "Q2", "title": "B轮融资", "desc": "估值翻倍"},
            ],
        },
        {
            "layout": "chart",
            "title": "季度营收对比",
            "chart": {
                "type": "column",
                "categories": ["Q1", "Q2", "Q3", "Q4"],
                "series": [
                    {"name": "2025", "values": [100, 120, 140, 160]},
                    {"name": "2026", "values": [120, 150, 0, 0]},
                ],
                "y_axis": {"title": "营收（百万元）", "min": 0},
            },
        },
        {"layout": "thanks"},
    ],
}

# 预检（v1.2: 含密度评估；v1.5: 审美/手感预检）
report = ppt_jsx.estimate_length(content)
craft = ppt_jsx.taste_check(content, theme="tech", lang="cn")
print("overflow:", report["overflow_slides"],
      "dense:", report["dense_slides"],
      "sparse:", report["sparse_slides"])
print("taste:", craft["score"], craft["passed"])

# 渲染（v1.4: theme 支持自然语言 / dict 自定义 / 14 套主题名）
out = ppt_jsx.generate(content, "./output/demo.pptx", theme="tech", lang="cn")
print("已生成：", out)
```

### 随包 Smoke Tests

解包后可在包根目录运行：

```bash
python smoke_tests/run_smoke_tests.py
```

该脚本当前只做随包结构与导入级冒烟检查：确认关键目录和脚本存在，并验证临时 `skills/` 布局下 `shared` 兼容别名可用。

### 交付质量门

正式交付前建议运行：

```bash
python quality_gates/run_quality_gate.py output/deck.pptx --json-report output/quality_report.json
```

质量门当前只检查 OOXML 包结构、DOCX/PPTX 主文档 XML 是否存在，以及占位符文本（如 `TODO`、`TBD`、`[Image not found]`、`[Image unavailable]`）。若需要把警告也视为失败，可加 `--strict`。

---

## 8. 模块架构速览

```
pro_ppt_gen/
├── SKILL.md                # 本文档
├── __init__.py             # 版本号 v1.6.3
├── ppt_jsx.py              # 对外 API：generate/outline/estimate_length/auto_layout/
│                           #            taste_check/update_slide/add_slide/delete_slide/ThemeFactory
├── tokens/
│   ├── design_tokens.py    # 设计令牌（字号/间距/颜色/圆角/图表尺寸）
│   └── themes.py           # 14 套主题 + ThemeFactory（NL 解析/自定义合并）
├── engine/
│   ├── layout.py           # 12 列网格系统 + 文本高度估算 + 自动缩放
│   ├── renderer.py         # python-pptx 封装（形状/文字/表格/远程图片下载/备注/页码）
│   └── chart_renderer.py   # 原生 OOXML 图表渲染
├── compiler/
│   ├── parser.py           # 语义 JSON → 布局节点分发 + notes 集中处理
│   └── validators.py       # 字段校验 + 字数预警 + 图表/timeline 结构校验
├── smoke_tests/            # 随包冒烟测试
├── quality_gates/          # 交付前结构/渲染质量门
└── layouts/                # 14 个版式处理器
    ├── title.py / toc.py / section.py
    ├── content.py / two_column.py / image_text.py / full_image.py
    ├── table.py / chart.py / kpi.py / quote.py / summary.py / thanks.py
    ├── timeline.py         # [v1.2] 时间轴版式
    ├── _helpers.py         # pt/emu 转换工具
    └── _titlebar.py        # 通用标题栏绘制
```

**核心保证：**
- 所有坐标 = 网格系统计算，无硬编码 left/top。
- 所有字号/颜色/图表色板 = 从 `TOKENS` 读取，按主题 deep-merge 覆盖。
- 所有中文文字（含图表/备注）= 通过 `a:ea` 标签设置微软雅黑。
- 所有图表 = `slide.shapes.add_chart` 原生 OOXML，可双击编辑。
- 远程图片 6 秒超时自动降级为占位框，不中断渲染。
- v1.1 的 13 版式 + 三主题 + 所有 JSON 100% 兼容，无需迁移。
