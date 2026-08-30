"""内容生成：按大纲逐节撰写 Markdown 正文。

省 token 设计（适配推理模型）：
- 短文档（<=3 节）一次调用生成全文，避免多次“思考-输出”的固定开销
- 多章节时上下文只保留最近 2 节摘要（每节 180 字符），防止历史无限膨胀
- system 提示保持完全一致，最大化 DeepSeek 提示词缓存命中（缓存 token 更便宜）
"""
from __future__ import annotations

import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Callable, Optional

from .llm import LLMClient, LLMError

STYLE_GUIDE = {
    "business": "商务汇报风格：结构清晰、结论先行、数据支撑，语言专业简洁。",
    "report": "分析报告风格：客观严谨、论证充分，适当使用图表化的表格呈现数据。",
    "academic": "学术风格：行文规范、逻辑严密，术语准确，引用格式规范。",
    "creative": "创意文案风格：有感染力、语言生动，适合宣传与文案类文档。",
    "default": "通用专业风格：条理清晰、表达流畅，适合日常办公文档。",
}

# 文档类型专项写作指南（提高实验报告等专业文档的质量）
DOC_TYPE_GUIDE = {
    "experiment_report": (
        "这是一份实验/测试报告，必须遵循以下专业要求：\n"
        "- 完整覆盖：实验目的、实验原理、实验环境、实验步骤、数据记录、结果分析、实验结论（硬件/物理类实验可含误差分析，软件/数据类实验可含问题与调试记录）。\n"
        "- 章节名称与措辞严格贴合实验学科：数据库/软件类实验用「实验环境（软件与工具）」「数据记录（表结构/运行结果）」等，不要用“器材”“误差”这类物理实验措辞；硬件/物理类实验才写“器材”“误差”。\n"
        "- 数据必须用 Markdown 表格呈现（表头含单位或列名，单位用普通括号如 L (m)，不要用 LaTeX 符号），不要用纯文字罗列；表格行数要与实验数据条数一致。\n"
        "- 结果分析要定量：给出关键数值、计算过程、与理论值/预期值的偏差（百分比）。\n"
        "- 误差分析要具体：区分系统误差与随机误差，说明来源与改进措施（软件类实验如无误差概念，改为说明问题排查与结果验证）。\n"
        "- 步骤用 1. 有序列表，环境/工具用无序列表；结论简洁、可验证。"
    ),
    "weekly_report": (
        "这是周报/月报：按「本周工作 → 数据与进展 → 问题与风险 → 下周计划」组织，"
        "进展用表格或要点，成果量化，问题给出解决方案与责任人。"
    ),
    "meeting_notes": (
        "这是会议纪要：包含会议信息（时间/地点/参会人）、议题、讨论要点、决议、待办事项（责任人+期限）。"
        "待办用表格：事项｜负责人｜截止时间。"
    ),
    "prd": (
        "这是需求/方案文档：背景与目标 → 用户/场景分析 → 功能需求（优先级）→ 非功能需求 → 里程碑与验收标准。"
        "功能需求用表格：编号｜需求｜优先级。"
    ),
    "analysis": (
        "这是分析报告：结论先行，数据支撑，分维度展开，最后给出建议。关键数据用表格呈现。"
    ),
    "general": (
        "通用专业风格：条理清晰、表达流畅，重要信息用要点或表格呈现，适合正式办公文档。"
    ),
}

# 参考文件（图片/文本/docx）的使用规则
REF_GUIDE = """【用户提供了参考文件，必须遵守】
- 正文中要实际引用参考材料：文本/表格数据要写进数据记录、结果分析等对应章节，不要空泛带过。
- 图片条目形如 [IMAGE_FILE 文件名 ABS_PATH=C:/.../xx.png]，需要展示该图时，单独一行输出：![图片说明](图片绝对路径)
  必须原样保留 ABS_PATH 里的完整路径（把 \\ 换成 / 即可），图片说明要写清图的内容。
- 参考数据不完整时如实说明，不要编造数据。"""

# 每节生成的输出预算（推理模型思考占大头，需要给足空间）
SECTION_MAX_TOKENS = 8192
# 每节摘要注入的长度与保留条数（省 token）
SUMMARY_CHARS = 300  # 每节摘要长度：内容衔接质量与 token 的平衡
SUMMARY_KEEP = 2


def _build_system(language: str, style: str, doc_type: str, reference_context: str) -> str:
    """构建完全稳定的 system 提示：内容固定，便于命中提示词缓存。"""
    style_note = STYLE_GUIDE.get(style, STYLE_GUIDE["default"])
    doc_note = DOC_TYPE_GUIDE.get(doc_type, DOC_TYPE_GUIDE["general"])
    system = (
        "你是一名资深的专业文档写手。你的任务是按照给定大纲撰写文档正文。\n"
        f"写作语言：{language}。\n风格要求：{style_note}\n\n"
        f"【文档类型专项要求】\n{doc_note}\n\n"
        "Markdown 写作规则：\n"
        "- 一级章节第一行写 `# 标题`，二级章节第一行写 `## 标题`，之后不得再使用 `#`/`##`/`###` 级别的标题。\n"
        "- 段落之间用空行分隔；适当使用 `-` 无序列表、`1.` 有序列表、`| 表头 | 表头 |` 表格、`> 引用` 等结构。\n"
        "- 正文不要使用 **加粗**、*斜体*、`反引号` 等行内标记（正式文档中会造成字体粗细不一、排版混乱）。\n"
        "- 不要使用 LaTeX 数学符号（如 \\( \\)、\\frac、\\rho、\\bar 等），单位与变量用普通写法，例如 `L (m)`、`ρ (Ω·m)`。\n"
        "- 内容要具体详实：给出关键数值、条件、过程与结论，避免「综上所述」「等等」这类空话套话。\n"
        "- 内容量要求：每节正文一般不少于 150 字（实验步骤、操作流程类至少 4 个编号步骤）；整篇文档内容要充实完整，禁止只写标题和一两句空话。\n"
        "- 数据必须真实可信：表格中的数据若无参考文件依据，须标注「示例」并在对应位置说明（例如“表 1 为示例数据，请按实际记录替换”），严禁凭空伪造具体测量结果与结论。\n"
        "- 若提供了参考文件：相关章节必须引用其中的数据/表格/图片（图片按 [IMAGE_FILE ...] 规则嵌入），不得另起炉灶。\n"
        "- 只输出该小节正文 Markdown，不要输出任何多余解释。"
    )
    if reference_context:
        system += "\n\n" + REF_GUIDE + "\n\n【参考文件内容摘要】\n" + reference_context
    return system


def ensure_section_headings(markdown: str, sections: list[dict]) -> str:
    """保证每个章节在正文中都有对应标题行（按大纲顺序补位）。

    模型偶发漏写 `# 标题` 时，在正确位置补上标题：
    - 缺失标题插入到「大纲中它之后第一个已存在标题」之前；
    - 若后面没有已存在标题，则追加到文档末尾。
    已有标题保持不变，新标题严格按大纲顺序排列。
    """
    import re as _re
    heading_re = _re.compile(r"^(#{1,4})\s+(.*)$")
    lines = markdown.splitlines()

    # 标题归一化：忽略「一、」「1.」等序号前缀，用于容错匹配
    num_prefix_re = _re.compile(r"^[（(]?[0-9一二三四五六七八九十]+[、．.）)]?\s*")

    def _norm_title(t: str) -> str:
        t = t.strip()
        t2 = num_prefix_re.sub("", t).strip()
        return t2 or t

    # 文档中已出现的标题（归一化）-> 首次出现行号
    existing: dict[str, int] = {}
    for idx, ln in enumerate(lines):
        m = heading_re.match(ln.strip())
        if m:
            existing.setdefault(_norm_title(m.group(2)), idx)

    present = {s["heading"]: _norm_title(s["heading"]) in existing for s in sections}

    # 按大纲顺序：缺失标题累积，遇到已存在标题时把累积项插到它之前
    by_pos: "dict[int, list[str]]" = {}
    pending: list[str] = []
    for sec in sections:
        h = sec["heading"]
        if present.get(h):
            pos = existing[_norm_title(h)]
            if pending:
                by_pos.setdefault(pos, []).extend(pending)
                pending = []
        else:
            pending.append("#" * sec.get("level", 1) + " " + h)
    if pending:  # 大纲尾部缺失 -> 文档末尾
        by_pos.setdefault(len(lines), []).extend(pending)

    if not by_pos:
        return markdown
    # 同一位置按大纲顺序（组内正序），从后往前落位：先插组内靠后的
    for pos in sorted(by_pos, reverse=True):
        group = by_pos[pos]
        for txt in reversed(group):
            lines.insert(pos, txt)
    return "\n".join(lines)

def generate_document(
    plan: dict[str, Any],
    llm: LLMClient,
    on_section: Optional[Callable[[int, int, str], None]] = None,
    max_parallel: Optional[int] = None,
    reference_context: str = "",
) -> str:
    """生成正文 Markdown，返回完整文档（不含标题页）。

    - 短文档（<=3 节）：一次调用生成全文，省去多次推理调用的固定开销
    - 长文档：波次并行，每节一次调用；后波次只看到最近 SUMMARY_KEEP 节的摘要
    - reference_context：附加参考文件摘要文本
    - on_section(index, total, heading) 每节开始前回调，用于界面展示进度
    """
    title = plan["title"]
    language = plan.get("language", "zh-CN")
    style = plan.get("style", "default")
    doc_type = plan.get("doc_type", "general")
    sections = plan["sections"]
    total = len(sections)

    if max_parallel is None:
        max_parallel = getattr(llm.config, "concurrency", 3)
    max_parallel = max(1, int(max_parallel))

    outline_lines = [
        f"{'#' * sec['level']} {sec['heading']}" for sec in sections
    ]
    outline_text = "\n".join(outline_lines)
    system = _build_system(language, style, doc_type, reference_context)

    # ---- 短文档：一次调用生成全文 ----
    if total <= 3:
        if on_section:
            on_section(1, total, title)
        req_lines = [
            f"{'#' * sec['level']} {sec['heading']}\n"
            f"  本节要求：{sec.get('description') or '按该标题合理展开'}"
            for sec in sections
        ]
        user = (
            f"文档标题《{title}》。\n"
            "请一次性输出完整正文：每一节都以对应级别的 Markdown 标题开头"
            f"（一级 `#`、二级 `##`），节与节之间用空行分隔。\n"
            "内容必须充实完整：每节正文不少于 150 字；步骤/流程类用编号列出完整流程（不少于 4 步）；"
            "数据用表格呈现。禁止只写标题和一两句空话。\n\n"
            f"【大纲与每节要求】\n" + "\n\n".join(req_lines)
        )
        if reference_context:
            user += "\n\n【参考文件内容摘要】\n" + reference_context
        content = llm.chat(
            [{"role": "system", "content": system}, {"role": "user", "content": user}],
            temperature=0.6,
            max_tokens=SECTION_MAX_TOKENS,
        ).strip()
        return content

    # ---- 长文档：波次并行，逐节生成 ----
    def write_one(index: int, sec: dict, prev_summaries: list[str]) -> tuple[int, str]:
        heading = sec["heading"]
        level = sec["level"]
        description = sec.get("description", "")
        md_heading = "#" if level == 1 else "##"
        if on_section:
            on_section(index, total, heading)
        user = (
            f"这是第 {index}/{total} 节，文档标题《{title}》。\n"
            f"本节标题：{heading}（输出时第一行写 `{md_heading} {heading}`）\n"
            f"本节要求：{description or '按该标题合理展开'}\n"
            "本节内容必须充实：正文不少于 150 字；步骤/流程类至少 4 个编号步骤；数据用表格。\n\n"
            f"【全文大纲】\n{outline_text}\n\n"
            f"【已完成的章节摘要（仅作衔接参考，不要重复已有内容）】\n"
            + ("\n\n".join(prev_summaries) if prev_summaries else "（这是第一节）")
        )
        messages = [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ]
        try:
            content = llm.chat(messages, temperature=0.6, max_tokens=SECTION_MAX_TOKENS).strip()
        except LLMError:
            # 单节失败重试一次（如偶发截断/空响应），仍失败才上抛
            content = llm.chat(messages, temperature=0.6, max_tokens=SECTION_MAX_TOKENS).strip()
        return index, content

    parts: list[str] = [""] * total
    prev_summaries: list[str] = []
    wave_start = 0
    while wave_start < total:
        wave = list(enumerate(sections, 1))[wave_start: wave_start + max_parallel]
        snapshot = list(prev_summaries)
        with ThreadPoolExecutor(max_workers=len(wave)) as pool:
            futures = {
                pool.submit(write_one, idx, sec, snapshot): idx for idx, sec in wave
            }
            outcomes: dict[int, Any] = {}
            for fut in as_completed(futures):
                idx = futures[fut]
                try:
                    outcomes[idx] = fut.result()
                except Exception as exc:  # noqa: BLE001
                    outcomes[idx] = exc
        for idx, sec in wave:
            result = outcomes.get(idx)
            if isinstance(result, Exception):
                raise result
            content = result[1] if isinstance(result, tuple) else result
            parts[idx - 1] = content
            prev_summaries.append(f"[{idx}] {sec['heading']}\n{content[:SUMMARY_CHARS]}")
            if len(prev_summaries) > SUMMARY_KEEP:
                prev_summaries = prev_summaries[-SUMMARY_KEEP:]
        wave_start += len(wave)

    return "\n\n".join(part for part in parts if part)

# ---------------------------------------------------------------- 内容充实度补全
# 与 verify.py 的 MIN_SECTION_BODY 保持一致：低于此值视为「半截内容」
MIN_SECTION_BODY_CHARS = 40


def _split_sections(markdown: str) -> tuple[list, list]:
    """把 markdown 拆成节列表，返回 (行列表, 节列表)。

    节结构：{"level": 1, "heading": "标题", "start": 行号, "end": 行号(不含)}。
    """
    heading_re = re.compile(r"^(#{1,4})\s+(.*)$")
    lines = markdown.splitlines()
    sections: list[dict] = []
    cur = None
    for i, ln in enumerate(lines):
        m = heading_re.match(ln)
        if m:
            if cur is not None:
                cur["end"] = i
                sections.append(cur)
            cur = {"level": len(m.group(1)), "heading": m.group(2).strip(),
                   "start": i, "end": None}
    if cur is not None:
        cur["end"] = len(lines)
        sections.append(cur)
    return lines, sections


def expand_short_sections(
    markdown: str,
    plan: dict[str, Any],
    llm: LLMClient,
    log: Optional[Callable[[str], None]] = None,
    reference_context: str = "",
) -> str:
    """把正文过短的章节做一次批量扩写（仅 1 次 LLM 调用），返回补全后的 markdown。

    只扩写「已生成但正文 < MIN_SECTION_BODY_CHARS 字」的章节；没有过短章节时
    原样返回，不消耗任何 token。失败时静默保留当前内容，不影响主流程。
    """
    lines, sections = _split_sections(markdown)
    short: list[dict] = []
    for sec in sections:
        body = "\n".join(lines[sec["start"] + 1: sec["end"]])
        if len(re.sub(r"\s", "", body)) < MIN_SECTION_BODY_CHARS:
            short.append({**sec, "body": body})
    if not short:
        return markdown

    doc_type = plan.get("doc_type", "general")
    style = plan.get("style", "default")
    title = plan.get("title", "")
    brief = "\n\n".join(
        "### 章节：" + s["heading"] + "\n当前内容：\n" + ((s["body"].strip() or "（空）")[:400])
        for s in short
    )
    system = (
        "你是一名严谨的文档写手，负责把文档中「内容过短、没写完」的章节补全完整。\n"
        "要求：\n"
        "- 严格只输出 JSON，格式：sections 数组，每项含 heading 与 content 两个字段。\n"
        "- content 只写该章节正文（不含标题行），可用 `-` 无序列表、`1.` 有序列表、`| 表格 |`、`> 引用` 等结构。\n"
        "- 正文不要使用 **加粗**、*斜体*、`反引号` 等行内标记；不要使用 LaTeX 数学符号。\n"
        "- 每节内容要充实完整：说明/分析/总结类 120~250 字；步骤/流程类用编号列出完整流程（不少于 4 步，写明具体操作、参数与观察/结果）。\n"
        "- 数据若无参考文件依据必须标注「示例」，严禁凭空伪造具体测量结果与结论。\n"
        "- 保持与文档整体风格一致，不新增章节标题，不重复其他章节的内容。"
    )
    user = (
        "文档标题《" + title + "》。文档类型：" + (doc_type or "general")
        + "。\n下面这些章节正文过短，请补全（保留原意，只做扩写，不要改章节标题）：\n\n"
        + brief
        + (("\n\n【参考文件内容摘要】\n" + reference_context) if reference_context else "")
    )
    try:
        data = llm.chat_json(
            [{"role": "system", "content": system}, {"role": "user", "content": user}],
            temperature=0.4, max_tokens=8192,
        )
    except Exception as exc:  # noqa: BLE001
        if log:
            log("   ⚠ 过短章节补全失败（" + str(exc) + "），保留当前内容")
        return markdown

    replacements: dict[str, str] = {}
    for item in data.get("sections") or []:
        if not isinstance(item, dict):
            continue
        h = str(item.get("heading", "")).strip()
        c = str(item.get("content", "")).strip()
        if h and c:
            replacements[h] = c
    if not replacements:
        return markdown

    # 标题容错匹配：容忍「1.」「一、」等序号前缀
    num_prefix_re = re.compile(r"^[（(]?[0-9一二三四五六七八九十]+[、．.）)]?\s*")

    def _norm(t: str) -> str:
        t2 = num_prefix_re.sub("", t.strip()).strip()
        return t2 or t.strip()

    target = {_norm(h): h for h in replacements}
    new_lines = list(lines)
    fixed = 0
    # 从后往前拼接，避免替换后行号偏移
    for sec in reversed(sections):
        key = _norm(sec["heading"])
        if key not in target:
            continue
        content = replacements[target[key]]
        block = [ln for ln in content.splitlines() if ln.strip()]
        new_lines[sec["start"] + 1: sec["end"]] = block
        fixed += 1
    if log and fixed:
        names = ", ".join(replacements[target[k]] for k in list(target)[:3])
        log("   ⚙ 已自动补全 " + str(fixed) + " 个过短章节（" + names + "…）")
    return "\n".join(new_lines)
