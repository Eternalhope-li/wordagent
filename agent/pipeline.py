"""高层编排：指令 -> 大纲 -> 正文 -> docx（控制台与桌面版共用）。"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

from .config import Config
from .llm import LLMClient
from .memory import Memory
from .planner import parse_request
from .refdocs import collect_reference_context
from .renderer import render_markdown_to_docx
from .standards import get_standard, standard_to_preset
from .template_reader import extract_template, structure_text, to_renderer_preset
from .writer import ensure_section_headings, generate_document


def safe_filename(text: str) -> str:
    """去除文件名非法字符。"""
    cleaned = re.sub(r'[\\/:*?"<>|\r\n]+', "", text).strip()
    return cleaned or "未命名文档"


def _extract_section_body(markdown: str, heading: str) -> str:
    """从生成的分节 markdown 中抽取指定章节标题后的正文（直到下一个同级标题）。"""
    import re as _re
    norm = _re.sub(r"\s+", "", heading)
    lines = markdown.splitlines()
    start = -1
    for i, ln in enumerate(lines):
        m = _re.match(r"^#{1,4}\s+(.*)$", ln.strip())
        if m:
            t = _re.sub(r"\s+", "", m.group(1))
            if t == norm or norm in t or t in norm:
                start = i
                break
    if start == -1:
        return ""
    body = []
    for ln in lines[start + 1:]:
        if _re.match(r"^#{1,4}\s+", ln.strip()):
            break
        body.append(ln)
    # 清理 markdown 标记（正式文档正文不需要）
    text = "\n".join(body).strip()
    text = _re.sub(r"^#{1,6}\s+", "", text, flags=_re.M)
    return text


def run_pipeline(
    command: str,
    config: Config,
    memory: Memory,
    log: Callable[[str], None] = print,
    style_override: Optional[str] = None,
    output_dir_override: Optional[Path] = None,
    llm: Optional[LLMClient] = None,
    reference_files: Optional[list] = None,
    template_path: Optional[Path] = None,
    confirm_plan: Optional[Callable[[dict], bool]] = None,
    standard_override: Optional[str] = None,
) -> Optional[Path]:
    """执行完整流程，返回生成的 docx 路径。

    llm 可注入自定义客户端（测试或代理场景使用），默认按 config 创建。
    reference_files：可选的参考文件列表（图片/文本/csv/docx），内容会注入规划与写作提示。
    template_path：可选模板 .docx。提供时，章节结构严格按模板，排版格式也按模板。
    未提供时，会自动把 reference_files 中的 .docx 当作模板（模板就是参考文件）。
    
    """
    memory.add_user(command)
    llm = llm or LLMClient(config)

    refs = collect_reference_context(reference_files, log=log)

    template_preset: Optional[dict] = None
    template_sections: Optional[list] = None
    tpl_struct_text = ""
    # 模板即参考文件：reference_files 里的 docx 自动作为模板（格式+结构），无需单独指定
    if not template_path and reference_files:
        for _f in reference_files:
            if str(_f).lower().endswith(".docx"):
                template_path = Path(_f)
                break
    if template_path and Path(template_path).exists():
        tpl_info = extract_template(template_path)
        template_preset = to_renderer_preset(tpl_info)
        template_sections = tpl_info.get("sections")
        tpl_struct_text = structure_text(tpl_info)
        if template_preset:
            log(f"📋 已读取模板：{Path(template_path).name}（页面/字体/标题样式将按模板）")
        if template_sections:
            log(f"📋 章节结构将严格按模板：共 {len(template_sections)} 个标题")
        # 模板正文里写的要求/说明也必须传给 AI（生成模式）
        try:
            from .template_reader import full_text as _tpl_full_text
            _full = _tpl_full_text(template_path)
            if _full:
                _blk = "\n\n### 模板正文全文（模板里写的要求/说明/注意事项必须遵守）\n" + _full[:6000]
                refs["planner_summary"] += _blk
                refs["writer_context"] += _blk
        except Exception:
            pass

    # 图片/PDF 模板（可能是模糊扫描件）：自动增强识别章节结构
    if not template_sections and reference_files:
        from .template_scanner import scan_template_file
        scan_path = next(
            (Path(f) for f in reference_files if str(f).lower().endswith(
                (".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff", ".pdf"))),
            None,
        )
        if scan_path is not None and scan_path.exists():
            scan = scan_template_file(scan_path, log=log)
            if scan and scan.get("sections"):
                template_sections = scan["sections"]
                tpl_struct_text = "\n".join(
                    f"- [{s['level']}] {s['heading']}" for s in template_sections
                )
                log(f"📋 已识别图片/PDF 模板结构：共 {len(template_sections)} 个章节")
                if scan.get("lines"):
                    refs["writer_context"] += (
                        "\n\n### 模板识别全文（OCR）\n" + "\n".join(scan["lines"])
                    )

    # 模板驱动模式：提供 docx 模板时，不再新建文档、不生成大纲，
    # 而是解析模板填写位 -> AI 填内容 -> 写入模板副本（格式 100% 继承模板，省一次 API 调用）
    if template_path and Path(template_path).exists():
        from .template_engine import analyze_template, blueprint_text, apply_fill
        from .templater import plan_fill as _plan_fill_template

        log("① 以模板为母版解析并填写（格式继承模板，不重建文档）...")
        tinfo = analyze_template(template_path)
        tfill = _plan_fill_template(command, blueprint_text(tinfo), llm)
        if tfill.get("error"):
            log(f"   ⚠ AI 分析模板失败：{tfill['error']}")
        log(f"   计划填写 {len(tfill['sections'])} 个章节、{len(tfill['tables'])} 个表格、{len(tfill['cells'])} 个填写位")

        from .templater import _log_fill_report, _save_fill_report
        import shutil as _sh
        work_path = Path(str(template_path) + ".gen.docx")
        _sh.copy(template_path, work_path)
        fill_report: dict = {}
        applied = apply_fill(work_path, tfill, log=log, report=fill_report)
        for item in applied:
            log(f"   ✓ {item}")
        if not applied:
            log("⚠ 没有可填写的位置：模板可能已完整，或章节标题/表头未匹配上。")
        output_dir = Path(output_dir_override or config.output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        title = command.strip() or Path(template_path).stem or "未命名文档"
        output_path = output_dir / f"{stamp}_{safe_filename(title)}.docx"
        _sh.move(str(work_path), str(output_path))
        # 程序化质量门禁：修复安全项并复检（officecli Verification Gate，不消耗 token）
        try:
            from .verify import fix_document as _fix_doc
            log("   ⚙ 质量检查并自动修复 ...")
            _fix_doc(output_path, log=log)
        except Exception as _exc:  # noqa: BLE001
            log(f"   ⚠ 质量检查跳过（{_exc}）")
        memory.add_result({
            "title": title, "file": str(output_path),
            "style": "template", "doc_type": "general",
            "sections": len(tfill.get("sections", [])), "mode": "template-driven",
        })
        if hasattr(llm, "usage_text"):
            log(f"   ⚙ {llm.usage_text()}")
        _log_fill_report(fill_report, log)
        _save_fill_report(output_path, fill_report, log)
        log(f"✔ 生成完成：{output_path}")
        return output_path

    log("① 解析需求，生成文档大纲 ...")
    plan = parse_request(command, memory.context_text(), llm,
                         reference_summary=refs["planner_summary"],
                         template_structure=tpl_struct_text,
                         log=log)
    if style_override:
        plan["style"] = style_override
    if standard_override:
        plan["standard"] = standard_override
    std = get_standard(plan.get("standard") or "general")
    std_preset = standard_to_preset(std)
    if std.get("name"):
        log(f"   📐 排版标准：{std['name']}")
    if template_sections:  # 兜底：结构 100% 按模板，AI 只补充各节写作说明
        desc_map = {s.get("heading", ""): s.get("description", "") for s in plan.get("sections", [])}
        plan["sections"] = [
            {"heading": _sec["heading"], "level": _sec["level"],
             "description": desc_map.get(_sec["heading"])
             or f"围绕「{_sec['heading']}」展开，内容要具体、贴合模板结构与用户要求。"}
            for _sec in template_sections
        ]
    log(f"   标题《{plan['title']}》｜ 风格：{plan['style']}｜ 类型：{plan['doc_type']}｜ 共 {len(plan['sections'])} 节")
    for sec in plan["sections"]:
        indent = "　" * (sec["level"] - 1)
        log(f"   {indent}· {sec['heading']}")

    # 大纲确认：不满意可直接取消，避免正文 token 白烧
    if confirm_plan is not None:
        if not confirm_plan(plan):
            log("已取消：大纲未确认，未生成文档。可在指令中补充章节/风格要求后重试。")
            return None

    log("② 逐节生成正文（自动衔接上下文，保持文风一致）...")
    markdown = generate_document(
        plan,
        llm,
        on_section=lambda i, total, heading: log(f"   [{i}/{total}] 撰写「{heading}」..."),
        reference_context=refs["writer_context"],
    )

    output_dir = Path(output_dir_override or config.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"{stamp}_{safe_filename(plan['title'])}.docx"

    # 兜底：确保每个章节标题都出现在正文中（模型偶发漏写标题时自动补齐）
    markdown = ensure_section_headings(markdown, plan.get("sections") or [])

    # 内容充实度补全：正文过短的章节一次批量扩写（默认开启，可用 WORDAGENT_AUTO_EXPAND=0 关闭）
    if getattr(config, "auto_expand", True):
        try:
            from .writer import expand_short_sections as _expand_short
            _md2 = _expand_short(markdown, plan, llm, log=log,
                                reference_context=refs["writer_context"])
            if _md2 and _md2 != markdown:
                markdown = _md2
        except Exception as _exc:  # noqa: BLE001
            log(f"   ⚠ 章节补全跳过（{_exc}）")

    log("③ 排版并保存 Word 文档 ...")
    render_markdown_to_docx(
        markdown,
        output_path,
        plan["title"],
        plan.get("style", "default"),
        plan.get("toc", False),
        plan.get("language", "zh-CN"),
        extra_images=refs["images"],
        template_preset=template_preset,
        standard_preset=std_preset,
    )
 
    # 程序化质量门禁：自动修复可安全修复的问题并复检（officecli Verification Gate，不消耗 token）
    try:
        from .verify import fix_document as _fix_doc
        log("④ 质量检查并自动修复 ...")
        _fix_doc(output_path, log=log)
    except Exception as _exc:  # noqa: BLE001
        log(f"   ⚠ 质量检查跳过（{_exc}）")

    memory.add_result(
        {
            "title": plan["title"],
            "file": str(output_path),
            "style": plan["style"],
            "doc_type": plan.get("doc_type", "general"),
            "sections": len(plan["sections"]),
        }
    )
    if hasattr(llm, "usage_text"):
        log(f"   ⚙ {llm.usage_text()}")
    if not refs["names"]:
        log("⚠ 本次未提供参考数据文件，文档中的具体数值为 AI 生成的示例，正式使用前请务必核对/替换。")
    log(f"✔ 生成完成：{output_path}")
    return output_path
