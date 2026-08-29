"""需求解析：把用户指令拆解为结构化文档大纲。"""
from __future__ import annotations

import re
from typing import Any

from .llm import LLMClient

PLANNER_PROMPT = """你是一名专业的文档规划专家。请把用户的指令解析为一份结构化的 Word 文档大纲。

要求：
1. 识别文档类型、目标读者、语气风格，并给出合适的标题。
2. 拆解出 3~10 个章节，重要内容用 level=1，细分内容用 level=2。
3. style 取值只能是以下之一：business（商务/汇报）、report（报告）、academic（学术）、creative（创意文案）、default（通用）。
4. doc_type 取值只能是以下之一：experiment_report（实验/测试报告）、weekly_report（周报/月报）、meeting_notes（会议纪要）、prd（需求/方案文档）、analysis（分析报告）、general（通用）。
5. 用户指令中出现"实验报告/实验/实验数据/实验结果/测试报告"等关键词时，doc_type 必须为 experiment_report。
   章节组织必须贴合实验所属学科方向（用户指令明确提到方向时严格按指令，禁止生搬硬套物理实验结构）：
   - 计算机/软件/数据库/编程/网络类（指令含 数据库/SQL/MySQL/Oracle/程序/代码/Java/Python/C语言/软件/接口/网络 等）：
     「实验目的 → 实验原理 → 实验环境（软件与工具）→ 实验步骤 → 数据记录（表结构/运行结果用表格）→ 结果分析 → 问题与调试记录 → 实验结论」，
     章节名称要用对应学科的措辞，例如"数据库实验报告"应叫「实验环境（数据库版本与工具）」「数据记录（表结构与查询结果）」而不是"器材"。
   - 物理/化学/电子/生物等硬件实验类：按「实验目的 → 实验原理 → 实验环境与器材 → 实验步骤 → 数据记录 → 结果分析 → 误差分析 → 实验结论」。
   - 无法判断时按通用实验报告组织；"数据记录"与"结果分析"必须包含用表格呈现的数据，描述里写明需要几列、记录哪些数据。
6. standard 字段按文档适用的排版标准选择，取值只能是：gb_9704（党政机关公文，指令含公文/国标公文/GB/T 9704/红头文件时必选）、paper（学术论文/参考文献，含论文/学术/GB/T 7714时必选）、experiment（实验/测试报告）、meeting（会议纪要）、weekly（周报/月报）、business（商务/方案/汇报）、general（通用）。用户明确提到标准时优先按用户要求。
7. 严格只输出 JSON，不要输出任何解释文字。JSON 格式如下：
8. 如果【模板章节结构】非空：必须严格沿用给定结构的全部标题与层级、保持顺序，不得新增、删除或改名标题；每个标题仍需给出 description 说明这一节要写什么。
{{
  "title": "文档标题",
  "style": "business",
  "doc_type": "experiment_report",
  "standard": "experiment",
  "language": "zh-CN",
  "toc": true,
  "experiment_field": "database",
  "sections": [
    {{"heading": "章节标题", "level": 1, "description": "这一章要写什么的简短说明"}}
  ]
}}
"""

# 兜底识别实验/测试类文档的关键词（LLM 漏给 doc_type 时使用）
EXPERIMENT_KEYWORDS = ("实验报告", "实验数据", "实验结果", "实验分析", "测试报告", "实验")

# 实验报告学科方向识别：决定章节结构用「软件/数据类」还是「硬件/物理类」
EXPERIMENT_FIELD_RULES = [
    ("database", ("数据库", "sql", "mysql", "oracle", "sqlserver", "sql server", "postgres", "sqlite", "表结构", "查询", "触发器", "存储过程", "数据库设计")),
    ("software", ("软件", "程序", "编程", "代码", "java", "python", "c语言", "c++", "开发", "接口", "系统设计", "算法", "数据结构", "前端", "后端")),
    ("network", ("网络", "通信", "tcp", "http", "协议", "路由", "组网")),
    ("hardware", ("电路", "电子", "单片机", "嵌入式", "数字电路", "模拟电路", "物理", "光学", "力学", "电磁", "化学", "溶液", "生物", "细胞", "基因")),
]


def detect_experiment_field(command: str) -> str:
    """按指令关键词推断实验学科方向：database/software/network/hardware/general。"""
    low = command.lower()
    for field, words in EXPERIMENT_FIELD_RULES:
        for w in words:
            if w.lower() in low:
                return field
    return "general"


def parse_request(
    command: str,
    memory_context: str,
    llm: LLMClient,
    reference_summary: str = "",
    template_structure: str = "",
    log=None,
) -> dict[str, Any]:
    """解析用户指令，返回 {title, style, doc_type, language, toc, sections}。"""
    ref_block = ""
    if reference_summary:
        ref_block = (
            "\n\n【用户附加的参考文件（文件名 + 内容摘要，规划章节时应确保有章节引用/呈现这些材料）】\n"
            f"{reference_summary}"
        )
    tpl_block = ""
    if template_structure:
        tpl_block = (
            "\n\n【模板文档的章节结构（必须严格沿用，不得新增/删除/改名标题）】\n"
            f"{template_structure}"
        )
    messages = [
        {"role": "system", "content": PLANNER_PROMPT},
        {
            "role": "user",
            "content": (
                f"【历史记忆（可参考风格、格式、主题，但不要照抄）】\n{memory_context}\n"
                f"{ref_block}{tpl_block}\n\n"
                f"【本次指令】\n{command}"
            ),
        },
    ]
    try:
        plan = llm.chat_json(messages, temperature=0.3, max_tokens=4096)
        plan = _normalize_plan(plan, command)
    except Exception:
        if log is not None:
            log("   ⚠ 大纲解析失败，已使用默认大纲结构（可重新描述指令重试）。")
        plan = _fallback_plan(command)
    return plan


def _normalize_plan(plan: dict, command: str) -> dict:
    sections = []
    for sec in plan.get("sections") or []:
        if not isinstance(sec, dict):
            continue
        heading = str(sec.get("heading", "")).strip()
        if not heading:
            continue
        level = int(sec.get("level", 1))
        level = 1 if level < 1 else (2 if level > 2 else level)
        sections.append(
            {
                "heading": heading,
                "level": level,
                "description": str(sec.get("description", "")).strip(),
            }
        )
    if not sections:
        raise ValueError("大纲为空")

    doc_type = str(plan.get("doc_type") or "").strip().lower()
    if doc_type not in ("experiment_report", "weekly_report", "meeting_notes", "prd", "analysis", "general"):
        if any(k in command for k in EXPERIMENT_KEYWORDS):
            doc_type = "experiment_report"
        else:
            doc_type = "general"
    standard = str(plan.get("standard") or "").strip().lower()
    allowed = ("gb_9704", "paper", "experiment", "meeting", "weekly", "business", "general")
    if standard not in allowed:
        standard = _guess_standard(command, doc_type)

    # 实验学科方向：本地关键词判定优先（LLM 可能被误导，如把数据库实验判成物理实验）。
    # 本地明确识别到方向时强制覆盖 LLM 值；本地无法判断时才信任 LLM 的值。
    local_field = detect_experiment_field(command)
    llm_field = str(plan.get("experiment_field") or "").strip().lower()
    if local_field != "general":
        experiment_field = local_field
    elif llm_field in ("database", "software", "network", "hardware"):
        experiment_field = llm_field
    else:
        experiment_field = local_field

    return {
        "title": str(plan.get("title") or command.strip() or "未命名文档")[:80],
        "style": str(plan.get("style") or "default").strip().lower(),
        "doc_type": doc_type,
        "standard": standard,
        "language": str(plan.get("language") or "zh-CN").strip(),
        "toc": bool(plan.get("toc", True)),
        "experiment_field": experiment_field,
        "sections": sections,
    }


def _guess_standard(command: str, doc_type: str) -> str:
    """关键词 + 文档类型推断标准（LLM 未给出或给错时兜底）。"""
    from .standards import guess_standard as gs
    return gs(command, doc_type)


def _fallback_plan(command: str) -> dict[str, Any]:
    """LLM 解析失败时的保底方案：标题 + 通用三段式大纲。"""
    title = re.sub(r"\s+", " ", command.strip())[:50] or "未命名文档"
    if any(k in command for k in EXPERIMENT_KEYWORDS):
        # 实验/测试类文档使用专用保底大纲（按学科方向区分措辞）
        field = detect_experiment_field(command)
        if field in ("database", "software", "network"):
            return {
                "title": title,
                "style": "report",
                "doc_type": "experiment_report",
                "standard": "experiment",
                "language": "zh-CN",
                "toc": True,
                "experiment_field": field,
                "sections": [
                    {"heading": "实验目的", "level": 1, "description": "本实验要完成的功能或验证的目标"},
                    {"heading": "实验原理", "level": 1, "description": "所依据的理论、技术或方法说明"},
                    {"heading": "实验环境与工具", "level": 1, "description": "操作系统、软件/数据库版本、开发工具等环境信息"},
                    {"heading": "实验步骤", "level": 1, "description": "按编号列出操作步骤、SQL/代码要点与注意事项"},
                    {"heading": "数据记录", "level": 1, "description": "用表格记录表结构、查询/运行结果等数据"},
                    {"heading": "结果分析", "level": 1, "description": "对运行结果进行分析、与预期对比"},
                    {"heading": "问题与调试记录", "level": 1, "description": "实验中遇到的问题、原因与解决办法"},
                    {"heading": "实验结论", "level": 1, "description": "结论、收获与后续改进方向"},
                ],
            }
        return {
            "title": title,
            "style": "report",
            "doc_type": "experiment_report",
            "standard": "experiment",
            "language": "zh-CN",
            "toc": True,
            "experiment_field": field,
            "sections": [
                {"heading": "实验目的", "level": 1, "description": "本实验要验证或测量的目标与预期结果"},
                {"heading": "实验原理", "level": 1, "description": "所依据的理论公式、定律与推导过程"},
                {"heading": "实验环境与器材", "level": 1, "description": "设备型号、软件版本、材料清单"},
                {"heading": "实验步骤", "level": 1, "description": "按编号列出操作步骤与注意事项"},
                {"heading": "数据记录", "level": 1, "description": "用表格记录原始数据与观测现象"},
                {"heading": "结果分析", "level": 1, "description": "数据处理、图表化结果、定量分析"},
                {"heading": "误差分析", "level": 1, "description": "系统误差与随机误差来源及改进方法"},
                {"heading": "实验结论", "level": 1, "description": "结论、与理论值的对比及后续展望"},
            ],
        }
    return {
        "title": title,
        "style": "default",
        "doc_type": "general",
        "standard": _guess_standard(command, "general"),
        "language": "zh-CN",
        "toc": False,
        "sections": [
            {"heading": "概述", "level": 1, "description": "主题背景与核心内容概述"},
            {"heading": "正文", "level": 1, "description": "围绕主题展开详细论述"},
            {"heading": "总结", "level": 1, "description": "要点归纳与后续建议"},
        ],
    }
