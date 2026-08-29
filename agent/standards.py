"""文档标准库：按国家标准/行业规范精确控制 Word 排版。

优先级：用户提供的模板 docx > 指定/识别出的标准 > 通用风格预设。

标准字段说明：
- body_font / body_size       正文中文字体与字号（磅）
- body_font_latin             正文西文字体（缺省同 body_font）
- line_spacing                行距：float=倍数（如 1.5），int=固定磅值（如 28）
- first_line_indent_chars     首行缩进字符数（0=不缩进）
- justify                     正文是否两端对齐
- heading_fonts               各级标题字体 {1: 黑体, 2: 楷体_GB2312, ...}
- heading_sizes               各级标题字号（磅）
- heading_bold                各级标题是否加粗
- heading_color               标题颜色（十六进制，如 000000）
- title_font / title_size     封面标题字体/字号（封面关闭时忽略）
- title_page                  是否生成 AI 封面页（公文等关闭，直接正文）
- toc                         是否生成目录
- page                        页边距 {top_cm,bottom_cm,left_cm,right_cm}
"""
from __future__ import annotations

import json
from pathlib import Path

STANDARDS_DIR = Path(__file__).resolve().parent.parent / "standards"

BUILTIN_STANDARDS: dict[str, dict] = {
    # 党政机关公文格式 GB/T 9704-2012
    "gb_9704": {
        "name": "国标公文（GB/T 9704）",
        "body_font": "仿宋_GB2312",
        "body_font_latin": "Times New Roman",
        "body_size": 16.0,           # 三号
        "line_spacing": 28,          # 固定 28 磅
        "line_spacing_rule": "exact",
        "first_line_indent_chars": 2,
        "justify": True,
        "heading_fonts": {1: "黑体", 2: "楷体_GB2312", 3: "仿宋_GB2312", 4: "仿宋_GB2312"},
        "heading_sizes": {1: 16.0, 2: 16.0, 3: 16.0, 4: 16.0},
        "heading_bold": {1: False, 2: False, 3: True, 4: True},
        "heading_color": "000000",
        "title_font": "方正小标宋简体",
        "title_size": 22.0,          # 二号
        "title_bold": False,
        "title_page": False,
        "toc": False,
        "code_font": "Consolas",
        "page": {"top_cm": 3.7, "bottom_cm": 3.5, "left_cm": 2.8, "right_cm": 2.6},
    },
    # 学术论文（正文部分参照 GB/T 7714-2015 引用规范）
    "paper": {
        "name": "学术论文（GB/T 7714）",
        "body_font": "宋体",
        "body_font_latin": "Times New Roman",
        "body_size": 12.0,           # 小四
        "line_spacing": 1.5,
        "line_spacing_rule": "multiple",
        "first_line_indent_chars": 2,
        "justify": True,
        "heading_fonts": {1: "黑体", 2: "黑体", 3: "黑体", 4: "黑体"},
        "heading_sizes": {1: 16.0, 2: 14.0, 3: 12.5, 4: 12.0},
        "heading_bold": {1: False, 2: False, 3: False, 4: False},
        "heading_color": "000000",
        "title_font": "黑体",
        "title_size": 22.0,
        "title_bold": True,
        "title_page": True,
        "toc": True,
        "code_font": "Consolas",
        "page": {"top_cm": 2.54, "bottom_cm": 2.54, "left_cm": 3.17, "right_cm": 3.17},
    },
    # 实验报告
    "experiment": {
        "name": "实验报告",
        "body_font": "宋体",
        "body_font_latin": "Times New Roman",
        "body_size": 12.0,
        "line_spacing": 1.5,
        "line_spacing_rule": "multiple",
        "first_line_indent_chars": 2,
        "justify": True,
        "heading_fonts": {1: "黑体", 2: "黑体", 3: "黑体", 4: "黑体"},
        "heading_sizes": {1: 16.0, 2: 14.0, 3: 12.5, 4: 12.0},
        "heading_bold": {1: False, 2: False, 3: False, 4: False},
        "heading_color": "000000",
        "title_font": "黑体",
        "title_size": 22.0,
        "title_bold": True,
        "title_page": True,
        "toc": False,
        "code_font": "Consolas",
        "page": {"top_cm": 2.5, "bottom_cm": 2.5, "left_cm": 2.8, "right_cm": 2.8},
    },
    # 商务汇报文档
    "business": {
        "name": "商务文档",
        "body_font": "微软雅黑",
        "body_size": 11.0,
        "line_spacing": 1.4,
        "line_spacing_rule": "multiple",
        "first_line_indent_chars": 0,
        "justify": True,
        "heading_fonts": {1: "微软雅黑", 2: "微软雅黑", 3: "微软雅黑", 4: "微软雅黑"},
        "heading_sizes": {1: 18.0, 2: 15.0, 3: 13.0, 4: 12.0},
        "heading_bold": {1: True, 2: True, 3: True, 4: True},
        "heading_color": "1F4E79",
        "title_font": "微软雅黑",
        "title_size": 26.0,
        "title_bold": True,
        "title_page": True,
        "toc": True,
        "code_font": "Consolas",
        "page": {"top_cm": 2.5, "bottom_cm": 2.5, "left_cm": 2.8, "right_cm": 2.8},
    },
    # 会议纪要
    "meeting": {
        "name": "会议纪要",
        "body_font": "宋体",
        "body_font_latin": "Times New Roman",
        "body_size": 12.0,
        "line_spacing": 1.4,
        "line_spacing_rule": "multiple",
        "first_line_indent_chars": 2,
        "justify": True,
        "heading_fonts": {1: "黑体", 2: "黑体", 3: "黑体", 4: "黑体"},
        "heading_sizes": {1: 16.0, 2: 14.0, 3: 12.5, 4: 12.0},
        "heading_bold": {1: False, 2: False, 3: False, 4: False},
        "heading_color": "000000",
        "title_font": "黑体",
        "title_size": 20.0,
        "title_bold": True,
        "title_page": True,
        "toc": False,
        "code_font": "Consolas",
        "page": {"top_cm": 2.5, "bottom_cm": 2.5, "left_cm": 2.8, "right_cm": 2.8},
    },
    # 周报/月报
    "weekly": {
        "name": "周报/月报",
        "body_font": "微软雅黑",
        "body_size": 11.0,
        "line_spacing": 1.4,
        "line_spacing_rule": "multiple",
        "first_line_indent_chars": 0,
        "justify": True,
        "heading_fonts": {1: "微软雅黑", 2: "微软雅黑", 3: "微软雅黑", 4: "微软雅黑"},
        "heading_sizes": {1: 16.0, 2: 14.0, 3: 12.5, 4: 12.0},
        "heading_bold": {1: True, 2: True, 3: True, 4: True},
        "heading_color": "2E74B5",
        "title_font": "微软雅黑",
        "title_size": 22.0,
        "title_bold": True,
        "title_page": True,
        "toc": False,
        "code_font": "Consolas",
        "page": {"top_cm": 2.5, "bottom_cm": 2.5, "left_cm": 2.8, "right_cm": 2.8},
    },
    # 通用默认
    "general": {
        "name": "通用文档",
        "body_font": "微软雅黑",
        "body_size": 12.0,
        "line_spacing": 1.5,
        "line_spacing_rule": "multiple",
        "first_line_indent_chars": 2,
        "justify": True,
        "heading_fonts": {1: "微软雅黑", 2: "微软雅黑", 3: "微软雅黑", 4: "微软雅黑"},
        "heading_sizes": {1: 16.0, 2: 14.0, 3: 12.5, 4: 12.0},
        "heading_bold": {1: True, 2: True, 3: True, 4: True},
        "heading_color": "1F3864",
        "title_font": "微软雅黑",
        "title_size": 26.0,
        "title_bold": True,
        "title_page": True,
        "toc": True,
        "code_font": "Consolas",
        "page": {"top_cm": 2.5, "bottom_cm": 2.5, "left_cm": 2.8, "right_cm": 2.8},
    },
}

# 指令关键词 -> 标准 id（LLM 识别失败时的兜底）
STANDARD_KEYWORDS = [
    ("gb_9704", ("公文", "红头", "GB/T 9704", "国标公文", "党政机关")),
    ("paper", ("论文", "学术", "参考文献", "GB/T 7714", "毕业论文", "学报")),
    ("experiment", ("实验报告", "实验", "试验报告", "测试报告")),
    ("meeting", ("会议纪要", "会议记录", "纪要")),
    ("weekly", ("周报", "月报", "周总结", "工作周报")),
    ("business", ("商务", "商业计划", "融资", "方案", "汇报", "提案")),
]


def guess_standard(command: str, doc_type: str = "") -> str:
    """按关键词 + 文档类型推断标准 id。"""
    for std_id, keywords in STANDARD_KEYWORDS:
        for kw in keywords:
            if kw.lower() in command.lower():
                return std_id
    mapping = {
        "experiment_report": "experiment",
        "meeting_notes": "meeting",
        "weekly_report": "weekly",
        "prd": "business",
        "analysis": "business",
    }
    return mapping.get(doc_type, "general")


def load_custom_standards() -> dict[str, dict]:
    """加载 standards/*.json 自定义标准（id 用文件名）。"""
    out: dict[str, dict] = {}
    if not STANDARDS_DIR.is_dir():
        return out
    for path in sorted(STANDARDS_DIR.glob("*.json")):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict) and data.get("name"):
                data.setdefault("_custom", True)
                out[path.stem] = data
        except (json.JSONDecodeError, OSError):
            continue
    return out


def get_standard(std_id: str) -> dict:
    """按 id 取标准；找不到回退 general。"""
    if not std_id:
        return BUILTIN_STANDARDS["general"]
    custom = load_custom_standards()
    if std_id in custom:
        return custom[std_id]
    return BUILTIN_STANDARDS.get(std_id, BUILTIN_STANDARDS["general"])


def standard_to_preset(std: dict) -> dict:
    """把标准转成 renderer 可用的 preset。"""
    preset: dict = {}
    for key in ("body_font", "body_font_latin", "body_size", "heading_fonts",
                "heading_sizes", "heading_bold", "heading_color", "page",
                "line_spacing", "line_spacing_rule", "first_line_indent_chars",
                "justify", "title_font", "title_size", "title_bold",
                "title_page", "toc", "code_font"):
        if std.get(key) is not None:
            preset[key] = std[key]
    return preset


def list_standards() -> list[dict]:
    """内置 + 自定义标准清单（GUI 下拉用）。"""
    items = [{"id": k, "name": v["name"]} for k, v in BUILTIN_STANDARDS.items()]
    for k, v in load_custom_standards().items():
        items.append({"id": k, "name": v["name"] + "（自定义）"})
    return items