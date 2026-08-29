"""WordAgent：AI 文档生成 Agent（PC 端软件核心包）。

- 生成模式：需求解析 -> 大纲 -> 分节撰写 -> docx 排版（planner/writer/renderer）
- 编辑模式：docx 提取 -> 修改计划 -> 三级匹配落笔 -> AI 复核 -> 备份另存（extractor/editor）
- 质量门禁：verify.py 程序化校验（标题层级/空段/残留/字体/表格/内容完整度），生成与编辑后自动执行
- 上下文记忆：memory.py ｜ 编排：pipeline.py
- 控制台入口：python main.py ｜ 桌面版入口：python gui.py
"""
from .config import Config
from .editor import (
    AmbiguousTargetError,
    EditValidationError,
    edit_document,
    finalize_edit,
    prepare_edit,
)
from .llm import LLMClient, LLMError
from .memory import Memory
from .pipeline import run_pipeline, safe_filename
from .templater import fill_template
from .verify import QualityReport, fix_document, verify_document

__version__ = "1.7.5"
__all__ = [
    "Config", "LLMClient", "LLMError", "Memory",
    "run_pipeline", "safe_filename",
    "prepare_edit", "finalize_edit", "edit_document",
    "fill_template",
    "verify_document", "fix_document", "QualityReport",
    "EditValidationError", "AmbiguousTargetError",
    "__version__",
]
