"""参考文件收集：把用户附加的 图片/文本/csv/docx 整理为生成上下文。

生成模式下，用户可选择若干文件作为写作参考：
- 图片：本地 OCR 提取文字作为写作参考，并提示模型用 `![说明](绝对路径)` 嵌入正文
- 文本类（txt/md/csv/log/json）：读取前 N 字符作为摘要注入提示词
- docx：用 extractor 提取文本与表格
- 其他类型：仅记录文件名
"""
from __future__ import annotations

from pathlib import Path
from typing import Callable, Optional

from docx import Document

from .extractor import docx_to_markdown
from .image_reader import ocr_image

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff"}
TEXT_SUFFIXES = {".txt", ".md", ".markdown", ".csv", ".log", ".json", ".xml", ".py", ".ini", ".conf"}
MAX_FILE_EXCERPT = 2000  # 单个文本文件注入提示词的最大字符数
MAX_REF_FILES = 8        # 一次最多处理的参考文件数


def _read_text(path: Path, limit: int = MAX_FILE_EXCERPT) -> str:
    """带编码兜底读取文本文件前 limit 字符。"""
    for encoding in ("utf-8", "utf-8-sig", "gbk", "gb18030", "latin-1"):
        try:
            text = path.read_text(encoding=encoding)
            break
        except (UnicodeDecodeError, OSError):
            continue
    else:
        return ""
    return text.strip().replace("\r\n", "\n")[:limit]


def _csv_to_text(path: Path) -> str:
    import csv
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as fh:
            rows = list(csv.reader(fh))
    except (UnicodeDecodeError, OSError, csv.Error):
        return _read_text(path)
    if not rows:
        return ""
    lines = ["| " + " | ".join(str(c).strip() for c in row) + " |" for row in rows[:40]]
    return "\n".join(lines)[:MAX_FILE_EXCERPT]


def _docx_to_text(path: Path) -> str:
    try:
        md, _ = docx_to_markdown(Document(str(path)))
        return md.strip().replace("\r\n", "\n")[:MAX_FILE_EXCERPT]
    except Exception:
        return ""


def _describe_one(path: Path, log: Optional[Callable[[str], None]] = None) -> dict:
    """读取单个文件，返回 {name, kind, excerpt}。"""
    suffix = path.suffix.lower()
    if suffix in IMAGE_SUFFIXES:
        abs_path = str(path.resolve())
        ocr_text = ocr_image(path, log=log)
        if ocr_text:
            excerpt = "[IMAGE_FILE %s ABS_PATH=%s]\n[图片OCR识别文字]\n%s" % (path.name, abs_path, ocr_text)
        else:
            excerpt = "[IMAGE_FILE %s ABS_PATH=%s]" % (path.name, abs_path)
        return {"name": path.name, "kind": "image",
                "excerpt": excerpt}
    if suffix in TEXT_SUFFIXES:
        excerpt = _csv_to_text(path) if suffix == ".csv" else _read_text(path)
        return {"name": path.name, "kind": "text", "excerpt": excerpt or f"[文件内容为空]"}
    if suffix == ".docx":
        excerpt = _docx_to_text(path)
        return {"name": path.name, "kind": "docx", "excerpt": excerpt or f"[docx 提取为空]"}
    return {"name": path.name, "kind": "other", "excerpt": f"[暂不支持解析该类型，仅作参考：{path.name}]"}


def collect_reference_context(
    paths: Optional[list],
    log: Optional[Callable[[str], None]] = None,
) -> dict:
    """收集参考文件，返回：
    {
      "names": [...],            # 文件名列表
      "images": [...],           # 图片绝对路径（供渲染器嵌入）
      "planner_summary": str,    # 供规划阶段使用的简短摘要
      "writer_context": str,     # 供写作阶段使用的完整摘要
    }
    """
    result = {"names": [], "images": [], "planner_summary": "", "writer_context": ""}
    items: list[dict] = []
    for raw in paths or []:
        path = Path(raw).expanduser()
        if not path.is_file():
            if log:
                log(f"⚠ 参考文件不存在，已跳过：{path}")
            continue
        if len(items) >= MAX_REF_FILES:
            if log:
                log(f"⚠ 参考文件超过 {MAX_REF_FILES} 个，其余已忽略")
            break
        try:
            item = _describe_one(path, log=log)
            items.append(item)
            if item["kind"] == "image":
                result["images"].append(str(path.resolve()))
            if log:
                log(f"   📎 参考文件：{path.name}（{item['kind']}）")
        except Exception as exc:  # noqa: BLE001
            if log:
                log(f"⚠ 读取参考文件失败，已跳过：{path.name}（{exc}）")

    if not items:
        return result

    result["names"] = [it["name"] for it in items]
    planner_parts = []
    for it in items:
        excerpt = it["excerpt"].replace("\n", " ")[:160]
        planner_parts.append(f"- {it['name']}：{excerpt}")
    result["planner_summary"] = "\n".join(planner_parts)

    writer_parts = []
    for it in items:
        writer_parts.append(f"### 文件《{it['name']}》\n{it['excerpt']}")
    result["writer_context"] = "\n\n".join(writer_parts)[:MAX_REF_FILES * (MAX_FILE_EXCERPT + 300)]
    return result
