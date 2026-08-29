"""模糊模板识别：图片/PDF/扫描件模板 -> 自动增强 -> OCR -> 结构推断。

适用场景：用户传入的模板是模糊的扫描件、照片或 PDF（看不清、不是 docx）。
处理流程：
1. 图片/PDF 渲染为图像（PDF 用 PyMuPDF，200 DPI）
2. 图像增强：放大 + 灰度 + CLAHE 对比度 + 锐化（模糊模板识别率大幅提升）
3. OCR（rapidocr）：按坐标排序还原版面
4. 结构推断：用居中/序号/关键词启发式识别章节标题
"""
from __future__ import annotations

from pathlib import Path
from typing import Callable, Optional

from PIL import Image, ImageEnhance, ImageFilter, ImageOps

from .image_reader import _get_engine

ENHANCE_MIN_SIDE = 1600   # 短边小于该值先放大
OCR_MIN_LINES = 6         # 原图 OCR 少于该行数判定为模糊/低清，走增强


def enhance_image(img: Image.Image) -> Image.Image:
    """增强模糊模板图像：放大 -> 灰度 -> 对比度(CLAHE) -> 锐化。"""
    w, h = img.size
    if max(w, h) < ENHANCE_MIN_SIDE:
        scale = ENHANCE_MIN_SIDE / max(w, h)
        img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
    img = ImageOps.grayscale(img)
    # CLAHE：用 PIL 的对比度拉伸近似（ImageOps.autocontrast + 增强）
    img = ImageOps.autocontrast(img, cutoff=1)
    img = ImageEnhance.Contrast(img).enhance(1.6)
    img = ImageEnhance.Sharpness(img).enhance(2.0)
    img = img.filter(ImageFilter.UnsharpMask(radius=2, percent=150, threshold=2))
    return img


def pdf_to_image(path: Path) -> Optional[Image.Image]:
    """用 PyMuPDF 渲染 PDF 首页为高分辨率图像（模板结构一般在前几页）。"""
    try:
        import pymupdf  # PyMuPDF 1.28+ 推荐入口
    except ImportError:
        import fitz as pymupdf  # 旧版兼容
    try:
        doc = pymupdf.open(str(path))
        if doc.page_count < 1:
            return None
        page = doc[0]
        pix = page.get_pixmap(matrix=pymupdf.Matrix(200 / 72, 200 / 72), alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        doc.close()
        return img
    except Exception:
        return None


def _line_features(box, img_width: int) -> dict:
    """从 OCR 边框计算版面特征（居中程度、纵向位置）。"""
    x1, y1 = box[0][0], box[0][1]
    x2, y2 = box[2][0], box[2][1]
    cx = (x1 + x2) / 2.0
    return {
        "x1": x1, "y1": y1, "x2": x2, "y2": y2,
        "center": cx / max(1, img_width),
        "centered": 0.32 < cx / max(1, img_width) < 0.68,
        "width": x2 - x1,
    }


def _looks_like_heading(text: str, feat: dict) -> bool:
    """标题启发式：居中 / 序号开头 / 章节关键词。"""
    t = text.strip()
    if not t or len(t) > 32:
        return False
    if feat.get("centered") and len(t) <= 24:
        return True
    if t[0] in "一二三四五六七八九十" and "、" in t[:4]:
        return True
    import re
    if re.match(r"^第[一二三四五六七八九十百]+[章节篇部分][\s、．.：:]", t):
        return True
    if re.match(r"^实验[一二三四五六七八九十0-9]+[、：:.．\s]", t):
        return True
    if re.match(r"^[1-9][\.、)]\s*\S", t):
        return True
    return False


def _heading_level(text: str) -> int:
    t = text.strip()
    import re
    if re.match(r"^[1-9][\.、)]", t) and not (t[0] in "一二三四五六七八九十" and "、" in t[:4]):
        return 2
    return 1


def _ocr(img: Image.Image) -> list[dict]:
    """OCR 返回 [{text, box}]，按 y 坐标排序。"""
    import numpy as np
    result, _ = _get_engine()(np.array(img))
    lines = []
    for item in result or []:
        text = (item[1] or "").strip()
        if not text:
            continue
        lines.append({"text": text, "box": item[0]})
    lines.sort(key=lambda it: (it["box"][0][1], it["box"][0][0]))
    return lines


def scan_template_image(path: Path, log: Optional[Callable[[str], None]] = None) -> dict:
    """识别图片模板，返回 {sections, lines, enhanced}。"""
    img = Image.open(path).convert("RGB")
    w0, h0 = img.size

    lines = _ocr(img)
    used_enhance = False
    if len(lines) < OCR_MIN_LINES:
        enhanced = enhance_image(img)
        lines2 = _ocr(enhanced)
        if len(lines2) > len(lines):
            lines = lines2
            used_enhance = True
            img = enhanced
    if log and used_enhance:
        log(f"   🔍 模板较模糊，已自动增强画质后识别（{len(lines)} 行文字）")
    elif log:
        log(f"   🔍 模板识别完成（{len(lines)} 行文字）")

    img_width = img.size[0]
    sections = []
    for it in lines:
        feat = _line_features(it["box"], img_width)
        if _looks_like_heading(it["text"], feat):
            sections.append({"heading": it["text"].strip(), "level": _heading_level(it["text"])})
    # 过滤文档标题（如居中的「实 验 报 告」），避免当成章节
    _TITLE_KEYWORDS = ("报告", "说明书", "设计", "方案", "课程作业", "论文")
    filtered = []
    for idx, s in enumerate(sections):
        t = s["heading"]
        is_title = (
            idx == 0
            and len(t) <= 14
            and any(k in t for k in _TITLE_KEYWORDS)
        )
        if not is_title:
            filtered.append(s)
    return {
        "sections": filtered,
        "lines": [it["text"] for it in lines],
        "enhanced": used_enhance,
    }


def scan_template_file(path: Path, log: Optional[Callable[[str], None]] = None) -> Optional[dict]:
    """按文件类型识别模板（图片 / PDF）。docx 请走 template_reader。"""
    suffix = path.suffix.lower()
    if suffix in (".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"):
        try:
            return scan_template_image(path, log=log)
        except Exception as exc:  # noqa: BLE001
            if log:
                log(f"⚠ 图片模板识别失败：{exc}")
            return None
    if suffix == ".pdf":
        img = pdf_to_image(path)
        if img is None:
            if log:
                log("⚠ PDF 渲染失败，无法识别模板结构。")
            return None
        # 用临时图片走统一识别流程
        try:
            from PIL import Image
            import tempfile, os
            fd, tmp = tempfile.mkstemp(suffix=".png")
            os.close(fd)
            img.save(tmp)
            try:
                return scan_template_image(Path(tmp), log=log)
            finally:
                os.unlink(tmp)
        except Exception as exc:  # noqa: BLE001
            if log:
                log(f"⚠ PDF 模板识别失败：{exc}")
            return None
    return None