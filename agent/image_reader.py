"""图片 OCR 读取：模型非多模态时，用本地 RapidOCR 提取图片文字作为写作参考。"""
from __future__ import annotations

from pathlib import Path
from typing import Callable, Optional

IMAGE_MAX_SIDE = 2000   # 超长边先等比缩放，加速 OCR
OCR_EXCERPT_LIMIT = 4000  # 单张图片最多注入的识别文字数
_engine = None


def _get_engine():
    global _engine
    if _engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _engine = RapidOCR()
    return _engine


def ocr_image(path: str | Path, log: Optional[Callable[[str], None]] = None) -> str:
    """对图片执行 OCR，返回识别文字；环境不支持或失败时返回空串（不阻断流程）。"""
    try:
        from PIL import Image
        import numpy as np
    except Exception as exc:
        if log:
            log(f"⚠ OCR 依赖缺失（{exc}），图片仅嵌入不提取文字")
        return ""
    try:
        img = Image.open(path).convert("RGB")
        w, h = img.size
        if max(w, h) > IMAGE_MAX_SIDE:
            scale = IMAGE_MAX_SIDE / max(w, h)
            img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
        arr = np.array(img)
        result, _ = _get_engine()(arr)
        if not result:
            return ""
        lines, seen = [], set()
        for item in result:
            text = (item[1] or "").strip()
            if text and text not in seen:
                seen.add(text)
                lines.append(text)
        return "\n".join(lines)[:OCR_EXCERPT_LIMIT]
    except Exception as exc:  # noqa: BLE001
        if log:
            log(f"⚠ 图片 OCR 失败，已跳过文字提取：{exc}")
        return ""
