"""WordAgent 配置加载。"""
from __future__ import annotations

import os
import sys
from dataclasses import dataclass
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:  # 未安装 python-dotenv 时静默跳过
    pass

DEFAULT_BASE_URL = "https://api.deepseek.com"
DEFAULT_MODEL = "deepseek-v4-flash"
DEFAULT_OUTPUT_DIR = "output"


def _user_data_dir() -> "Path | None":
    """安装版（PyInstaller 冻结）运行时，把输出与记忆放到「我的文档/WordAgent」，
    避免卸载软件时丢失用户的重要文件；开发版保持相对路径（cwd/output）不变。"""
    try:
        if not getattr(sys, "frozen", False):
            return None
        home = Path.home()
        docs = home / "Documents"
        return (docs if docs.exists() else home) / "WordAgent"
    except Exception:
        return None


def _resolve_data_path(raw: str) -> Path:
    """相对路径在安装版下解析到用户数据目录，绝对路径原样保留。"""
    p = Path(raw.strip() or ".")
    base = _user_data_dir()
    if base is not None and not p.is_absolute():
        return base / p
    return p


@dataclass
class Config:
    api_key: str = ""
    base_url: str = DEFAULT_BASE_URL
    model: str = DEFAULT_MODEL
    output_dir: Path = Path(DEFAULT_OUTPUT_DIR)
    memory_file: Path = Path("memory.json")
    temperature: float = 0.7
    max_tokens: int = 8192
    request_timeout: int = 180
    max_retries: int = 3
    concurrency: int = 3  # 分节生成时的并发请求数（不要太大，避免限流）

    @classmethod
    def from_env(cls) -> "Config":
        return cls(
            api_key=os.getenv("DEEPSEEK_API_KEY", "").strip(),
            base_url=os.getenv("DEEPSEEK_BASE_URL", DEFAULT_BASE_URL).strip().rstrip("/"),
            model=os.getenv("DEEPSEEK_MODEL", DEFAULT_MODEL).strip(),
            output_dir=_resolve_data_path(os.getenv("OUTPUT_DIR", DEFAULT_OUTPUT_DIR).strip()),
            memory_file=_resolve_data_path(os.getenv("MEMORY_FILE", "memory.json").strip()),
            temperature=float(os.getenv("DEEPSEEK_TEMPERATURE", "0.7")),
            max_tokens=int(os.getenv("DEEPSEEK_MAX_TOKENS", "8192")),
            request_timeout=int(os.getenv("DEEPSEEK_TIMEOUT", "180")),
            max_retries=int(os.getenv("DEEPSEEK_MAX_RETRIES", "3")),
            concurrency=max(1, int(os.getenv("DEEPSEEK_CONCURRENCY", "3"))),
        )
