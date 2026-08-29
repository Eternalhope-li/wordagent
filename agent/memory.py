"""持久化上下文记忆：记录历史指令与生成结果，供后续对话引用。"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any, Optional


class Memory:
    def __init__(self, path: Path, max_entries: int = 60):
        self.path = Path(path)
        self.max_entries = max_entries
        self.settings: dict[str, Any] = {}
        self.entries: list[dict] = []
        self.load()

    # ---------- 持久化 ----------
    def load(self) -> None:
        if not self.path.exists():
            return
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                self.entries = data.get("entries", [])
                self.settings = data.get("settings", {})
            else:
                self.entries = []
        except (json.JSONDecodeError, OSError):
            self.entries = []

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "settings": self.settings,
            "entries": self.entries[-self.max_entries :],
        }
        self.path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    # ---------- 写入 ----------
    def add_user(self, text: str) -> None:
        self.entries.append(
            {"role": "user", "time": self._now(), "text": text.strip()}
        )
        self.save()

    def add_result(self, result: dict) -> None:
        self.entries.append({"role": "assistant", "time": self._now(), **result})
        self.save()

    # ---------- 读取 ----------
    def context_text(self, max_entries: int = 8) -> str:
        """构造注入系统提示的近期记忆摘要。"""
        recent = [e for e in self.entries if e.get("role") == "user"][-max_entries:]
        if not recent:
            return "（暂无历史）"
        lines = []
        for i, item in enumerate(recent, 1):
            text = str(item.get("text", "")).replace("\n", " ")[:200]
            lines.append(f"{i}. [{item.get('time', '')}] {text}")
        return "\n".join(lines)

    def history_text(self) -> str:
        """完整历史（供 /memory 命令展示）。"""
        if not self.entries:
            return "（记忆为空）"
        lines = []
        for i, item in enumerate(self.entries, 1):
            role = "用户" if item.get("role") == "user" else "助手"
            if item.get("role") == "user":
                body = str(item.get("text", "")).replace("\n", " ")[:120]
            else:
                body = f"生成文档《{item.get('title', '')}》 -> {item.get('file', '')}"
            lines.append(f"{i}. [{item.get('time', '')}] {role}: {body}")
        return "\n".join(lines)

    def clear(self) -> None:
        self.entries = []
        self.save()

    def set_setting(self, key: str, value: Any) -> None:
        self.settings[key] = value
        self.save()

    def get_setting(self, key: str, default: Any = None) -> Any:
        return self.settings.get(key, default)

    @staticmethod
    def _now() -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
