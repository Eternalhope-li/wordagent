"""DeepSeek（OpenAI 兼容）对话客户端。

适配推理模型（deepseek-v4-flash：先输出 reasoning_content 再输出正文）：
- 自动统计每次调用的 token 消耗（输入/输出/推理/缓存命中），供界面与日志展示
- 输出被 max_tokens 截断（finish_reason=length）时自动加大预算重试，避免同参数盲重试反复烧 token
- 网络错误/限流/5xx 退避重试；4xx（Key/模型/地址错误）直接报错不重试
- JSON 模式单次请求 + 文本兜底提取，不重复烧 token
"""
from __future__ import annotations

import json
import time
from typing import Any, Optional

import requests

from .config import Config

# 单次请求的输出预算上限：推理模型思考会占大头，给足空间避免截断
MAX_OUTPUT_CAP = 32768


class LLMError(RuntimeError):
    """LLM 调用相关错误。"""


def _strip_fences(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        lines = text.splitlines()
        if lines and lines[0].strip().lower().startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        text = "\n".join(lines).strip()
    return text


class LLMClient:
    def __init__(self, config: Config):
        self.config = config
        if not config.api_key:
            raise LLMError("缺少 DeepSeek API Key：请设置 DEEPSEEK_API_KEY 环境变量、.env 文件或 --api-key 参数")
        self._session = requests.Session()  # keep-alive 复用连接
        self.reset_usage()

    # ---------------- token 统计 ----------------
    def reset_usage(self) -> None:
        self._usage = {
            "calls": 0,
            "prompt_tokens": 0,
            "completion_tokens": 0,
            "reasoning_tokens": 0,
            "cached_tokens": 0,
        }

    @property
    def usage(self) -> dict:
        self._ensure_usage()
        return dict(self._usage)

    def usage_text(self) -> str:
        self._ensure_usage()
        u = self._usage
        return (
            f"本次共调用 {u['calls']} 次｜输入 {u['prompt_tokens']}｜输出 {u['completion_tokens']}"
            f"（含推理 {u['reasoning_tokens']}）｜缓存命中 {u['cached_tokens']} tokens"
        )

    def _ensure_usage(self) -> None:
        if not hasattr(self, "_usage"):
            self.reset_usage()

    def _record_usage(self, data: dict) -> None:
        self._ensure_usage()
        usage = data.get("usage") or {}
        if not usage:
            return
        self._usage["calls"] += 1
        self._usage["prompt_tokens"] += int(usage.get("prompt_tokens") or 0)
        self._usage["completion_tokens"] += int(usage.get("completion_tokens") or 0)
        comp = usage.get("completion_tokens_details") or {}
        self._usage["reasoning_tokens"] += int(comp.get("reasoning_tokens") or 0)
        prompt = usage.get("prompt_tokens_details") or {}
        cached = int(prompt.get("cached_tokens") or 0) or int(usage.get("prompt_cache_hit_tokens") or 0)
        self._usage["cached_tokens"] += cached

    def _headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self.config.api_key}",
            "Content-Type": "application/json",
        }

    # ---------------- 对话 ----------------
    def chat(
        self,
        messages: list[dict],
        temperature: Optional[float] = None,
        max_tokens: Optional[int] = None,
        json_mode: bool = False,
    ) -> str:
        """发送一轮对话，返回助手正文（自动重试网络错误/限流/截断）。"""
        budget = max_tokens or self.config.max_tokens
        payload: dict[str, Any] = {
            "model": self.config.model,
            "messages": messages,
            "temperature": self.config.temperature if temperature is None else temperature,
            "max_tokens": budget,
            "stream": False,
        }
        if json_mode:
            payload["response_format"] = {"type": "json_object"}

        url = f"{self.config.base_url}/chat/completions"
        last_error: Optional[Exception] = None
        for attempt in range(1, self.config.max_retries + 1):
            try:
                resp = self._session.post(
                    url, headers=self._headers(), json=payload, timeout=self.config.request_timeout
                )
            except requests.RequestException as exc:
                last_error = exc
                if attempt < self.config.max_retries:
                    time.sleep(1.5 * attempt)
                continue

            if resp.status_code == 429:  # 限流：退避重试
                last_error = RuntimeError("HTTP 429 限流")
                if attempt < self.config.max_retries:
                    time.sleep(2 * attempt)
                continue

            if resp.status_code >= 400:
                detail = (resp.text or "")[:300]
                if resp.status_code in (400, 401, 403, 404):
                    raise LLMError(
                        f"LLM 接口返回 {resp.status_code}：{detail}\n"
                        f"提示：请检查 API Key、Base URL 与模型名称（当前 {self.config.model}）"
                    )
                last_error = RuntimeError(f"HTTP {resp.status_code}：{detail}")
                if attempt < self.config.max_retries:
                    time.sleep(1.5 * attempt)
                continue

            try:
                data = resp.json()
                choice = data["choices"][0]
                content = (choice["message"].get("content") or "").strip()
                finish_reason = choice.get("finish_reason") or ""
            except (KeyError, IndexError, ValueError) as exc:
                last_error = exc
                if attempt < self.config.max_retries:
                    time.sleep(1.5 * attempt)
                continue

            self._record_usage(data)

            if finish_reason == "length":
                # 输出被截断（无论是否有正文）：加大预算重试一次，避免静默产出不完整章节
                new_budget = min(budget * 2, MAX_OUTPUT_CAP)
                if new_budget > budget and attempt < self.config.max_retries:
                    budget = new_budget
                    payload["max_tokens"] = budget
                    last_error = RuntimeError(
                        "output truncated by max_tokens, retrying with larger budget"
                    )
                    time.sleep(0.3)
                    continue
                raise LLMError(
                    f"model output still truncated at {budget} token cap; "
                    "raise max_tokens in API settings (max 32768)."
                )
            if not content:
                last_error = RuntimeError("model returned empty content")
                if attempt < self.config.max_retries:
                    time.sleep(0.3)
                continue
            return content

        raise LLMError(f"调用 LLM 失败（模型 {self.config.model}）: {last_error}")

    # ---------------- JSON 对话 ----------------
    def chat_json(
        self,
        messages: list[dict],
        temperature: Optional[float] = None,
        max_tokens: Optional[int] = None,
    ) -> dict:
        """请求 JSON 输出并解析；失败时从文本中兜底提取 JSON 片段，不重复请求。"""
        text = self.chat(
            messages,
            temperature=0.2 if temperature is None else temperature,
            max_tokens=max_tokens,
            json_mode=True,
        )
        try:
            return json.loads(_strip_fences(text))
        except json.JSONDecodeError:
            pass
        text = _strip_fences(text)
        start, end = text.find("{"), text.rfind("}")
        if start != -1 and end > start:
            try:
                return json.loads(text[start : end + 1])
            except json.JSONDecodeError:
                pass
        raise LLMError(f"模型未返回合法 JSON，前 500 字符: {text[:500]}")