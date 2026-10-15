"""Client gọi Qwen qua OpenRouter: bất đồng bộ, retry, đếm token."""
from __future__ import annotations

import asyncio
import re
from dataclasses import dataclass, field

import httpx

from . import config


@dataclass
class Usage:
    calls: int = 0
    prompt_tokens: int = 0
    completion_tokens: int = 0
    cost: float = 0.0

    def add(self, u: dict):
        self.calls += 1
        self.prompt_tokens += u.get("prompt_tokens", 0) or 0
        self.completion_tokens += u.get("completion_tokens", 0) or 0
        self.cost += u.get("cost", 0) or 0


@dataclass
class LLM:
    model: str = config.MODEL
    client: httpx.AsyncClient = field(default=None)

    def __post_init__(self):
        if not config.API_KEY:
            raise RuntimeError("Thiếu OPENROUTER_API_KEY (đặt trong file .env hoặc biến môi trường).")
        self.client = httpx.AsyncClient(
            base_url=config.BASE_URL, timeout=config.LLM_TIMEOUT,
            headers={"Authorization": f"Bearer {config.API_KEY}", "X-Title": "excel-table-qa"})

    def _payload(self, messages, temperature, max_tokens):
        body = {"model": self.model, "messages": messages, "temperature": temperature,
                "max_tokens": max_tokens or config.MAX_TOKENS, "usage": {"include": True}}
        r = config.REASONING.lower()
        if r == "off":
            body["reasoning"] = {"enabled": False}
        elif r in ("low", "medium", "high"):
            body["reasoning"] = {"effort": r, "exclude": True}
        return body

    async def chat(self, messages: list[dict], usage: Usage, temperature: float = 0.0,
                   max_tokens: int | None = None) -> str:
        """Trả về nội dung trả lời, '' nếu thất bại."""
        body = self._payload(messages, temperature, max_tokens)
        for attempt in range(config.LLM_RETRIES + 1):
            try:
                resp = await self.client.post("/chat/completions", json=body)
                if resp.status_code == 429 or resp.status_code >= 500:
                    raise httpx.HTTPStatusError(f"HTTP {resp.status_code}", request=resp.request, response=resp)
                if resp.status_code >= 400:              # sai key/tham số: không thử lại
                    print(f"[llm] HTTP {resp.status_code}: {resp.text[:300]}")
                    return ""
                data = resp.json()
                usage.add(data.get("usage") or {})
                text = (data["choices"][0]["message"].get("content") or "").strip()
                return re.sub(r"<think>.*?</think>", "", text, flags=re.S).strip()
            except (httpx.HTTPError, KeyError, IndexError, ValueError) as e:
                if attempt == config.LLM_RETRIES:
                    print(f"[llm] lỗi sau {attempt + 1} lần thử: {e}")
                    return ""
                await asyncio.sleep(1.5 * 2 ** attempt)
        return ""

    async def close(self):
        await self.client.aclose()
