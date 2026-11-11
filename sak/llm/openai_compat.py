# -*- coding: utf-8 -*-
"""OpenAI 兼容适配器：DeepSeek / 通义千问 / Kimi / 智谱 / OpenAI / vLLM / Ollama 通用。"""
from __future__ import annotations

import json
import urllib.error
import urllib.request

from .base import AssistantTurn, LLMClient, ToolCall, ToolSpec


class OpenAICompatClient(LLMClient):
    def __init__(self, model: str, base_url: str, api_key: str, **kw):
        super().__init__(model, **kw)
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key

    def _endpoint(self) -> str:
        # 允许用户填 https://host 或 https://host/v1
        if self.base_url.endswith("/chat/completions"):
            return self.base_url
        return f"{self.base_url}/chat/completions"

    def chat(self, system: str, messages: list[dict], tools: list[ToolSpec]) -> AssistantTurn:
        payload: dict = {
            "model": self.model,
            "temperature": self.temperature,
            "max_tokens": self.max_tokens,
            "messages": [{"role": "system", "content": system}] + self._to_wire(messages),
        }
        if tools:
            payload["tools"] = [
                {
                    "type": "function",
                    "function": {
                        "name": t.name,
                        "description": t.description,
                        "parameters": t.input_schema or {"type": "object", "properties": {}},
                    },
                }
                for t in tools
            ]
            payload["tool_choice"] = "auto"

        data = _post_json(self._endpoint(), payload, self.api_key)
        choice = (data.get("choices") or [{}])[0]
        message = choice.get("message") or {}
        calls = []
        for item in message.get("tool_calls") or []:
            fn = item.get("function") or {}
            calls.append(
                ToolCall(
                    id=str(item.get("id") or fn.get("name")),
                    name=str(fn.get("name") or ""),
                    arguments=self._parse_arguments(fn.get("arguments")),
                )
            )
        return AssistantTurn(content=message.get("content") or "", tool_calls=calls, raw=data)

    def _to_wire(self, messages: list[dict]) -> list[dict]:
        wire = []
        for msg in messages:
            role = msg["role"]
            if role == "tool":
                wire.append(
                    {
                        "role": "tool",
                        "tool_call_id": msg.get("tool_call_id", ""),
                        "content": msg.get("content", ""),
                    }
                )
            elif role == "assistant" and msg.get("tool_calls"):
                wire.append(
                    {
                        "role": "assistant",
                        "content": msg.get("content") or None,
                        "tool_calls": [
                            {
                                "id": c.id,
                                "type": "function",
                                "function": {
                                    "name": c.name,
                                    "arguments": json.dumps(c.arguments, ensure_ascii=False),
                                },
                            }
                            for c in msg["tool_calls"]
                        ],
                    }
                )
            else:
                wire.append({"role": role, "content": msg.get("content", "")})
        return wire


def _post_json(url: str, payload: dict, api_key: str, headers: dict | None = None) -> dict:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="POST")
    req.add_header("Content-Type", "application/json")
    if api_key:
        req.add_header("Authorization", f"Bearer {api_key}")
    for k, v in (headers or {}).items():
        req.add_header(k, v)
    try:
        with urllib.request.urlopen(req, timeout=180) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")[:2000]
        raise RuntimeError(f"模型接口返回 {exc.code}：{detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"无法连接模型接口 {url}：{exc.reason}") from exc
