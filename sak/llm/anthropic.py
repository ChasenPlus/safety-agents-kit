# -*- coding: utf-8 -*-
"""Anthropic Messages API 适配器。"""
from __future__ import annotations

import json
import urllib.error
import urllib.request

from .base import AssistantTurn, LLMClient, ToolCall, ToolSpec

API_VERSION = "2023-06-01"


class AnthropicClient(LLMClient):
    def __init__(self, model: str, base_url: str, api_key: str, **kw):
        super().__init__(model, **kw)
        self.base_url = (base_url or "https://api.anthropic.com").rstrip("/")
        self.api_key = api_key

    def chat(self, system: str, messages: list[dict], tools: list[ToolSpec]) -> AssistantTurn:
        payload: dict = {
            "model": self.model,
            "max_tokens": self.max_tokens,
            "temperature": self.temperature,
            "system": system,
            "messages": self._to_wire(messages),
        }
        if tools:
            payload["tools"] = [
                {"name": t.name, "description": t.description,
                 "input_schema": t.input_schema or {"type": "object", "properties": {}}}
                for t in tools
            ]

        url = f"{self.base_url}/v1/messages"
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(url, data=body, method="POST")
        req.add_header("Content-Type", "application/json")
        req.add_header("x-api-key", self.api_key)
        req.add_header("anthropic-version", API_VERSION)
        try:
            with urllib.request.urlopen(req, timeout=180) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")[:2000]
            raise RuntimeError(f"Anthropic 返回 {exc.code}：{detail}") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"无法连接 Anthropic：{exc.reason}") from exc

        text_parts, calls = [], []
        for block in data.get("content") or []:
            if block.get("type") == "text":
                text_parts.append(block.get("text", ""))
            elif block.get("type") == "tool_use":
                calls.append(
                    ToolCall(
                        id=str(block.get("id")),
                        name=str(block.get("name")),
                        arguments=self._parse_arguments(block.get("input")),
                    )
                )
        return AssistantTurn(content="".join(text_parts), tool_calls=calls, raw=data)

    def _to_wire(self, messages: list[dict]) -> list[dict]:
        wire: list[dict] = []
        for msg in messages:
            role = msg["role"]
            if role == "tool":
                block = {
                    "type": "tool_result",
                    "tool_use_id": msg.get("tool_call_id", ""),
                    "content": msg.get("content", ""),
                }
                if wire and wire[-1]["role"] == "user" and isinstance(wire[-1]["content"], list):
                    wire[-1]["content"].append(block)
                else:
                    wire.append({"role": "user", "content": [block]})
            elif role == "assistant":
                blocks: list[dict] = []
                if msg.get("content"):
                    blocks.append({"type": "text", "text": msg["content"]})
                for c in msg.get("tool_calls") or []:
                    blocks.append(
                        {"type": "tool_use", "id": c.id, "name": c.name, "input": c.arguments}
                    )
                wire.append({"role": "assistant", "content": blocks or [{"type": "text", "text": ""}]})
            else:
                wire.append({"role": "user", "content": [{"type": "text", "text": msg.get("content", "")}]})
        return wire
