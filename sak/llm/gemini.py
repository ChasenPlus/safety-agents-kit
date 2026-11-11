# -*- coding: utf-8 -*-
"""Google Gemini 适配器（generateContent + functionDeclarations）。"""
from __future__ import annotations

import json
import urllib.error
import urllib.request

from .base import AssistantTurn, LLMClient, ToolCall, ToolSpec


class GeminiClient(LLMClient):
    def __init__(self, model: str, base_url: str, api_key: str, **kw):
        super().__init__(model, **kw)
        self.base_url = (base_url or "https://generativelanguage.googleapis.com/v1beta").rstrip("/")
        self.api_key = api_key

    def chat(self, system: str, messages: list[dict], tools: list[ToolSpec]) -> AssistantTurn:
        contents = []
        for msg in messages:
            role = msg["role"]
            if role == "tool":
                contents.append(
                    {
                        "role": "user",
                        "parts": [
                            {
                                "functionResponse": {
                                    "name": msg.get("name", "tool"),
                                    "response": {"result": msg.get("content", "")},
                                }
                            }
                        ],
                    }
                )
            elif role == "assistant":
                parts: list[dict] = []
                if msg.get("content"):
                    parts.append({"text": msg["content"]})
                for c in msg.get("tool_calls") or []:
                    parts.append({"functionCall": {"name": c.name, "args": c.arguments}})
                contents.append({"role": "model", "parts": parts or [{"text": ""}]})
            else:
                contents.append({"role": "user", "parts": [{"text": msg.get("content", "")}]})

        payload: dict = {
            "contents": contents,
            "generationConfig": {"temperature": self.temperature, "maxOutputTokens": self.max_tokens},
        }
        if system:
            payload["systemInstruction"] = {"parts": [{"text": system}]}
        if tools:
            payload["tools"] = [
                {
                    "functionDeclarations": [
                        {
                            "name": t.name,
                            "description": t.description,
                            # Gemini 不接受 $schema 等字段，只保留 type/properties/required
                            "parameters": _clean_schema(t.input_schema),
                        }
                        for t in tools
                    ]
                }
            ]

        url = f"{self.base_url}/models/{self.model}:generateContent?key={self.api_key}"
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(url, data=body, method="POST")
        req.add_header("Content-Type", "application/json")
        try:
            with urllib.request.urlopen(req, timeout=180) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")[:2000]
            raise RuntimeError(f"Gemini 返回 {exc.code}：{detail}") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"无法连接 Gemini：{exc.reason}") from exc

        candidate = (data.get("candidates") or [{}])[0]
        parts = (candidate.get("content") or {}).get("parts") or []
        text_parts, calls = [], []
        for i, part in enumerate(parts):
            if "text" in part:
                text_parts.append(part["text"])
            if "functionCall" in part:
                fc = part["functionCall"]
                calls.append(
                    ToolCall(
                        id=f"gemini_call_{i}",
                        name=str(fc.get("name")),
                        arguments=self._parse_arguments(fc.get("args")),
                    )
                )
        return AssistantTurn(content="".join(text_parts), tool_calls=calls, raw=data)


def _clean_schema(schema: dict | None) -> dict:
    if not schema:
        return {"type": "OBJECT", "properties": {}}
    out = {}
    for key in ("type", "description", "properties", "required", "items", "enum"):
        if key in schema:
            out[key] = schema[key]
    if "type" in out and isinstance(out["type"], str):
        out["type"] = out["type"].upper()
    if "properties" in out and isinstance(out["properties"], dict):
        out["properties"] = {k: _clean_schema(v) for k, v in out["properties"].items()}
    if "items" in out and isinstance(out["items"], dict):
        out["items"] = _clean_schema(out["items"])
    return out
