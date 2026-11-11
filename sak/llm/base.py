# -*- coding: utf-8 -*-
"""模型适配层：把各家的 tool calling 差异收敛成同一组数据结构。"""
from __future__ import annotations

import json
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Any


@dataclass
class ToolSpec:
    """暴露给模型的工具定义（JSON Schema）。"""
    name: str
    description: str
    input_schema: dict[str, Any]
    server: str = ""


@dataclass
class ToolCall:
    id: str
    name: str
    arguments: dict[str, Any]


@dataclass
class AssistantTurn:
    content: str = ""
    tool_calls: list[ToolCall] = field(default_factory=list)
    raw: Any = None


class LLMClient(ABC):
    """统一的对话接口。messages 为中立格式：
    {"role":"user","content":str}
    {"role":"assistant","content":str,"tool_calls":[ToolCall...]}
    {"role":"tool","tool_call_id":str,"name":str,"content":str}
    """

    def __init__(self, model: str, temperature: float = 0.2, max_tokens: int = 4096):
        self.model = model
        self.temperature = temperature
        self.max_tokens = max_tokens

    @abstractmethod
    def chat(self, system: str, messages: list[dict], tools: list[ToolSpec]) -> AssistantTurn:
        ...

    # ---- 通用小工具 ----
    @staticmethod
    def _parse_arguments(raw: Any) -> dict[str, Any]:
        if isinstance(raw, dict):
            return raw
        if not raw:
            return {}
        try:
            return json.loads(raw)
        except (TypeError, ValueError):
            return {"_raw": str(raw)}
