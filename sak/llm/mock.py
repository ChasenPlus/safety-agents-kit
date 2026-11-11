# -*- coding: utf-8 -*-
"""离线 Mock 模型：不联网，用于自测编排链路。

行为：
- 若问题里出现 "TOOL:" 前缀，则按 `TOOL:<工具名> <JSON参数>` 调一次工具，再输出总结；
- 否则直接回显系统提示词末尾 + 用户问题，证明提示词与工具描述已正确注入。
"""
from __future__ import annotations

import json
import re

from .base import AssistantTurn, LLMClient, ToolCall, ToolSpec

TOOL_DIRECTIVE = re.compile(r"TOOL:([A-Za-z0-9_\-]+)\s*(\{.*\})?", re.S)


class MockClient(LLMClient):
    def __init__(self, model: str = "mock", **kw):
        super().__init__(model, **kw)
        self.calls = 0

    def chat(self, system: str, messages: list[dict], tools: list[ToolSpec]) -> AssistantTurn:
        self.calls += 1
        last_user = next((m for m in reversed(messages) if m["role"] == "user"), {})
        question = str(last_user.get("content", ""))

        match = TOOL_DIRECTIVE.search(question)
        already_called = any(m["role"] == "tool" for m in messages)
        if match and not already_called:
            name = match.group(1)
            args = self._parse_arguments(match.group(2) or "{}")
            return AssistantTurn(
                content="",
                tool_calls=[ToolCall(id=f"mock_{self.calls}", name=name, arguments=args)],
            )

        tool_names = ", ".join(t.name for t in tools) or "（无）"
        return AssistantTurn(
            content=(
                "[mock] 分支提示词已加载，共 "
                f"{len(system)} 字符；可用工具：{tool_names}。\n"
                f"问题：{question}"
            )
        )
