# -*- coding: utf-8 -*-
"""按配置创建模型客户端。新增供应商只需实现 LLMClient.chat() 并在这里注册。"""
from __future__ import annotations

from ..config import Config
from ..presets import PRESETS, protocol_of
from .anthropic import AnthropicClient
from .base import AssistantTurn, LLMClient, ToolCall, ToolSpec
from .gemini import GeminiClient
from .mock import MockClient
from .openai_compat import OpenAICompatClient

PROTOCOLS = ("openai", "anthropic", "gemini", "mock")


def build_client(cfg: Config) -> LLMClient:
    protocol = protocol_of(cfg.provider)
    if protocol == "mock":
        return MockClient()

    kwargs = {"temperature": cfg.temperature, "max_tokens": cfg.max_tokens}
    if protocol == "anthropic":
        return AnthropicClient(cfg.model, cfg.base_url, cfg.api_key, **kwargs)
    if protocol == "gemini":
        return GeminiClient(cfg.model, cfg.base_url, cfg.api_key, **kwargs)
    if protocol == "openai":
        return OpenAICompatClient(cfg.model, cfg.base_url, cfg.api_key, **kwargs)
    raise SystemExit(
        f"不支持的 provider：{cfg.provider}，可选：{', '.join([*PROTOCOLS, *PRESETS])}"
    )


__all__ = [
    "AssistantTurn",
    "LLMClient",
    "ToolCall",
    "ToolSpec",
    "PROTOCOLS",
    "PRESETS",
    "build_client",
]
