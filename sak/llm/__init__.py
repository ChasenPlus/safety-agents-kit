# -*- coding: utf-8 -*-
"""按配置创建模型客户端。新增供应商只需在这里注册。"""
from __future__ import annotations

from ..config import Config
from .anthropic import AnthropicClient
from .base import LLMClient
from .gemini import GeminiClient
from .mock import MockClient
from .openai_compat import OpenAICompatClient

PROVIDERS = ("openai", "anthropic", "gemini", "mock")

# 常见供应商的默认接入点，用户只填 provider + api_key 即可
PRESETS = {
    "deepseek": ("openai", "https://api.deepseek.com/v1", "deepseek-chat"),
    "qwen": ("openai", "https://dashscope.aliyuncs.com/compatible-mode/v1", "qwen-max"),
    "kimi": ("openai", "https://api.moonshot.cn/v1", "moonshot-v1-32k"),
    "zhipu": ("openai", "https://open.bigmodel.cn/api/paas/v4", "glm-4-plus"),
    "openai": ("openai", "https://api.openai.com/v1", "gpt-4o-mini"),
    "ollama": ("openai", "http://localhost:11434/v1", "qwen2.5:14b"),
    "vllm": ("openai", "http://localhost:8000/v1", "local-model"),
    "anthropic": ("anthropic", "https://api.anthropic.com", "claude-sonnet-4-20250514"),
    "gemini": ("gemini", "https://generativelanguage.googleapis.com/v1beta", "gemini-2.0-flash"),
}


def build_client(cfg: Config) -> LLMClient:
    provider = cfg.provider.lower()
    if provider == "mock":
        return MockClient()
    if provider == "anthropic":
        return AnthropicClient(cfg.model, cfg.base_url, cfg.api_key,
                               temperature=cfg.temperature, max_tokens=cfg.max_tokens)
    if provider == "gemini":
        return GeminiClient(cfg.model, cfg.base_url, cfg.api_key,
                            temperature=cfg.temperature, max_tokens=cfg.max_tokens)
    if provider in ("openai", "openai_compat", "compatible"):
        return OpenAICompatClient(cfg.model, cfg.base_url, cfg.api_key,
                                  temperature=cfg.temperature, max_tokens=cfg.max_tokens)
    raise SystemExit(f"不支持的 provider：{cfg.provider}，可选：{', '.join(PROVIDERS)}")
