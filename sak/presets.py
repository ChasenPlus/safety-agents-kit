# -*- coding: utf-8 -*-
"""常见模型供应商的接入点。只影响默认值，不改变协议。

用法：SAK_PROVIDER=ollama 时，若没有显式指定 base_url / model / api_key_env，就取这里的值。
"""
from __future__ import annotations

# provider 别名 -> (协议, base_url, 默认模型, 默认密钥环境变量)
PRESETS: dict[str, tuple[str, str, str, str]] = {
    "deepseek": ("openai", "https://api.deepseek.com/v1", "deepseek-chat", "DEEPSEEK_API_KEY"),
    "qwen": ("openai", "https://dashscope.aliyuncs.com/compatible-mode/v1", "qwen-max", "DASHSCOPE_API_KEY"),
    "kimi": ("openai", "https://api.moonshot.cn/v1", "moonshot-v1-32k", "MOONSHOT_API_KEY"),
    "zhipu": ("openai", "https://open.bigmodel.cn/api/paas/v4", "glm-4-plus", "ZHIPU_API_KEY"),
    "ollama": ("openai", "http://localhost:11434/v1", "qwen2.5:14b", "OLLAMA_API_KEY"),
    "vllm": ("openai", "http://localhost:8000/v1", "local-model", "VLLM_API_KEY"),
    "anthropic": ("anthropic", "https://api.anthropic.com", "claude-sonnet-4-20250514", "ANTHROPIC_API_KEY"),
    "gemini": ("gemini", "https://generativelanguage.googleapis.com/v1beta", "gemini-2.0-flash", "GEMINI_API_KEY"),
}


def protocol_of(provider: str) -> str:
    """把 provider 别名映射到实际协议实现。"""
    preset = PRESETS.get(provider.lower())
    if preset:
        return preset[0]
    if provider.lower() in ("openai", "openai_compat", "compatible"):
        return "openai"
    return provider.lower()
