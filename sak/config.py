# -*- coding: utf-8 -*-
"""配置加载：仓库内 sak.json + 环境变量覆盖。

查找顺序（先找到先用，后面的只补缺省值）：
    1. --config 显式指定的文件
    2. 当前目录 ./sak.json
    3. 用户目录 ~/.sak/config.json
    4. 仓库自带 config.example.json
环境变量：
    SAK_PROVIDER / SAK_MODEL / SAK_BASE_URL / SAK_API_KEY / SAK_KB_ROOT / SAK_MAX_STEPS
"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .presets import PRESETS

REPO_ROOT = Path(__file__).resolve().parents[1]

DEFAULTS: dict[str, Any] = {
    "provider": "openai",
    "model": "deepseek-chat",
    "base_url": "https://api.deepseek.com/v1",
    "api_key_env": "DEEPSEEK_API_KEY",
    "temperature": 0.2,
    "max_tokens": 4096,
    "max_steps": 24,
    "kb_root": "./kb",
    "prompts_dir": "./prompts",
    "mcp_servers": {},
}

ENV_OVERRIDES = {
    "SAK_PROVIDER": "provider",
    "SAK_MODEL": "model",
    "SAK_BASE_URL": "base_url",
    "SAK_API_KEY": "api_key",
    "SAK_KB_ROOT": "kb_root",
    "SAK_MAX_STEPS": "max_steps",
}


@dataclass
class ServerConfig:
    """一个 MCP 服务（stdio）的启动参数。"""
    name: str
    command: str
    args: list[str] = field(default_factory=list)
    env: dict[str, str] = field(default_factory=dict)
    cwd: str | None = None


@dataclass
class Config:
    provider: str
    model: str
    base_url: str
    api_key_env: str
    api_key: str
    temperature: float
    max_tokens: int
    max_steps: int
    kb_root: Path
    prompts_dir: Path
    mcp_servers: dict[str, ServerConfig]
    source: Path | None = None

    def server(self, name: str) -> ServerConfig | None:
        return self.mcp_servers.get(name)


def _resolve_arg(value: str, base: Path) -> str:
    """只对形如脚本路径的参数做相对路径解析，命令开关（-y 等）原样保留。"""
    if value.endswith((".py", ".js", ".mjs")) or value.startswith(("./", "../", ".\\", "..\\")):
        path = Path(value).expanduser()
        if not path.is_absolute():
            return str((base / path).resolve())
    return value


def _candidate_paths(explicit: str | None) -> list[Path]:
    cands = []
    if explicit:
        cands.append(Path(explicit))
    cands += [
        Path.cwd() / "sak.json",
        Path.home() / ".sak" / "config.json",
        REPO_ROOT / "sak.json",
    ]
    return cands


def load_config(explicit: str | None = None) -> Config:
    data: dict[str, Any] = dict(DEFAULTS)
    source: Path | None = None
    explicit_keys: set[str] = set()

    if explicit and not Path(explicit).is_file():
        raise SystemExit(f"配置文件不存在：{explicit}")

    for path in _candidate_paths(explicit):
        if not path.is_file():
            continue
        source = path
        try:
            loaded = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise SystemExit(f"配置文件不是合法 JSON：{path}（{exc}）")
        for k, v in loaded.items():
            if k == "mcp_servers":
                data["mcp_servers"] = {**data["mcp_servers"], **v}
            else:
                data[k] = v
                explicit_keys.add(k)
        break

    for env_key, cfg_key in ENV_OVERRIDES.items():
        if os.environ.get(env_key):
            data[cfg_key] = os.environ[env_key]
            explicit_keys.add(cfg_key)

    # provider 写成预设名（deepseek / ollama / qwen …）时，补齐未显式指定的接入点
    preset = PRESETS.get(str(data["provider"]).lower())
    if preset:
        _, preset_base, preset_model, preset_key_env = preset
        if "base_url" not in explicit_keys:
            data["base_url"] = preset_base
        if "model" not in explicit_keys:
            data["model"] = preset_model
        if "api_key_env" not in explicit_keys:
            data["api_key_env"] = preset_key_env

    api_key = str(data.get("api_key") or os.environ.get(str(data["api_key_env"]), ""))

    base = source.parent if source else REPO_ROOT

    def _resolve(p: str | Path) -> Path:
        path = Path(p).expanduser()
        return path if path.is_absolute() else (base / path).resolve()

    servers: dict[str, ServerConfig] = {}
    for name, raw in (data.get("mcp_servers") or {}).items():
        if not raw:
            continue
        servers[name] = ServerConfig(
            name=name,
            command=str(raw.get("command", "python")),
            args=[_resolve_arg(str(a), base) for a in raw.get("args", [])],
            env={str(k): str(v) for k, v in (raw.get("env") or {}).items()},
            cwd=str(_resolve(raw["cwd"])) if raw.get("cwd") else None,
        )

    return Config(
        provider=str(data["provider"]),
        model=str(data["model"]),
        base_url=str(data["base_url"]).rstrip("/"),
        api_key_env=str(data["api_key_env"]),
        api_key=api_key,
        temperature=float(data["temperature"]),
        max_tokens=int(data["max_tokens"]),
        max_steps=int(data["max_steps"]),
        kb_root=_resolve(data["kb_root"]),
        prompts_dir=_resolve(data["prompts_dir"]),
        mcp_servers=servers,
        source=source,
    )
