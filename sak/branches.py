# -*- coding: utf-8 -*-
"""分支清单：读取 branches/*.json，拼装系统提示词。"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
BRANCHES_DIR = REPO_ROOT / "branches"


@dataclass
class Branch:
    id: str
    name: str
    summary: str
    prompt_file: str
    servers: list[str] = field(default_factory=list)
    handoff_to: list[str] = field(default_factory=list)

    def prompt(self, prompts_dir: Path) -> str:
        path = (prompts_dir / self.prompt_file)
        if not path.is_file():
            path = REPO_ROOT / "prompts" / self.prompt_file
        if not path.is_file():
            raise SystemExit(f"分支 {self.id} 的提示词文件不存在：{path}")
        return path.read_text(encoding="utf-8")


def load_branches(branches_dir: Path | None = None) -> dict[str, Branch]:
    directory = branches_dir or BRANCHES_DIR
    if not directory.is_dir():
        raise SystemExit(
            f"找不到分支目录：{directory}\n"
            "分支清单与提示词随源码分发，请用 `pip install -e .` 或 `python -m sak` 在仓库目录内运行。"
        )
    out: dict[str, Branch] = {}
    for path in sorted(directory.glob("*.json")):
        raw = json.loads(path.read_text(encoding="utf-8"))
        branch = Branch(
            id=str(raw["id"]),
            name=str(raw["name"]),
            summary=str(raw.get("summary", "")),
            prompt_file=str(raw["prompt"]),
            servers=[str(s) for s in raw.get("servers", [])],
            handoff_to=[str(s) for s in raw.get("handoff_to", [])],
        )
        out[branch.id] = branch
    return out


def resolve_branch(branches: dict[str, Branch], key: str) -> Branch:
    """支持字母、名称、别名（不区分大小写、允许包含匹配）。"""
    k = key.strip()
    if not k:
        raise SystemExit("分支不能为空")
    for b in branches.values():
        if k == b.id or k.lower() == b.id.lower():
            return b
    for b in branches.values():
        if k == b.name or k in b.name:
            return b
    available = "、".join(f"{b.id}={b.name}" for b in branches.values())
    raise SystemExit(f"未知分支：{key}。可用：{available}")
