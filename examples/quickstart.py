# -*- coding: utf-8 -*-
"""库调用示例：不经过 CLI，直接把 SAK 当 Python 包用。

运行前准备：
    pip install -e .
    export DEEPSEEK_API_KEY=sk-xxx      # 或任意一家，改 config 里的 provider/base_url

    python examples/quickstart.py
"""
from __future__ import annotations

from pathlib import Path

from sak.agent import BranchAgent, MCPManager
from sak.branches import load_branches, resolve_branch
from sak.config import load_config
from sak.llm import build_client

REPO = Path(__file__).resolve().parents[1]


def main() -> None:
    cfg = load_config(str(REPO / "sak.json")) if (REPO / "sak.json").is_file() else load_config()
    branch = resolve_branch(load_branches(), "A")

    print(f"分支：{branch.id} {branch.name}")
    print(f"模型：{cfg.provider}/{cfg.model}")

    manager = MCPManager(cfg, progress=lambda m: print(f"  {m}"))
    manager.start(branch.servers)
    try:
        agent = BranchAgent(cfg, branch, build_client(cfg))
        answer = agent.run(
            "GB 50194-2014 对施工现场配电箱的漏电保护有什么要求？"
            "请给出标准号、条款号和原文。",
            manager,
        )
    finally:
        manager.close()

    print("\n===== 结果 =====")
    print(answer)


if __name__ == "__main__":
    main()
