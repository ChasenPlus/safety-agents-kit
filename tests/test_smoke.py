# -*- coding: utf-8 -*-
"""离线自测：不联网、不依赖任何真实模型。

    python tests/test_smoke.py

覆盖：
1. 8 个分支清单与提示词可加载
2. 配置解析（相对路径按配置文件所在目录解析）
3. MCP 客户端能启动真实 stdio 服务并调用工具（用 server/agent_switch）
4. 编排循环：模型发工具调用 -> 工具执行 -> 模型收尾
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))

from sak.agent import BranchAgent, MCPManager  # noqa: E402
from sak.branches import load_branches, resolve_branch  # noqa: E402
from sak.config import Config, ServerConfig, load_config  # noqa: E402
from sak.mcp_client import StdioMCPClient  # noqa: E402

ok = 0
fail = 0


def check(name: str, cond: bool, detail: str = "") -> None:
    global ok, fail
    if cond:
        ok += 1
        print(f"  PASS  {name}")
    else:
        fail += 1
        print(f"  FAIL  {name} {detail}")


def test_branches() -> None:
    print("[1] 分支清单")
    branches = load_branches()
    check("分支数量为 8", len(branches) == 8, f"实际 {len(branches)}")
    for bid, branch in branches.items():
        try:
            text = branch.prompt(REPO / "prompts")
            check(f"{bid} 提示词非空", len(text) > 500, f"{len(text)} 字符")
        except SystemExit as exc:
            check(f"{bid} 提示词可读", False, str(exc))
    check("按名称解析分支", resolve_branch(branches, "条款检索与判定").id == "A")


def test_config() -> None:
    print("[2] 配置解析")
    cfg = load_config(str(REPO / "tests" / "config.mcp.json"))
    check("provider=mock", cfg.provider == "mock", cfg.provider)
    check("kb_root 解析为绝对路径", cfg.kb_root.is_absolute(), str(cfg.kb_root))
    check("mcp 服务器已解析", "sak-handoff" in cfg.mcp_servers, str(list(cfg.mcp_servers)))


def test_mcp() -> None:
    print("[3] MCP 客户端")
    client = StdioMCPClient("sak-handoff", sys.executable, [str(REPO / "server" / "agent_switch" / "server.py")])
    try:
        client.start()
    except Exception as exc:  # noqa: BLE001
        check("stdio 服务启动", False, str(exc))
        return
    try:
        names = {t.get("name") for t in client.tools}
        check("暴露 3 个工具", len(names) == 3, str(sorted(names)))
        text = client.call_tool("list_branches", {})
        check("调用 list_branches 有返回", "示例" in text or "分支" in text, text[:80])
        packet = client.call_tool("get_handoff_packet", {"branch": "A", "task": "自测任务"})
        check("交接包含任务与提示词", "自测任务" in packet and len(packet) > 1000, str(len(packet)))
    finally:
        client.close()


def test_agent_loop() -> None:
    print("[4] 编排循环（mock 模型 + 真实 MCP）")
    cfg = load_config(str(REPO / "tests" / "config.mcp.json"))
    cfg = Config(**{**cfg.__dict__, "mcp_servers": {"sak-handoff": cfg.mcp_servers["sak-handoff"]}})
    branch = resolve_branch(load_branches(), "A")
    manager = MCPManager(cfg)
    try:
        manager.start(["sak-handoff"])
        check("MCP 服务已连接", "sak-handoff" in manager.clients, str(list(manager.clients)))
        from sak.llm.mock import MockClient

        agent = BranchAgent(cfg, branch, MockClient())
        answer = agent.run("TOOL:list_branches {}", manager)
        check("模型发起工具调用并收尾", "[mock]" in answer, answer[:120])
    finally:
        manager.close()


def main() -> int:
    test_branches()
    test_config()
    test_mcp()
    test_agent_loop()
    print(f"\n结果：PASS {ok} / FAIL {fail}")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
