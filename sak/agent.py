# -*- coding: utf-8 -*-
"""MCP 服务管理与单分支执行循环。"""
from __future__ import annotations

import datetime
import sys
from dataclasses import dataclass
from typing import Callable

from .branches import Branch
from .config import Config
from .llm.base import AssistantTurn, LLMClient, ToolSpec
from .mcp_client import MCPError, StdioMCPClient

Progress = Callable[[str], None]


def _noop(_: str) -> None:
    pass


def _command(name: str) -> str:
    """配置里写 python/python3 时，用当前解释器，避免 PATH 里没有 python。"""
    if name in {"python", "python3", "py"}:
        return sys.executable
    return name


@dataclass
class ToolRoute:
    spec: ToolSpec
    client: StdioMCPClient
    remote_name: str


class MCPManager:
    """按分支声明的服务名启动 MCP 子进程，并把工具汇总成一份可暴露给模型的清单。"""

    def __init__(self, cfg: Config, progress: Progress = _noop):
        self.cfg = cfg
        self.progress = progress
        self.clients: dict[str, StdioMCPClient] = {}
        self._routes: dict[str, ToolRoute] = {}

    def start(self, servers: list[str]) -> None:
        for name in servers:
            if name in self.clients:
                continue
            conf = self.cfg.server(name)
            if conf is None:
                self.progress(f"! 跳过未配置的 MCP 服务：{name}")
                continue
            client = StdioMCPClient(conf.name, _command(conf.command), conf.args,
                                    env={**conf.env, "SAK_KB_ROOT": str(self.cfg.kb_root)},
                                    cwd=conf.cwd)
            try:
                client.start()
            except MCPError as exc:
                self.progress(f"! MCP 服务 {name} 启动失败：{exc}")
                continue
            self.clients[name] = client
            self.progress(f"· MCP 服务已连接：{name}（{len(client.tools)} 个工具）")
            self._register(name, client)

    def _register(self, server: str, client: StdioMCPClient) -> None:
        for raw in client.tools:
            remote = raw.get("name", "")
            local = remote
            if local in self._routes:
                local = f"{server}__{remote}"
            self._routes[local] = ToolRoute(
                spec=ToolSpec(
                    name=local,
                    description=raw.get("description", "") or f"{server} 提供的工具 {remote}",
                    input_schema=raw.get("inputSchema") or {"type": "object", "properties": {}},
                    server=server,
                ),
                client=client,
                remote_name=remote,
            )

    def tool_specs(self) -> list[ToolSpec]:
        return [r.spec for r in self._routes.values()]

    def call(self, name: str, arguments: dict) -> str:
        route = self._routes.get(name)
        if route is None:
            return f"[错误] 未知工具：{name}"
        try:
            return route.client.call_tool(route.remote_name, arguments)
        except MCPError as exc:
            return f"[错误] 工具 {name} 执行失败：{exc}"

    def close(self) -> None:
        for client in self.clients.values():
            client.close()
        self.clients.clear()
        self._routes.clear()


class BranchAgent:
    def __init__(self, cfg: Config, branch: Branch, llm: LLMClient, progress: Progress = _noop):
        self.cfg = cfg
        self.branch = branch
        self.llm = llm
        self.progress = progress

    # ---- 提示词 ----
    def system_prompt(self) -> str:
        prompt = self.branch.prompt(self.cfg.prompts_dir)
        today = datetime.date.today().isoformat()
        runtime = (
            "\n\n---\n运行环境（由 Safety Agents Kit 注入，不属于分支规则）：\n"
            f"- 知识库根目录：{self.cfg.kb_root}\n"
            f"- 今天日期：{today}\n"
            "- 可用工具来自 MCP 服务，调用方式与提示词中描述一致；"
            "若提示词要求读取某份 Wiki 笔记而工具无法访问，按提示词的降级规则处理并在结果里注明。\n"
        )
        return prompt + runtime

    def run(self, task: str, manager: MCPManager) -> str:
        system = self.system_prompt()
        tools = manager.tool_specs()
        messages: list[dict] = [{"role": "user", "content": task}]

        for step in range(1, self.cfg.max_steps + 1):
            self.progress(f"  · 第 {step} 轮：请求模型 {self.llm.model}")
            turn: AssistantTurn = self.llm.chat(system, messages, tools)

            if not turn.tool_calls:
                return turn.content

            messages.append({
                "role": "assistant",
                "content": turn.content,
                "tool_calls": turn.tool_calls,
            })
            for call in turn.tool_calls:
                self.progress(f"    → 调用工具 {call.name}")
                output = manager.call(call.name, call.arguments)
                messages.append({
                    "role": "tool",
                    "tool_call_id": call.id,
                    "name": call.name,
                    "content": output,
                })

        return (
            f"[达到最大步数 {self.cfg.max_steps}，已停止] 最后一段模型输出：\n"
            + (turn.content or "（无文本输出）")
        )
