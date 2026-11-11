# -*- coding: utf-8 -*-
"""极简 MCP 客户端（stdio + 换行分隔 JSON-RPC 2.0）。

不依赖官方 SDK，避免不同版本 mcp/fastmcp 之间的依赖冲突：
只要对方是标准 MCP stdio 服务，就能被本客户端驱动。
"""
from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
import time
from typing import Any

PROTOCOL_VERSION = "2024-11-05"


class MCPError(RuntimeError):
    pass


class StdioMCPClient:
    def __init__(self, name: str, command: str, args: list[str],
                 env: dict[str, str] | None = None, cwd: str | None = None,
                 startup_timeout: float = 60.0):
        self.name = name
        self.command = command
        self.args = args
        self.extra_env = env or {}
        self.cwd = cwd
        self.startup_timeout = startup_timeout
        self._proc: subprocess.Popen | None = None
        self._lines: queue.Queue = queue.Queue()
        self._next_id = 0
        self._lock = threading.Lock()
        self.tools: list[dict] = []

    # ---------- 生命周期 ----------
    def start(self) -> None:
        env = os.environ.copy()
        env.update(self.extra_env)
        env.setdefault("PYTHONIOENCODING", "utf-8")
        env.setdefault("PYTHONUTF8", "1")
        cmd = [self.command] + self.args
        creationflags = 0
        if os.name == "nt":
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            self._proc = subprocess.Popen(
                cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                env=env, cwd=self.cwd, text=True, encoding="utf-8", errors="replace",
                bufsize=1, creationflags=creationflags,
            )
        except FileNotFoundError as exc:
            hint = "（需要先安装 Node.js 18+）" if self.command in {"npx", "node"} else ""
            raise MCPError(f"[{self.name}] 启动失败，找不到命令 {self.command!r}{hint}：{exc}") from exc

        threading.Thread(target=self._pump_stdout, daemon=True).start()
        threading.Thread(target=self._pump_stderr, daemon=True).start()

        self._request("initialize", {
            "protocolVersion": PROTOCOL_VERSION,
            "capabilities": {},
            "clientInfo": {"name": "safety-agents-kit", "version": "0.1.0"},
        })
        self._notify("notifications/initialized", {})
        self.tools = self._request("tools/list", {}).get("tools", [])

    def close(self) -> None:
        if not self._proc:
            return
        try:
            self._proc.stdin and self._proc.stdin.close()
            self._proc.terminate()
            self._proc.wait(timeout=5)
        except Exception:  # noqa: BLE001
            try:
                self._proc.kill()
            except Exception:  # noqa: BLE001
                pass
        self._proc = None

    # ---------- IO ----------
    def _pump_stdout(self) -> None:
        assert self._proc and self._proc.stdout
        for line in self._proc.stdout:
            line = line.strip()
            if not line:
                continue
            try:
                self._lines.put(json.loads(line))
            except json.JSONDecodeError:
                continue  # 非 JSON 输出（启动横幅等）直接忽略

    def _pump_stderr(self) -> None:
        assert self._proc and self._proc.stderr
        for line in self._proc.stderr:
            sys.stderr.write(f"[{self.name}] {line}")

    def _send(self, payload: dict) -> None:
        if not self._proc or not self._proc.stdin:
            raise MCPError(f"[{self.name}] 进程未启动")
        with self._lock:
            self._proc.stdin.write(json.dumps(payload, ensure_ascii=False) + "\n")
            self._proc.stdin.flush()

    def _notify(self, method: str, params: dict) -> None:
        self._send({"jsonrpc": "2.0", "method": method, "params": params})

    def _request(self, method: str, params: dict, timeout: float | None = None) -> dict:
        self._next_id += 1
        req_id = self._next_id
        self._send({"jsonrpc": "2.0", "id": req_id, "method": method, "params": params})
        deadline = time.time() + (timeout or self.startup_timeout)
        while time.time() < deadline:
            try:
                msg = self._lines.get(timeout=0.2)
            except queue.Empty:
                if self._proc and self._proc.poll() is not None:
                    raise MCPError(f"[{self.name}] 进程已退出（exit={self._proc.returncode}）")
                continue
            if msg.get("id") != req_id:
                continue  # 通知或其他请求的响应
            if "error" in msg:
                raise MCPError(f"[{self.name}] {method} 失败：{msg['error']}")
            return msg.get("result") or {}
        raise MCPError(f"[{self.name}] {method} 超时（{timeout or self.startup_timeout}s）")

    # ---------- 工具 ----------
    def call_tool(self, name: str, arguments: dict[str, Any], timeout: float = 300.0) -> str:
        result = self._request("tools/call", {"name": name, "arguments": arguments}, timeout=timeout)
        parts = []
        for item in result.get("content") or []:
            if item.get("type") == "text":
                parts.append(item.get("text", ""))
            elif item.get("type") == "resource":
                parts.append(json.dumps(item.get("resource", {}), ensure_ascii=False))
            else:
                parts.append(json.dumps(item, ensure_ascii=False))
        text = "\n".join(parts) if parts else json.dumps(result, ensure_ascii=False)
        if result.get("isError"):
            return f"[工具执行报错] {text}"
        return text
