# -*- coding: utf-8 -*-
"""sak-shell：受限的脚本执行 MCP 服务器。

分支提示词里有大量"运行 gate.py / append_ledger_entry.py / 生成 docx"的强制步骤，
独立运行时没有宿主 IDE 的 Bash 工具，所以由本服务提供最小可用的执行能力。

安全边界：
  - 只允许执行允许根目录内的 .py 脚本（默认：知识库根 + 仓库根）
  - 只允许 python 解释器，不接受任意 shell 命令
  - 有超时，输出有截断
它不是通用终端，也不需要通用终端。

环境变量：
    SAK_SHELL_ROOTS  允许执行脚本的根目录，多个用 os.pathsep 分隔；默认 SAK_KB_ROOT + 仓库根
    SAK_SHELL_TIMEOUT 单次执行超时秒数，默认 300
"""
import os
import subprocess
import sys
from pathlib import Path

from fastmcp import FastMCP

mcp = FastMCP("sak-shell")
MAX_OUTPUT = 20000


def _roots() -> list[Path]:
    raw = os.environ.get("SAK_SHELL_ROOTS") or ""
    items = [p for p in raw.split(os.pathsep) if p.strip()]
    if not items:
        kb = os.environ.get("SAK_KB_ROOT")
        if kb:
            items.append(kb)
        items.append(str(Path(__file__).resolve().parents[2]))
    return [Path(p).expanduser().resolve() for p in items]


def _resolve(path: str, base: Path) -> Path:
    p = Path(path).expanduser()
    if not p.is_absolute():
        p = base / p
    p = p.resolve()
    for root in _roots():
        if p == root or p.is_relative_to(root):
            return p
    raise PermissionError(f"路径不在允许范围内：{p}\n允许的根目录：{[str(r) for r in _roots()]}")


def _clip(text: str) -> str:
    if len(text) <= MAX_OUTPUT:
        return text
    return text[:MAX_OUTPUT] + f"\n…（已截断，原长 {len(text)} 字符）"


@mcp.tool()
def run_python(script: str, args: list[str] | None = None, cwd: str = "",
               timeout: int = 0) -> dict:
    """用当前 Python 解释器执行允许目录内的 .py 脚本，返回退出码与输出。

    参数：
        script: 脚本路径，绝对路径或相对 cwd
        args: 传给脚本的参数数组，如 ["--type", "auto", "报表.docx"]
        cwd: 工作目录，默认知识库根目录
        timeout: 秒，0 表示用 SAK_SHELL_TIMEOUT 或默认 300

    用途：交付门禁 gate.py、台账 append_ledger_entry.py、docx 生成脚本等。
    不执行 shell 命令，不支持管道与重定向。
    """
    base = Path(os.environ.get("SAK_KB_ROOT") or _roots()[0]).resolve()
    workdir = _resolve(cwd, base) if cwd else base
    path = _resolve(script, workdir)

    if not path.is_file():
        return {"status": "error", "message": f"脚本不存在：{path}"}
    if path.suffix.lower() != ".py":
        return {"status": "error", "message": f"只允许执行 .py 脚本：{path}"}
    if not workdir.is_dir():
        return {"status": "error", "message": f"工作目录不存在：{workdir}"}

    limit = int(timeout) or int(os.environ.get("SAK_SHELL_TIMEOUT", "300") or 300)
    env = os.environ.copy()
    env.setdefault("PYTHONUTF8", "1")
    env.setdefault("PYTHONIOENCODING", "utf-8")
    env.setdefault("SAK_KB_ROOT", str(base))

    try:
        proc = subprocess.run(
            [sys.executable, "-X", "utf8", str(path), *(args or [])],
            cwd=str(workdir), env=env, capture_output=True, text=True,
            encoding="utf-8", errors="replace", timeout=limit,
        )
    except subprocess.TimeoutExpired:
        return {"status": "timeout", "script": str(path), "timeout": limit,
                "message": f"执行超过 {limit} 秒被终止"}

    return {
        "status": "ok" if proc.returncode == 0 else "failed",
        "script": str(path),
        "cwd": str(workdir),
        "exit_code": proc.returncode,
        "stdout": _clip(proc.stdout),
        "stderr": _clip(proc.stderr),
    }


if __name__ == "__main__":
    mcp.run()
