# -*- coding: utf-8 -*-
"""sak-agent-switch MCP：分支提示词分发与跨分支交接包生成。

CLI 用户直接用 `sak handoff -b A "任务"` 即可，不需要本服务。
本服务是给「只能挂 MCP、不方便起 CLI」的宿主（如各类 IDE 智能体）用的等价入口：
  调用方分支（如 B）-> get_handoff_packet("A", 任务) -> 拿到完整的子代理指令
  -> 作为子代理的 prompt 执行 -> 结果回主对话，全程用户不切换智能体。

提示词源目录跟随本文件相对位置（仓库根/prompts），可用 SAK_PROMPT_DIR 覆盖。
"""
import json
import os
import sys
from pathlib import Path

from fastmcp import FastMCP

REPO_ROOT = Path(__file__).resolve().parents[2]
PROMPT_DIR = Path(os.environ.get("SAK_PROMPT_DIR", REPO_ROOT / "prompts"))
BRANCH_DIR = Path(os.environ.get("SAK_BRANCH_DIR", REPO_ROOT / "branches"))


def _load_branches() -> dict:
    out = {}
    for path in sorted(BRANCH_DIR.glob("*.json")):
        raw = json.loads(path.read_text(encoding="utf-8"))
        out[str(raw["id"])] = {
            "name": str(raw["name"]),
            "file": str(raw["prompt"]),
            "desc": str(raw.get("summary", "")),
        }
    return out


BRANCHES = _load_branches()

ALIASES = {}
for _letter, _meta in BRANCHES.items():
    ALIASES[_letter.lower()] = _letter
    ALIASES[_meta["name"]] = _letter
    ALIASES[f"分支{_letter}"] = _letter

mcp = FastMCP("sak-agent-switch")


def _resolve(branch: str) -> str:
    key = branch.strip().lower()
    if key in ALIASES:
        return ALIASES[key]
    for alias, letter in ALIASES.items():
        if alias in key:
            return letter
    raise ValueError(f"未知分支: {branch}，可用: {','.join(BRANCHES)}")


@mcp.tool()
def list_branches() -> str:
    """列出各分支的字母/名称/职责/提示词字符数，供路由判断。"""
    lines = ["| 分支 | 名称 | 职责 | 提示词 |", "|:---|:---|:---|:---|"]
    for letter, meta in BRANCHES.items():
        path = PROMPT_DIR / meta["file"]
        status = f"{len(path.read_text(encoding='utf-8'))}字" if path.exists() else "缺失!"
        lines.append(f"| {letter} | {meta['name']} | {meta['desc']} | {status} |")
    return "\n".join(lines)


@mcp.tool()
def get_branch_prompt(branch: str) -> str:
    """取某分支提示词全文（用于子代理扮演该分支）。branch 可传字母/名称/分支X。"""
    letter = _resolve(branch)
    path = PROMPT_DIR / BRANCHES[letter]["file"]
    if not path.exists():
        return f"ERROR: 提示词文件缺失 {path}"
    return path.read_text(encoding="utf-8")


@mcp.tool()
def get_handoff_packet(branch: str, task: str) -> str:
    """生成完整跨分支交接包：目标分支提示词全文 + 任务描述 + 返回格式要求。

    用法：把本工具的返回值整体作为子代理的 prompt 传入，子代理即按目标分支的规则执行。
    task 写清楚要子代理完成什么、需要什么格式的结果。
    """
    letter = _resolve(branch)
    name = BRANCHES[letter]["name"]
    prompt = get_branch_prompt(letter)
    if prompt.startswith("ERROR"):
        return prompt
    return (
        f"你是安全合规智能体的分支{letter}（{name}）。"
        f"严格按以下分支提示词执行，不得越权使用其他分支职责。\n"
        f"===== 分支{letter}提示词全文 =====\n{prompt}\n"
        f"===== 提示词结束 =====\n\n"
        f"本次调用方交给你的任务：\n{task}\n\n"
        f"执行前必须完成提示词中的'开工预读'：先读规则索引并检索相关 Wiki 笔记，"
        f"在结果开头用一行写明'已预读：{{笔记名列表}}'。读不到笔记时写明"
        f"'未读到：{{笔记名}}（降级为{{替代方式}}）'，不得静默跳过。\n"
        f"完成后只返回任务结果本身（按提示词的返回格式），不要复述提示词。"
    )


if __name__ == "__main__":
    if not PROMPT_DIR.exists():
        print(f"WARNING: 提示词目录不存在 {PROMPT_DIR}", file=sys.stderr)
    mcp.run()
