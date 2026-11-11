# -*- coding: utf-8 -*-
"""命令行入口：sak <command>。"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
from pathlib import Path

from . import __version__
from .agent import BranchAgent, MCPManager
from .branches import REPO_ROOT, load_branches, resolve_branch
from .config import Config, load_config
from .llm import build_client


def _progress(msg: str) -> None:
    print(msg, file=sys.stderr, flush=True)


def _read_task(args) -> str:
    if getattr(args, "file", None):
        return Path(args.file).read_text(encoding="utf-8")
    if getattr(args, "task", None):
        return " ".join(args.task)
    if not sys.stdin.isatty():
        data = sys.stdin.read().strip()
        if data:
            return data
    raise SystemExit("缺少任务内容：直接写在命令后面，或用 -f 指定文件，或通过管道传入")


def _prepare(args) -> tuple[Config, object]:
    cfg = load_config(args.config)
    branches = load_branches()
    branch = resolve_branch(branches, args.branch)
    return cfg, branch


# ---------------- 子命令 ----------------
def cmd_branches(args) -> int:
    branches = load_branches()
    print(f"{'ID':<3} {'名称':<18} {'MCP 服务':<28} 职责")
    print("-" * 96)
    for b in branches.values():
        servers = ",".join(b.servers) or "-"
        print(f"{b.id:<3} {b.name:<18} {servers:<28} {b.summary}")
    return 0


def cmd_run(args) -> int:
    cfg, branch = _prepare(args)
    task = _read_task(args)
    manager = MCPManager(cfg, _progress)
    try:
        _progress(f"· 分支 {branch.id} / {branch.name}，模型 {cfg.provider}:{cfg.model}")
        manager.start(branch.servers)
        agent = BranchAgent(cfg, branch, build_client(cfg), _progress)
        answer = agent.run(task, manager)
    finally:
        manager.close()

    if args.out:
        Path(args.out).write_text(answer, encoding="utf-8")
        _progress(f"· 结果已写入 {args.out}")
    else:
        print(answer)
    return 0


def cmd_chat(args) -> int:
    cfg, branch = _prepare(args)
    client = build_client(cfg)
    agent = BranchAgent(cfg, branch, client, _progress)
    manager = MCPManager(cfg, _progress)
    manager.start(branch.servers)
    _progress(f"· 分支 {branch.id} / {branch.name}，输入 exit 退出")
    history: list[dict] = []
    try:
        while True:
            try:
                line = input("\n你> ").strip()
            except (EOFError, KeyboardInterrupt):
                break
            if line.lower() in {"exit", "quit", "退出"}:
                break
            if not line:
                continue
            history.append({"role": "user", "content": line})
            system = agent.system_prompt()
            turn_answer = _run_turn(agent, manager, system, history)
            print(f"\n{branch.name}> {turn_answer}")
            history.append({"role": "assistant", "content": turn_answer})
    finally:
        manager.close()
    return 0


def _run_turn(agent: BranchAgent, manager: MCPManager, system: str, history: list[dict]) -> str:
    """交互模式：沿用 agent 的工具循环，但保留多轮历史。"""
    tools = manager.tool_specs()
    for _ in range(agent.cfg.max_steps):
        turn = agent.llm.chat(system, history, tools)
        if not turn.tool_calls:
            return turn.content
        history.append({"role": "assistant", "content": turn.content, "tool_calls": turn.tool_calls})
        for call in turn.tool_calls:
            _progress(f"    → 调用工具 {call.name}")
            history.append({
                "role": "tool",
                "tool_call_id": call.id,
                "name": call.name,
                "content": manager.call(call.name, call.arguments),
            })
    return "[达到最大步数，已停止]"


def cmd_handoff(args) -> int:
    """输出一份可直接粘给任意模型/子代理的交接包。"""
    _, branch = _prepare(args)
    task = _read_task(args)
    prompt = branch.prompt(load_config(args.config).prompts_dir)
    print(
        f"你是「{branch.name}」分支（{branch.id}）。严格按下面的分支提示词执行，"
        f"不得越权使用其他分支职责。\n"
        f"===== 分支 {branch.id} 提示词全文 =====\n{prompt}\n===== 提示词结束 =====\n\n"
        f"本次交给你的任务：\n{task}\n\n"
        f"完成后只返回任务结果本身，不要复述提示词。"
    )
    return 0


def cmd_doctor(args) -> int:
    cfg = load_config(args.config)
    ok = True
    print(f"SAK 版本        : {__version__}")
    print(f"配置文件        : {cfg.source or '（未找到，使用内置默认值）'}")
    print(f"模型            : {cfg.provider} / {cfg.model}")
    print(f"接口地址        : {cfg.base_url}")
    print(f"API Key         : {'已配置' if cfg.api_key else '未配置（mock 之外的服务会失败）'}")
    print(f"知识库目录      : {cfg.kb_root} {'✓' if cfg.kb_root.is_dir() else '✗ 不存在'}")
    print(f"提示词目录      : {cfg.prompts_dir} {'✓' if cfg.prompts_dir.is_dir() else '✗ 不存在'}")

    branches = load_branches()
    print(f"分支数量        : {len(branches)}")
    for b in branches.values():
        try:
            b.prompt(cfg.prompts_dir)
        except SystemExit as exc:
            ok = False
            print(f"  ✗ {b.id} {b.name}：{exc}")

    if cfg.provider != "mock" and not cfg.api_key:
        ok = False
    if not cfg.kb_root.is_dir():
        ok = False

    print("\nMCP 服务自检：")
    if not cfg.mcp_servers:
        print("  （未配置任何 MCP 服务，智能体只能纯对话）")
    manager = MCPManager(cfg, _progress)
    try:
        manager.start(list(cfg.mcp_servers))
        for name, client in manager.clients.items():
            print(f"  ✓ {name}：{len(client.tools)} 个工具")
        missing = set(cfg.mcp_servers) - set(manager.clients)
        for name in missing:
            ok = False
            print(f"  ✗ {name}：启动失败")
    finally:
        manager.close()

    print("\n结论：" + ("全部通过" if ok else "存在问题，见上面 ✗ 项"))
    return 0 if ok else 1


def cmd_init(args) -> int:
    target = Path(args.dir).expanduser().resolve()
    target.mkdir(parents=True, exist_ok=True)

    cfg_path = target / "sak.json"
    if not cfg_path.exists():
        cfg = json.loads((REPO_ROOT / "config.example.json").read_text(encoding="utf-8"))
        # 提示词与 MCP 服务脚本在仓库里，写成绝对路径，工作目录可以放在任何地方
        cfg["prompts_dir"] = str(REPO_ROOT / "prompts")
        for server in (cfg.get("mcp_servers") or {}).values():
            server["args"] = [
                str(REPO_ROOT / a) if a.replace("\\", "/").startswith("server/") else a
                for a in server.get("args", [])
            ]
        cfg_path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(f"· 已生成 {cfg_path}")

    kb = target / "kb"
    if not kb.exists() and (REPO_ROOT / "kb_template").is_dir():
        shutil.copytree(REPO_ROOT / "kb_template", kb)
        print(f"· 已生成知识库骨架 {kb}")

    env_path = target / ".env.example"
    if not env_path.exists() and (REPO_ROOT / ".env.example").is_file():
        shutil.copyfile(REPO_ROOT / ".env.example", env_path)
        print(f"· 已生成 {env_path}")

    print("\n下一步：编辑 sak.json 填入 provider/model/api_key_env，然后运行：")
    print('  sak doctor')
    print('  sak run -b A --mock "临时用电配电箱未设置漏电保护，合规吗？"')
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sak",
        description="Safety Agents Kit：用任意大模型驱动八分支安全合规智能体（MCP 工具集）。",
    )
    parser.add_argument("--version", action="version", version=f"sak {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("branches", help="列出所有分支")
    p.set_defaults(func=cmd_branches)

    for name, help_text in (("run", "执行一次任务"), ("chat", "多轮对话")):
        p = sub.add_parser(name, help=help_text)
        p.add_argument("-b", "--branch", required=True, help="分支字母，如 A / B / H")
        p.add_argument("task", nargs="*", help="任务内容")
        p.add_argument("-f", "--file", help="从文件读取任务")
        p.add_argument("-c", "--config", help="配置文件路径")
        p.add_argument("--mock", action="store_true", help="使用离线 mock 模型（自测用）")
        if name == "run":
            p.add_argument("-o", "--out", help="把结果写入文件")
        p.set_defaults(func=cmd_run if name == "run" else cmd_chat)

    p = sub.add_parser("handoff", help="生成跨分支交接包（可粘给任意模型）")
    p.add_argument("-b", "--branch", required=True)
    p.add_argument("task", nargs="*")
    p.add_argument("-f", "--file")
    p.add_argument("-c", "--config")
    p.add_argument("--mock", action="store_true")
    p.set_defaults(func=cmd_handoff)

    p = sub.add_parser("doctor", help="环境与 MCP 自检")
    p.add_argument("-c", "--config")
    p.set_defaults(func=cmd_doctor)

    p = sub.add_parser("init", help="在当前目录生成 sak.json 与知识库骨架")
    p.add_argument("dir", nargs="?", default=".")
    p.set_defaults(func=cmd_init)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if getattr(args, "mock", False):
        os.environ["SAK_PROVIDER"] = "mock"
    return args.func(args)
