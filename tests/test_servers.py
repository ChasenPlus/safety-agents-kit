# -*- coding: utf-8 -*-
"""随包 MCP 服务可用性测试：全部本地启动，不联网、不需要模型与 API Key。

    python tests/test_servers.py

验证的是"干净环境里服务能不能起来、工具在不在、基本读写通不通"，
分支编排逻辑由 test_smoke.py 覆盖。
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))

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


def boot(name: str, script: str, env: dict | None = None) -> StdioMCPClient:
    client = StdioMCPClient(name, sys.executable, [str(REPO / script)], env=env)
    client.start()
    return client


def test_basic_servers(root: str) -> None:
    print("[服务启动与工具清单]")
    expected = {
        "sak-fs": ("server/fs/server.py", {"read_file", "write_file", "list_directory", "search_files"}),
        "sak-excel": ("server/excel/server.py", {"describe_sheets", "read_sheet", "write_sheet"}),
        "sak-shell": ("server/shell/server.py", {"run_python"}),
        "sak-private": ("server/private/server.py", {"check_coverage", "match_regulation"}),
        "sak-rag": ("server/rag/server.py", {"rag_search", "rag_reindex", "rag_index_status"}),
        "sak-crawler": ("server/crawler/server.py", {"search_web", "fetch_url", "check_pdf_file"}),
        "sak-report": ("server/report/server.py", {"generate_rectification_notice", "get_data_schema"}),
        "sak-handoff": ("server/agent_switch/server.py", {"list_branches", "get_branch_prompt"}),
    }
    for name, (script, want) in expected.items():
        try:
            client = boot(name, script, env={"SAK_KB_ROOT": root, "SAK_FS_ROOTS": root})
        except Exception as exc:  # noqa: BLE001
            check(f"{name} 启动", False, f"{type(exc).__name__}: {exc}")
            continue
        try:
            got = {t["name"] for t in client.tools}
            check(f"{name} 启动（{len(got)} 个工具）", True)
            check(f"{name} 工具齐全", want <= got, f"缺 {sorted(want - got)}")
        finally:
            client.close()


def test_fs_roundtrip(root: str) -> None:
    print("[文件读写]")
    client = boot("sak-fs", "server/fs/server.py", env={"SAK_FS_ROOTS": root})
    try:
        wrote = client.call_tool("write_file", {"path": "demo.txt", "content": "第一行\n第二行"})
        check("write_file", "第二行" not in json.dumps(wrote) and "ok" in json.dumps(wrote))
        read = client.call_tool("read_file", {"path": "demo.txt"})
        check("read_file 内容回读", "第二行" in json.dumps(read, ensure_ascii=False))
        listed = client.call_tool("list_directory", {"path": "."})
        check("list_directory", "demo.txt" in json.dumps(listed, ensure_ascii=False))
        outside = client.call_tool("read_file", {"path": str(Path(root).parent / "outside.txt")})
        check("越界路径被拒绝", "不在允许范围" in json.dumps(outside, ensure_ascii=False))
    finally:
        client.close()


def test_excel_roundtrip(root: str) -> None:
    print("[表格读写]")
    client = boot("sak-excel", "server/excel/server.py", env={"SAK_FS_ROOTS": root})
    try:
        client.call_tool("write_sheet", {
            "path": "台账.xlsx", "sheet": "制度台账",
            "rows": [["制度编号", "制度名称"], ["ZD-001", "示例制度"]],
        })
        desc = json.dumps(client.call_tool("describe_sheets", {"path": "台账.xlsx"}), ensure_ascii=False)
        check("describe_sheets 读到新建表", "制度台账" in desc and "制度编号" in desc)
        client.call_tool("write_sheet", {
            "path": "台账.xlsx", "sheet": "制度台账",
            "rows": [["ZD-002", "示例制度二"]], "mode": "append",
        })
        rows = json.dumps(client.call_tool("read_sheet", {"path": "台账.xlsx", "sheet": "制度台账"}),
                          ensure_ascii=False)
        check("append 后共 3 行", "ZD-002" in rows and "示例制度二" in rows)
    finally:
        client.close()


def test_shell(root: str) -> None:
    print("[脚本执行]")
    script = Path(root) / "hello.py"
    script.write_text("import sys\nprint('arg=' + (sys.argv[1] if len(sys.argv) > 1 else ''))\n",
                      encoding="utf-8")
    client = boot("sak-shell", "server/shell/server.py",
                  env={"SAK_KB_ROOT": root, "SAK_SHELL_ROOTS": root})
    try:
        res = json.loads(client.call_tool("run_python", {"script": "hello.py", "args": ["gate"]}))
        check("run_python 执行成功", res["exit_code"] == 0 and "arg=gate" in res["stdout"], res.get("stderr", ""))
        bad = json.dumps(client.call_tool("run_python", {"script": "nope.py"}), ensure_ascii=False)
        check("不存在的脚本报错", "脚本不存在" in bad)
        outside = json.dumps(client.call_tool("run_python", {"script": str(REPO / "sak" / "cli.py")}),
                             ensure_ascii=False)
        check("越界脚本被拒绝", "不在允许范围" in outside)
    finally:
        client.close()


def test_report() -> None:
    print("[整改通知单]")
    client = boot("sak-report", "server/report/server.py")
    try:
        schema = json.dumps(client.call_tool("get_data_schema", {}), ensure_ascii=False)
        check("get_data_schema 返回字段说明", "categories" in schema and "示例数据" in schema)
        missing = json.dumps(client.call_tool("generate_rectification_notice",
                                              {"data_file": "not-exists.json"}), ensure_ascii=False)
        check("缺文件时如实报错", "数据文件不存在" in missing)
    finally:
        client.close()


def main() -> int:
    root = tempfile.mkdtemp(prefix="sak_srv_")
    test_basic_servers(root)
    test_fs_roundtrip(root)
    test_excel_roundtrip(root)
    test_shell(root)
    test_report()
    print(f"\n结果：PASS {ok} / FAIL {fail}")
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main())
