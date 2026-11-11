# -*- coding: utf-8 -*-
"""模型适配器的线格式自测：用本地假 HTTP 服务校验请求/响应格式，不联网。

    python tests/test_llm_wire.py
"""
from __future__ import annotations

import json
import sys
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))

from sak.llm.anthropic import AnthropicClient  # noqa: E402
from sak.llm.base import ToolCall, ToolSpec  # noqa: E402
from sak.llm.gemini import GeminiClient  # noqa: E402
from sak.llm.openai_compat import OpenAICompatClient  # noqa: E402

ok = 0
fail = 0
captured: dict[str, dict] = {}


def check(name: str, cond: bool, detail: str = "") -> None:
    global ok, fail
    if cond:
        ok += 1
        print(f"  PASS  {name}")
    else:
        fail += 1
        print(f"  FAIL  {name} {detail}")


TOOL = ToolSpec(
    name="rag_search",
    description="本地语义检索",
    input_schema={"type": "object", "properties": {"query": {"type": "string"}}, "required": ["query"]},
    server="sak-rag",
)


class Handler(BaseHTTPRequestHandler):
    def do_POST(self) -> None:  # noqa: N802
        length = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(length).decode("utf-8"))
        captured[self.path] = body
        captured[self.path + "#headers"] = {k.lower(): v for k, v in self.headers.items()}

        if "generateContent" in self.path:
            payload = {"candidates": [{"content": {"parts": [{"functionCall": {"name": "rag_search", "args": {"query": "漏电保护"}}}]}}]}
        elif self.path.startswith("/v1/messages"):
            payload = {"content": [
                {"type": "text", "text": "先查一下"},
                {"type": "tool_use", "id": "tu_1", "name": "rag_search", "input": {"query": "漏电保护"}},
            ]}
        else:
            payload = {"choices": [{"message": {
                "content": "先查一下",
                "tool_calls": [{"id": "call_1", "type": "function",
                                "function": {"name": "rag_search", "arguments": "{\"query\": \"漏电保护\"}"}}],
            }}]}

        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, *args) -> None:  # 静默
        pass


def start_server() -> tuple[ThreadingHTTPServer, str]:
    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return server, f"http://127.0.0.1:{server.server_address[1]}"


def test_openai(base: str) -> None:
    print("[1] OpenAI 兼容")
    client = OpenAICompatClient("deepseek-chat", base, "sk-test")
    history = [
        {"role": "user", "content": "查漏电保护"},
        {"role": "assistant", "content": "", "tool_calls": [ToolCall("call_0", "rag_search", {"query": "x"})]},
        {"role": "tool", "tool_call_id": "call_0", "name": "rag_search", "content": "结果"},
    ]
    turn = client.chat("你是分支A", history, [TOOL])
    path = next(p for p in captured if p.endswith("/chat/completions"))
    body = captured[path]
    check("system 放在 messages[0]", body["messages"][0] == {"role": "system", "content": "你是分支A"})
    check("工具按 function 结构下发", body["tools"][0]["function"]["name"] == "rag_search")
    check("工具结果带 tool_call_id", body["messages"][-1]["tool_call_id"] == "call_0")
    check("解析出工具调用", len(turn.tool_calls) == 1 and turn.tool_calls[0].arguments == {"query": "漏电保护"})
    check("保留文本内容", turn.content == "先查一下")
    check("Authorization 头正确", captured[path + "#headers"].get("authorization") == "Bearer sk-test")


def test_anthropic(base: str) -> None:
    print("[2] Anthropic")
    client = AnthropicClient("claude-x", base, "sk-ant")
    history = [
        {"role": "user", "content": "查漏电保护"},
        {"role": "assistant", "content": "", "tool_calls": [ToolCall("tu_0", "rag_search", {"query": "x"})]},
        {"role": "tool", "tool_call_id": "tu_0", "name": "rag_search", "content": "结果"},
    ]
    turn = client.chat("你是分支A", history, [TOOL])
    body = captured["/v1/messages"]
    check("system 是顶层字段", body["system"] == "你是分支A")
    check("工具用 input_schema", body["tools"][0]["input_schema"]["type"] == "object")
    check("assistant 用 tool_use 块", body["messages"][1]["content"][0]["type"] == "tool_use")
    check("工具结果用 tool_result 块", body["messages"][2]["content"][0]["type"] == "tool_result")
    check("解析出工具调用", turn.tool_calls[0].id == "tu_1" and turn.tool_calls[0].name == "rag_search")
    check("api-key 头正确", captured["/v1/messages#headers"].get("x-api-key") == "sk-ant")


def test_gemini(base: str) -> None:
    print("[3] Gemini")
    client = GeminiClient("gemini-x", base, "g-key")
    history = [
        {"role": "user", "content": "查漏电保护"},
        {"role": "assistant", "content": "", "tool_calls": [ToolCall("g0", "rag_search", {"query": "x"})]},
        {"role": "tool", "tool_call_id": "g0", "name": "rag_search", "content": "结果"},
    ]
    turn = client.chat("你是分支A", history, [TOOL])
    path = next(p for p in captured if "generateContent" in p)
    body = captured[path]
    check("systemInstruction 正确", body["systemInstruction"]["parts"][0]["text"] == "你是分支A")
    decl = body["tools"][0]["functionDeclarations"][0]
    check("schema 类型大写", decl["parameters"]["type"] == "OBJECT")
    check("工具结果用 functionResponse", body["contents"][-1]["parts"][0]["functionResponse"]["name"] == "rag_search")
    check("解析出函数调用", turn.tool_calls[0].name == "rag_search" and turn.tool_calls[0].arguments == {"query": "漏电保护"})
    check("key 走 query 参数", "key=g-key" in path)


def main() -> int:
    server, base = start_server()
    try:
        test_openai(base)
        test_anthropic(base)
        test_gemini(base)
    finally:
        server.shutdown()
    print(f"\n结果：PASS {ok} / FAIL {fail}")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
