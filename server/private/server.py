# -*- coding: utf-8 -*-
"""
安全合规判断服务 MCP Server

把 coverage_core.py（R11覆盖充分性核验）和 matching_core.py（制度归属匹配）
包装成MCP工具，供各分支调用。

依赖：
    pip install fastmcp
    （可选，用于LLM精化）pip install anthropic

启动方式（stdio）：
    python server/private/server.py

环境变量：
    SAK_DIRECTORY_XLSX   企业标准体系目录.xlsx 的完整路径（match_regulation 需要）
    ANTHROPIC_API_KEY    可选，设置后 check_coverage 会用 LLM 对启发式判断做二次精化
    SAK_JUDGMENT_LOG     可选，判定日志路径，默认写在本目录下
"""

import os
from fastmcp import FastMCP
import coverage_core as cc
import matching_core as mc

mcp = FastMCP("sak-private")

DEFAULT_XLSX = os.environ.get(
    "SAK_DIRECTORY_XLSX",
    os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                 "kb_template", "企业标准体系目录.xlsx"),
)
DEFAULT_LOG = os.environ.get(
    "SAK_JUDGMENT_LOG",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "judgment_log.jsonl"),
)


@mcp.tool()
def check_coverage(question: str, candidate_clause: str) -> dict:
    """R11覆盖充分性核验：判断候选条款是否针对问题实际问的具体要求（指标/频次/
    责任主体/适用范围/操作细节）给出了明确、可执行的规定，而不只是主题相关的泛泛提及。

    参数：
        question: 用户原始问题
        candidate_clause: 候选条款原文（一字不差摘录）

    返回：{"covered": bool, "reason_summary": str}
    - covered=true  → 判定为"已实质覆盖"，可作为正式依据引用
    - covered=false → 判定为"命中但未实质覆盖"，不得作为依据，reason_summary说明具体原因
    """
    return cc.check_coverage(question, candidate_clause, log_path=DEFAULT_LOG)


@mcp.tool()
def match_regulation(question: str) -> dict:
    """判断问题理论上应归属示例电力哪个企业标准/制度编号，用于R06b判断制度归属时调用。
    读取《示例电力企业标准体系目录.xlsx》做关键词匹配，未命中则退化为按"所属大类"推断。

    参数：
        question: 用户原始问题

    返回：{"matches": [{"code","name","confidence","matched_keyword","score"}], "directory_found": bool}
    - 按关键词具体程度加权排序，最多返回前 5 条；命中更多时 note 说明总命中数
    - confidence为"已确认"、"已确认（弱匹配，建议人工核实）"或"推断，建议人工核实所属制度编号"
    - directory_found=false时，note字段会提示目录文件缺失
    """
    return mc.match_regulation(question, DEFAULT_XLSX)


if __name__ == "__main__":
    mcp.run()
