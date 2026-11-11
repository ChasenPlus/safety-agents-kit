# -*- coding: utf-8 -*-
"""
sak-rag MCP 服务器：示例电力知识库本地语义检索

工具：
- rag_search：语义检索（可按层级过滤、默认排除定责文件）
- rag_index_status：索引状态
- rag_upsert_file：单文件增量入库（分支 D 下载/转换/OCR 后调用）
- rag_delete_file：单文件移除同步
- rag_reindex：全量/增量重建

配置（env，含默认值）：SAK_KB_ROOT、SAK_DIRECTORY_XLSX、
SAK_RAG_MODEL_DIR、SAK_RAG_INDEX_DIR、SAK_RAG_MODEL_FILE
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from fastmcp import FastMCP  # noqa: E402
import rag_core  # noqa: E402

mcp = FastMCP("sak-rag")

_INDEX = None


def get_index():
    global _INDEX
    if _INDEX is None:
        _INDEX = rag_core.RagIndex()
    return _INDEX


@mcp.tool()
def rag_search(
    query: str,
    layers: list[str] | None = None,
    top_k: int = 5,
    include_responsibility: bool = False,
    mode: str = "hybrid",
    rerank: bool = True,
) -> str:
    """语义检索知识库。返回 JSON 数组字符串，每个元素为一条命中：file/rel_path/layer/clause_hint(最近条款号)/heading/line_range/score/snippet/images(块内图片相对md目录的引用路径数组，空为无图)/is_ocr。
    - layers：可选层级过滤，取值 企业标准/行业标准/国家标准/部委规章/法律法规/清单/公文/Wiki/导航/制度目录；默认全部。
    - include_responsibility：是否包含"定责文件（不引用）"，默认 false；定责文件结果仅用于岗位定位，不作合规依据。
    - mode：检索模式。hybrid（默认，向量+关键词混合）；keyword（纯关键词精确匹配，适合"标准号+条款号"类精确查询）；semantic（纯向量语义，适合概念性问题）。
    - rerank：是否对召回候选做 cross-encoder 精排，默认 true；keyword 模式下自动跳过。
    - 召回后必须按 rel_path 回源文件、按 clause_hint/line_range 取逐字原文，不得直接把 snippet 当作正式引用。
    """
    import json

    if not query or not query.strip():
        raise ValueError("query 不能为空")
    if layers:
        bad = [x for x in layers if x not in rag_core.ALL_LAYERS]
        if bad:
            raise ValueError(f"未知层级: {bad}；可选: {rag_core.ALL_LAYERS}")
    results = get_index().search(
        query.strip(),
        layers=layers or None,
        top_k=max(1, min(top_k, 20)),
        include_responsibility=include_responsibility,
        mode=mode,
        rerank=rerank,
    )
    return json.dumps(results, ensure_ascii=False)


@mcp.tool()
def rag_index_status() -> dict:
    """返回索引状态：各层文件/chunk 数、定责 chunk 数、索引大小、模型文件、最后更新时间。"""
    return get_index().status()


@mcp.tool()
def rag_upsert_file(path: str) -> dict:
    """将知识库内一个 md 文件（或制度目录 xlsx）增量入库；新增/修改文件后调用。返回其库内相对路径。"""
    rel = get_index().upsert_file(path)
    return {"ok": True, "rel_path": rel}


@mcp.tool()
def rag_delete_file(path: str) -> dict:
    """文件从知识库删除后同步移除其索引。path 可为绝对路径或 __directory_xlsx__。"""
    ok = get_index().delete_file(path)
    return {"ok": bool(ok)}


@mcp.tool()
def rag_reindex(full: bool = False) -> dict:
    """扫描知识库同步索引。full=true 为全量重建（清空、压缩空位）；默认增量（按 hash 判定）。"""
    return get_index().reindex(full=full, progress=True)


def _startup_warmup():
    """同步预热（启动多花约 2-3 秒，常驻进程只付一次）：建元数据缓存、
    加载 embedding/reranker 并跑 dummy 推理，消除首个检索请求的冷启动尖峰。
    预热失败不阻断启动，首个请求退回惰性加载。"""
    try:
        if not (Path(rag_core.MODEL_DIR) / "tokenizer.json").is_file():
            print(f"[sak-rag] 未找到向量模型：{rag_core.MODEL_DIR}", file=sys.stderr)
            print("[sak-rag] 请放入 bge-small-zh 模型目录（含 tokenizer.json 与 onnx 权重），"
                  "然后运行 python server/rag/rag_core.py reindex；未放模型时检索不可用。",
                  file=sys.stderr)
            return
        get_index().warmup()
    except Exception as exc:  # noqa: BLE001
        print(f"[sak-rag] 启动预热跳过（首个请求会惰性加载）: {exc}", file=sys.stderr)


if __name__ == "__main__":
    _startup_warmup()
    mcp.run()
