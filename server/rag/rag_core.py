# -*- coding: utf-8 -*-
"""
示例电力知识库 · 本地 RAG 核心库

能力：
- 文档发现：按层级规则扫描知识库 md（排除"待查文件（不引用）"），制度目录 xlsx 作为特殊数据源
- 结构切块：markdown 标题/条款感知，目标约 320 字/块、硬上限 450 字、50 字重叠；记录行号/标题/最近条款号/OCR 标记
- Embedding：bge-small-zh-v1.5 ONNX（onnxruntime 推理，默认 fp32；设 SAK_RAG_MODEL_FILE=model_quantized.onnx 用量化版）
- 向量存储：SQLite 元数据 + numpy 分片 memmap（每片 50000×512），支持增量 upsert/delete 与空位回收

CLI：
    python rag_core.py status
    python rag_core.py reindex [--full] [--batch 16]
    python rag_core.py upsert --path "<知识库内md绝对路径>"
    python rag_core.py delete --path "<md绝对路径或__directory_xlsx__>"

环境变量（均有默认值）：
    SAK_KB_ROOT / SAK_DIRECTORY_XLSX / SAK_RAG_MODEL_DIR /
    SAK_RAG_INDEX_DIR / SAK_RAG_MODEL_FILE
"""
import argparse
import hashlib
import os
import re
import sqlite3
import sys
import time
from pathlib import Path

import numpy as np

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# ============ 路径与常量 ============
HERE = Path(__file__).parent
DEFAULT_KB_ROOT = Path(__file__).resolve().parents[2] / "kb"
KB_ROOT = Path(os.environ.get("SAK_KB_ROOT", DEFAULT_KB_ROOT))
DIRECTORY_XLSX = Path(
    os.environ.get(
        "SAK_DIRECTORY_XLSX",
        HERE / ".." / ".." / "kb_template" / "制度目录.xlsx",
    )
).resolve()
MODEL_DIR = Path(os.environ.get("SAK_RAG_MODEL_DIR", HERE / "models" / "bge-small-zh-v1.5"))
INDEX_DIR = Path(os.environ.get("SAK_RAG_INDEX_DIR", HERE / "index"))
MODEL_FILE = os.environ.get("SAK_RAG_MODEL_FILE", "model.onnx")
RERANKER_DIR = Path(
    os.environ.get("SAK_RAG_RERANKER_DIR", HERE / "models" / "bge-reranker-base")
)
RERANKER_FILE = os.environ.get("SAK_RAG_RERANKER_FILE", "model_quantized.onnx")
RERANK_TOP_N = int(os.environ.get("SAK_RAG_RERANK_TOP_N", "24"))
# 单次检索最多执行的 FTS 短语数：条款号/英文词在前优先；尾部 trigram 由向量兜底，
# 防止长 query 切成数十个 trigram 后触发上百次 FTS 扫描（实测曾占端到端 30s 中的大头）
FTS_MAX_PHRASES = int(os.environ.get("SAK_RAG_FTS_MAX_PHRASES", "24"))

EMBED_DIM = 512
PART_ROWS = 50000
TARGET_CHARS = 320
HARD_CHARS = 450
OVERLAP_CHARS = 50
# 切块器版本：升级切块逻辑时 +1，reindex 据此强制重建受影响文件（meta 表记录）。
# v2：表格感知切块（管道表格按行自洽切分、列名展开、续行合并）
# v3：图片保真（独占行图片标记不被硬切/overlap 截断、图注与图绑定、chunk 带 images）
CHUNKER_VERSION = "3"
QUERY_INSTRUCTION = "为这个句子生成表示以用于检索相关文章："
DIRECTORY_KEY = "__directory_xlsx__"

# 层级规则（顺序敏感：特殊目录先匹配）。layer=None 表示排除
# 2026-10-10 起知识库改用数字编号目录（00国家标准/01法律法规/02部委规章/03行业标准/
# 04企业标准/05问题分析参考清单/06公文要求/07定责依据/08培训资料），新旧前缀双兼容：
# 旧索引/历史脚本里的无编号路径仍可识别，新文件一律落编号目录。
LAYER_RULES = [
    ("02-原始资料层/07定责依据/", "定责文件", True),
    ("02-原始资料层/定责文件（不引用）/", "定责文件", True),
    ("02-原始资料层/待查文件（不引用）/", None, False),
    ("02-原始资料层/04企业标准/", "企业标准", False),
    ("02-原始资料层/企业标准/", "企业标准", False),
    ("02-原始资料层/03行业标准/", "行业标准", False),
    ("02-原始资料层/行业标准/", "行业标准", False),
    ("02-原始资料层/00国家标准/", "国家标准", False),
    ("02-原始资料层/国家标准/", "国家标准", False),
    ("02-原始资料层/02部委规章/", "部委规章", False),
    ("02-原始资料层/部委规章/", "部委规章", False),
    ("02-原始资料层/01法律法规/", "法律法规", False),
    ("02-原始资料层/法律法规/", "法律法规", False),
    ("02-原始资料层/05问题分析参考清单/", "清单", False),
    ("02-原始资料层/问题分析参考清单/", "清单", False),
    ("02-原始资料层/06公文要求/", "公文", False),
    ("02-原始资料层/其他/", None, False),
    ("01-Wiki层/", "Wiki", False),
    ("00-导航层/", "导航", False),
]
ALL_LAYERS = ["企业标准", "行业标准", "国家标准", "部委规章", "法律法规", "清单", "公文", "Wiki", "导航", "制度目录"]


def classify(rel_posix: str):
    """按相对路径返回 (layer, is_responsibility)；layer=None 表示排除"""
    for prefix, layer, is_resp in LAYER_RULES:
        if rel_posix.startswith(prefix):
            return layer, is_resp
    return None, False


# ============ 文本清洗与切块 ============
_WIKILINK_RE = re.compile(r"\[\[([^\]|]+)(?:\|([^\]]+))?\]\]")
_HEADING_RE = re.compile(r"^(#{1,6})\s*(.+)")
_CLAUSE_CN_RE = re.compile(r"第\s*\d+(?:\.\d+)*\s*条")
_CLAUSE_NUM_RE = re.compile(r"^[\s\-*（(]*(\d{1,2}(?:\.\d+){1,5})(?![\d.])")
# 独占一行的 markdown 图片：![](./xxx_images/p001_01.png)；图注以"图1/图 2/图A"开头
_IMG_LINE_RE = re.compile(r"^!\[[^\]]*\]\(([^)]+)\)\s*$")
_IMG_CAP_RE = re.compile(r"^图\s*[0-9０-９A-Za-z一二三四五六七八九十]")
_IMG_MD_RE = re.compile(r"!\[[^\]]*\]\([^)]+\)")
_IMG_MD_PATH_RE = re.compile(r"!\[[^\]]*\]\(([^)]+)\)")


def _strip_img_md(s: str) -> str:
    """embedding/FTS/snippet 用：剥图片路径标记，保留同单元图注文字。"""
    return _IMG_MD_RE.sub("", s or "")


_SKIP_DIRS = {".obsidian", ".git", "node_modules", ".trash"}

# 标准号提取：GB/T 50057、DL/T 853、Q/BJP 2141、NB/T 31026、AQ 1008 等
# 注意必须前缀+号码成对提取——只比数字会把 GBT 853-1988 错配成 DL/T 853（库内真件）
_STD_PREFIXES = [
    "GB/T", "GB_T", "GBZ", "GB",
    "DL/T", "DL_T", "DLT", "DL",
    "NB/T", "NB_T", "NBT", "NB",
    "Q/BJP", "Q_BJP", "QBJP", "Q/BJH", "Q_BJH", "QBJH",
    "JGJ/T", "JGJT", "JGJ", "CJJ/T", "CJJ",
    "JT/T", "JTJ", "JT", "TSG", "CECS", "AQ",
    "HG/T", "HG", "SH/T", "SH", "MT/T", "MT",
    "YS/T", "YS", "YB/T", "YB", "SL/T", "SL",
    "HJ/T", "HJ", "GA/T", "GA",
]
_STD_REF_RE = re.compile(
    r"(" + "|".join(re.escape(p) for p in _STD_PREFIXES) + r")"
    r"[\s\-—_／/]*(\d{2,6})(?:\.(\d+))?(?:\s*[-—]\s*(\d{4}))?",
    re.IGNORECASE,
)


def _std_pfx_norm(s: str) -> str:
    """前缀归一：DL/T、DL_T、DLT、DL T 全部 → DLT。"""
    return re.sub(r"[\s_／/\\]+", "", (s or "").upper())


def extract_standard_refs(query):
    """提取标准号引用，返回 [(前缀归一, 数字, 分部分号, 年份)]，如
    'DL/T 853-2004' → [('DLT','853','','2004')]，
    'GB/T 150.1' → [('GBT','150','1','')]。"""
    out, seen = [], set()
    for m in _STD_REF_RE.finditer(query or ""):
        ref = (_std_pfx_norm(m.group(1)), m.group(2), m.group(3) or "", m.group(4) or "")
        if ref not in seen:
            seen.add(ref)
            out.append(ref)
    return out


def extract_standard_numbers(query):
    """从 query 提取标准号数字主体列表（向后兼容旧调用）。"""
    return [r[1] for r in extract_standard_refs(query)]


def _std_name_norm(path_or_name: str) -> str:
    """文件名/路径归一：大写、全角斜杠/破折号统一、去空白与 _ / \\\\。
    库内实测命名：GB_T150.1-2024、DL_T 317—2025、GBT 39001、DLT 1364。"""
    s = (path_or_name or "").upper().replace("／", "/").replace("—", "-")
    return re.sub(r"[\s_/\\]+", "", s)


def _std_ref_match(ref, norm_name: str):
    """归一文件名中是否含该标准号。返回 'year'（含指明年份）/ 'core'（仅前缀+号码）
    / None。分部分号规则：query 指定 .1 则文件必须带 .1（150.1≠150.2）；
    query 未指定时接受同系列任一部分（查 GB/T 150 能命中 150.1~150.4）。"""
    pfx, num, suffix, year = ref
    if suffix:
        pat = r"(?<![A-Z0-9])" + re.escape(pfx + num) + r"\." + suffix + r"(?!\d)"
    else:
        pat = r"(?<![A-Z0-9])" + re.escape(pfx + num) + r"(?:\.\d+)?(?!\d)"
    if not re.search(pat, norm_name):
        return None
    if year and re.search(r"(?<!\d)" + year + r"(?!\d)", norm_name):
        return "year"
    return "core"


def clean_line(line: str) -> str:
    s = _WIKILINK_RE.sub(lambda m: m.group(2) or m.group(1), line)
    s = re.sub(r"[*`>~]", "", s)
    s = re.sub(r"^#{1,6}\s*", "", s)
    s = s.replace("|", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _split_long_unit(s: str, limit: int = HARD_CHARS):
    """超长单行（常见于 xlsx 拍平行）按句末标点累积切分，仍超长则硬切。
    避免单块远超 HARD_CHARS 后被 embedding(512token)/reranker(384字) 截断尾部。"""
    if len(s) <= limit:
        return [s]
    parts = re.split(r"(?<=[。；;！？])", s)
    out, cur = [], ""
    for p in parts:
        if not p:
            continue
        while len(p) > limit:
            if cur:
                out.append(cur)
                cur = ""
            out.append(p[:limit])
            p = p[limit:]
        if len(cur) + len(p) <= limit:
            cur += p
        else:
            if cur:
                out.append(cur)
            cur = p
    if cur:
        out.append(cur)
    return out


# ============ 管道表格解析（v2 表格感知切块） ============
_TBL_SEP_RE = re.compile(r"^\s*\|?[\s:|-]*-{2,}[\s:|-]*\|?\s*$")


def _clean_cell(s: str) -> str:
    """表格单元格清洗：wikilink 取显示文本，去 markdown 装饰，空白归一。
    保留 | 的处理由调用方先 split 完成。"""
    s = _WIKILINK_RE.sub(lambda m: m.group(2) or m.group(1), s)
    s = re.sub(r"[*`>~]", "", s)
    return re.sub(r"\s+", " ", s).strip()


def _split_pipe_cells(stripped: str):
    """拆一行管道表格为单元格列表（去掉首尾边界管道符）。
    [[目标|显示]] wikilink 内部的 | 不作为列分隔（真实库表格单元格里
    大量嵌标准号 wikilink，朴素 split 会把单元格炸成带残括号的碎片）。"""
    s = stripped.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    raw_cells, buf = [], []
    i = 0
    n = len(s)
    while i < n:
        if s[i:i + 2] == "[[":
            j = s.find("]]", i + 2)
            if j != -1:
                buf.append(s[i:j + 2])
                i = j + 2
                continue
        if s[i] == "|":
            raw_cells.append("".join(buf))
            buf = []
        else:
            buf.append(s[i])
        i += 1
    raw_cells.append("".join(buf))
    return [_clean_cell(x) for x in raw_cells]


def _looks_pipe_row(stripped: str) -> bool:
    """该行是否为管道表格行（≥2 个非边界单元格）；wikilink 内部的 | 不计。"""
    return "|" in stripped and len(_split_pipe_cells(stripped)) >= 2


def _split_row_pieces(text: str, lead_cell: str, limit: int = HARD_CHARS):
    """超长表格行按句末标点切片；续片带首列实体前缀，保证每片语义自洽。"""
    if len(text) <= limit:
        return [text]
    parts = _split_long_unit(text, limit)
    if len(parts) > 1 and lead_cell:
        lead = lead_cell[:16]
        parts = [parts[0]] + [f"{lead}（续） {p}" for p in parts[1:]]
    return parts


def _render_table_rows(data, headers, reserve=0):
    """把表格数据行渲染为 [(start_line, end_line, [piece,...])]。
    data=[(lineno, rawline)]（lineno 为 1-based）；headers 为表头单元格列表或 None。
    reserve：切块时块首固定前缀（caption/列名行）占用的字符预算，
    超长行切片上限相应缩小，保证最终块不超 HARD_CHARS。
    - 宽表（≥4 个非空列名）：每格展开"列名：值"，跳过空格；
    - 窄表/无表头：只拼非空值；
    - 首列空且其余非空列都在第 2 列及以后 → 上一行的续行，文本并入前行。
    """
    named = [h for h in (headers or []) if h]
    wide = len(named) >= 4
    limit = max(120, HARD_CHARS - reserve)
    out = []
    for ln, raw in data:
        cells = _split_pipe_cells(raw)
        nonempty = [k for k, c in enumerate(cells) if c]
        if not nonempty:
            continue
        # 续行判定：首列与第二列均空，且内容都在第 3 列及以后。
        # cqs 双依据排版：|  |  |  |  |  | 《第二依据》… |（第二列空）；
        # 巡查汇总表的空序号独立记录：|  | 十堰热电 | 2021年 | … |（第二列有值），
        # 不得误并入前行
        is_cont = (not cells[0] and (len(cells) < 2 or not cells[1])
                   and all(k > 1 for k in nonempty))
        if wide:
            segs = []
            for k in nonempty:
                h = headers[k] if k < len(headers) else ""
                segs.append(f"{h}：{cells[k]}" if h else cells[k])
            text = " ".join(segs)
        else:
            text = " ".join(cells[k] for k in nonempty)
        pieces = _split_row_pieces(text, cells[0], limit)
        if is_cont and out:
            sln, eln, pp = out[-1]
            # 并入首片后立即重切（必须在拼入后续片之前，否则会切错位置）；
            # 旧实现先 extend 再判断 pp[-1]，超长合并片沉到列表中部永不被切，
            # 曾在真实库累积出 5000+ 字的单块
            merged = (pp[-1] + " " + pieces[0]).strip() if pp else pieces[0]
            pp = pp[:-1] + _split_long_unit(merged, limit) + pieces[1:]
            out[-1] = (sln, ln, pp)
        else:
            out.append((ln, ln, pieces))
    return out


def parse_and_chunk(text: str):
    """结构感知切块 v2（表格感知）。返回
    [{start_line,end_line,heading,clause,body,is_ocr}]

    管道表格特殊处理（库内 1311 个文件含 3242 张规范管道表）：
    - 旧版 clean_line 把 | 换成空格，多行表格被揉成普通长块，单行答案
      （如"绝缘手套 工频耐压试验 1次/半年"）被相邻行稀释并受 bm25 长度惩罚；
    - v2 识别"表头行 + --- 分隔行 + 连续数据行"，宽表（≥4 个有列名的列）
      每行渲染为"列名：值"自洽单元，窄表保留紧凑值序列并在块首带一次列名；
    - 首列为空的行视为上一行的续行（cqs 一条数据挂两个依据的排版）；
    - 表格块与正文之间硬边界、不做 overlap；超长行按句切分并带首列实体前缀。
    """
    is_ocr = bool(re.search(r"OCR", "\n".join(text.splitlines()[:30])))
    lines = text.splitlines()
    n = len(lines)

    # ---- 行分类：H 标题 / I 图片行 / S 表格分隔行 / T 管道行 / P 普通行 / B 空行 ----
    kinds = ["B"] * n
    cells_of = {}
    for i, raw in enumerate(lines):
        st = raw.strip()
        if not st:
            continue
        if _HEADING_RE.match(st):
            kinds[i] = "H"
            continue
        if _IMG_LINE_RE.match(st):
            kinds[i] = "I"
            continue
        if "|" in st and _TBL_SEP_RE.match(st):
            kinds[i] = "S"
            continue
        if "|" in st:
            cells = _split_pipe_cells(st)
            if len(cells) >= 2:
                kinds[i] = "T"
                cells_of[i] = cells
                continue
        kinds[i] = "P"

    # ---- 表格区间：belong[i] = 所属 span 序号（-1 为无）----
    spans = []  # (head_i或-1, data_start, data_end, headers或None)
    belong = [-1] * n
    for i, k in enumerate(kinds):
        if k == "S" and i > 0 and kinds[i - 1] == "T" and i + 1 < n:
            j = i + 1
            while j < n and kinds[j] == "T":
                j += 1
            if j > i + 1:  # 至少 1 行数据才算表
                sid = len(spans)
                spans.append((i - 1, i + 1, j, cells_of[i - 1]))
                for r in range(i - 1, j):
                    belong[r] = sid
    # 无分隔行的 ≥3 连续管道行兜底视为无表头表格（库内 213 个文件为此形态）
    i = 0
    while i < n:
        if kinds[i] == "T" and belong[i] < 0:
            j = i
            while j < n and kinds[j] == "T" and belong[j] < 0:
                j += 1
            if j - i >= 3:
                sid = len(spans)
                spans.append((-1, i, j, None))
                for r in range(i, j):
                    belong[r] = sid
            i = j
        else:
            i += 1

    # ---- 每个 span 预渲染：caption 检测 + 数据行渲染/续行合并/超长切片 ----
    span_ctx = {}
    for sid, (head, d0, d1, headers) in enumerate(spans):
        caption = ""
        cap_line = -1
        if head >= 0:
            # caption 可紧贴表头，或间隔 1 个空行
            for back in (head - 1, head - 2):
                if back >= 0 and kinds[back] == "P":
                    if back == head - 2 and kinds[head - 1] != "B":
                        continue
                    cap = clean_line(lines[back])
                    if len(cap) <= 40 and re.search(
                        r"表\s*[0-9０-９一二三四五六七八九十A-Za-z]", cap
                    ):
                        caption, cap_line = cap, back
                    break
        named = [h for h in (headers or []) if h]
        wide = len(named) >= 4
        header_line = " ".join(named)
        # OCR 拍碎的表格偶发超长"表头单元格"（实测 2600 字的数字垃圾），
        # 作为每块前缀会把所有块顶破硬限；异常即不用列名前缀
        if len(header_line) > 80:
            header_line = ""
        # 前缀字符预算：宽表首块带 caption（≤40+1）；窄表每块带列名行；
        # 窄表首块若还有 caption，超限则在组装时丢弃 caption（见 tflush）
        reserve = 48 if wide else (len(header_line) + 1 if header_line else 0)
        rows = _render_table_rows(
            [(r + 1, lines[r]) for r in range(d0, d1)], headers, reserve=reserve
        )
        span_ctx[sid] = {
            "caption": caption,
            "caption_line": cap_line if caption else -1,
            "header_line": header_line,
            "wide": wide,
            "reserve": reserve,
            "rows": rows,
        }

    # ---- 图片行预扫：绑定图注（紧邻或隔 1 空行、≤40字、"图1…"开头）----
    # img_caps[i] = (图注文本, 图注行0基)；cap_lines 内的 P 行不再单独产出
    img_caps = {}
    cap_lines = set()
    for i, k in enumerate(kinds):
        if k != "I":
            continue
        j = i + 1
        while j < n and kinds[j] == "B":
            j += 1
        if j - i > 2:
            continue
        if j < n and kinds[j] == "P":
            cap = clean_line(lines[j])
            if cap and len(cap) <= 40 and _IMG_CAP_RE.match(cap):
                img_caps[i] = (cap, j)
                cap_lines.add(j)

    # ---- 有序产出单元：heading / prose / image / table ----
    units = []
    stack = []
    heading_path = ""
    cur_clause = ""
    for i, raw in enumerate(lines):
        sid = belong[i]
        if sid >= 0:
            head, d0, d1, _ = spans[sid]
            if i == d0:
                units.append({
                    "kind": "table",
                    "heading": heading_path,
                    "clause": cur_clause,
                    "ctx": span_ctx[sid],
                })
            continue
        st = raw.strip()
        if kinds[i] == "H":
            m = _HEADING_RE.match(st)
            level = len(m.group(1))
            title = clean_line(m.group(2))
            while stack and stack[-1][0] >= level:
                stack.pop()
            stack.append((level, title))
            heading_path = " / ".join(t for _, t in stack)
            cur_clause = ""
            if title:
                units.append({
                    "kind": "h", "line": i + 1, "text": title,
                    "heading": heading_path, "clause": cur_clause,
                })
            continue
        if kinds[i] == "I":
            m = _IMG_LINE_RE.match(st)
            cap_info = img_caps.get(i)
            cap, cap_i = cap_info if cap_info else ("", -1)
            text = st + (" " + cap if cap else "")
            units.append({
                "kind": "image", "line": i + 1, "text": text,
                "heading": heading_path, "clause": cur_clause,
                "path": m.group(1).strip(),
                "end_line": (cap_i + 1) if cap_i >= 0 else i + 1,
            })
            continue
        if kinds[i] != "P" or i in cap_lines:
            continue
        clean = clean_line(raw)
        if not clean:
            continue
        if _CLAUSE_CN_RE.search(clean):
            cur_clause = _CLAUSE_CN_RE.search(clean).group(0).replace(" ", "")
        else:
            mm = _CLAUSE_NUM_RE.match(clean)
            if mm:
                cur_clause = mm.group(1)
        for piece in _split_long_unit(clean):
            units.append({
                "kind": "p", "line": i + 1, "text": piece,
                "heading": heading_path, "clause": cur_clause,
            })

    # ---- 组装：正文流（带 overlap）+ 表格块（硬边界、无 overlap）----
    chunks = []
    pbuf = []
    pbuf_len = 0

    def flush_prose():
        nonlocal pbuf, pbuf_len
        if not pbuf:
            return
        body = " ".join(u["text"] for u in pbuf)
        start_line = pbuf[0]["line"]
        end_line = max(u.get("end_line", u["line"]) for u in pbuf)
        chunks.append({
            "start_line": start_line,
            "end_line": end_line,
            "heading": pbuf[-1]["heading"],
            "clause": next((u["clause"] for u in reversed(pbuf) if u["clause"]), ""),
            "body": body,
            "images": [u["path"] for u in pbuf if u["kind"] == "image"],
            "is_ocr": is_ocr,
        })
        # overlap 只取自文字部分：图片标记是原子单元，绝不能被 50 字尾切断，
        # 也不重复带进下一块（图随当前块走）
        prose_tail_src = " ".join(u["text"] for u in pbuf if u["kind"] != "image")
        tail = prose_tail_src[-OVERLAP_CHARS:]
        cut = tail.find(" ")
        if cut > 0:
            tail = tail[cut + 1:]
        if tail:
            pbuf = [{
                "kind": "p", "line": end_line, "text": tail,
                "heading": pbuf[-1]["heading"], "clause": pbuf[-1]["clause"],
            }]
            pbuf_len = len(tail)
        else:
            pbuf = []
            pbuf_len = 0

    def emit_table(u):
        nonlocal pbuf, pbuf_len
        ctx = u["ctx"]
        # caption 行刚作为普通行进了 pbuf：弹出避免与表内首块重复
        if ctx["caption_line"] >= 0 and pbuf and pbuf[-1]["text"] == ctx["caption"]:
            pbuf_len -= len(pbuf[-1]["text"])
            pbuf.pop()
        tbuf = []
        tlen = 0
        first = True

        def tflush():
            nonlocal tbuf, tlen, first
            if not tbuf:
                return
            prefix = []
            if first and ctx["caption"]:
                prefix.append(ctx["caption"])
            # 紧凑窄表每块带一次列名；宽表行内已自带"列名：值"
            if not ctx["wide"] and ctx["header_line"]:
                prefix.append(ctx["header_line"])
            body = " ".join(prefix + [t for _, _, t in tbuf])
            # 窄表首块 caption + 列名双前缀顶破硬限时，丢弃 caption 保行内容
            if len(body) > HARD_CHARS and first and ctx["caption"]:
                prefix = prefix[1:]
                body = " ".join(prefix + [t for _, _, t in tbuf])
            chunks.append({
                "start_line": tbuf[0][0],
                "end_line": tbuf[-1][1],
                "heading": u["heading"],
                "clause": u["clause"],
                "body": body,
                "images": [],
                "is_ocr": is_ocr,
            })
            first = False
            tbuf = []
            tlen = 0

        reserve = ctx["reserve"]
        for row_i, (sln, eln, pieces) in enumerate(ctx["rows"]):
            # 宽表每行已自含列名，逐行成块（跨行聚合会重新稀释答案密度、
            # 重新引入 bm25 长度惩罚）；仅窄表紧凑行聚合到 TARGET_CHARS
            if ctx["wide"] and tbuf:
                tflush()
            for pc in pieces:
                # 宽表行可能切成多片：append 前先判硬限，否则 300+402 会拼出 700+ 块；
                # 窄表按 TARGET 聚合
                if tbuf and reserve + tlen + len(pc) > (
                    TARGET_CHARS if not ctx["wide"] else HARD_CHARS
                ):
                    tflush()
                tbuf.append((sln, eln, pc))
                tlen += len(pc)
                if reserve + tlen >= HARD_CHARS:
                    tflush()
        tflush()

    for u in units:
        if u["kind"] == "table":
            flush_prose()
            emit_table(u)
            continue
        if u["kind"] == "image":
            # 图片是原子单元：不参与字符预算，不被硬切，随当前正文块走
            pbuf.append(u)
            continue
        if u["kind"] == "h" and pbuf:
            flush_prose()
        # append 前先判限：旧逻辑 append 后才判，50 字 overlap 尾+小单元+大片
        # 曾拼出 700+ 字段落块（v1 既有缺陷），现封顶 HARD+OVERLAP；
        # len(pbuf) 预留 join 空格
        nsep = sum(1 for x in pbuf if x["kind"] != "image")
        if pbuf and pbuf_len + len(u["text"]) + nsep > HARD_CHARS + OVERLAP_CHARS:
            flush_prose()
        pbuf.append(u)
        pbuf_len += len(u["text"])
        if pbuf_len >= TARGET_CHARS:
            flush_prose()
    flush_prose()
    return chunks


# ============ FTS5 关键词查询构造（trigram，无需中文分词） ============
# 仅保留纯功能虚词；关/把/对/行/地/次/间/按/据 等业务核心字不得停用
_FTS_STOP = set(
    "的了吗呢啊呀是在和与或不没为也都就还又有什么怎哪应该可以能会要须需"
    "得着之其此即些个"
)


def _norm(s):
    """FTS 规范化：去全部空白。标准 md 常含排版空格（金 属 导 体），
    不规范化则 trigram 全部失效；chunks 表原文保持不动。"""
    return re.sub(r"\s+", "", s or "")


def _norm_fts(chunk: dict) -> str:
    """FTS 索引文本：标题路径+条款号+正文一起入索引。
    正文未重复标题/条款号时（常见于标准条文），关键词只出现在标题也能命中。
    图片路径标记不进 FTS（英文路径 trigram 是纯噪声），图注文字保留。"""
    return _norm(" ".join([
        chunk.get("heading") or "",
        chunk.get("clause") or "",
        _strip_img_md(chunk["body"]),
    ]))


def _fts_add_seg(seg, phrases):
    n = len(seg)
    if n < 3:
        return
    if n <= 6:
        # 短片段整体保留（3-6 字：一个或两个 trigram）
        phrases.append(seg)
        return
    # 长片段：相邻 3gram 全枚举，保住自然词边界
    # （如"绝缘胶垫试验周期" -> 绝缘胶/缘胶垫/胶垫试...，含"绝缘胶垫"的 chunk 必命中）
    for i in range(n - 2):
        phrases.append(seg[i : i + 3])


_CN_DIGITS = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5",
              "六": "6", "七": "7", "八": "8", "九": "9", "〇": "0", "零": "0"}


def _digit_variants(p, max_variants=8):
    """含中文数字的短语补充阿拉伯数字变体（如"少于五根"->"少于5根"）。
    不全局替换，仅在短语内枚举数字位组合，避免误伤"一般/统一"等业务字。
    数字位过多时组合数 2^k 爆炸（每个变体都触发 FTS 扫描），上限截断。"""
    positions = [i for i, c in enumerate(p) if c in _CN_DIGITS]
    if not positions:
        return [p]
    out = [p]
    chars = list(p)
    for r in range(1, len(positions) + 1):
        from itertools import combinations

        for combo in combinations(positions, r):
            c2 = chars[:]
            for i in combo:
                c2[i] = _CN_DIGITS[c2[i]]
            v = "".join(c2)
            if v not in out:
                out.append(v)
            if len(out) >= max_variants:
                return out
    return out


def build_fts_phrases(query):
    """把自然提问切成候选短语列表（停用字切碎，长段滑窗）。"""
    phrases = []
    # 条款号/章节号原样保留为短语（如 4.2.2、6.1.8），供精确命中
    phrases.extend(re.findall(r"\d+(?:\.\d+){1,4}", query))
    for seg in re.findall(r"[一-鿿A-Za-z0-9]+", query):
        if re.fullmatch(r"[A-Za-z0-9]+", seg):
            if len(seg) >= 3:
                phrases.append(seg)
            continue
        buf = ""
        for ch in seg:
            if ch in _FTS_STOP:
                _fts_add_seg(buf, phrases)
                buf = ""
            else:
                buf += ch
        _fts_add_seg(buf, phrases)
    # 去重保序
    seen = set()
    out = []
    for p in phrases:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out


# 语料 30.5 万块时 df≤500 ≈ 0.16%，视为"区分性实体"
_ENTITY_DF_MAX = 500


def extract_key_entities(query, conn, filter_sql="", params=(), max_check=16):
    """提取查询中的语料级稀有实体，用于精排后的"区分性词覆盖"保护。

    背景：bge-reranker-base 对"只共享泛词"的块会虚高（query=绝缘手套试验周期，
    含"绝缘电阻表/预防性试验宜为三年"的变压器块得分 2.8，高于含
    "绝缘手套 工频耐压试验 1次/半年"的真答案块 1.0）。含查询稀有实体的块
    必须排在不含的块之前。4-gram 在 trigram 索引上直接可查 df，无需分词。

    实体分两级：
    - strong：停用字边界切出的完整短片段（3-6字），通常是查询主体名词
      （"绝缘手套"在"的"之前成段；"预防性试验周期"是长段=提问意图，不算）；
    - weak：长段（>6字）内滑出的稀有 4-gram。
    覆盖排序以 strong 命中数优先、weak 次之，避免泛词碎片淹没主体实体。
    返回 (strong, weak)；均空表示不干预排序。
    """
    runs = []
    for seg in re.findall(r"[一-鿿A-Za-z0-9]+", query):
        if re.fullmatch(r"[A-Za-z0-9]+", seg):
            if len(seg) >= 3:
                runs.append(seg)
            continue
        buf = ""
        for ch in seg:
            if ch in _FTS_STOP:
                if len(buf) >= 3:
                    runs.append(buf)
                buf = ""
            else:
                buf += ch
        if len(buf) >= 3:
            runs.append(buf)

    def _rare(grams):
        kept = []
        sql = (
            "SELECT COUNT(*) FROM chunks_fts"
            " JOIN chunks c ON c.id=chunks_fts.rowid"
            " JOIN files f ON f.id=c.file_id"
            + filter_sql + " AND chunks_fts MATCH ?"
        )
        for g in grams:
            try:
                df = conn.execute(sql, list(params) + ['"%s"' % g]).fetchone()[0]
            except sqlite3.OperationalError:
                continue
            if 1 <= df <= _ENTITY_DF_MAX:
                kept.append(g)
        return kept

    strong_cands, weak_cands = [], []
    seen_s, seen_w = set(), set()
    for run in runs:
        if 3 <= len(run) <= 6 and run not in seen_s:
            seen_s.add(run)
            strong_cands.append(run)
        if len(run) > 6:
            for i in range(len(run) - 3):
                g = run[i : i + 4]
                if g not in seen_w:
                    seen_w.add(g)
                    weak_cands.append(g)
        if len(strong_cands) + len(weak_cands) >= max_check:
            break
    strong = _rare(strong_cands[:8])
    weak = _rare(weak_cands[:max_check])
    if not strong and not weak:
        # 退回稀有 3-gram（query 全为长段且无稀有 4-gram 时的最后兜底）
        tri, seen3 = [], set()
        for run in runs:
            for i in range(len(run) - 2):
                g = run[i : i + 3]
                if g not in seen3:
                    seen3.add(g)
                    tri.append(g)
        weak = _rare(tri[:8])
    return strong, weak


# ============ Embedding ============
class Embedder:
    _instance = None

    def __init__(self):
        import onnxruntime as ort
        from tokenizers import Tokenizer

        self.tok = Tokenizer.from_file(str(MODEL_DIR / "tokenizer.json"))
        self.tok.enable_padding(pad_id=0, pad_token="[PAD]")
        self.tok.enable_truncation(max_length=512)
        so = ort.SessionOptions()
        so.intra_op_num_threads = max(1, (os.cpu_count() or 2) - 1)
        model_path = MODEL_DIR / MODEL_FILE
        if not model_path.exists():
            raise FileNotFoundError(f"embedding 模型不存在: {model_path}")
        self.sess = ort.InferenceSession(str(model_path), so, providers=["CPUExecutionProvider"])

    @classmethod
    def get(cls):
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def embed(self, texts, batch=16, is_query=False):
        if is_query:
            texts = [QUERY_INSTRUCTION + t for t in texts]
        out = []
        for i in range(0, len(texts), batch):
            batch_texts = texts[i : i + batch]
            enc = self.tok.encode_batch(batch_texts)
            ids = np.asarray([e.ids for e in enc], dtype=np.int64)
            mask = np.asarray([e.attention_mask for e in enc], dtype=np.int64)
            tt = np.zeros_like(ids)
            hidden = self.sess.run(
                None,
                {"input_ids": ids, "attention_mask": mask, "token_type_ids": tt},
            )[0]
            vec = hidden[:, 0].astype(np.float32)
            vec = vec / np.maximum(np.linalg.norm(vec, axis=1, keepdims=True), 1e-12)
            out.append(vec)
        if not out:
            return np.zeros((0, EMBED_DIM), dtype=np.float32)
        return np.concatenate(out, axis=0)


# ============ Cross-Encoder 精排（bge-reranker-base） ============
class Reranker:
    """对 (query, passage) 对打相关性分数。模型缺失/加载失败时优雅降级为无精排。"""

    _instance = None
    _unavailable = False

    def __init__(self):
        import onnxruntime as ort
        from tokenizers import Tokenizer

        self.tok = Tokenizer.from_file(str(RERANKER_DIR / "tokenizer.json"))
        self.tok.enable_padding(pad_id=0, pad_token="[PAD]")
        self.tok.enable_truncation(max_length=512)
        so = ort.SessionOptions()
        so.intra_op_num_threads = max(1, (os.cpu_count() or 2) - 1)
        model_path = RERANKER_DIR / RERANKER_FILE
        if not model_path.exists():
            raise FileNotFoundError(f"reranker 模型不存在: {model_path}")
        self.sess = ort.InferenceSession(str(model_path), so, providers=["CPUExecutionProvider"])

    @classmethod
    def get(cls):
        if cls._unavailable:
            return None
        if cls._instance is None:
            try:
                cls._instance = cls()
            except Exception as e:
                print(f"[警告] reranker 加载失败，退回无精排: {e}", file=sys.stderr)
                cls._unavailable = True
                return None
        return cls._instance

    def score(self, query, texts, batch=4, max_passage=384):
        """对 (query, text) 对逐个打分，返回 logits 数组（越大越相关）。
        max_passage: passage 截断字数，与 query(128字) 合计在 512token 预算内；
        batch=4 为本机实测最优（bge-reranker-base CPU：24条 0.9s；
        batch=8 需 2.6s、batch=16 需 12.6s，大 batch 内存带宽劣化）。"""
        out = []
        q = query[:128]
        for i in range(0, len(texts), batch):
            pairs = [(q, (t or "")[:max_passage]) for t in texts[i : i + batch]]
            enc = self.tok.encode_batch(pairs)
            ids = np.asarray([e.ids for e in enc], dtype=np.int64)
            mask = np.asarray([e.attention_mask for e in enc], dtype=np.int64)
            logits = self.sess.run(
                None,
                {"input_ids": ids, "attention_mask": mask},
            )[0]
            out.append(logits.reshape(-1).astype(np.float32))
        if not out:
            return np.zeros(0, dtype=np.float32)
        return np.concatenate(out, axis=0)


# ============ 分片向量存储 ============
class VectorStore:
    def __init__(self, index_dir: Path):
        self.dir = index_dir
        self.dir.mkdir(parents=True, exist_ok=True)
        self._parts = {}

    def _part_path(self, pid):
        return self.dir / f"vectors_part{pid:04d}.npy"

    def _part(self, pid):
        if pid not in self._parts:
            path = self._part_path(pid)
            if not path.exists():
                np.lib.format.open_memmap(
                    path, mode="w+", dtype=np.float32, shape=(PART_ROWS, EMBED_DIM)
                )
            self._parts[pid] = np.load(path, mmap_mode="r+")
        return self._parts[pid]

    def allocate(self, conn, n):
        positions = []
        if n:
            rows = conn.execute(
                "SELECT pos FROM free_positions ORDER BY pos LIMIT ?", (n,)
            ).fetchall()
            positions = [r[0] for r in rows]
            if positions:
                conn.executemany(
                    "DELETE FROM free_positions WHERE pos=?", [(p,) for p in positions]
                )
        while len(positions) < n:
            next_pos = int(
                conn.execute("SELECT value FROM meta WHERE key='next_pos'").fetchone()[0]
            )
            positions.append(next_pos)
            conn.execute(
                "UPDATE meta SET value=? WHERE key='next_pos'", (str(next_pos + 1),)
            )
        return positions

    def put(self, positions, vectors):
        # 按分片分组整片花式写入（一次 NumPy 赋值替代逐位置 Python 循环）
        pos = np.asarray(positions, dtype=np.int64)
        vec = np.asarray(vectors, dtype=np.float32)
        part_ids = pos // PART_ROWS
        offsets = pos % PART_ROWS
        for pid in np.unique(part_ids):
            mask = part_ids == pid
            self._part(int(pid))[offsets[mask]] = vec[mask]
        for m in self._parts.values():
            m.flush()

    def close(self):
        """关闭所有 memmap 句柄。Windows 上删除/重建 .npy 前必须调用，
        否则 unlink 触发 PermissionError（文件仍被映射占用）。"""
        for m in self._parts.values():
            try:
                mm = getattr(m, "_mmap", None)
                if mm is not None:
                    mm.close()
            except Exception:
                pass
        self._parts.clear()

    def get(self, positions):
        """按 positions 取向量。按 part 分组批量花式索引，避免逐行 Python 循环。"""
        n = len(positions)
        if n == 0:
            return np.zeros((0, EMBED_DIM), dtype=np.float32)
        result = np.empty((n, EMBED_DIM), dtype=np.float32)
        pos_arr = np.asarray(positions, dtype=np.int64)
        parts = pos_arr // PART_ROWS
        offs = pos_arr % PART_ROWS
        for pid in np.unique(parts):
            mask = parts == pid
            result[mask] = self._part(int(pid))[offs[mask]]
        return result

    def score_part(self, pid, qvec):
        """对单个分片做连续 GEMV，返回 (PART_ROWS,) 相似度。
        直接在 memmap 上顺序读计算，不物化整片副本（单片 102MB），
        比 get(全量positions) 拼 700MB 大数组再矩阵乘快且省内存。"""
        return self._part(pid) @ qvec

    def free(self, conn, positions):
        conn.executemany(
            "INSERT OR IGNORE INTO free_positions(pos) VALUES(?)", [(int(p),) for p in positions]
        )


# ============ 索引器 ============
SCHEMA = """
CREATE TABLE IF NOT EXISTS files(
    id INTEGER PRIMARY KEY, path TEXT UNIQUE, layer TEXT, is_resp INTEGER DEFAULT 0,
    hash TEXT, mtime REAL, size INTEGER, n_chunks INTEGER DEFAULT 0, updated REAL
);
CREATE TABLE IF NOT EXISTS chunks(
    id INTEGER PRIMARY KEY, file_id INTEGER, idx INTEGER, start_line INTEGER, end_line INTEGER,
    heading TEXT, clause TEXT, is_ocr INTEGER DEFAULT 0, text TEXT, pos INTEGER
);
CREATE TABLE IF NOT EXISTS free_positions(pos INTEGER PRIMARY KEY);
CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY, value TEXT);
CREATE INDEX IF NOT EXISTS idx_chunks_file ON chunks(file_id);
CREATE VIRTUAL TABLE IF NOT EXISTS chunks_fts USING fts5(text, tokenize='trigram');
"""


class RagIndex:
    def __init__(self):
        INDEX_DIR.mkdir(parents=True, exist_ok=True)
        self.db_path = INDEX_DIR / "rag_index.sqlite"
        self.conn = sqlite3.connect(str(self.db_path))
        self.conn.executescript(SCHEMA)
        # 旧库平滑升级：补 size 列（mtime+size 预检可免全库 md5）
        have_cols = {r[1] for r in self.conn.execute("PRAGMA table_info(files)")}
        if "size" not in have_cols:
            self.conn.execute("ALTER TABLE files ADD COLUMN size INTEGER")
        self.conn.execute("INSERT OR IGNORE INTO meta(key,value) VALUES('next_pos','0')")
        self.conn.commit()
        self.vstore = VectorStore(INDEX_DIR)
        # 进程内全量元数据缓存：search 每次都要把 30万+ chunks 元数据拉成
        # Python 行（实测 650ms+），而 MCP 是常驻进程、绝大多数请求只读。
        # 失效不靠各写入路径手动清缓存，而是读 meta.meta_rev 修订号——
        # 任何改变检索结果的提交都递增它，跨进程（CLI reindex）同样生效。
        self._rows_cache = None
        self._cache_rev = None
        # 过滤签名 -> 已构造行列表（随主缓存一起按修订号失效）
        self._filter_rows = {}

    def close(self):
        self.conn.commit()
        self.conn.close()
        self.vstore.close()

    # ---- 元数据缓存（search 热路径）----
    # 主缓存固定 c.id 序（无 WHERE 旧 SQL 的实测计划顺序）；末列追加 f.id 供
    # _rows_for 处理 files 驱动计划的 (file_id,c.id) 序，下游不取末列。
    _ROWS_SQL = (
        "SELECT c.id,c.pos,c.start_line,c.end_line,c.heading,c.clause,c.is_ocr,"
        "f.path,f.layer,f.is_resp,f.id FROM chunks c JOIN files f ON c.file_id=f.id"
        " ORDER BY c.id"
    )

    def _bump_meta_rev(self):
        """检索可见数据发生变化的事务里调用（随该事务一起提交）。"""
        self.conn.execute(
            "INSERT OR REPLACE INTO meta(key,value) VALUES('meta_rev',"
            "COALESCE((SELECT CAST(value AS INTEGER)+1 FROM meta "
            "WHERE key='meta_rev'),1))"
        )

    def _meta_rev(self):
        row = self.conn.execute(
            "SELECT value FROM meta WHERE key='meta_rev'"
        ).fetchone()
        return int(row[0]) if row is not None else None

    def _ensure_rows_cache(self):
        """返回全量元组行缓存（c.id 序）；修订号变化时重建。
        path/heading/clause 做字符串池：path 全库仅约 3800 个不同值、heading
        同标题下所有块重复，朴素 fetchall 每行新建 str 约 200MB，归一后约 1/4；
        字符串同一性不影响 ==/拼接/显示，语义零变化。"""
        rev = self._meta_rev()
        if self._rows_cache is None or rev != self._cache_rev:
            hpool, cpool, ppool = {}, {}, {}
            rows = []
            for r in self.conn.execute(self._ROWS_SQL):
                h = r[4]
                if h is not None:
                    h = hpool.setdefault(h, h)
                c = r[5]
                if c is not None:
                    c = cpool.setdefault(c, c)
                p = ppool.setdefault(r[7], r[7])
                rows.append((r[0], r[1], r[2], r[3], h, c, r[6], p, r[8], r[9], r[10]))
            self._rows_cache = rows
            self._cache_rev = rev
            self._filter_rows = {}
        return self._rows_cache

    @staticmethod
    def _filter_clause(layers, include_responsibility):
        where, params = [], []
        if layers:
            where.append("f.layer IN (%s)" % ",".join("?" for _ in layers))
            params.extend(layers)
        if not include_responsibility:
            where.append("f.is_resp=0")
        filter_sql = (" WHERE " + " AND ".join(where)) if where else ""
        return where, params, filter_sql

    # 与旧版 search 完全相同的宽列查询（无 ORDER BY）：不能改成只查 c.id 的
    # 窄查询——窄查询命中 idx_chunks_file 覆盖索引会换连接计划、行序不同。
    _LEGACY_WIDE_SQL = (
        "SELECT c.id,c.pos,c.start_line,c.end_line,c.heading,c.clause,c.is_ocr,"
        "f.path,f.layer,f.is_resp FROM chunks c JOIN files f ON c.file_id=f.id"
    )

    def _rows_for(self, layers, include_responsibility):
        """返回某过滤条件下的元组行，行序与旧版"无 ORDER BY 宽列 SQL"逐位一致。
        每个过滤签名进程生命周期只取一次真实行序（大集合约 0.6s），之后缓存
        命中近乎零成本；行映射回主缓存元组，字符串对象不重复物化。"""
        master = self._ensure_rows_cache()
        key = (tuple(layers) if layers else None, bool(include_responsibility))
        cached = self._filter_rows.get(key)
        if cached is not None:
            return cached
        if layers is None:
            # 无 layers：旧宽列计划实测走 chunks PK（c.id 序），与 master 相同
            sub = list(master) if include_responsibility else [r for r in master if r[9] == 0]
            self._filter_rows[key] = sub
            return sub
        _, params, filter_sql = self._filter_clause(layers, include_responsibility)
        # 用旧版同款宽列 SQL 拿真实行序，随即只保留 id 序、映射回主缓存元组
        legacy_ids = [
            r[0]
            for r in self.conn.execute(self._LEGACY_WIDE_SQL + filter_sql, params)
        ]
        id2master = {r[0]: r for r in master}
        sub = [id2master[cid] for cid in legacy_ids]
        self._filter_rows[key] = sub
        return sub

    # 启动预热时预建的过滤签名：E 分支每个工作包固定走这些层；
    # 上位两层/四层首次建序各约 0.6s（宽列 JOIN 30 万行），其余几十 ms。
    _WARMUP_LAYER_SIGS = [
        ["企业标准"],
        ["清单"],
        ["制度目录"],
        ["行业标准", "国家标准"],
        ["行业标准", "国家标准", "部委规章", "法律法规"],
    ]

    def warmup(self):
        """MCP 启动时同步预热（常驻进程只付一次，约 3-5 秒）：
        1) 主元数据缓存 + 常用过滤签名（省每次 30 万行 fetchall/建序）
        2) embedding/reranker 加载并 dummy 推理（省首问约 1 秒冷推理）
        3) 向量分片触页（省首次 GEMV 的 memmap 缺页）
        任一步失败不阻断启动。"""
        self._ensure_rows_cache()
        for sig in self._WARMUP_LAYER_SIGS:
            self._rows_for(sig, False)
        emb = Embedder.get()
        qvec = emb.embed(["预热"], is_query=True)[0]
        for p in sorted(INDEX_DIR.glob("vectors_part*.npy")):
            pid = int(p.stem.split("part")[-1])
            self.vstore.score_part(pid, qvec)
        reranker = Reranker.get()
        if reranker is not None:
            reranker.score("预热", ["预热用的条文片段"])

    # ---- 文档发现 ----
    def discover(self):
        """廉价扫描：只 stat 不读文件，返回
        {rel_posix: (abspath, layer, is_resp, mtime, size)}。
        md5 推迟到 reindex 确认 mtime/size 有变时才算——全库扫描从
        "读一遍所有 md（OCR 大 md 单个数 MB）"降为纯元数据遍历。"""
        found = {}
        for root, dirs, files in os.walk(KB_ROOT):
            dirs[:] = [d for d in dirs if d not in _SKIP_DIRS]
            for name in files:
                if not name.lower().endswith(".md") or name.startswith("~$"):
                    continue
                abspath = Path(root) / name
                rel = abspath.relative_to(KB_ROOT).as_posix()
                layer, is_resp = classify(rel)
                if layer is None:
                    continue
                st = abspath.stat()
                found[rel] = (str(abspath), layer, 1 if is_resp else 0,
                              st.st_mtime, st.st_size)
        # 制度目录 xlsx 特殊数据源
        if DIRECTORY_XLSX.exists():
            st = DIRECTORY_XLSX.stat()
            rel = f"{DIRECTORY_KEY}/{DIRECTORY_XLSX.name}"
            found[rel] = (str(DIRECTORY_XLSX), "制度目录", 0, st.st_mtime, st.st_size)
        # 定责文件层去重：同名岗位清单在"07定责依据/根目录"与
        # "07定责依据/…/全员安全岗位责任制2026/"下各存一份（275 组，正文 273 组完全相同），
        # 全量索引会让同一岗位在检索结果里占多个名额。只保留路径最短的一份，
        # 其余不进索引（docx/md 原文件不动）；同名但内容确实不同的版本会失去索引，
        # 需要时把差异版本改名再入库。
        best = {}
        for rel, (_, layer, _r, _m, _s) in found.items():
            if layer != "定责文件":
                continue
            key = rel.rsplit("/", 1)[-1]
            if key not in best or rel.count("/") < best[key].count("/"):
                best[key] = rel
        keep = set(best.values())
        found = {rel: v for rel, v in found.items()
                 if v[1] != "定责文件" or rel in keep}
        return found

    @staticmethod
    def _file_sig(abspath: Path):
        h = hashlib.md5(abspath.read_bytes()).hexdigest()
        return h, abspath.stat().st_mtime

    @staticmethod
    def _read_text(abspath: str):
        if abspath.lower().endswith(".xlsx"):
            return RagIndex._xlsx_to_text(abspath)
        return Path(abspath).read_text(encoding="utf-8", errors="replace")

    @staticmethod
    def _xlsx_to_text(abspath: str):
        import openpyxl

        wb = openpyxl.load_workbook(abspath, read_only=True, data_only=True)
        lines = []
        for sheet in wb.sheetnames:
            if sheet.strip() in ("使用说明", "说明", "README"):
                continue
            for row in wb[sheet].iter_rows(values_only=True):
                cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
                if cells:
                    lines.append(f"{sheet} " + " ".join(cells))
        wb.close()
        return "\n".join(lines)

    # ---- 单文件同步 ----
    def upsert_file(self, abspath: str, batch=16):
        abspath = Path(abspath)
        if not abspath.exists():
            raise FileNotFoundError(abspath)
        if abspath.suffix.lower() == ".xlsx":
            rel = f"{DIRECTORY_KEY}/{abspath.name}"
            layer, is_resp = "制度目录", False
        else:
            rel = abspath.relative_to(KB_ROOT).as_posix()
            layer, is_resp = classify(rel)
            if layer is None:
                raise ValueError(f"路径不在索引范围（可能是待查文件）: {rel}")
        h, mtime = self._file_sig(abspath)
        size = abspath.stat().st_size
        self._upsert(rel, str(abspath), layer, 1 if is_resp else 0,
                     mtime, size, h=h, batch=batch)
        return rel

    def _upsert(self, rel, abspath, layer, is_resp, mtime, size, h=None, batch=16):
        if h is None:
            h = hashlib.md5(Path(abspath).read_bytes()).hexdigest()
        text = self._read_text(abspath)
        chunked = parse_and_chunk(text)
        # 同文件内 (标题,条款,正文) 完全相同的块只保留首次出现：
        # OCR/转换残留的重复页眉水印、重复表格行会产生相同 embedding，
        # 去重后检索结果更多样，也省向量存储与精排配额
        seen, deduped = set(), []
        for c in chunked:
            key = (c["heading"], c["clause"], c["body"])
            if key in seen:
                continue
            seen.add(key)
            deduped.append(c)
        chunked = deduped
        # 删除旧 chunks 并回收向量位
        row = self.conn.execute("SELECT id FROM files WHERE path=?", (rel,)).fetchone()
        file_id = row[0] if row else None
        if file_id is not None:
            old_pos = [
                r[0]
                for r in self.conn.execute(
                    "SELECT pos FROM chunks WHERE file_id=?", (file_id,)
                ).fetchall()
                if r[0] is not None
            ]
            if old_pos:
                self.vstore.free(self.conn, old_pos)
            self.conn.execute(
                "DELETE FROM chunks_fts WHERE rowid IN"
                " (SELECT id FROM chunks WHERE file_id=?)", (file_id,)
            )
            self.conn.execute("DELETE FROM chunks WHERE file_id=?", (file_id,))
        else:
            cur = self.conn.execute("SELECT MAX(id)+1 FROM files").fetchone()[0]
            file_id = cur or 1
        self.conn.execute(
            "INSERT OR REPLACE INTO files(id,path,layer,is_resp,hash,mtime,size,n_chunks,updated)"
            " VALUES(?,?,?,?,?,?,?,?,?)",
            (file_id, rel, layer, is_resp, h, mtime, int(size), len(chunked), time.time()),
        )
        if chunked:
            embedder = Embedder.get()
            # 向量输入剥掉图片路径标记（图注文字保留，对检索有价值），body 原文入库
            vectors = embedder.embed(
                [_strip_img_md(c["body"]) for c in chunked], batch=batch
            )
            positions = self.vstore.allocate(self.conn, len(chunked))
            self.vstore.put(np.asarray(positions), vectors)
            # 显式分配 chunk id，chunks 与 FTS 各一条 executemany，
            # 替代旧版每块 2 次 execute 的 Python 往返（大文件上千块时差几十倍）
            next_id = self.conn.execute("SELECT MAX(id)+1 FROM chunks").fetchone()[0] or 1
            ids = list(range(next_id, next_id + len(chunked)))
            chunk_rows = [
                (
                    cid, file_id, idx, c["start_line"], c["end_line"],
                    c["heading"], c["clause"], 1 if c["is_ocr"] else 0,
                    c["body"], int(pos),
                )
                for idx, (c, cid, pos) in enumerate(zip(chunked, ids, positions))
            ]
            self.conn.executemany(
                "INSERT INTO chunks(id,file_id,idx,start_line,end_line,heading,clause,"
                "is_ocr,text,pos) VALUES(?,?,?,?,?,?,?,?,?,?)",
                chunk_rows,
            )
            self.conn.executemany(
                "INSERT INTO chunks_fts(rowid,text) VALUES(?,?)",
                [(cid, _norm_fts(c)) for (cid, c) in zip(ids, chunked)],
            )
        self._bump_meta_rev()
        self.conn.commit()

    def delete_file(self, path: str):
        """path 可为知识库内 md 绝对路径、rel 路径或 __directory_xlsx__"""
        if path == DIRECTORY_KEY or path.startswith(DIRECTORY_KEY + "/"):
            rel = path if "/" in path else None
            if rel is None:
                row = self.conn.execute(
                    "SELECT path FROM files WHERE layer='制度目录' LIMIT 1"
                ).fetchone()
                if not row:
                    return False
                rel = row[0]
        else:
            p = Path(path)
            try:
                rel = p.relative_to(KB_ROOT).as_posix()
            except ValueError:
                rel = path
        return self._delete_rel(rel)

    def _delete_rel(self, rel):
        row = self.conn.execute("SELECT id FROM files WHERE path=?", (rel,)).fetchone()
        if not row:
            return False
        file_id = row[0]
        old_pos = [
            r[0]
            for r in self.conn.execute(
                "SELECT pos FROM chunks WHERE file_id=?", (file_id,)
            ).fetchall()
            if r[0] is not None
        ]
        if old_pos:
            self.vstore.free(self.conn, old_pos)
        self.conn.execute(
            "DELETE FROM chunks_fts WHERE rowid IN"
            " (SELECT id FROM chunks WHERE file_id=?)", (file_id,)
        )
        self.conn.execute("DELETE FROM chunks WHERE file_id=?", (file_id,))
        self.conn.execute("DELETE FROM files WHERE id=?", (file_id,))
        self._bump_meta_rev()
        self.conn.commit()
        return True

    @staticmethod
    def _has_pipe_table(abspath: str) -> bool:
        """流式快速判定 md 是否含管道表格：规范表（表头+---分隔行+数据行）
        或 ≥3 连续管道行。命中即停，供切块器版本迁移只重建受影响文件。"""
        if not str(abspath).lower().endswith(".md"):
            return False
        consec = 0
        prev_pipe = False
        try:
            with open(abspath, encoding="utf-8", errors="replace") as fh:
                for raw in fh:
                    st = raw.strip()
                    if not st:
                        consec = 0
                        prev_pipe = False
                        continue
                    is_sep = ("|" in st and _TBL_SEP_RE.match(st) is not None)
                    if is_sep and prev_pipe:
                        nxt = fh.readline().strip()
                        if nxt and _looks_pipe_row(nxt):
                            return True
                        # 分隔行后无数据：当作普通管道行计数处理
                        consec = 0
                        prev_pipe = False
                        continue
                    is_pipe = _looks_pipe_row(st)
                    if is_pipe:
                        consec += 1
                        if consec >= 3:
                            return True
                        prev_pipe = True
                    else:
                        consec = 0
                        prev_pipe = False
        except OSError:
            return False
        return False

    @staticmethod
    def _has_image_ref(abspath: str) -> bool:
        """流式快速判定 md 是否含独占行图片标记。供 v2→v3 切块器迁移
        只重建含图文件（旧版把图片语法当普通行，可能被硬切/overlap 截断）。"""
        if not str(abspath).lower().endswith(".md"):
            return False
        try:
            with open(abspath, encoding="utf-8", errors="replace") as fh:
                for raw in fh:
                    if _IMG_LINE_RE.match(raw.strip()):
                        return True
        except OSError:
            return False
        return False

    # ---- 全量/增量重建 ----
    def reindex(self, full=False, batch=16, progress=True):
        import gc

        if full:
            self.conn.close()
            # Windows 必须先关 memmap 句柄再删分片，否则 unlink 抛 PermissionError
            self.vstore.close()
            gc.collect()
            for p in INDEX_DIR.glob("*"):
                if p.is_file():
                    p.unlink()
            self.conn = sqlite3.connect(str(self.db_path))
            self.conn.executescript(SCHEMA)
            self.conn.execute("INSERT OR IGNORE INTO meta(key,value) VALUES('next_pos','0')")
            self.conn.commit()
            self.vstore = VectorStore(INDEX_DIR)

        found = self.discover()
        # 切块器版本迁移：旧库用 v1 切块（表格被揉进普通块），v2 只重建含表格文件
        ver = self.conn.execute(
            "SELECT value FROM meta WHERE key='chunker_version'"
        ).fetchone()
        stale_chunker = full or ver is None or ver[0] != CHUNKER_VERSION
        # (hash, layer, mtime, size)；旧库 size 可能为 NULL → 该文件走 md5 兜底
        db_files = {
            r[0]: (r[1], r[2], r[3], r[4])
            for r in self.conn.execute("SELECT path,hash,layer,mtime,size FROM files")
        }

        added = changed = deleted = skipped = 0
        t0 = time.time()
        # 版本迁移期间每个文件的"受切块器版本影响"判定（管道表 or 图片引用）
        rebuild_cache = {}

        def needs_rebuild(p):
            if p not in rebuild_cache:
                rebuild_cache[p] = (
                    self._has_pipe_table(p) or self._has_image_ref(p)
                )
            return rebuild_cache[p]

        keys = sorted(found.keys())
        for n, rel in enumerate(keys, 1):
            abspath, layer, is_resp, mtime, size = found[rel]
            try:
                old = db_files.get(rel)
                if old is None:
                    self._upsert(rel, abspath, layer, is_resp, mtime, size, batch=batch)
                    added += 1
                elif (
                    old[3] is not None and old[2] == mtime and old[3] == size
                    and not (stale_chunker and needs_rebuild(abspath))
                ):
                    # mtime+size 均未变：跳过且不读文件、不算 md5（增量扫描主要耗时来源）；
                    # 切块器版本过旧时，含表格/图片文件即使未变也要强制重建
                    skipped += 1
                else:
                    # mtime/size 缺失（旧库）或已变：算 md5 最终确认，避免 touch 触发无意义重建
                    h = hashlib.md5(Path(abspath).read_bytes()).hexdigest()
                    if old[0] == h and not (
                        stale_chunker and str(abspath).lower().endswith(".md")
                        and needs_rebuild(abspath)
                    ):
                        # 内容相同：仅刷新 stat 缓存，否则每次扫描都会重算 md5
                        if old[2] != mtime or old[3] != size:
                            self.conn.execute(
                                "UPDATE files SET mtime=?, size=? WHERE path=?",
                                (mtime, int(size), rel),
                            )
                            self.conn.commit()
                        skipped += 1
                    else:
                        self._upsert(rel, abspath, layer, is_resp, mtime, size, h=h, batch=batch)
                        changed += 1
            except (FileNotFoundError, OSError) as ex:
                # 扫描快照与实际文件之间存在并发移动/删除（用户正在整理目录）时，
                # 跳过该文件继续，不让整轮 reindex 崩溃；下轮增量自动补齐/清理。
                print(f"  [跳过] 文件在扫描后消失: {rel} ({ex.__class__.__name__})",
                      file=sys.stderr)
                continue
            if progress and n % 20 == 0:
                rate = n / max(time.time() - t0, 0.1)
                eta = (len(keys) - n) / max(rate, 0.1)
                print(
                    f"  进度 {n}/{len(keys)}  新增{added} 更新{changed} 跳过{skipped}"
                    f"  {rate:.1f}文件/秒 ETA {eta/60:.1f}分钟",
                    flush=True, file=sys.stderr,
                )
        # 磁盘已删除的文件同步移除（制度目录源缺失时保留并警告）
        for rel in list(db_files.keys()):
            if rel in found:
                continue
            if rel.startswith(DIRECTORY_KEY):
                print(f"  [警告] 制度目录 xlsx 不存在，保留旧索引: {rel}", file=sys.stderr)
                continue
            self._delete_rel(rel)
            deleted += 1
        elapsed = time.time() - t0
        self.conn.execute(
            "INSERT OR REPLACE INTO meta(key,value) VALUES('chunker_version',?)",
            (CHUNKER_VERSION,),
        )
        # full 重建后 meta_rev 行已随库删除，这里补递增；全部跳过的增量
        # 场景无任何检索可见变化，也顺带刷一个号，代价可忽略
        self._bump_meta_rev()
        self.conn.commit()
        print(
            f"完成：新增 {added}，更新 {changed}，跳过 {skipped}，删除 {deleted}，"
            f"耗时 {elapsed/60:.1f} 分钟（切块器 v{CHUNKER_VERSION}）", file=sys.stderr,
        )
        return {"added": added, "changed": changed, "skipped": skipped, "deleted": deleted,
                "elapsed_sec": elapsed}

    # ---- 状态 ----
    def status(self):
        rows = self.conn.execute(
            "SELECT layer, COUNT(*), COALESCE(SUM(n_chunks),0), COALESCE(SUM(is_resp),0)"
            " FROM files GROUP BY layer ORDER BY layer"
        ).fetchall()
        layers = [
            {"layer": r[0], "files": r[1], "chunks": r[2], "is_responsibility_files": r[3]}
            for r in rows
        ]
        total_files = sum(l["files"] for l in layers)
        total_chunks = sum(l["chunks"] for l in layers)
        resp_chunks = self.conn.execute(
            "SELECT COUNT(*) FROM chunks c JOIN files f ON c.file_id=f.id WHERE f.is_resp=1"
        ).fetchone()[0]
        updated = self.conn.execute("SELECT MAX(updated) FROM files").fetchone()[0]
        size_mb = 0.0
        for p in INDEX_DIR.glob("vectors_part*.npy"):
            size_mb += p.stat().st_size
        return {
            "kb_root": str(KB_ROOT),
            "index_dir": str(INDEX_DIR),
            "model_file": MODEL_FILE,
            "total_files": total_files,
            "total_chunks": total_chunks,
            "responsibility_chunks": resp_chunks,
            "layers": layers,
            "index_size_mb": round(size_mb / 1024 / 1024, 1),
            "last_updated": updated,
        }

    # ---- FTS 关键词索引回填 ----
    def ensure_fts(self, progress=True, force=False):
        """chunks_fts 与 chunks 行数不一致时重建 FTS 索引（全量回填）。
        force=True 时行数一致也强制重建（用于 FTS 文本格式升级，如标题/条款入索引）。"""
        n_chunks = self.conn.execute("SELECT COUNT(*) FROM chunks").fetchone()[0]
        n_fts = self.conn.execute("SELECT COUNT(*) FROM chunks_fts").fetchone()[0]
        if n_fts == n_chunks and not force:
            return {"chunks": n_chunks, "rebuilt": False}
        t0 = time.time()
        self.conn.execute("DELETE FROM chunks_fts")
        rows = [
            (cid, _norm(" ".join([h or "", cl or "", t or ""])))
            for cid, h, cl, t in self.conn.execute(
                "SELECT id,heading,clause,text FROM chunks ORDER BY id"
            ).fetchall()
        ]
        B = 5000
        for i in range(0, len(rows), B):
            self.conn.executemany(
                "INSERT INTO chunks_fts(rowid,text) VALUES(?,?)", rows[i : i + B]
            )
            self.conn.commit()
            if progress:
                print(f"  FTS 回填 {min(i+B,len(rows))}/{len(rows)}", file=sys.stderr)
        return {"chunks": n_chunks, "rebuilt": True,
                "elapsed_sec": round(time.time() - t0, 1)}

    # ---- 检索（向量 + FTS 关键词混合，RRF 融合；可选 cross-encoder 精排） ----
    def search(self, query, layers=None, top_k=5, include_responsibility=False, batch=16,
               mode="hybrid", rerank=True):
        if mode not in ("hybrid", "keyword", "semantic"):
            raise ValueError(f"未知 mode: {mode}（可选 hybrid/keyword/semantic）")
        where, params, filter_sql = self._filter_clause(layers, include_responsibility)
        # 主查询不含 c.text（358k 条 ×320 字≈115MB，按需取即可）
        # 行索引：0=id 1=pos 2=start_line 3=end_line 4=heading 5=clause 6=is_ocr
        #         7=path 8=layer 9=is_resp 10=file_id（仅缓存排序用）
        # 元组行进进程缓存（meta_rev 驱动失效），行序与旧无 ORDER BY SQL 逐位一致；
        # filter_sql/params 仍供下面 FTS、实体召回、标准号加权等 SQL 使用。
        rows = self._rows_for(layers, include_responsibility)
        if not rows:
            return []
        id2row = {rows[i][0]: i for i in range(len(rows))}

        def _fetch_texts(cids):
            """按需取一批 chunk 的 text，返回 {cid: text}。"""
            if not cids:
                return {}
            placeholders = ",".join("?" for _ in cids)
            return dict(
                self.conn.execute(
                    f"SELECT id, text FROM chunks WHERE id IN ({placeholders})",
                    list(cids),
                ).fetchall()
            )

        rrf = {}
        fts_phrases = []  # FTS 通道填充，强实体强制召回复用做覆盖打分
        scores = None
        # 标准号引用提前解析：FTS 通道要用它抑制标准号数字碎片（"T85"/"853"会经 trigram
        # 跨前缀命中 GBT853 的标题），精准加权通道在其后
        refs = extract_standard_refs(query)
        std_nums = [r[1] for r in refs]
        fts_suppress = set()
        for pfx, num, _suffix, year in refs:
            targets = [num] + ([year] if year else [])
            for ph in build_fts_phrases(_norm(query)):
                if not re.search(r"\d", ph):
                    continue
                digits = "".join(re.findall(r"\d+", ph))
                letters = re.sub(r"[^A-Z]", "", ph)
                # 1) 含号码≥2位连续数字子串的 trigram 碎片（T85/853/53带/T150）抑制；
                # 2) 字母部分与完整前缀一致的（DLT853）保留；
                # 3) 无字母的年份碎片抑制
                is_prefixed = letters == pfx and digits in num
                if is_prefixed:
                    continue
                if (len(digits) >= 2 and any(digits in t for t in targets)) or ph == num:
                    fts_suppress.add(ph)
        # 向量候选：分片在 memmap 上直接打分，不再物化全语料矩阵（35.8万×512≈700MB）
        if mode in ("hybrid", "semantic"):
            positions = np.asarray([r[1] for r in rows], dtype=np.int64)
            qvec = Embedder.get().embed([query], batch=1, is_query=True)[0]
            n_rows = len(rows)
            cand_scores = np.empty(n_rows, dtype=np.float32)
            total_chunks = self.conn.execute("SELECT COUNT(*) FROM chunks").fetchone()[0]
            if n_rows >= total_chunks:
                # 无层过滤（全语料）：按 pos 排序，逐分片连续 GEMV + 映射回收分
                order_pos = np.argsort(positions, kind="stable")
                sorted_pos = positions[order_pos]
                max_pid = int(sorted_pos[-1] // PART_ROWS)
                edges = np.searchsorted(
                    sorted_pos, [p * PART_ROWS for p in range(max_pid + 2)]
                )
                for pid in range(max_pid + 1):
                    lo, hi = int(edges[pid]), int(edges[pid + 1])
                    if hi <= lo:
                        continue
                    part_scores = self.vstore.score_part(pid, qvec)
                    local = sorted_pos[lo:hi] - pid * PART_ROWS
                    cand_scores[order_pos[lo:hi]] = part_scores[local]
            else:
                # 有层过滤：分片花式索引，只物化被选行（内存峰值随过滤选择性下降）
                part_ids = positions // PART_ROWS
                offsets = positions % PART_ROWS
                for pid in np.unique(part_ids):
                    ri = np.nonzero(part_ids == pid)[0]
                    cand_scores[ri] = (
                        self.vstore._part(int(pid))[offsets[ri]] @ qvec
                    )
            scores = cand_scores
            pool_n = min(max(top_k * 16, 80), n_rows)
            # argpartition 只取池不做全排序（O(n) 对 35.8 万元素）
            if pool_n >= n_rows:
                top_idx = np.argsort(-scores)
            else:
                top_idx = np.argpartition(-scores, pool_n - 1)[:pool_n]
                top_idx = top_idx[np.argsort(-scores[top_idx])]
            for rank, i in enumerate(top_idx):
                rrf[rows[int(i)][0]] = 1.0 / (60 + rank)
        # 关键词候选：所有短语 OR 联合，一条 MATCH + 全局 bm25 排序取 top50。
        # 旧实现"每短语单独 MATCH 取本地 top10"有两个缺陷：
        # 1) 答案块若在任一短语下都排不进本地前10就漏召回（如"绝缘手套"df=220，
        #    真正的试验周期块只提一次，被重复提及多的块挤出）；
        # 2) 最多 24 条 SQL。联合 MATCH 的全局 bm25 天然给"命中多个稀有短语"的
        #    块更高分（等价多短语加权一次到位），SQL 也降到 1 条。
        # 注意：FTS rowid = chunks.id，须经 chunks 才能关联到 files；
        # FTS 权重加倍确保原词精确命中可翻盘。
        if mode in ("hybrid", "keyword"):
            W_FTS = 2.0
            fts_phrases = []
            seen = set()
            # 标准号数字碎片跨前缀泄漏由精准加权通道处理，FTS 不再重复加分
            for raw_phrase in build_fts_phrases(_norm(query))[:FTS_MAX_PHRASES]:
                if raw_phrase in fts_suppress:
                    continue
                for phrase in _digit_variants(raw_phrase):
                    if phrase not in seen:
                        seen.add(phrase)
                        fts_phrases.append(phrase)
            if fts_phrases:
                match_expr = " OR ".join('"%s"' % p.replace('"', '""') for p in fts_phrases)
                fsql = (
                    "SELECT chunks_fts.rowid FROM chunks_fts"
                    " JOIN chunks c ON c.id=chunks_fts.rowid"
                    " JOIN files f ON f.id=c.file_id"
                    + filter_sql + " AND chunks_fts MATCH ?"
                    " ORDER BY chunks_fts.rank LIMIT 50"
                )
                try:
                    fts_rows = self.conn.execute(fsql, params + [match_expr]).fetchall()
                except sqlite3.OperationalError:
                    fts_rows = []
                for rank, (cid,) in enumerate(fts_rows):
                    if cid in id2row:
                        rrf[cid] = rrf.get(cid, 0.0) + W_FTS / (60 + rank)
        # 标准号精准加权（refs/std_nums 已在 FTS 通道前解析）：
        # 前缀必须一致：DL/T 853 不得命中库内的 GBT 853-1988（旧逻辑只比数字会错配）；
        # query 带年份且库内有同年版本时只取同年；前缀匹配全空才退回纯数字边界匹配兜底
        std_match_cids = []
        if refs:
            f_where = list(where)
            f_params = list(params)
            f_filter = (" WHERE " + " AND ".join(f_where)) if f_where else ""
            all_files = self.conn.execute(
                "SELECT f.id, f.path FROM files f" + f_filter, f_params
            ).fetchall()
            core_ids, year_ids = set(), set()
            for fid, path in all_files:
                nn = _std_name_norm(path)
                kinds = {_std_ref_match(ref, nn) for ref in refs}
                if "year" in kinds:
                    year_ids.add(fid)
                if "year" in kinds or "core" in kinds:
                    core_ids.add(fid)
            matched_file_ids = year_ids or core_ids
            if not matched_file_ids:
                # 兜底：罕见/怪异命名（手写前缀等）退回数字主体边界匹配，保证召回
                std_pattern = re.compile(
                    r"(?:^|[^\d])(" + "|".join(re.escape(n) for n in std_nums) + r")(?:[^\d]|$)"
                )
                matched_file_ids = {fid for fid, path in all_files if std_pattern.search(path)}
            if matched_file_ids:
                placeholders = ",".join("?" for _ in matched_file_ids)
                std_match_cids = [
                    r[0]
                    for r in self.conn.execute(
                        f"SELECT id FROM chunks WHERE file_id IN ({placeholders})",
                        list(matched_file_ids),
                    ).fetchall()
                ]
            std_bonus = 0.1  # 远大于 RRF 常规分(0.017~0.05)，确保标准号文件置顶
            for cid in std_match_cids:
                rrf[cid] = rrf.get(cid, 0.0) + std_bonus
        # 标准层小幅加权：企业标准/行业标准/国家标准/部委规章/法律法规 的 chunk 略加分，
        # 避免 Wiki/经验文件因语义相似长期压制标准原文（0.005 远小于 FTS 0.033，不翻盘）
        STD_LAYERS = {"企业标准", "行业标准", "国家标准", "部委规章", "法律法规"}
        STD_LAYER_BONUS = 0.005
        for cid in rrf:
            if rows[id2row[cid]][8] in STD_LAYERS:
                rrf[cid] += STD_LAYER_BONUS
        # RRF 排序
        ordered = sorted(rrf, key=rrf.get, reverse=True)
        # cross-encoder 精排：对前 RERANK_TOP_N 个候选重打分重排序
        # keyword 模式跳过；标准号查询跳过（用户意图精确，rerank 可能压掉目标文件）；
        # 条件精排：top1 有强 FTS 信号且明显领先 top2 时跳过（省 3.5s）
        rerank_scores = {}
        if rerank and mode != "keyword" and not std_nums and ordered:
            # 区分性实体（含强实体强制召回，见下）
            strong_ent, weak_ent = extract_key_entities(
                query, self.conn, filter_sql, params
            )
            # 强实体强制召回：bm25 对长表格块有长度惩罚，含主体名词的真答案块
            # 可能排不进联合 MATCH top50（实测"绝缘手套 工频耐压试验 1次/半年"
            # 表格块在"绝缘手套"单短语 bm25 排序中落到 40 名开外，候选 24 块
            # 无一含"绝缘手套"）。对策：AND 直取 bm25 前 60，再用查询短语覆盖数
            # 在 Python 侧重排取前 10 补进精排候选，由 rerank+覆盖分层定真伪。
            extra_cids = []
            if strong_ent:
                expr = " AND ".join('"%s"' % g.replace('"', '""') for g in strong_ent)
                fs = (
                    "SELECT chunks_fts.rowid FROM chunks_fts"
                    " JOIN chunks c ON c.id=chunks_fts.rowid"
                    " JOIN files f ON f.id=c.file_id"
                    + filter_sql + " AND chunks_fts MATCH ?"
                    " ORDER BY chunks_fts.rank LIMIT 200"
                )
                try:
                    fr = self.conn.execute(fs, params + [expr]).fetchall()
                    head = set(ordered[:RERANK_TOP_N])
                    pool = [
                        (cid, rk)
                        for rk, (cid,) in enumerate(fr)
                        if cid in id2row and cid not in head
                    ]
                    if pool:
                        ptexts = _fetch_texts([c for c, _ in pool])
                        cov_phrases = [p for p in fts_phrases if p not in strong_ent]
                        scored = []
                        for cid, rk in pool:
                            r = rows[id2row[cid]]
                            full = _norm(
                                (r[4] or "") + (ptexts.get(cid, "") or "")
                            )
                            s_hit = sum(1 for g in strong_ent if g in full)
                            w_hit = sum(1 for p in cov_phrases if p in full)
                            scored.append((s_hit, w_hit, -rk, cid))
                        scored.sort(reverse=True)
                        extra_cids = [cid for _, _, _, cid in scored[:10]]
                except sqlite3.OperationalError:
                    pass
            # 联合 MATCH 下单块 FTS 最高分=2/60≈0.033；再叠加向量 RRF 与标准层
            # 加权后达 0.045 以上，说明"关键词第一且语义同池"，可跳过精排省 ~3.5s；
            # 但有强制召回块时必须精排（融合分代表不了它们）
            top1_score = rrf[ordered[0]]
            top2_score = rrf[ordered[1]] if len(ordered) > 1 else 0.0
            strong_signal = top1_score >= 0.045
            clear_lead = (top1_score - top2_score) >= 0.01
            if extra_cids or not (strong_signal and clear_lead):
                rr = Reranker.get()
                if rr is not None:
                    base = ordered[:RERANK_TOP_N]
                    cand = base + extra_cids
                    cand_texts = _fetch_texts(cand)
                    # 标题拼在 passage 最前：标题是高信息密度线索，且补偿正文截断
                    texts = [
                        ((rows[id2row[c]][4] or "") + "\n" + (cand_texts.get(c, "") or "")).strip()
                        for c in cand
                    ]
                    try:
                        r_scores = rr.score(query, texts)
                        for c, s in zip(cand, r_scores):
                            rerank_scores[c] = float(s)
                        # 区分性实体覆盖保护：先按 strong（主体名词）命中数分层，
                        # 再按 weak（提问意图碎片）是否命中做二元分层，层内完全
                        # 交给 rerank 分排序。weak 只做有/无二元而不计数，避免
                        # 碎片偶然多命中的块压过 rerank 强烈否定的块。
                        if os.environ.get("RAG_DEBUG"):
                            print(
                                f"[RAG_DEBUG] strong={strong_ent} weak={weak_ent} "
                                f"extra={len(extra_cids)}",
                                file=sys.stderr,
                            )
                        if strong_ent or weak_ent:
                            covers = []
                            for i in range(len(cand)):
                                full = _norm(
                                    (rows[id2row[cand[i]]][4] or "")
                                    + (cand_texts.get(cand[i], "") or "")
                                )
                                s_hit = sum(1 for g in strong_ent if g in full)
                                w_bin = 1 if any(g in full for g in weak_ent) else 0
                                covers.append((s_hit, w_bin))
                            if any(s or w for s, w in covers):
                                o = sorted(
                                    range(len(cand)),
                                    key=lambda i: (
                                        covers[i][0],
                                        covers[i][1],
                                        float(r_scores[i]),
                                    ),
                                    reverse=True,
                                )
                                cand = [cand[i] for i in o]
                            else:
                                cand = [cand[i] for i in np.argsort(-r_scores)]
                        else:
                            cand = [cand[i] for i in np.argsort(-r_scores)]
                        cand_set = set(cand)
                        ordered = cand + [c for c in ordered if c not in cand_set]
                    except Exception as e:
                        print(f"[警告] rerank 推理失败，退回无精排: {e}", file=sys.stderr)
        # 每文件限量组装：保证来源多样性
        max_per_file = max(2, (top_k + 1) // 2)
        # 预取前 top_k*4 候选的 text（覆盖 per_file 跳过情况）
        need_cids = ordered[: top_k * 4]
        text_map = _fetch_texts(need_cids)
        results = []
        per_file = {}
        for cid in ordered:
            r = rows[id2row[cid]]
            path = r[7]
            if per_file.get(path, 0) >= max_per_file:
                continue
            per_file[path] = per_file.get(path, 0) + 1
            file_name = path.split("/")[-1]
            if cid in rerank_scores:
                sc = round(rerank_scores[cid], 4)
            elif scores is not None:
                sc = round(float(scores[id2row[cid]]), 4)
            else:
                sc = round(float(rrf[cid]), 4)
            body_text = text_map.get(cid, "") or ""
            results.append(
                {
                    "file": file_name,
                    "rel_path": path,
                    "layer": r[8],
                    "clause_hint": r[5],
                    "heading": r[4],
                    "line_range": [r[2], r[3]],
                    "score": sc,
                    "snippet": _strip_img_md(body_text)[:200],
                    # 块内图片（相对 md 所在目录的引用路径），供下游按位插图
                    "images": [
                        m.group(1).strip()
                        for m in _IMG_MD_PATH_RE.finditer(body_text)
                    ],
                    "is_ocr": bool(r[6]),
                    "is_responsibility": bool(r[9]),
                }
            )
            if len(results) >= top_k:
                break
        return results


# ============ CLI ============
def main():
    parser = argparse.ArgumentParser(description="示例电力本地 RAG 核心库")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_reindex = sub.add_parser("reindex", help="增量重建（--full 全量重建）")
    p_reindex.add_argument("--full", action="store_true")
    p_reindex.add_argument("--batch", type=int, default=16)

    p_up = sub.add_parser("upsert", help="单文件入库")
    p_up.add_argument("--path", required=True)
    p_up.add_argument("--batch", type=int, default=16)

    p_del = sub.add_parser("delete", help="单文件移除")
    p_del.add_argument("--path", required=True)

    sub.add_parser("status", help="查看索引状态")

    sub.add_parser("rebuild_fts", help="强制重建 FTS 关键词索引（标题/条款入索引后对旧库生效）")

    args = parser.parse_args()
    idx = RagIndex()
    try:
        if args.cmd == "status":
            import json

            print(json.dumps(idx.status(), ensure_ascii=False, indent=2))
        elif args.cmd == "reindex":
            idx.reindex(full=args.full, batch=args.batch)
        elif args.cmd == "upsert":
            rel = idx.upsert_file(args.path, batch=args.batch)
            print(f"已入库: {rel}")
        elif args.cmd == "delete":
            ok = idx.delete_file(args.path)
            print("已删除" if ok else "未找到对应记录")
        elif args.cmd == "rebuild_fts":
            import json

            print(json.dumps(idx.ensure_fts(force=True), ensure_ascii=False, indent=2))
    finally:
        idx.close()


if __name__ == "__main__":
    main()
