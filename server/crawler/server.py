# -*- coding: utf-8 -*-
"""
示例电力标准爬虫 MCP Server

把 crawler_core.py 的能力包装成MCP工具，供分支A/E以及用户手动触发的"标准库#"命令调用。

依赖：
    pip install fastmcp curl_cffi python-docx pypdf openpyxl xlrd markdownify

启动方式（stdio，供 .trae/mcp.json 里的 sak-crawler 服务器调用）：
    python sak_crawler_mcp.py

设计原则（对应"调用最少算力"的要求）：
    - search_standards 只查询、不下载，用于"先看有哪些候选，再决定下载哪条"，
      这是默认应该被调用的工具，而不是 download_industry_bulk。
    - download_standard 只下载单条记录，且沿用原脚本的去重逻辑。
    - download_industry_bulk 保留，但工具描述里明确标注"高开销，仅用于一次性建库"。
"""

import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fastmcp import FastMCP
import crawler_core as cc
import web_fetcher as wf
import gb_standard as gbs

mcp = FastMCP("sak-crawler")


@mcp.tool()
def list_industries() -> dict:
    """列出全部支持的行业标准分类（代码+中文名），供用户选择或前端展示菜单使用。"""
    return cc.list_industries()


@mcp.tool()
def search_standards(industry: str, keyword: str = "", limit: int = 20) -> dict:
    """按行业（可传中文名或代码，如"电力"或"DL"）查询标准列表，可选按关键词/标准号过滤标题。

    只查询、不下载任何文件，开销很低。这是日常问答场景应该优先调用的工具——先看看有哪些
    候选标准，确认要哪一条，再调用 download_standard 精确下载那一条，而不是直接调用
    download_industry_bulk 下载一整个行业。

    返回候选列表，每条含 std_no（标准号）、title（标准名称）、pk（用于download_standard的标识）、status。
    """
    return cc.search_industry_records(industry, keyword, limit)


@mcp.tool()
def download_standard(industry: str, pk: str, std_no: str, title: str, out_root: str) -> dict:
    """下载并转换单条标准为Markdown，供检索到具体候选后精确获取正文全文时调用。

    参数里的 pk/std_no/title 应该直接取自 search_standards 返回结果里的对应字段，不要臆造。
    自带去重：目标.md文件已存在时直接跳过、不重复请求网络，返回 status=skipped。

    out_root 建议固定指向工作目录下的标准库根目录，例如
    "{工作目录}/标准库/DownloadMD"，这样重复调用时的去重判断才有意义。
    """
    return cc.download_one_standard(industry, pk, std_no, title, out_root)


@mcp.tool()
def download_industry_bulk(industries: list[str], out_root: str) -> dict:
    """【高开销，谨慎调用】批量下载一个或多个行业的全部标准并转换为Markdown。

    这是"一次性建库"或"定期全量补齐"场景用的工具，会对该行业的每一条标准逐一发起
    下载请求（可能是几十到几百次网络请求），耗时长、流量大。
    日常问答/工作包生成场景里，不要调用这个工具——应该用 search_standards +
    download_standard 组合，只精确获取当前问题真正需要的那一两条标准。

    仅在用户明确要求"批量下载/建库/补齐某行业全部标准"时才调用本工具。
    """
    sess = cc._new_session()
    results = []
    for ind in industries:
        results.append(cc.process_industry(sess, ind, out_root))
    return {"results": results}


@mcp.tool()
def convert_local_dir(src_dir: str, program_root: str) -> dict:
    """把本地目录下的 doc/docx/pdf/xls/xlsx/html 等文档批量转换为Markdown。

    用于把用户手头已有的、还没入库的制度/标准文件批量转成系统能检索的Markdown格式，
    不涉及网络请求，只有本地IO和转换计算开销。
    """
    return cc.process_local_dir(src_dir, program_root)


@mcp.tool()
def compare_files(path_a: str, path_b: str) -> dict:
    """比较两个文件内容是否完全一致（逐字节比对），用于人工核实"这两份文件是不是同一版本"。"""
    return cc.compare_files(path_a, path_b)


# ===========================================================================
# 以下为 v2 新增工具：网页搜索 / 任意URL抓取 / "不择手段"获取标准
# ===========================================================================


def _infer_industry_from_std_no(std_no: str):
    """从标准号前缀推断 hbba 行业代码。
    返回 (industry_code, industry_name) 或 (None, None)。
    GB/GB-T/Q/ 开头的是国标/企标，不在 hbba 行业标准库，返回 None。
    """
    import re
    s = std_no.strip().upper().replace("／", "/").replace("—", "-")
    # 国标/企标不在 hbba
    if s.startswith("GB") or s.startswith("Q/") or s.startswith("QBEH") or s.startswith("DB"):
        return None, None
    # 取斜杠前的前缀，再提取其中的字母部分（应对 "AQ 3009" 这种无斜杠格式）
    raw_prefix = s.split("/")[0].split("-")[0]
    m = re.match(r"^([A-Z]+)", raw_prefix.strip())
    if not m:
        return None, None
    prefix = m.group(1)
    # 直接匹配（如 DL, NB, AQ）
    if prefix in cc.INDUSTRY_MAP:
        return prefix, cc.INDUSTRY_MAP[prefix]
    # 尝试去掉末尾字母（如 DLT→DL, NBT→NB）
    if len(prefix) >= 3:
        core = prefix.rstrip("TZ")
        if core in cc.INDUSTRY_MAP:
            return core, cc.INDUSTRY_MAP[core]
    return None, None


@mcp.tool()
def search_web(query: str, limit: int = 10) -> dict:
    """通用网页搜索（免API密钥）：Google（浏览器点开取真实URL）+ Bing/百度/DDG 精确URL。

    各引擎真实行为（2026 实测）：
    - Google：结果最全、无验证码。其结果链接是不透明令牌（/goto?url=…），静态解不出，
      但用 Scrapling 的 page_action 在真实浏览器里新开标签点开即可还原真实 URL（见
      _resolve_google_link）——还原成功的条目 exact=True；少数被 Google 掐断的退回
      cite 域名级 URL（exact=False）。
    - Bing / 百度 / DuckDuckGo：稳定补充，给出可直接抓取的精确URL（Bing 实测能直出
      openstd 详情页 hcno 与地方政府直链 PDF），条目 exact=True。

    用于：
    - 搜索标准号找原文下载链接（优先用 exact=True 的条目）
    - 搜索条款原文、任意网页信息

    返回：{"engine": "google+bing" 等, "total": N,
           "results": [{"title", "url", "snippet", "via", "exact"}, ...]}
    （精确URL条目在前，仅有域名级的 Google 覆盖条目在后。）
    """
    return wf.search_web(query, limit=limit)


@mcp.tool()
def fetch_url(url: str, save_dir: str = "", mode: str = "auto") -> dict:
    """抓取任意URL，自动识别内容类型并提取文本。

    - HTML → Scrapling 提取正文（markdown，自动剔除脚本/隐藏内容）
    - PDF → pymupdf 提取 + 扫描件OCR回退
    - docx/xlsx/doc/xls/md/txt → 自动转换为 Markdown
    - 其他二进制 → 保存原始文件

    用于：拿到一个网页/PDF链接后，把内容抓下来转成可检索的文本。
    save_dir 传目录路径时，原始文件会保存到该目录。

    mode 控制抓取通道（默认 "auto"，逐级升级）：
    - "http"   ：只用 HTTP+cloudscraper（最快，禁浏览器；批量抓普通页时用）
    - "auto"   ：HTTP → cloudscraper → 反检测浏览器 → 渲染JS浏览器（默认）
    - "stealth"：优先走反检测浏览器（适合 Cloudflare/强反爬站）
    - "dynamic"：优先走渲染JS的完整浏览器（适合 JS 单页应用/需执行脚本才出内容的站）
    返回里的 method 字段会告诉你实际生效的通道。
    """
    return wf.fetch_url(url, save_path=None if not save_dir else
                        os.path.join(save_dir, cc.safe_filename(url.rsplit("/", 1)[-1].split("?")[0])),
                        mode=mode)


@mcp.tool()
def acquire_standard(std_no: str, title: str = "",
                     out_root: str = "", industry: str = "",
                     source_urls: list = None, mode: str = "auto") -> dict:
    """【终极获取】"不择手段"拿到标准全文。

    按以下顺序尝试，任一成功即返回：
    1. 若传了 source_urls（URL列表）→ 直接逐个抓取，第一个成功就用
    2. GB/GB-T/GB-Z 国标 → 国家标准全文公开系统(openstd.samr.gov.cn)下载
    3. 若传了 industry 或能从 std_no 推断出行业 → hbba 搜索 + download_standard
       （download_standard 内部已内置"hbba失败→web搜索回退"逻辑）
    4. 都不行 → search_web 搜标准号 + fetch_with_fallback 逐个抓

    参数：
    - std_no: 标准号（如 "DL/T 408—2023"、"GB 50263-2007"）
    - title: 标准名称（可选，提高搜索命中率）
    - out_root: 输出根目录，Markdown 文件会存到 out_root/{子目录}/ 下
    - industry: 行业名或代码（可选，如"电力"或"DL"）
    - source_urls: 已知的下载URL列表（可选，最高优先级）
    - mode: 抓取通道模式，见 fetch_url（默认 "auto"）

    返回：{"status": "ok"/"failed", "md_file": 路径,
           "source": "direct_url"/"gb_openstd"/"hbba"/"web_fallback"/"web_search",
           "message": 说明, "url": 来源URL（web时）}
    """
    if not out_root:
        out_root = os.path.join(os.path.expanduser("~"), "Desktop")

    # 1) 用户给了已知URL → 直接抓
    if source_urls:
        folder = os.path.join(out_root, _safe_dir_name(std_no, title))
        os.makedirs(folder, exist_ok=True)
        md_file = os.path.join(folder, cc.safe_filename("%s %s.md" % (std_no, title or "标准")))
        fb = wf.fetch_with_fallback(source_urls, save_dir=folder, mode=mode)
        if fb.get("status") == "ok" and fb.get("text", "").strip():
            text = fb["text"]
            used_url = source_urls[(fb.get("tried") or 1) - 1] if fb.get("tried") else source_urls[0]
            header = "# %s %s\n\n> 来源：直接指定URL\n> 抓取URL: %s\n> 正式引用前请核对原文（R00铁律）\n\n" % (
                std_no, title, used_url)
            with open(md_file, "w", encoding="utf-8") as f:
                f.write(header + text)
            return {"status": "ok", "md_file": md_file, "source": "direct_url",
                    "url": used_url, "message": "通过指定URL获取成功"}
        return {"status": "failed", "md_file": None,
                "message": "指定URL全部抓取失败: %s" % fb.get("error", "")}

    # 2) GB 国标 → 国家标准全文公开系统
    std_no_upper = std_no.strip().upper().replace("／", "/").replace("—", "-")
    if std_no_upper.startswith("GB") and not std_no_upper.startswith("GBE") and not std_no_upper.startswith("GBT/") or std_no_upper.startswith("GB/T") or std_no_upper.startswith("GBZ"):
        # 2026-10-10 目录编号化：00国家标准 优先，旧"国家标准"惯例自动纠正
        folder = cc.resolve_layer_root(out_root, "国家标准", "00国家标准")
        os.makedirs(folder, exist_ok=True)
        pdf_path = os.path.join(folder, cc.safe_filename("%s %s.pdf" % (std_no, title or "")))
        md_file = os.path.join(folder, cc.safe_filename("%s %s.md" % (std_no, title or "标准")))

        result = gbs.download_gb(std_no, out_path=pdf_path)
        if result.get("status") == "ok":
            # PDF 下载成功 → 转换为 Markdown（含OCR回退）
            try:
                with open(result["pdf_file"], "rb") as f:
                    pdf_data = f.read()
                md_text, img_n = cc.to_markdown(
                    pdf_data, "doc.pdf", "application/pdf",
                    assets_dir=cc.assets_dir_for_md(md_file))
                method = result.get("method", "")
                header = "# %s %s\n\n> 来源：国家标准全文公开系统(%s方式)\n" % (
                    std_no, title or result.get("title", ""), method)
                if method == "preview":
                    header += "> ⚠️ 预览方式下载（图片拼接PDF），已自动OCR转换，可能有错字。请对照原文核对（R00铁律）。\n"
                header += "\n"
                if not md_text.strip():
                    md_text = "> 未提取到可用文本，请直接查看PDF原件。"
                with open(md_file, "w", encoding="utf-8") as f:
                    f.write(header + md_text)
                return {"status": "ok", "md_file": md_file, "pdf_file": result["pdf_file"],
                        "source": "gb_openstd", "method": method, "images": img_n,
                        "message": "从国家标准全文公开系统下载成功(%s)" % method}
            except Exception as exc:
                # PDF 转换失败，但 PDF 文件已下载
                return {"status": "ok", "md_file": None, "pdf_file": result["pdf_file"],
                        "source": "gb_openstd", "method": method,
                        "message": "PDF下载成功但转换Markdown失败: %s，PDF文件可用" % exc}
        # 国标系统失败 → 继续走下面的 hbba / web 搜索

    # 3) hbba 路径（行业标准）
    ind_inp = industry or _infer_industry_from_std_no(std_no)[0]
    if ind_inp:
        # 先搜索
        sr = cc.search_industry_records(ind_inp, keyword=std_no.split("—")[0].split("-")[0].strip(),
                                        limit=20)
        candidates = sr.get("records", [])
        # 用标准号精确匹配（去掉空格和符号差异后比对）
        std_no_clean = _normalize_std_no(std_no)
        matched = None
        for r in candidates:
            if _normalize_std_no(r.get("std_no", "")) == std_no_clean:
                matched = r
                break
        if not matched and candidates:
            # 退而求其次：标题包含
            for r in candidates:
                if title and title in r.get("title", ""):
                    matched = r
                    break
        if not matched and len(candidates) == 1:
            matched = candidates[0]

        if matched:
            result = cc.download_one_standard(
                ind_inp, matched["pk"], matched["std_no"], matched["title"], out_root)
            if result.get("status") in ("downloaded", "skipped"):
                result["source"] = result.get("source", "hbba")
                return result
            # hbba 失败但 download_one_standard 内部已尝试 web 回退，
            # 走到这里说明回退也失败了，继续走下面的纯 web 搜索

    # 3) 纯 web 搜索回退（国标/企标/hbba未命中/行业推断失败）
    folder = os.path.join(out_root, _safe_dir_name(std_no, title))
    os.makedirs(folder, exist_ok=True)
    md_file = os.path.join(folder, cc.safe_filename("%s %s.md" % (std_no, title or "标准")))

    query = '"%s" %s 全文' % (std_no, title) if title else '"%s" 全文' % std_no
    sr = wf.search_web(query, limit=8)
    urls = [r["url"] for r in sr.get("results", []) if r.get("url")]
    if not urls:
        return {"status": "failed", "md_file": None,
                "message": "所有途径均失败：hbba未命中或无行业归属，web搜索无结果（%s）" % sr.get("error", "")}

    fb = wf.fetch_with_fallback(urls, save_dir=folder, mode=mode)
    if fb.get("status") == "ok" and len(fb.get("text", "").strip()) >= 200:
        text = fb["text"]
        used_url = urls[(fb.get("tried") or 1) - 1] if fb.get("tried") else urls[0]
        header = "# %s %s\n\n> ⚠️ 来源：网络搜索（hbba未提供或未命中）。\n> 搜索引擎: %s\n> 抓取URL: %s\n> 正式引用前请对照官方出版物核对原文（R00铁律）\n\n" % (
            std_no, title, sr.get("engine"), used_url)
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(header + text)
        return {"status": "ok", "md_file": md_file, "source": "web_search",
                "url": used_url, "message": "通过网络搜索获取成功"}
    return {"status": "failed", "md_file": None,
            "message": "web搜索有结果但抓取失败或内容过短: %s" % fb.get("error", "文本不足200字符")}


def _normalize_std_no(s: str) -> str:
    """标准化标准号用于比对：去空格、统一斜杠/破折号。"""
    s = s.upper().strip()
    s = s.replace("／", "/").replace("—", "-").replace("–", "-")
    s = s.replace(" ", "").replace("\t", "")
    # 去掉 /T 和 /Z 后缀差异（DL/T 和 DLT 视为相同）
    s = s.replace("/T", "T").replace("/Z", "Z")
    return s


def _safe_dir_name(std_no: str, title: str) -> str:
    """生成安全的子目录名。"""
    base = std_no.split("/")[0].split("-")[0].strip().replace(" ", "")
    return base if base else "STD"


@mcp.tool()
def acquire_standard_free(std_no: str, title: str = "",
                           out_root: str = "", mode: str = "auto") -> dict:
    """【免费获取收费标准】从官方/协会公开页、GitHub 存档、Internet Archive 取标准全文。

    适用于：
    - 标准在国内官方渠道收费（如部分 DL/T、NB/T、GB/T 在 hbba/openstd 仅预览）
    - acquire_standard 失败后，想再试公开免费渠道
    - 标准号已知但官方未公开，希望从协会公开页/开源存档中找

    渠道纪律与自动抓取顺序（按域名可信度排序）：
    1. 官方/政府/行业协会公开页（openstd、mohurd、cwec、cnass 等）
    2. GitHub/Gitee 存档（raw 直链优先）
    3. Internet Archive 历史快照
    4. 其他普通网页
    付费墙/盗版文档站（道客巴巴/豆丁/百度文库/book118 等）只记录到
    blocked_candidates，不自动抓取。
    PDF 下载后强制：封面标准号身份校验（防错配）+ 魔数/体积校验 +
    页数/双页拼版/末页完整体检，不合格自动删件换源。

    参数：
    - std_no: 标准号（必填）
    - title: 标准名称（可选，提高命中率）
    - out_root: 输出目录（可选，默认桌面）
    - mode: 抓取通道模式，见 fetch_url（默认 "auto"）

    返回：{"status", "md_file"/"saved_file", "source", "source_url",
           "identity_check", "health_check", "blocked_candidates", "tried", "message"}
    """
    if not out_root:
        out_root = os.path.join(os.path.expanduser("~"), "Desktop")
    folder = os.path.join(out_root, _safe_dir_name(std_no, title))
    os.makedirs(folder, exist_ok=True)
    md_file = os.path.join(folder, cc.safe_filename("%s %s.md" % (std_no, title or "标准")))

    r = wf.acquire_standard_via_free_channels(std_no, title, save_dir=folder,
                                              max_attempts=10, mode=mode)
    if r.get("status") == "ok":
        text = r.get("text", "")
        source_type = r.get("source_type", "unknown")
        source_url = r.get("source_url", "")
        # 处理只有 saved_file 没有 text 的情况（比如纯 PDF 下载）
        if not text.strip() and r.get("saved_file"):
            try:
                with open(r["saved_file"], "rb") as f:
                    data = f.read()
                conv = cc.to_markdown(
                    data, os.path.basename(r["saved_file"]),
                    "application/octet-stream",
                    assets_dir=cc.assets_dir_for_md(md_file))
                text = conv[0] if isinstance(conv, tuple) else conv
            except Exception as e:
                text = f"> 抓取到原始文件但转换文本失败: {e}\n> 文件路径: {r['saved_file']}"
        # 体检告警（双页拼版/末页空白/扫描件等）写入 md 头并在 message 里提示
        health = r.get("health_check")
        warn_line = ""
        if health and health.get("warnings"):
            warn_line = "> 下载件体检告警：%s\n" % "；".join(health["warnings"])
        if text.strip():
            header = ("# %s %s\n\n> 来源：公开免费渠道(%s)\n> 抓取URL: %s\n"
                      "> 封面身份校验：%s\n%s"
                      "> ⚠️ 非官方出版物，正式引用前必须对照官方出版物核对原文（R00铁律）\n\n") % (
                std_no, title, source_type, source_url,
                r.get("identity_check", "未知"), warn_line)
            with open(md_file, "w", encoding="utf-8") as f:
                f.write(header + text)
            return {"status": "ok", "md_file": md_file, "source": source_type,
                    "source_url": source_url,
                    "identity_check": r.get("identity_check"),
                    "health_check": health,
                    "blocked_candidates": r.get("blocked_candidates", []),
                    "tried": r.get("tried", []),
                    "message": "通过公开免费渠道(%s)获取成功%s" % (
                        source_type,
                        "；体检告警：" + "；".join(health["warnings"]) if health and health.get("warnings") else "")}
        # 有 saved_file 但没转出文本，也返回让用户自己处理
        if r.get("saved_file"):
            return {"status": "ok", "md_file": None,
                    "saved_file": r["saved_file"], "source": source_type,
                    "source_url": source_url,
                    "identity_check": r.get("identity_check"),
                    "health_check": health,
                    "blocked_candidates": r.get("blocked_candidates", []),
                    "tried": r.get("tried", []),
                    "message": "公开渠道抓到原始文件但未转为MD，请手工处理: %s" % r["saved_file"]}
    return {"status": "failed", "md_file": None,
            "tried": r.get("tried", []),
            "blocked_candidates": r.get("blocked_candidates", []),
            "message": "公开免费渠道全部失败: %s" % r.get("error", "未获取到有效文本")}


@mcp.tool()
def diagnose_ocr() -> dict:
    """诊断 OCR 引擎依赖与配置状态。

    用于排查"PDF转MD出来是乱码/空文本"问题，返回各依赖项的就绪状态：
    rapidocr / pymupdf / pdfplumber / pillow / opencv / numpy
    """
    import ocr_engine
    return ocr_engine.diagnose()


@mcp.tool()
def check_pdf_file(path: str, min_pages: int = 3) -> dict:
    """本地 PDF 下载件完整体检（毫秒级，不做OCR）。

    用于批量下载后/入库前快速发现：
    - 页数过少（预览件/截断件）
    - 双页拼版扫描（横向页占比高，需拆页后重新OCR）
    - 末页空白（下载被截断）
    - 扫描件无文字层（转MD需走OCR）
    - 加密PDF

    参数：
    - path: PDF 文件绝对路径（或目录，目录时递归检查其中全部 .pdf）
    - min_pages: 页数下限，低于则告警（默认3）

    返回：单文件返回体检 dict；目录返回 {"files":[...], "problems":[有告警的路径...]}
    """
    if not os.path.isabs(path):
        return {"status": "failed", "error": "请提供绝对路径"}
    if os.path.isdir(path):
        import glob
        pdfs = sorted(glob.glob(os.path.join(path, "**", "*.pdf"), recursive=True))
        files = []
        problems = []
        for p in pdfs:
            h = cc.pdf_health_check(p, min_pages=min_pages)
            h["path"] = p
            files.append(h)
            if h.get("warnings"):
                problems.append({"path": p, "warnings": h["warnings"]})
        return {"status": "ok", "total": len(files), "files": files,
                "problem_count": len(problems), "problems": problems}
    if not os.path.isfile(path):
        return {"status": "failed", "error": "文件不存在: %s" % path}
    h = cc.pdf_health_check(path, min_pages=min_pages)
    h["status"] = "ok" if "error" not in h else "failed"
    h["path"] = path
    return h


if __name__ == "__main__":
    mcp.run()
