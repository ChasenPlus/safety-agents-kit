# -*- coding: utf-8 -*-
"""
标准文档处理核心库（原 download_std.py 的非交互式重构版）

改动说明（相对原脚本）：
    - 原脚本是交互式命令行工具（用 input() 接收参数），无法被MCP server直接调用。
      本文件把所有业务逻辑改成参数化函数，去掉全部 input()。
    - 转换逻辑（docx_to_md / doc_to_md / pdf_to_md / xlsx_to_md / xls_to_md / html_to_md /
      raw_to_md / to_markdown / normalize_md_text / safe_filename）与原脚本基本一致。
    - 新增 search_industry_records()：只查询、按关键词过滤，不下载——用于"先搜索预览、
      再挑真正需要的下载"，避免每次提问都整段行业标准全量下载，这是控制算力/流量开销的关键改动。
    - 新增 download_one_standard()：只下载单条记录（按pk），供按需精确下载使用。
    - process_industry()（整行业批量下载）、process_local_dir()（本地目录转换）保留，
      供"一次性建库"场景使用，但不建议在日常问答里触发。
    - runduibi()（交互式两文件对比）重构为 compare_files(path_a, path_b)。
    - 网络请求库用 curl_cffi（session.impersonate="chrome"）：requests的TLS握手特征和
      真实浏览器不一样，容易被部分网站的反爬检测拦截；curl_cffi带真实Chrome的TLS/JA3
      指纹发请求，规避这类检测。
    - process_industry() 的限速延迟(sleep_rand)时机已对照原脚本核对：只在真正发起过
      网络请求（下载成功/空文件/下载或转换失败）时才延迟，命中本地已存在文件而跳过时
      不延迟——这一点在早期重构版本中曾经写成"每条记录都延迟"，已按原脚本行为修正回来，
      避免批量重跑时对已下载文件也白白等待。
"""

import io
import os
import re
import sys
import time
import random
import datetime
import filecmp
import urllib.parse
from pathlib import Path

from curl_cffi import requests

BASE = "https://hbba.sacinfo.org.cn"

INDUSTRY_MAP = {
    "AQ": "安全生产", "BB": "包装", "CB": "船舶", "CH": "测绘", "CJ": "城镇建设",
    "CY": "新闻出版", "DA": "档案", "DB": "地震", "DL": "电力", "DY": "电影",
    "DZ": "地质矿产", "EJ": "核工业", "FZ": "纺织", "GA": "公共安全", "GC": "国家物资储备",
    "GF": "国防工业", "GH": "供销合作", "GM": "国密", "GY": "广播电视和网络视听",
    "HB": "航空", "HG": "化工", "HJ": "环境保护", "HS": "海关", "HY": "海洋",
    "JB": "机械", "JC": "建材", "JG": "建筑工程", "JR": "金融", "JS": "机关事务",
    "JT": "交通", "JY": "教育", "KA": "矿山安全", "LB": "旅游", "LD": "劳动和劳动安全",
    "LS": "粮食", "LY": "林业", "MH": "民用航空", "MR": "市场监管", "MT": "煤炭",
    "MZ": "民政", "NB": "能源", "NY": "农业", "QB": "轻工", "QC": "汽车",
    "QJ": "航天", "QX": "气象", "RB": "认证认可", "RF": "人民防空", "SB": "国内贸易",
    "SC": "水产", "SF": "司法", "SH": "石油化工", "SJ": "电子", "SL": "水利",
    "SN": "出入境检验检疫", "SW": "税务", "SY": "石油天然气", "TB": "铁路", "TD": "土地管理",
    "TY": "体育", "WB": "物资管理", "WH": "文化", "WJ": "兵工民品", "WM": "外经贸",
    "WS": "卫生", "WW": "文物保护", "XB": "稀土", "XF": "消防救援", "YB": "黑色冶金",
    "YC": "烟草", "YD": "通信", "YJ": "减灾救灾与综合性应急管理", "YS": "有色金属",
    "YY": "医药", "YZ": "邮政", "ZY": "中医药",
}
NAME2CODE = {v: k for k, v in INDUSTRY_MAP.items()}

HEADERS = {
    # 不手写 User-Agent：impersonate="chrome" 会自动生成与其TLS指纹匹配的真实Chrome
    # User-Agent；手动覆盖成写死的字符串反而会造成"TLS指纹说是Chrome，UA头却对不上"的
    # 不一致，部分风控系统专门检测这种不一致。
    "X-Requested-With": "XMLHttpRequest",
    "Referer": BASE + "/stdList",
}


def sleep_rand():
    time.sleep(random.uniform(1, 3))


def normalize_md_text(raw):
    lines = [ln.rstrip() for ln in raw.splitlines()]
    out, blank = [], False
    for ln in lines:
        if not ln.strip():
            if out and not blank:
                out.append("")
            blank = True
            continue
        blank = False
        out.append(ln.strip())
    while out and out[-1] == "":
        out.pop()
    return "\n".join(out)


def safe_filename(s):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", s)
    return s.strip(" .") or "unnamed"


def docx_to_md(data, assets_dir=None):
    """docx → markdown。

    assets_dir 给定时，内嵌图片（DrawingML + 老 VML）按文档顺序抽取落盘，
    并在对应段落/表格单元格位置插入 ``![](./{目录名}/d01.png)`` 引用；
    rId→落盘文件走显式映射（同一图多处引用复用同一文件），不靠正则猜名。
    返回 markdown 文本（与旧签名兼容）。
    """
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    import image_extract as ie

    document = docx.Document(io.BytesIO(data))

    # 图片资产准备
    asset_ctx = None
    if assets_dir:
        rels = ie.docx_rels_images(document)
        if rels:
            os.makedirs(assets_dir, exist_ok=True)
            ref_dir = os.path.basename(assets_dir.rstrip("/\\"))
            asset_ctx = {"rels": rels, "map": {}, "seq": 0,
                         "ref_dir": ref_dir, "dir": assets_dir}

    def emit_images(element):
        """把元素内按序出现的图片转成 md 图片行（仅在 asset_ctx 启用时）。"""
        if not asset_ctx:
            return []
        lines = []
        for rid in ie.iter_element_rids(element):
            rel = asset_ctx["map"].get(rid)
            if rel is None:
                spec = asset_ctx["rels"].get(rid)
                if not spec:
                    continue  # 关系指向非图片部件（ole 对象等），跳过
                ext, blob = spec
                asset_ctx["seq"] += 1
                noext = os.path.join(asset_ctx["dir"], "d%02d" % asset_ctx["seq"])
                fname = ie.save_normalized(blob, ext, noext)
                if not fname:
                    continue
                rel = fname
                asset_ctx["map"][rid] = fname
            lines.append("![](./%s/%s)" % (asset_ctx["ref_dir"], rel))
        return lines

    out = []

    def iter_block(parent):
        for child in parent.element.body.iterchildren():
            if child.tag.endswith("}p"):
                yield Paragraph(child, parent)
            elif child.tag.endswith("}tbl"):
                yield Table(child, parent)

    for block in iter_block(document):
        if isinstance(block, Paragraph):
            img_lines = emit_images(block._p)
            if not block.text.strip():
                out.extend(img_lines)
                continue
            name = (block.style.name if block.style else "") or ""
            txt = block.text.strip()
            if "Heading 1" in name:
                out.append("# " + txt)
            elif "Heading 2" in name:
                out.append("## " + txt)
            elif "Heading 3" in name:
                out.append("### " + txt)
            else:
                out.append(txt)
            # 图在文字段落之后（标题段一般不含图，img_lines 多为空）
            out.extend(img_lines)
        elif isinstance(block, Table):
            md_rows = []
            # 单元格内图片：追加在该单元格文字之后，随表格行保留位置
            for row in block.rows:
                cells = []
                for c in row.cells:
                    txt = c.text.strip().replace("|", "\\|")
                    imgs = emit_images(c._tc)
                    if imgs:
                        txt = (txt + " " + " ".join(imgs)).strip()
                    cells.append(txt)
                md_rows.append("| " + " | ".join(cells) + " |")
            if md_rows:
                n = len(md_rows[0].split("|")) - 2
                out.append(md_rows[0])
                out.append("|" + " --- |" * n)
                out.extend(md_rows[1:])
                out.append("")
    return normalize_md_text("\n".join(out))


def _doc_via_com(data):
    """Windows：用 WPS/Word COM 把老式 OLE .doc 字节流转成 .docx 字节流。

    通过独立 PowerShell 进程（带 UTF-8 BOM 脚本）调用，避免非交互会话挂起；
    任一 COM 程序不可用或超时返回 None。
    """
    if os.name != "nt":
        return None
    import tempfile
    import subprocess
    import shutil
    tmpdir = tempfile.mkdtemp(prefix="doc2md_")
    src = os.path.join(tmpdir, "in.doc")
    dst = os.path.join(tmpdir, "out.docx")
    ps = os.path.join(tmpdir, "conv.ps1")
    with open(src, "wb") as f:
        f.write(data)
    sq = lambda p: p.replace("'", "''")
    script = (
        "$ErrorActionPreference='Stop'\n"
        "$app=$null\n"
        "foreach($pn in @('kwps.Application','wps.Application','Word.Application')){"
        "try{$app=New-Object -ComObject $pn;break}catch{}}\n"
        "if(-not $app){exit 2}\n"
        "try{$app.Visible=$false}catch{}\n"
        "try{$app.DisplayAlerts=0}catch{}\n"
        "try{Unblock-File -LiteralPath '" + sq(src) + "' -ErrorAction SilentlyContinue}catch{}\n"
        "try{\n"
        "$d=$app.Documents.Open('" + sq(src) + "',$false,$true);"
        "$d.SaveAs('" + sq(dst) + "',12);$d.Close($false)\n"
        "}finally{try{$app.Quit()}catch{}}\n"
    )
    with open(ps, "w", encoding="utf-8-sig") as f:
        f.write(script)
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps],
            timeout=120, capture_output=True,
        )
        if r.returncode == 0 and os.path.exists(dst):
            with open(dst, "rb") as f:
                return f.read()
    except Exception:
        for proc in ("wps", "WINWORD"):
            try:
                subprocess.run(["taskkill", "/im", proc + ".exe", "/f"],
                               capture_output=True)
            except Exception:
                pass
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
    return None


def _flat_xml_doc_to_md(data):
    """Word 2003 Flat OPC XML（.doc 扩展名但内容为 <?xml ... wordprocessingml）。

    结构为 <pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document>…。
    先取出 document.xml 部件，再逐段落提取 w:t 文本。
    """
    import html
    text = data.decode("utf-8", errors="ignore")
    m = re.search(
        r'<pkg:part\b[^>]*pkg:name="/word/document\.xml"[^>]*>([\s\S]*?)</pkg:part>',
        text,
    )
    doc = m.group(1) if m else text
    paras = []
    for p in re.findall(r"<w:p[\s>][\s\S]*?</w:p>", doc):
        body = re.sub(r"<w:tab\s*/>", "\t", p)
        body = re.sub(r"<w:br\s*/>|<w:cr\s*/>", "\n", body)
        body = "".join(re.findall(r"<w:t(?=[\s>])[^>]*>([\s\S]*?)</w:t>", body))
        body = html.unescape(body).strip()
        if body:
            paras.append(body)
    return normalize_md_text("\n\n".join(paras))


def doc_to_md(data, assets_dir=None):
    # OLE2 复合文档（真正的老式 .doc）：裸字节强解只会得到乱码，必须走 COM 真转换
    if data[:4] == b"\xd0\xcf\x11\xe0":
        docx_bytes = _doc_via_com(data)
        if docx_bytes:
            return docx_to_md(docx_bytes, assets_dir=assets_dir)
        return ("> 该文件为旧版 .doc 二进制文档，本机 WPS/Word 自动转换未成功，"
                "请用 WPS/Word 打开原件另存为 .docx 后重新转换。")
    # Word 2003 Flat OPC XML（.doc 扩展名、内容为 XML：
    # 头部特征 progid="Word.Document" + <pkg:package>/<pkg:part>，
    # wordprocessingml 命名空间要到 2 万字节后的 document.xml 部件才出现，不能作门头判据）
    if b"<?xml" in data[:200] and (b"mso-application" in data[:300] or b"pkg:part" in data[:300]):
        md = _flat_xml_doc_to_md(data)
        if len(md.strip()) >= 20:
            return md
    # 其余按 UTF-8 文本兜底（纯文本误命中等）
    try:
        text = data.decode("utf-8", errors="ignore")
    except Exception:
        text = ""
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    printable = re.findall(r"[一-龥a-zA-Z0-9\s，。；：、（）()《》\"'\-.%/]", text)
    cleaned = "".join(printable)
    cleaned = re.sub(r"\s{2,}", "\n", cleaned)
    if len(cleaned.strip()) < 20:
        return ("> 该文件为旧版 .doc 二进制文档，当前环境缺少转换工具，"
                "无法可靠提取文本，请使用 Word 打开原文。")
    # 兜底守护：合法文档文本不超过 500KB。超大=误把 Flat OPC XML/base64 图片
    # 当文本捞了（旧缺陷曾产出 10~23MB 垃圾），此时拒绝产出、提示走专门转换。
    if len(cleaned.encode("utf-8")) > 500_000:
        return ("> 该文件疑似 Word/XML 格式文档，自动文本提取失败（内容含大量非正文数据），"
                "请用 WPS/Word 另存为 .docx 后重新转换。")
    return normalize_md_text(cleaned)


def _pdf_to_md_with_images(data, assets_dir):
    """数字版 PDF 的图文交错提取（仅在 pdf_to_md 收到 assets_dir 时走这里）。

    - 用 fitz dict 模式把文本块/图片块按版面 y 坐标交错输出，图片插在对应段落后；
    - 图标/水印（<120px 或面积<2%）过滤；同一 xref 全文档只抽一次；
    - 无图页继续用 pdfplumber 文本（保表格管道符），有图页用 dict 交错版；
    - 几乎无文字层的整页扫描块（占面≥50%）不抽图片对象，改对该页补 OCR 文字；
    - 扫描件/整篇位图直接返回 None，交回旧的纯文本（OCR）路径。
    """
    import fitz
    import image_extract as ie
    try:
        from ocr_engine import is_scanned_pdf, _maybe_strip_watermark
    except Exception:
        return None
    try:
        if is_scanned_pdf(data):
            return None  # 扫描件不抽图对象，由调用方走 OCR 路径
    except Exception:
        return None

    os.makedirs(assets_dir, exist_ok=True)
    ref_dir = os.path.basename(assets_dir.rstrip("/\\"))
    doc = fitz.open(stream=data, filetype="pdf")
    xref_files: dict = {}   # xref -> 已落盘文件名（全文档去重）
    page_seq = 0            # 页内图片序号
    pages_out = []          # [(页码1基, 是否有图, 页文本)]
    weak_pages = []         # 0基页索引：整页扫描型附图页，待补 OCR
    total_imgs = 0
    try:
        for pno, page in enumerate(doc):
            page_seq = 0
            d = page.get_text("dict")
            # 统一的版面流：文本块与图片块按 (y0, x0) 交错，保证图在对应段落位置
            flow = []
            for b in d.get("blocks", []):
                if b.get("type") == 0 and b.get("bbox"):
                    flow.append((float(b["bbox"][1]), float(b["bbox"][0]), "text", b))
            for it in ie.page_image_blocks(page):
                kind = "scanimg" if it["is_fullpage"] else "img"
                flow.append((it["bbox"][1], it["bbox"][0], kind, it))
            flow.sort(key=lambda x: (round(x[0], 1), x[1]))

            page_has_img = False
            parts = []
            for _, _, kind, item in flow:
                if kind == "text":
                    text_lines = []
                    for ln in item.get("lines", []):
                        line_txt = "".join(sp.get("text", "") for sp in ln.get("spans", []))
                        line_txt = _maybe_strip_watermark(line_txt).strip()
                        if line_txt:
                            text_lines.append(line_txt)
                    if text_lines:
                        parts.append("\n".join(text_lines))
                elif kind == "scanimg":
                    # 整页位图：数字件中的扫描附图插页。登记补 OCR，不抽大图对象。
                    if pno not in weak_pages:
                        weak_pages.append(pno)
                else:
                    xref = item["xref"]
                    fname = xref_files.get(xref)
                    if fname is None:
                        spec = ie.extract_xref_image(doc, xref)
                        if not spec:
                            continue
                        ext, blob = spec
                        page_seq += 1
                        noext = os.path.join(assets_dir, "p%03d_%02d" % (pno + 1, page_seq))
                        fname = ie.save_normalized(blob, ext, noext)
                        if not fname:
                            continue
                        xref_files[xref] = fname
                    parts.append("![](./%s/%s)" % (ref_dir, fname))
                    page_has_img = True
                    total_imgs += 1
            pages_out.append((pno, page_has_img, "\n".join(parts).strip()))
    finally:
        doc.close()

    if total_imgs == 0 and not weak_pages:
        return None  # 没有可保的图：交回旧路径，输出与旧版完全一致

    # 无图页用 pdfplumber 文本替换（保表格结构）；有图页保留交错版
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(data)) as pp:
            for idx, (pno, has_img, txt) in enumerate(pages_out):
                if has_img or pno >= len(pp.pages):
                    continue
                p_txt = (pp.pages[pno].extract_text() or "").strip()
                if p_txt and len(p_txt) > len(txt):
                    pages_out[idx] = (pno, False, p_txt)
    except Exception:
        pass

    # 整页扫描附图页补 OCR 文字（逐页调用，保证与页码对齐；封顶 30 页）
    if weak_pages:
        try:
            from ocr_engine import ocr_pdf_bytes
            for pno in weak_pages[:30]:
                t = ocr_pdf_bytes(data, dpi=200, pages=[pno], preprocess=True)
                # 替换或追加到对应页（dict 版该页通常无文字层）
                idx = next((k for k, v in enumerate(pages_out) if v[0] == pno), None)
                if idx is not None and t.strip():
                    base = pages_out[idx][2]
                    merged = (base + "\n\n" + t.strip()) if base else t.strip()
                    pages_out[idx] = (pno, pages_out[idx][1], merged)
        except Exception:
            pass

    body_parts = [t for _, _, t in pages_out if t.strip()]
    return normalize_md_text("\n\n".join(body_parts))


def pdf_to_md(data, assets_dir=None):
    # assets_dir 给定时：数字件走图文交错提取；扫描件/无图件结果与旧路径一致。
    if assets_dir:
        try:
            md = _pdf_to_md_with_images(data, assets_dir)
            if md is not None:
                return md
        except Exception:
            pass  # 图片管线任何异常都不得影响文字入库，落回旧路径
    # 优先用 OCR 引擎（pymupdf 提取 + 扫描件自动 OCR 回退），
    # 不可用时回退到 pypdf，保证环境缺包时不会整个爬虫挂掉。
    ocr_failed = None
    try:
        from ocr_engine import extract_pdf_with_ocr
        text, used_ocr = extract_pdf_with_ocr(data)
        if used_ocr and text.strip():
            text = ("> ⚠️ 该PDF为扫描件，以下文字由OCR识别转换，可能存在个别错字/空格。"
                    "正式引用前请对照PDF原件核对（R00铁律）。\n\n" + text)
        if text.strip():
            return normalize_md_text(text)
    except Exception as e:
        # 旧版静默 pass：OCR 管线整体失败时用户拿到的是 pypdf 残文却毫不知情。
        # 记录原因并在最终 md 顶部留痕（扫描件 pypdf 基本提不出正文，必须可见）。
        ocr_failed = str(e)
    # 回退：pypdf
    import pypdf
    reader = pypdf.PdfReader(io.BytesIO(data))
    chunks = []
    for page in reader.pages:
        try:
            txt = page.extract_text() or ""
        except Exception:
            txt = ""
        if txt.strip():
            chunks.append(txt.strip())
    body = normalize_md_text("\n\n".join(chunks))
    if ocr_failed is not None:
        import logging
        logging.getLogger(__name__).warning("OCR 管线失败，已降级 pypdf：%s", ocr_failed)
        warn = ("> ⚠️ OCR 转换管线异常（%s），以下为 pypdf 降级提取结果，"
                "扫描件可能为空或严重残缺，引用前请对照PDF原件。\n\n"
                % ocr_failed[:80].replace("\n", " "))
        return warn + body
    return body


def pdf_health_check(pdf, min_pages: int = 3):
    """下载件轻量体检（不做OCR、不渲染，毫秒级），用于落盘后自动发现残缺件。

    pdf: 文件路径(str) 或 PDF 字节。
    返回 dict：
      pages 页数 / is_encrypted 是否加密 / scanned 是否扫描件（前3页无文字层但有图）
      two_up_ratio 横向双页拼版页占比（>0.5 基本可判定是双页扫描，需拆页重OCR）
      last_blank 末页是否疑似空白 / warnings 问题清单
    无法解析时返回 {"error": ...}，不抛异常（体检失败不阻塞主流程）。
    """
    result = {"pages": 0, "is_encrypted": False, "scanned": False,
              "two_up_ratio": 0.0, "last_blank": False, "warnings": []}
    try:
        import fitz  # PyMuPDF
        if isinstance(pdf, (str, bytes)) and isinstance(pdf, str):
            doc = fitz.open(pdf)
        else:
            doc = fitz.open(stream=pdf, filetype="pdf")
    except Exception as exc:
        return {"error": "PDF无法打开: %s" % exc}
    try:
        n = doc.page_count
        result["pages"] = n
        result["is_encrypted"] = bool(doc.is_encrypted)
        if n < min_pages:
            result["warnings"].append("页数过少（%d页 < %d），疑似预览件/截断件" % (n, min_pages))
        if doc.is_encrypted:
            result["warnings"].append("PDF已加密，文字提取可能为空")

        # 双页拼版：扫描时把左右两页拍成一页时，页面宽明显大于高
        wide = 0
        sample = range(min(n, 10))  # 抽前10页足够判定版式
        for i in sample:
            pg = doc.load_page(i)
            r = pg.rect
            if r.width > r.height * 1.15:
                wide += 1
        ratio = wide / max(1, len(list(sample)))
        result["two_up_ratio"] = round(ratio, 2)
        if ratio > 0.5:
            result["warnings"].append(
                "疑似双页拼版扫描（横向页占比%.0f%%），需拆页后重新OCR" % (ratio * 100))

        # 扫描件识别：前3页文字层极少但有整页位图
        head_text = ""
        head_images = 0
        for i in range(min(3, n)):
            pg = doc.load_page(i)
            head_text += pg.get_text() or ""
            head_images += len(pg.get_images(full=True))
        if len(head_text.strip()) < 20 and head_images > 0:
            result["scanned"] = True

        # 末页空白检测：无文字且无图/矢量内容
        last = doc.load_page(n - 1)
        lt = (last.get_text() or "").strip()
        has_img = len(last.get_images(full=True)) > 0
        has_draw = len(last.get_drawings()) > 5
        if len(lt) < 10 and not has_img and not has_draw:
            result["last_blank"] = True
            result["warnings"].append("末页疑似空白，下载可能被截断")
    finally:
        doc.close()
    return result


def xlsx_to_md(data):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    parts = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        parts.append("## " + ws.title)
        ncols = max(len(r) for r in rows)
        lines = []
        for r in rows:
            cells = ["" if v is None else str(v).replace("|", "\\|").replace("\n", " ")
                     for v in r]
            cells += [""] * (ncols - len(cells))
            lines.append("| " + " | ".join(cells) + " |")
        lines.insert(1, "|" + " --- |" * ncols)
        parts.append("\n".join(lines))
    return normalize_md_text("\n\n".join(parts))


def xls_to_md(data):
    import xlrd
    wb = xlrd.open_workbook(file_contents=data)
    parts = []
    for ws in wb.sheets():
        if ws.nrows == 0:
            continue
        parts.append("## " + ws.name)
        ncols = ws.ncols
        lines = []
        for ri in range(ws.nrows):
            cells = []
            for ci in range(ncols):
                v = ws.cell_value(ri, ci)
                if isinstance(v, float) and v == int(v):
                    v = int(v)
                cells.append("" if v is None else str(v).replace("|", "\\|").replace("\n", " "))
            lines.append("| " + " | ".join(cells) + " |")
        lines.insert(1, "|" + " --- |" * ncols)
        parts.append("\n".join(lines))
    return normalize_md_text("\n\n".join(parts))


def html_to_md(data):
    from markdownify import markdownify as mdify
    text = data.decode("utf-8", errors="ignore")
    return normalize_md_text(mdify(text, heading_style="ATX"))


def raw_to_md(data):
    text = data.decode("utf-8", errors="ignore")
    return normalize_md_text(text)


def assets_dir_for_md(md_path):
    """按统一约定返回 md 对应的图片目录（与 md 同目录的 {stem}_images）。"""
    d = os.path.dirname(md_path)
    stem = os.path.splitext(os.path.basename(md_path))[0]
    return os.path.join(d, safe_filename(stem) + "_images")


def to_markdown(data, fname, content_type="", assets_dir=None):
    """文档转 markdown。

    assets_dir=None：返回 str（完全兼容旧调用）。
    assets_dir 给定时：返回 (md_text, images_count)，docx/pdf 内嵌图抽取落盘
    并在 md 对应位置插入相对引用；不支持图片的格式 images_count=0。
    """
    ext = (os.path.splitext(fname)[1] or "").lower()
    ctype = (content_type or "").lower()
    if assets_dir and (
        "wordprocessingml" in ctype or ext == ".docx"
        or ext == ".pdf" or "pdf" in ctype
    ):
        before = _count_files(assets_dir)
        if "wordprocessingml" in ctype or ext == ".docx":
            md = docx_to_md(data, assets_dir=assets_dir)
        else:
            md = pdf_to_md(data, assets_dir=assets_dir)
        after = _count_files(assets_dir)
        return md, max(0, after - before)
    if "wordprocessingml" in ctype:
        return docx_to_md(data)
    if ext == ".docx":
        return docx_to_md(data)
    if ext == ".doc":
        return doc_to_md(data, assets_dir=assets_dir)
    if ext in (".xlsx", ".xlsm") or "spreadsheetml" in ctype:
        return xlsx_to_md(data)
    if ext == ".xls":
        return xls_to_md(data)
    if ext in (".pdf",) or "pdf" in ctype:
        return pdf_to_md(data)
    if ext in (".html", ".htm") or "text/html" in ctype:
        return html_to_md(data)
    if ext in (".md", ".markdown", ".txt", ".text"):
        return raw_to_md(data)
    return raw_to_md(data)


def _count_files(path):
    try:
        return len(os.listdir(path))
    except OSError:
        return 0


def resolve_industry(inp):
    inp = inp.strip()
    if inp in INDUSTRY_MAP:
        return INDUSTRY_MAP[inp], inp
    if inp in NAME2CODE:
        return inp, NAME2CODE[inp]
    return inp, NAME2CODE.get(inp, inp)


def _new_session():
    # impersonate="chrome"：带上真实Chrome浏览器的TLS/JA3指纹发请求，躲开
    # "请求看起来不像真实浏览器"这类反爬检测。如果后续该指纹也被识别，可以把
    # "chrome"换成更具体的版本号（如"chrome124"）重新试。
    sess = requests.Session(impersonate="chrome")
    sess.headers.update(HEADERS)
    return sess


def query_all_records(sess, industry_name):
    """分页查询某行业全部标准记录，返回记录列表。有页数上限防死循环。"""
    records = []
    current = 1
    size = 100
    max_pages = 10000
    while current <= max_pages:
        data = {
            "current": str(current), "size": str(size), "key": "",
            "ministry": "", "industry": industry_name, "pubdate": "", "date": "",
        }
        resp = sess.post(BASE + "/stdQueryList", data=data, timeout=60)
        resp.raise_for_status()
        body = resp.json()
        recs = body.get("records") or []
        records.extend(recs)
        pages = int(body.get("pages") or 0)
        if current >= pages or not recs:
            break
        current += 1
        sleep_rand()
    return records


def search_industry_records(industry_inp, keyword="", limit=20):
    """【轻量级】只查询该行业标准列表并按关键词过滤标题/标准号，不下载任何文件。

    用于"按需检索"场景：先看有哪些候选标准，再由调用方决定要不要下载其中一两条，
    避免每次提问都触发整行业批量下载。

    返回：{"industry": 中文名, "code": 行业代码, "total": 匹配数, "records": [
        {"std_no":..., "title":..., "pk":..., "status":...}, ...]}（最多 limit 条）
    """
    name, code = resolve_industry(industry_inp)
    sess = _new_session()
    try:
        records = query_all_records(sess, name)
    except Exception as exc:
        return {"industry": name, "code": code, "total": 0, "records": [], "error": str(exc)}

    if keyword:
        kw = keyword.strip()
        records = [r for r in records
                   if kw in (r.get("chName") or "") or kw in (r.get("code") or "")]

    out = []
    for r in records[:limit]:
        out.append({
            "std_no": (r.get("code") or "unknown").strip(),
            "title": (r.get("chName") or "").strip(),
            "pk": r.get("pk") or "",
            "status": r.get("status", ""),
        })
    return {"industry": name, "code": code, "total": len(records), "records": out}


def get_download(sess, pk, retries=2):
    """下载文件，返回 (bytes, filename, content_type)。失败返回 (None, None, None)。"""
    url = BASE + "/portal/download/" + pk
    for attempt in range(retries + 1):
        try:
            resp = sess.get(url, timeout=120)
            resp.raise_for_status()
            cd = resp.headers.get("content-disposition", "")
            fname = ""
            if "filename" in cd:
                part = cd.split("filename=", 1)[1].split(";")[0].strip().strip('"')
                fname = urllib.parse.unquote(part).replace("+", " ")
            return resp.content, fname, resp.headers.get("content-type", "")
        except Exception:
            if attempt >= retries:
                return None, None, None
            sleep_rand()
    return None, None, None


def _web_fallback(std_no, title, folder, md_file):
    """hbba 下载失败/空文件时，通过网络搜索回退获取标准全文。

    策略：用标准号+标题搜索，取前5条结果逐个抓取，第一个能提取到
    实质文本内容（非空、长度>200）的就采用，写入 md_file。
    返回 (success: bool, url_used: str or None, message: str)
    """
    try:
        from web_fetcher import search_web, fetch_with_fallback
    except Exception as exc:
        return False, None, "web_fetcher 模块不可用: %s" % exc

    query = '"%s" %s 全文' % (std_no, title)
    try:
        sr = search_web(query, limit=5)
    except Exception as exc:
        return False, None, "搜索失败: %s" % exc

    urls = [r.get("url", "") for r in sr.get("results", []) if r.get("url")]
    if not urls:
        return False, None, "搜索无结果（engine=%s, error=%s）" % (
            sr.get("engine"), sr.get("error", ""))

    fb = fetch_with_fallback(urls, save_dir=folder)
    if fb.get("status") != "ok":
        return False, None, "全部候选URL抓取失败: %s" % fb.get("error", "")

    text = fb.get("text", "").strip()
    if len(text) < 200:
        return False, None, "抓取到的文本过短（%d字符），不可用" % len(text)

    used_url = urls[(fb.get("tried") or 1) - 1] if fb.get("tried") else urls[0]
    header = ("# %s %s\n\n"
              "> ⚠️ 来源：网络搜索回退（hbba官方接口未提供可下载电子文本）。\n"
              "> 搜索引擎: %s\n"
              "> 抓取URL: %s\n"
              "> 正式引用前请对照官方出版物核对原文（R00铁律）。\n\n"
              ) % (std_no, title, sr.get("engine"), used_url)
    with open(md_file, "w", encoding="utf-8") as f:
        f.write(header + text)
    return True, used_url, "通过网络搜索回退获取成功"


def resolve_layer_root(out_root, legacy_name, numbered_name):
    """知识库 2026-10-10 目录编号化后的落盘自适应，防止 makedirs 造出平行伪目录。

    覆盖三种调用惯例：
    1. out_root 末尾仍是旧层名（如 .../02-原始资料层/行业标准），且同级编号目录
       （03行业标准）已存在 → 纠正到编号目录；
    2. out_root 是上层（如 02-原始资料层）→ 优先返回其下已存在的编号目录，
       其次旧名目录；
    3. 都不存在时，返回编号目录路径（由调用方 makedirs 新建）。
    """
    root = os.path.normpath(out_root)
    base = os.path.basename(root)
    parent = os.path.dirname(root)
    if base == legacy_name:
        numbered = os.path.join(parent, numbered_name)
        return numbered if os.path.isdir(numbered) else root
    numbered_child = os.path.join(root, numbered_name)
    if os.path.isdir(numbered_child):
        return numbered_child
    legacy_child = os.path.join(root, legacy_name)
    if os.path.isdir(legacy_child):
        return legacy_child
    return numbered_child


def download_one_standard(industry_inp, pk, std_no, title, out_root):
    """【按需下载】只下载并转换单条标准记录，而不是整个行业。

    去重规则：目标 .md 文件已存在则直接跳过（视为已下载），不重复请求。

    下载失败/空文件时自动触发网络搜索回退（_web_fallback），
    尽一切可能拿到标准全文。

    返回：{"status": "downloaded"/"skipped"/"empty"/"failed"/"convert_failed",
           "md_file": 路径或None, "message": 说明,
           "source": "hbba"/"web_fallback"（仅downloaded时）}
    """
    name, code = resolve_industry(industry_inp)
    layer_root = resolve_layer_root(out_root, "行业标准", "03行业标准")
    folder = os.path.join(layer_root, name)
    os.makedirs(folder, exist_ok=True)

    md_file = os.path.join(folder, safe_filename("%s %s.md" % (std_no, title)))
    if os.path.exists(md_file):
        return {"status": "skipped", "md_file": md_file, "message": "该标准已下载过，直接复用本地文件"}

    sess = _new_session()
    data, fname, ctype = get_download(sess, pk)
    if data is None:
        # hbba 下载失败 → 立即走 web 回退
        ok, url, msg = _web_fallback(std_no, title, folder, md_file)
        if ok:
            return {"status": "downloaded", "md_file": md_file, "message": "hbba下载失败，%s" % msg,
                    "source": "web_fallback", "url": url}
        return {"status": "failed", "md_file": None,
                "message": "下载失败（网络或接口异常），web回退也未成功：%s" % msg}
    if len(data) == 0:
        # hbba 返回空文件 → 走 web 回退；回退也失败才写占位说明
        ok, url, msg = _web_fallback(std_no, title, folder, md_file)
        if ok:
            return {"status": "downloaded", "md_file": md_file, "message": "hbba空文件，%s" % msg,
                    "source": "web_fallback", "url": url}
        note = "# %s %s\n\n> 该标准未提供可下载的电子文本（下载接口返回空文件），web回退也未成功。\n> 回退失败原因: %s\n\n- 标准号: %s\n" % (
            std_no, title, msg, std_no)
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(note)
        return {"status": "empty", "md_file": md_file,
                "message": "标准存在但无可下载电子文本，web回退失败，已写入占位说明"}

    try:
        md_text, img_n = to_markdown(
            data, fname or "", ctype, assets_dir=assets_dir_for_md(md_file))
        header = "# %s %s\n\n" % (std_no, title)
        if not md_text.strip():
            md_text = "> 未提取到可用的文本内容。"
        with open(md_file, "w", encoding="utf-8") as f:
            f.write(header + md_text)
        return {"status": "downloaded", "md_file": md_file, "message": "下载并转换成功",
                "source": "hbba", "images": img_n}
    except Exception as exc:
        return {"status": "convert_failed", "md_file": None, "message": "转换失败: %s" % exc}


def process_industry(sess, industry_inp, out_root):
    """整行业批量下载——开销明显高于 download_one_standard()。"""
    name, code = resolve_industry(industry_inp)
    try:
        records = query_all_records(sess, name)
    except Exception as exc:
        return {"industry": name, "ok": 0, "skipped": 0, "failed": 0, "empty": 0, "error": str(exc)}
    if not records:
        return {"industry": name, "ok": 0, "skipped": 0, "failed": 0, "empty": 0, "total": 0}

    ok = skipped = failed = empty = 0
    for rec in records:
        std_no = (rec.get("code") or "unknown").strip()
        title = (rec.get("chName") or "").strip()
        pk = rec.get("pk") or ""
        if not pk:
            failed += 1
            continue
        result = download_one_standard(industry_inp, pk, std_no, title, out_root)
        if result["status"] == "downloaded":
            ok += 1
            sleep_rand()  # 只在真正发起过网络请求（成功下载）时限速，跳过/无pk的情况不需要延迟
        elif result["status"] == "skipped":
            skipped += 1
        elif result["status"] == "empty":
            empty += 1
            sleep_rand()  # 空文件也是真实发起过下载请求的，同样要限速
        else:
            failed += 1
            sleep_rand()  # 下载失败同样是发起过网络请求，仍要限速，避免失败后立刻重试下一条把请求打得更密

    return {"industry": name, "ok": ok, "skipped": skipped, "failed": failed, "empty": empty, "total": len(records)}


SCAN_EXTS = {".doc", ".docx", ".pdf", ".xls", ".xlsx", ".xlsm", ".html", ".htm", ".txt", ".md", ".markdown"}


def collect_files(root):
    found = []
    seen = set()
    for base, dirs, files in os.walk(root):
        real = os.path.realpath(base)
        if real in seen:
            dirs[:] = []
            continue
        seen.add(real)
        dirs[:] = [d for d in dirs if os.path.realpath(os.path.join(base, d)) not in seen]
        for fn in files:
            if os.path.splitext(fn)[1].lower() in SCAN_EXTS:
                found.append(os.path.join(base, fn))
    return found


def process_local_dir(src_dir, program_root):
    if not os.path.isdir(src_dir):
        return {"ok": 0, "failed": 0, "error": "目录不存在: %s" % src_dir}
    today = datetime.datetime.now().strftime("%Y%m%d")
    out_root = os.path.join(program_root, "TOMD", today)
    os.makedirs(out_root, exist_ok=True)

    files = collect_files(src_dir)
    if not files:
        return {"ok": 0, "failed": 0, "out_root": out_root, "message": "未找到可转换的文档文件"}

    ok = failed = images_total = 0
    for fp in files:
        rel = os.path.relpath(fp, src_dir)
        name = os.path.basename(rel)
        try:
            with open(fp, "rb") as f:
                data = f.read()
            if not data:
                continue
            rel_md = os.path.splitext(rel)[0] + ".md"
            out_file = os.path.join(out_root, rel_md)
            os.makedirs(os.path.dirname(out_file), exist_ok=True)
            ext = os.path.splitext(name)[1].lower()
            if ext in (".md", ".markdown"):
                md_text = raw_to_md(data)
            else:
                result = to_markdown(data, name,
                                     assets_dir=assets_dir_for_md(out_file))
                if isinstance(result, tuple):
                    md_text, img_n = result
                    images_total += img_n
                else:
                    md_text = result
            if not md_text.strip():
                md_text = "> 未提取到可用的文本内容。"
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(md_text)
            ok += 1
        except Exception:
            failed += 1

    return {"ok": ok, "failed": failed, "out_root": out_root,
            "total": len(files), "images_total": images_total}


def compare_files(path_a, path_b):
    """比较两个文件内容是否完全一致（原脚本 runduibi() 的非交互版本）。"""
    pa, pb = Path(path_a), Path(path_b)
    if not pa.exists() or not pb.exists():
        return {"same": None, "error": "文件路径不存在，请检查"}
    same = filecmp.cmp(str(pa), str(pb), shallow=False)
    return {"same": same}


def list_industries():
    """返回行业代码表，供菜单展示使用。"""
    return {"industries": [{"code": k, "name": v} for k, v in INDUSTRY_MAP.items()]}
