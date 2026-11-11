# -*- coding: utf-8 -*-
"""
国家标准下载模块 —— 从"国家标准全文公开系统"(openstd.samr.gov.cn) 获取 GB/GB-T/GB-Z 标准。

2026-10-03 重写（旧版依赖百度搜索找hcno+猜测端点，实测全失败）：
    真实链路（以 GB 30871-2022 实测，全部服务端渲染，无需JS执行）：
      1. 站内检索  GET /bzgk/std/std_list?p.p2=<标准号>
                   结果行内联 onclick="showInfo('32位HCNO')"，直接正则提取
      2. 详情页    GET /bzgk/std/newGbInfo?hcno=<HCNO>
                   解析"下载标准"按钮(.xz_btn)/"在线预览"按钮(.ck_btn)的 data-value
      3. 下载中转  GET /bzgk/std/showGb?type=download&hcno=<HCNO>&request_locale=zh
                   （Referer 必须是详情页）中转页内 var isValid='true' 时直接放行，
                   'false' 时需要图形验证码(verifyCode)，此时放弃自动下载
      4. PDF本体  GET /bzgk/std/viewGb?hcno=<HCNO>（Referer 为中转页）
                   返回 application/octet-stream + attachment;filename=xxx.pdf
    官方公开范围：强制性国标非采标可在线阅读和下载；推荐性国标非采标同；采标件只给在线阅读。

    下载的PDF常见两种形态：文字版 / 扫描件；另有"私有CMap字形乱码"文字版（能提取字数但汉字
    错位，如"犌犅３０８７１"）。后两类统一由 ocr_engine.extract_pdf_with_ocr 整页OCR，
    乱码检测(_looks_like_garbled)与水印识别(_maybe_strip_watermark)已在管线内处理。
"""

import os
import re
from typing import Optional

from curl_cffi import requests as cf_requests

BASE = "https://openstd.samr.gov.cn"
STD_PATH = "/bzgk/std"

GB_HEADERS = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}


def _new_session():
    # verify=False：项目部署路径含中文（非ASCII），curl_cffi加载certifi CA bundle会报
    # CURLE_SSL_CACERT(77)；openstd为政府公开站、仅读取，关闭校验可接受。
    sess = cf_requests.Session(impersonate="chrome", verify=False)
    sess.headers.update(GB_HEADERS)
    return sess


def _norm(s: str) -> str:
    """标准号归一化：去空格、全角转半角、大写，用于精确匹配。"""
    if not s:
        return ""
    out = []
    for ch in s:
        code = ord(ch)
        if 0xFF01 <= code <= 0xFF5E:
            out.append(chr(code - 0xFEE0))
        elif ch == "　":
            continue
        else:
            out.append(ch)
    return "".join(out).replace(" ", "").upper()


def search_gb(keyword: str, limit: int = 10) -> dict:
    """openstd 站内检索。

    返回：{"total": N, "records": [{"std_no", "title", "hcno", "url"}, ...]}
    """
    sess = _new_session()
    resp = sess.get(BASE + STD_PATH + "/std_list",
                    params={"p.p1": "0", "p.p90": "circulation_date",
                            "p.p91": "desc", "p.p2": keyword},
                    timeout=30)
    if resp.status_code != 200:
        return {"total": 0, "records": [], "error": "检索页HTTP %s" % resp.status_code}

    # 结果行：onclick="showInfo('HCNO');">LABEL</a>，标准号与名称各一条、同hcno
    pairs = re.findall(r"showInfo\('([0-9A-Fa-f]{32})'\);?\">([^<]{1,120})</a>", resp.text)
    by_hcno = {}
    order = []
    for hcno, label in pairs:
        hcno = hcno.upper()
        label = label.strip()
        if hcno not in by_hcno:
            by_hcno[hcno] = {"hcno": hcno, "std_no": "", "title": ""}
            order.append(hcno)
        if re.match(r"^(GB|GB/T|GB/Z)\s*\d", label.upper().replace("　", " ")):
            by_hcno[hcno]["std_no"] = re.sub(r"\s+", " ", label).strip()
        elif len(label) >= 4:
            by_hcno[hcno]["title"] = label

    records = []
    for hcno in order:
        item = by_hcno[hcno]
        item["url"] = BASE + STD_PATH + "/newGbInfo?hcno=" + hcno
        records.append(item)
        if len(records) >= limit:
            break
    return {"total": len(records), "records": records}


def get_gb_info(std_no: str) -> dict:
    """按标准号定位详情页，解析下载/预览按钮。"""
    target = _norm(std_no)
    # 关键词用标准号数字部分命中率更高（GB/T 150.1 等带点号也能搜到）
    num_m = re.search(r"(\d+(?:\.\d+)*)(?:[-—](\d{4}))?$", target.replace("/", ""))
    candidates = []
    for kw in dict.fromkeys([std_no, num_m.group(1) if num_m else std_no]):
        sr = search_gb(kw, limit=20)
        for r in sr.get("records", []):
            if r not in candidates:
                candidates.append(r)

    matched = None
    for r in candidates:
        if _norm(r.get("std_no", "")) == target:
            matched = r
            break
    if not matched:
        # 年份缺省时前缀匹配（如只要 GB 30871 而结果带 -2022）
        for r in candidates:
            n = _norm(r.get("std_no", ""))
            if n.startswith(target) or target.startswith(n.rstrip("-")):
                matched = r
                break

    if not matched:
        return {"found": False, "error": "openstd未收录或未检索到该标准",
                "searched": target, "candidates": [c.get("std_no") for c in candidates][:8]}

    hcno = matched["hcno"]
    sess = _new_session()
    detail_url = BASE + STD_PATH + "/newGbInfo?hcno=" + hcno
    resp = sess.get(detail_url, timeout=30)
    if resp.status_code != 200:
        return {"found": True, "hcno": hcno, "error": "详情页HTTP %s" % resp.status_code}

    html = resp.text
    info = {"found": True, "std_no": matched.get("std_no") or std_no,
            "title": matched.get("title", ""), "hcno": hcno, "url": detail_url,
            "can_download": bool(re.search(r'class="[^"]*xz_btn', html)),
            "can_preview": bool(re.search(r'class="[^"]*ck_btn', html))}
    # 状态字段
    m = re.search(r"标准状态[：:\s]*</?[^>]+>?\s*([现行废止即将实施]{2,4})", html)
    if m:
        info["status"] = m.group(1)
    return info


def download_gb(std_no: str, out_path: str = "") -> dict:
    """下载国标 PDF。返回 {status: ok/failed, pdf_file, method, message, ...}。"""
    if not out_path:
        safe_name = re.sub(r'[\\/:*?"<>|]', "_", std_no).strip()
        out_path = os.path.join(os.getcwd(), safe_name + ".pdf")

    info = get_gb_info(std_no)
    if not info.get("found"):
        return {"status": "failed", "message": "未找到该标准: %s" % info.get("error", ""),
                "std_no": std_no, "candidates": info.get("candidates", [])}

    hcno = info["hcno"]
    sess = _new_session()
    detail_url = info["url"]
    sess.get(detail_url, timeout=30)  # 建立会话cookie

    # 优先官方下载按钮；没有则退在线预览通道（采标件）
    types_to_try = ["download"] if info.get("can_download") else []
    if not types_to_try or not info.get("can_download"):
        types_to_try.append("online")
    if info.get("can_download"):
        types_to_try.append("online")

    last_msg = ""
    for typ in types_to_try:
        bridge_url = BASE + STD_PATH + "/showGb"
        rb = sess.get(bridge_url,
                      params={"type": typ, "hcno": hcno, "request_locale": "zh"},
                      headers={"Referer": detail_url},
                      timeout=60, allow_redirects=False)
        if rb.status_code != 200 or len(rb.content) < 1500:
            last_msg = "中转页异常 HTTP%s/%s字节" % (rb.status_code, len(rb.content))
            continue
        bridge_html = rb.text
        if "var isValid = 'false'" in bridge_html or "var isValid='false'" in bridge_html:
            last_msg = "该标准要求图形验证码，需浏览器人工下载"
            continue

        rv = sess.get(BASE + STD_PATH + "/viewGb", params={"hcno": hcno},
                      headers={"Referer": bridge_url + "?type=" + typ + "&hcno=" + hcno + "&request_locale=zh"},
                      timeout=180, allow_redirects=False)
        ctype = rv.headers.get("content-type", "").lower()
        if rv.status_code == 200 and (rv.content[:4] == b"%PDF" or "octet-stream" in ctype
                                      or "pdf" in ctype):
            if rv.content[:4] != b"%PDF" or len(rv.content) < 5000:
                last_msg = "viewGb返回非PDF或过小(%s字节)" % len(rv.content)
                continue
            with open(out_path, "wb") as f:
                f.write(rv.content)
            return {"status": "ok", "pdf_file": out_path, "method": "openstd_" + typ,
                    "message": "openstd官方公开系统下载成功（%s字节）；文本层若为乱码/扫描件由OCR管线处理" % len(rv.content),
                    "std_no": info.get("std_no", std_no), "title": info.get("title", ""),
                    "hcno": hcno, "source_url": detail_url, "bytes": len(rv.content)}
        last_msg = "viewGb未返回PDF: HTTP%s %s" % (rv.status_code, ctype[:40])

    return {"status": "failed", "message": last_msg or "官方通道未取得PDF（可能仅题录公开/采标件）",
            "std_no": std_no, "info": {k: v for k, v in info.items() if k != "url"},
            "detail_url": detail_url}
