# -*- coding: utf-8 -*-
"""
网页抓取与搜索模块 —— 让爬虫"不择手段"地拿到需要的标准和网页内容。

（2026-10-09 用 Scrapling 重做抓取/解析层，见下"改造说明"）

能力：
    1. search_web(query)：免API密钥的网页搜索（Google 浏览器点开取真实URL + Bing/百度
       精确URL），返回标题/链接/摘要列表。
    2. fetch_url(url)：抓取任意URL，自动识别内容类型：HTML→正文提取（Scrapling
       markdown）+ markdownify；PDF→交给 crawler_core.to_markdown（含OCR回退）；
       其他二进制→保存原始文件。
    3. fetch_with_fallback(urls)：按顺序尝试多个URL，第一个成功就返回。

改造说明（Scrapling 集成，硬校验铁律一律保留）：
    - 网络层用 Scrapling 的 `FetcherSession`（底层即 curl_cffi，带真实 Chrome TLS/JA3
      指纹、可选 http3、stealthy_headers），替代原先手搓的 curl_cffi 会话。
    - 新增"逐级升级"抓取阶梯：Scrapling HTTP → cloudscraper（Cloudflare/JS挑战）
      → StealthyFetcher（反检测无头浏览器，可解 Cloudflare Turnstile）
      → DynamicFetcher（Playwright 完整浏览器，渲染 JS 站）。
      浏览器两级仅在前面都失败、且环境中已安装浏览器依赖时才会启用；
      未安装则自动跳过，不报错（保证"没装浏览器也能跑"）。
    - HTML 正文提取改用 Scrapling 的 `Response.markdown(main_content_only=True)`：
      自动剔除 script/style/隐藏内容/疑似提示注入文本，产出更干净的 RAG-ready Markdown；
      过短时回退到全文文本。
    - 搜索结果解析改用 Scrapling 的 `Selector`（CSS 选择器 + get_all_text），
      比原先 BeautifulSoup 硬编码 class 更抗改版。
    - 全部既有硬校验保留不动：付费墙/盗版站拦截、候选源可信度分层排序、MD5 去重、
      %PDF 魔数/体积校验、封面标准号身份校验、PDF 体检、Referer 防盗链透传、
      四档随机延时（fetch/search/engine/retry）。

设计原则：
    - 所有网络请求都带真实浏览器指纹，躲开"请求不像真实浏览器"这类反爬检测。
    - Google 走无头浏览器，并用 Scrapling 的 `page_action`（真实 Playwright page）把
      不透明结果链接 `/goto?url=…` 新开标签点开、读回真实 URL（部分结果会被 Google
      掐断，退回 cite 域名级URL）；Bing/百度/DuckDuckGo 作为稳定补充取精确URL。
    - 证书：与项目既有约定一致，依赖环境变量 CURL_CA_BUNDLE 指向 ASCII 路径的
      cacert.pem（部署路径含中文会让 curl 加载 certifi 内置 CA 报 CURLE_SSL_CACERT(77)）。
"""

import io
import os
import re
import time
import random
import logging
import urllib.parse
from typing import Optional

from scrapling.parser import Selector

import crawler_core as cc

logger = logging.getLogger(__name__)


class _DropCloudflareNoise(logging.Filter):
    """Scrapling 在 solve_cloudflare=True 时，对没有 Cloudflare 挑战的普通页面
    也会打一条 ERROR 'No Cloudflare challenge found.'——属正常情况（只是没有挑战可解），
    这里静音掉，避免日志噪音掩盖真实错误。"""

    def filter(self, record):
        return "No Cloudflare challenge found" not in record.getMessage()


def _silence_scrapling_cf_noise():
    try:
        from scrapling.core.utils import log as _sc_log
        if not any(isinstance(f, _DropCloudflareNoise) for f in _sc_log.filters):
            _sc_log.addFilter(_DropCloudflareNoise())
    except Exception:
        pass


_silence_scrapling_cf_noise()


# ---------------------------------------------------------------------------
# Scrapling 底层封装
# ---------------------------------------------------------------------------

def _sc_session(timeout: int = 60, verify: bool = True, stealthy_headers: bool = True,
                impersonate: str = "chrome"):
    """创建 Scrapling 的 HTTP 会话（curl_cffi 内核 + 真实 Chrome TLS 指纹）。

    注意：FetcherSession 必须作为上下文管理器使用（with），否则内部 curl 会话未创建。
    """
    from scrapling.engines.static import FetcherSession
    return FetcherSession(impersonate=impersonate, timeout=timeout, verify=verify,
                          stealthy_headers=stealthy_headers, retries=2, retry_delay=1)


def _sc_get(url: str, params: Optional[dict] = None, referer: Optional[str] = None,
            timeout: int = 60, verify: bool = True, headers: Optional[dict] = None):
    """用 Scrapling HTTP 会话发一次 GET，返回 Response。

    传了 referer 时关闭 stealthy_headers（其默认会覆盖一套 Google Referer），
    改为显式设置 Referer + 成套 Sec-Fetch-*，满足政府/协会附件的防盗链校验。
    """
    hdrs = dict(headers or {})
    stealthy = True
    if referer:
        hdrs["Referer"] = referer
        try:
            same = urllib.parse.urlparse(referer).netloc.lower() == \
                   urllib.parse.urlparse(url).netloc.lower()
        except Exception:
            same = False
        hdrs["Sec-Fetch-Site"] = "same-origin" if same else "cross-site"
        hdrs["Sec-Fetch-Dest"] = "document"
        hdrs["Sec-Fetch-Mode"] = "navigate"
        stealthy = False   # 交给 curl_cffi 的 impersonate 生成其余头，避免覆盖我们的 Referer
    with _sc_session(timeout=timeout, verify=verify, stealthy_headers=stealthy) as s:
        return s.get(url, params=params, headers=hdrs or None, timeout=timeout, verify=verify)


def _wrap_response(url: str, data: bytes, status: int, headers: dict, encoding: str = "utf-8"):
    """把非 Scrapling 通道（cloudscraper）抓到的字节包成 Scrapling Response，
    这样上层处理（.markdown()/.css()/.body）可以统一走同一套代码。"""
    from scrapling.engines.toolbelt.custom import Response
    return Response(url=url, content=data, status=status, reason="", cookies={},
                    headers=dict(headers or {}), request_headers={},
                    encoding=encoding or "utf-8", method="GET")


_BROWSER_AVAILABLE = None


def _browser_available() -> bool:
    """环境里是否装了 Scrapling 的浏览器抓取依赖（playwright/patchright + 浏览器）。
    未装时返回 False，抓取阶梯自动跳过浏览器两级，不影响 HTTP 通道。"""
    global _BROWSER_AVAILABLE
    if _BROWSER_AVAILABLE is None:
        try:
            from scrapling.fetchers import StealthyFetcher, DynamicFetcher  # noqa: F401
            _BROWSER_AVAILABLE = True
        except Exception as exc:
            logger.info("浏览器抓取依赖不可用，抓取阶梯跳过 stealth/dynamic：%s", exc)
            _BROWSER_AVAILABLE = False
    return _BROWSER_AVAILABLE


def _sc_stealth_fetch(url: str, referer: Optional[str] = None,
                      timeout_ms: int = 45000, solve_cloudflare: bool = True,
                      proxy: Optional[str] = None, page_action=None):
    """StealthyFetcher：反检测无头浏览器，可绕过 Cloudflare Turnstile/Interstitial。

    page_action(page)：可选，传入一个接收 Playwright `page` 对象的回调，在导航完成后
    执行（Scrapling 原生支持）——用于点击/前进等真实浏览器交互（如解开 Google 结果链接）。
    """
    if not _browser_available():
        return None
    try:
        from scrapling.fetchers import StealthyFetcher
        kw = dict(headless=True, network_idle=True, timeout=timeout_ms,
                  solve_cloudflare=solve_cloudflare, google_search=not bool(referer))
        if referer:
            kw["extra_headers"] = {"Referer": referer}
        if proxy:
            kw["proxy"] = proxy
        if page_action is not None:
            kw["page_action"] = page_action
        return StealthyFetcher.fetch(url, **kw)
    except Exception as exc:
        logger.warning("StealthyFetcher 抓取失败 %s：%s", url, str(exc)[:160])
        return None


def _sc_dynamic_fetch(url: str, referer: Optional[str] = None, timeout_ms: int = 60000):
    """DynamicFetcher：Playwright 完整浏览器，抓取需执行 JS 才能渲染的页面。"""
    if not _browser_available():
        return None
    try:
        from scrapling.fetchers import DynamicFetcher
        kw = dict(headless=True, network_idle=True, timeout=timeout_ms,
                  google_search=not bool(referer))
        if referer:
            kw["extra_headers"] = {"Referer": referer}
        return DynamicFetcher.fetch(url, **kw)
    except Exception as exc:
        logger.warning("DynamicFetcher 抓取失败 %s：%s", url, str(exc)[:160])
        return None


def _ctype_of(headers) -> str:
    """从响应头（大小写不保证）里取 content-type。"""
    for k, v in (headers or {}).items():
        if k.lower() == "content-type":
            return (v or "").lower()
    return ""


def _try_cloudscraper(url: str, timeout: int = 60):
    """用 cloudscraper 绕过 Cloudflare 等反爬保护（HTTP 通道失败时的中间回退）。"""
    try:
        import cloudscraper
        scraper = cloudscraper.create_scraper(
            browser={"browser": "chrome", "platform": "windows", "mobile": False}
        )
        resp = scraper.get(url, timeout=timeout)
        resp.raise_for_status()
        return resp
    except Exception:
        return None


# 节奏集中管理：不同场景不同随机区间，避免固定周期特征
_SLEEP_RANGES = {
    "fetch":  (0.8, 2.5),    # 普通网页/文件抓取之间
    "search": (1.5, 3.5),    # 搜索引擎请求之间
    "engine": (3.0, 6.0),    # 切换搜索引擎/百度多轮查询之间
    "retry":  (5.0, 10.0),   # 被限流（429/403）后的冷却
}


def _sleep_rand(kind: str = "fetch"):
    lo, hi = _SLEEP_RANGES.get(kind, _SLEEP_RANGES["fetch"])
    time.sleep(random.uniform(lo, hi))


# 下载文件的最小合理体积（字节）。小于阈值的 PDF/DOC 基本是错误页/占位件
_MIN_DOC_BYTES = {"pdf": 50_000, "doc": 20_000, "docx": 5_000,
                  "xls": 10_000, "xlsx": 5_000}


def _first(sel_list):
    """从 Scrapling 的 Selectors（列表）里安全取第一个元素。"""
    try:
        return sel_list[0]
    except Exception:
        return None


def _text_of(node) -> str:
    try:
        return (node.get_all_text() or "").strip()
    except Exception:
        return ""


def _system_proxy() -> Optional[str]:
    """读取 Windows 系统代理（如 Clash/V2Ray 的 127.0.0.1:7897）。

    curl_cffi 不会自动读系统代理，而 Chromium 会自动读——所以直连 Google 被墙时，
    浏览器能到、HTTP 不能到。显式拿到代理地址喂给两条通道，Google 才可用。
    """
    try:
        import winreg
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                           r"Software\Microsoft\Windows\CurrentVersion\Internet Settings")
        enable, _ = winreg.QueryValueEx(k, "ProxyEnable")
        server, _ = winreg.QueryValueEx(k, "ProxyServer")
        if enable and server:
            if "://" not in server:
                server = "http://" + server
            return server
    except Exception:
        pass
    for ev in ("HTTPS_PROXY", "https_proxy", "HTTP_PROXY", "http_proxy"):
        if os.environ.get(ev):
            return os.environ[ev]
    return None


def _real_url_from_google(href: str) -> str:
    """Google 部分版式的结果链接形如 /url?q=<真实URL>&sa=...，可取 q 参数还原。

    注意：2026 起 Google 主版式已改成 `/goto?url=<不透明token>`，静态无法还原——这种
    链接由浏览器交互还原（见 `_google_search`/`_resolve_google_link`），此处只做
    显示用的 cite 兜底。"""
    if not href:
        return ""
    if "google.com/url" in href or href.startswith("/url?"):
        try:
            qs = urllib.parse.parse_qs(urllib.parse.urlparse(href).query)
            if qs.get("q"):
                href = qs["q"][0]
        except Exception:
            pass
    if href.startswith("/") or "google.com/goto" in href:
        return ""
    return href


def _cite_url(card) -> str:
    """从 Google 结果的 <cite>（展示URL，形如 `https://openstd.samr.gov.cn › bzgk › std`）重建真实URL。

    Google 已把点击链接换成不透明 token，唯一可用的真实地址就是 cite：
    `域名 + › 分隔的路径面包屑`。对 GitHub 这类「域+路径」纯结构站可精确还原。"""
    cite = _first(card.css("cite"))
    if cite is None:
        return ""
    txt = re.sub(r"\s+", " ", _text_of(cite)).strip()
    if not txt:
        return ""
    parts = [p.strip() for p in re.split(r"[›>]", txt) if p.strip()]
    if not parts:
        return ""
    # cite 里常混入标题文字，只取第一段的 URL 部分（空格前）
    base = parts[0].split(" ")[0]
    if not base.startswith("http"):
        base = "https://" + base
    if not re.match(r"^https?://[^\s/]+$", base):
        return ""
    base = base.rstrip("/")
    if len(parts) == 1:
        return base
    segs = [p.strip("/") for p in parts[1:] if p.strip("/")]
    # 面包屑可能是本地化中文（并非真实URL路径），此时只保留域名，避免拼出错误地址
    if not segs or any(not s.isascii() for s in segs):
        return base
    return base + "/" + "/".join(segs)


def _parse_google(page, sess=None) -> list[dict]:
    """解析 Google 搜索结果页（2026 现行结构）。

    结果卡片容器 `div.MjjYud`（旧的 `div.g` 已基本废弃）；标题是「a 包 h3」；
    点击链接是不可解读的 `/goto?url=token`——静态解析只能退回 cite 的域名级URL，
    故这里额外把原始 `/goto` 链接挂在 `_goto` 键上，交给浏览器交互去还原真实URL。"""
    results = []
    cards = page.css("div.MjjYud") or page.css("div.g")
    for card in cards:
        a = None
        for cand in card.css("a"):
            if cand.css("h3"):
                a = cand
                break
        if a is None:
            continue
        h3 = _first(a.css("h3"))
        title = _text_of(h3) if h3 is not None else _text_of(a)
        raw_href = (a.attrib or {}).get("href", "") or ""
        href = _real_url_from_google(raw_href)
        url = href if href.startswith("http") else _cite_url(card)
        url = url.split(" ")[0].strip()      # URL 不含裸空格，切掉混入的标题文字
        if not title or not url.startswith("http"):
            continue
        snippet = ""
        for cls in ("VwiC3b", "aCOpRe", "yXK7lf", "lEBKkf"):
            el = _first(card.css("div." + cls))
            if el is not None:
                snippet = _text_of(el)
                if snippet:
                    break
        if not snippet:
            texts = _text_of(card)
            if title in texts:
                snippet = texts.replace(title, "", 1).strip()[:200]
        item = {"title": title, "url": url, "snippet": snippet}
        # 浏览器解析用：Google 的不透明结果链接（需真浏览器交互才能还原）
        if raw_href and not raw_href.startswith("http") and "/goto?" in raw_href:
            item["_goto"] = raw_href
        results.append(item)
    return results


def _resolve_google_link(page, href: str, referer: str = "") -> str:
    """在真实浏览器里打开 Google 的不透明结果链接（`/goto?url=…`），读最终真实URL。

    这类链接由 Google 服务端在"带 Google 会话 Cookie + Referer 的真浏览器"发起时才解析出
    跳转；裸 HTTP 客户端（curl）解不开。做法：在浏览器上下文里新开标签导航该链接，
    `wait_until="commit"`（拿到响应即返回，不等正文下载完），读 `page.url` 即真实地址。
    实测部分结果仍会被 Google 掐断（ERR_CONNECTION_CLOSED），故重试一次、允许失败返回 ""。
    """
    if not href:
        return ""
    if href.startswith("http"):
        return href
    target = "https://www.google.com" + href
    for _ in range(2):
        np = None
        try:
            np = page.context.new_page()
            np.goto(target, timeout=10000, wait_until="commit",
                    referer=referer or "https://www.google.com/")
            u = np.url or ""
            if u.startswith("http"):
                return u
        except Exception as exc:
            logger.debug("解析 Google 结果链接失败：%s", str(exc)[:100])
        finally:
            try:
                if np is not None:
                    np.close()
            except Exception:
                pass
        time.sleep(random.uniform(1.0, 2.0))
    return ""


def _google_search(query: str, limit: int = 10) -> list[dict]:
    """Google 搜索（完整版）——走无头浏览器，并把结果链接在浏览器里点开还原真实URL。

    直连 Google 被墙时依赖系统代理（Chromium 自动读，另显式传入以防万一）。
    结果链接是 `/goto?url=<token>`：静态 HTML 里解不出来（cite 只有域名级），必须用
    Scrapling 的 `page_action` 拿到真实 Playwright page，逐条新开标签导航读最终 URL。
    成功还原的条目 `exact=True`；失败/未解析的退回 cite 域名级URL（`exact=False`）。"""
    if not _browser_available():
        return []
    proxy = _system_proxy()
    holder: list[dict] = []
    resolve_n = min(limit, 6)          # 只解析前若干条，控制耗时

    def _action(page):
        try:
            html = page.content()
        except Exception:
            return
        referer = page.url or "https://www.google.com/"
        for idx, item in enumerate(_parse_google(Selector(html))):
            goto = item.pop("_goto", "")
            if idx < resolve_n and goto:
                real = _resolve_google_link(page, goto, referer)
                if real:
                    item["url"] = real
                    item["exact"] = True
            holder.append(item)

    url = "https://www.google.com/search?q=%s&num=%d&hl=zh-CN" % (
        urllib.parse.quote(query), min(max(limit * 2, 10), 30))
    try:
        _sc_stealth_fetch(url, timeout_ms=30000, solve_cloudflare=False,
                          proxy=proxy, page_action=_action)
    except Exception as exc:
        logger.warning("Google 搜索失败：%s", str(exc)[:160])
    return holder


def _baidu_limited(resp) -> bool:
    """百度是否返回了安全验证/验证码页（短时间内查询过快触发限流）。

    命中时结果页里没有真正的搜索卡片，解析会得 0 条——必须区分"限流"与
    "真的没结果"，否则会误判为渠道失效。"""
    try:
        u = (getattr(resp, "url", "") or "")
        if "wappass.baidu.com" in u and "captcha" in u:
            return True
        title = resp.css("title::text").get() or ""
        return "百度安全验证" in title
    except Exception:
        return False


# ---------------------------------------------------------------------------
# 网页搜索
# ---------------------------------------------------------------------------

def _parse_duckduckgo(page, sess=None) -> list[dict]:
    """解析 DuckDuckGo HTML 搜索结果页。"""
    results = []
    snips = [_text_of(el) for el in page.css(".result__snippet")]
    idx = 0
    for a in page.css("a.result__a"):
        href = (a.attrib or {}).get("href", "")
        # DuckDuckGo 的链接是重定向，提取真实URL
        if "uddg=" in href:
            parsed = urllib.parse.urlparse(href)
            qs = urllib.parse.parse_qs(parsed.query)
            if "uddg" in qs:
                href = qs["uddg"][0]
        title = _text_of(a)
        if not title or not href.startswith("http"):
            continue
        snippet = snips[idx] if idx < len(snips) else ""
        idx += 1
        results.append({"title": title, "url": href, "snippet": snippet})
    return results


def _parse_bing(page, sess=None) -> list[dict]:
    """解析 Bing 搜索结果页。"""
    results = []
    for li in page.css("li.b_algo"):
        a = _first(li.css("h2 a"))
        if a is None:
            continue
        href = (a.attrib or {}).get("href", "")
        title = _text_of(a)
        if not title or not href.startswith("http"):
            continue
        p = _first(li.css(".b_caption p"))
        snippet = _text_of(p) if p is not None else ""
        results.append({"title": title, "url": href, "snippet": snippet})
    return results


def _parse_baidu(page, sess=None) -> list[dict]:
    """解析百度搜索结果页。百度链接是跳转链接，需要跟随重定向获取真实URL。"""
    results = []
    for item in page.css("div.result, div.c-container"):
        a = _first(item.css("h3 a"))
        if a is None:
            continue
        href = (a.attrib or {}).get("href", "")
        title = _text_of(a)
        if not title or not href:
            continue
        # 百度的链接是 http://www.baidu.com/link?url=... 跳转链接
        # 跟随重定向拿到真实URL（真假URL直接决定后续付费墙拦截/可信度分层是否命中）
        real_url = href
        if sess and "baidu.com/link" in href:
            try:
                resp = sess.get(href, timeout=10)
                if resp.url and resp.url != href:
                    real_url = resp.url
            except Exception:
                pass
        # 摘要
        snippet = ""
        for cls in ("c-abstract", "c-span-last", "content-right_8Zs40"):
            el = _first(item.css("." + cls))
            if el is not None:
                snippet = _text_of(el)
                if snippet:
                    break
        if not snippet:
            texts = _text_of(item)
            if title in texts:
                snippet = texts.replace(title, "", 1).strip()[:200]
        results.append({"title": title, "url": real_url, "snippet": snippet})
    return results


def _bing_search(query: str, limit: int = 10) -> list[dict]:
    """Bing 搜索：结果链接是可直接使用的真实URL（精确 URL 的主力引擎）。"""
    try:
        with _sc_session(timeout=30) as sess:
            resp = sess.get("https://www.bing.com/search",
                            params={"q": query, "count": str(min(max(limit * 2, 10), 30))})
        if resp.status == 200:
            return _parse_bing(resp)
    except Exception as exc:
        logger.debug("Bing 搜索失败：%s", exc)
    return []


def _baidu_search(query: str, limit: int = 10):
    """百度搜索：跳转链接跟随重定向取真实URL。返回 (results, limited)。"""
    try:
        with _sc_session(timeout=30) as sess:
            resp = sess.get("https://www.baidu.com/s",
                            params={"wd": query, "rn": str(min(max(limit * 2, 10), 30))})
            if _baidu_limited(resp):
                return [], True
            if resp.status == 200:
                return _parse_baidu(resp, sess=sess), False
    except Exception as exc:
        logger.debug("百度搜索失败：%s", exc)
    return [], False


def _duckduckgo_search(query: str, limit: int = 10) -> list[dict]:
    """DuckDuckGo HTML 版搜索（uddg 参数可还原真实URL），可能需翻墙。"""
    try:
        with _sc_session(timeout=30) as sess:
            resp = sess.post("https://html.duckduckgo.com/html/", data={"q": query})
        if resp.status == 200:
            return _parse_duckduckgo(resp)
    except Exception as exc:
        logger.debug("DuckDuckGo 搜索失败：%s", exc)
    return []


def _exact_url_search(query: str, limit: int = 10):
    """取"可下载候选"的精确URL搜索：Bing → 百度 → DuckDuckGo。

    Google 的结果链接是不透明令牌 `/goto?url=<token>`，静态解不出真网址（须靠浏览器
    点开还原，见 `_google_search`）；这里走另几个引擎作为稳定补充（Bing 实测直出
    openstd 详情页 hcno 与地方政府直链 PDF，百度靠跳转取真实URL）。返回 (results, engine, errors)。
    """
    errors = []
    results = _bing_search(query, limit=limit)
    if results:
        return results, "bing", errors
    errors.append("Bing: 返回空结果/非200")
    _sleep_rand("search")

    results, limited = _baidu_search(query, limit=limit)
    if limited:
        errors.append("百度: 命中百度安全验证（查询过快被限流），请冷却5-10秒后重试")
    elif results:
        return results, "baidu", errors
    else:
        errors.append("百度: 返回空结果/非200")
    _sleep_rand("search")

    results = _duckduckgo_search(query, limit=limit)
    if results:
        return results, "duckduckgo", errors
    errors.append("DuckDuckGo: 返回空结果/非200")
    return [], "none", errors


def search_web(query: str, limit: int = 10) -> dict:
    """免API密钥的网页搜索：**Google（浏览器点开取真实URL）+ Bing/百度/DDG 精确URL**。

    Google（完整版，走无头浏览器 + 系统代理）结果最全、无验证码。其结果链接是
    不透明令牌 `/goto?url=…`——静态解析拿不到真网址，但用 Scrapling 的 `page_action`
    在真实浏览器里点开即可还原（见 `_resolve_google_link`），还原成功的条目 `exact=True`；
    少数被 Google 掐断的退回 cite 域名级URL（`exact=False`）。Bing（→百度→DuckDuckGo）
    作为稳定补充，提供可直接抓取的精确URL。

    返回：{"engine": "google+bing" 等, "query": ..., "total": N,
           "results": [{"title", "url", "snippet", "via", "exact"}, ...]}
    （`results` 中精确URL条目在前，仅有域名级的 Google 覆盖条目在后。）
    """
    errors = []

    # 1) Google：浏览器点开取真实URL（部分成功），拿不到真链的退回域名级覆盖
    coverage = []
    if _browser_available():
        try:
            coverage = _google_search(query, limit=limit)
        except Exception as exc:
            errors.append("Google: %s" % str(exc)[:100])
        _sleep_rand("search")

    # 2) 精确URL：Bing → 百度 → DuckDuckGo（稳定补充）
    exact, exact_engine, errs = _exact_url_search(query, limit=limit)
    errors.extend(errs)

    # 3) 合并（按域名去重，优先级：Google 已还原真实URL → Bing/百度 精确URL → Google 域名级覆盖）
    merged, doms = [], set()

    def _add(it, via, exact):
        d = _domain_of(it.get("url", ""))
        if d and d in doms:
            return
        if d:
            doms.add(d)
        merged.append({**it, "via": via, "exact": exact})

    for it in coverage:
        if it.get("exact"):
            _add(it, "google", True)
    for it in exact:
        _add(it, exact_engine, True)
    for it in coverage:
        if not it.get("exact"):
            _add(it, "google", False)

    if merged:
        engines = []
        if coverage:
            engines.append("google")
        if exact:
            engines.append(exact_engine)
        return {"engine": "+".join(engines) or "google", "query": query,
                "total": len(merged[:limit]), "results": merged[:limit]}
    return {"engine": "none", "query": query, "total": 0, "results": [],
            "error": "; ".join(errors)}


# ---------------------------------------------------------------------------
# 网页抓取
# ---------------------------------------------------------------------------

def _extract_html_main_content(page, base_url: str = "") -> str:
    """把 HTML 页面转成干净 Markdown。

    优先用 Scrapling 的 markdown(main_content_only=True)：自动剔除
    script/style/noscript/svg 及 CSS 隐藏内容、疑似提示注入文本；过短时回退全文文本。
    """
    md = ""
    try:
        md = page.markdown(main_content_only=True) or ""
    except Exception as exc:
        logger.debug("markdown 转换失败：%s", exc)
    # markdown 过短（<200）时回退到整页纯文本
    if len(md.strip()) < 200:
        raw = ""
        try:
            raw = page.get_all_text() or ""
        except Exception:
            raw = ""
        if len(re.sub(r"\s+", "", raw)) > len(re.sub(r"\s+", "", md)):
            md = raw

    try:
        title = (page.css("title::text").get() or "").strip()
    except Exception:
        title = ""
    if title:
        md = f"# {title}\n\n{md}"
    if base_url:
        md = f"> 来源：{base_url}\n\n{md}"
    return cc.normalize_md_text(md)


def _do_request(url: str, params: Optional[dict] = None, referer: Optional[str] = None,
                timeout: int = 60, verify: bool = True, mode: str = "auto"):
    """按"逐级升级"阶梯抓取一个URL，返回 (Response, method, error)。

    mode:
      - "http"   ：只用 Scrapling HTTP + cloudscraper（快，禁浏览器）
      - "auto"   ：HTTP → cloudscraper → StealthyFetcher → DynamicFetcher（默认）
      - "stealth"：HTTP → StealthyFetcher（反爬站优先）
      - "dynamic"：HTTP → DynamicFetcher（JS 站优先）
    返回的 Response 统一为 Scrapling Response（云端通道也会被包成 Response）。
    """
    attempts = []

    # 1) Scrapling HTTP（curl_cffi + Chrome TLS 指纹）
    try:
        resp = _sc_get(url, params=params, referer=referer, timeout=timeout, verify=verify)
        if resp.status not in (403, 503, 429):
            return resp, "scrapling_http", None
        attempts.append("http:%s" % resp.status)
    except Exception as exc:
        attempts.append("http:%s" % str(exc)[:80])

    # 2) cloudscraper（绕过 Cloudflare JS 挑战；HTTP 被拦或报错时）
    _sleep_rand("retry")
    cs_resp = _try_cloudscraper(url, timeout=timeout)
    if cs_resp is not None:
        try:
            return (_wrap_response(url, cs_resp.content, cs_resp.status_code,
                                   cs_resp.headers, cs_resp.encoding),
                    "cloudscraper", None)
        except Exception as exc:
            attempts.append("cloudscraper:%s" % str(exc)[:60])
    else:
        attempts.append("cloudscraper:blocked/failed")

    if mode == "http":
        return None, None, " | ".join(attempts)

    # 3) StealthyFetcher（反检测无头浏览器，可解 Cloudflare Turnstile）
    if mode in ("auto", "stealth"):
        r = _sc_stealth_fetch(url, referer=referer)
        if r is not None and r.status not in (403, 503, 429):
            return r, "scrapling_stealth", None
        attempts.append("stealth:failed/blocked")

    # 4) DynamicFetcher（Playwright 完整浏览器，渲染 JS）
    if mode in ("auto", "dynamic"):
        r = _sc_dynamic_fetch(url, referer=referer)
        if r is not None:
            return r, "scrapling_dynamic", None
        attempts.append("dynamic:failed")

    return None, None, " | ".join(attempts)


def _looks_like_pdf(data: bytes) -> bool:
    """PDF 魔数检查：%PDF 出现在文件头 1KB 内（允许少量前导字节）。"""
    return data[:1024].lstrip().startswith(b"%PDF")


def _validate_doc_bytes(data: bytes, kind: str) -> Optional[str]:
    """下载件真伪/完整性前置校验。返回 None=通过，否则返回拒收原因。

    教训：政府/协会站直链在防盗链或会话失效时会回一个 HTML 错误页，
    Content-Type 却仍可能是 application/pdf——只看响应头会把错误页当 PDF 落盘。
    """
    kind = kind.lower()
    if kind == "pdf":
        if not _looks_like_pdf(data):
            head = data[:80].decode("utf-8", "ignore").replace("\n", " ")[:80]
            return "响应不是PDF（魔数%%PDF缺失，疑似错误页/登录页）：%s" % head
        if len(data) < _MIN_DOC_BYTES["pdf"]:
            return "PDF体积过小（%d 字节 < %d），疑似预览残页或占位件" % (
                len(data), _MIN_DOC_BYTES["pdf"])
    elif kind in _MIN_DOC_BYTES and len(data) < _MIN_DOC_BYTES[kind]:
        return "%s 体积过小（%d 字节 < %d），疑似下载不完整" % (
            kind.upper(), len(data), _MIN_DOC_BYTES[kind])
    return None


def fetch_url(url: str, save_path: Optional[str] = None,
              referer: Optional[str] = None, convert: bool = True,
              mode: str = "auto") -> dict:
    """抓取任意URL并转换为可用文本/文件。

    自动识别内容类型：
        - text/html → Scrapling 提取正文（markdown）
        - application/pdf → 魔数+体积校验 → crawler_core.to_markdown（含OCR回退）
        - 其他二进制 → 保存原始文件到 save_path

    referer: 来源页URL。部分政府/协会站（如水利工程协会附件）校验防盗链，
    必须带列表页 Referer 才返回真实文件。
    convert: False 时 PDF/Office 文档只做魔数校验+落盘、不做全文转换（扫描件
        可省去整篇 OCR），供"先封面身份校验/体检通过再转换"的流水线使用；
        HTML 提取不受影响。默认 True 保持旧行为。
    mode: 抓取阶梯模式，见 _do_request。默认 "auto"（HTTP→cloudscraper→浏览器）。

    返回：{"status": "ok"/"failed", "content_type": ..., "text": 提取的文本(HTML/PDF),
           "saved_file": 保存的文件路径或None, "error": ...,
           "method": 实际生效的通道（ok时）}
    """
    resp, method, err = _do_request(url, referer=referer, timeout=60, verify=True, mode=mode)
    if resp is None:
        return {"status": "failed", "error": "全部抓取通道失败: %s" % (err or "未知原因")}

    ctype = _ctype_of(resp.headers)
    data = resp.body or b""

    # HTML
    if "text/html" in ctype or "application/xhtml" in ctype:
        # URL 声称是 PDF，实际回了 HTML——多半是错误页/登录页，明确判失败，
        # 交上层换下一个候选源，而不是把错误页正文当标准文本
        if url.lower().split("?")[0].endswith(".pdf"):
            return {"status": "failed", "content_type": ctype,
                    "error": "PDF直链返回了HTML页面（疑似防盗链/登录失效），需带Referer或换源"}
        text = _extract_html_main_content(resp, url)
        return {"status": "ok", "content_type": ctype, "text": text, "saved_file": None,
                "method": method}

    # PDF（魔数+体积双校验，通过后才转换/落盘）
    if "pdf" in ctype or url.lower().endswith(".pdf"):
        bad = _validate_doc_bytes(data, "pdf")
        if bad:
            return {"status": "failed", "content_type": ctype, "error": bad}
        if not convert:
            # 延迟转换路径：只落盘，全文转换（含扫描件整篇 OCR）由上层
            # 在封面身份校验+体检通过后再做，错配件不再白花 OCR
            saved = None
            if save_path:
                with open(save_path, "wb") as f:
                    f.write(data)
                saved = save_path
            return {"status": "ok", "content_type": ctype, "text": "",
                    "saved_file": saved, "method": method, "bytes": len(data),
                    "converted": False}
        try:
            text = cc.to_markdown(data, "doc.pdf", ctype)
        except Exception as exc:
            return {"status": "failed", "content_type": ctype,
                    "error": f"PDF转换失败: {exc}"}
        saved = None
        if save_path:
            with open(save_path, "wb") as f:
                f.write(data)
            saved = save_path
        return {"status": "ok", "content_type": ctype, "text": text, "saved_file": saved,
                "method": method, "bytes": len(data), "converted": True}

    # 其他文档（docx/xlsx/doc/xls/html/md/txt）
    ext_match = re.search(r"\.(docx?|xlsx?|xlsm|html?|md|markdown|txt|text)(\?|$)", url, re.I)
    fname = urllib.parse.unquote(url.rsplit("/", 1)[-1].split("?")[0])
    if ext_match:
        kind = ext_match.group(1).lower().lstrip(".")
        bad = _validate_doc_bytes(data, kind)
        if bad:
            return {"status": "failed", "content_type": ctype, "error": bad}
    if ext_match or fname.lower().endswith((".doc", ".docx", ".xls", ".xlsx", ".xlsm",
                                            ".html", ".htm", ".md", ".markdown", ".txt")):
        if convert:
            try:
                text = cc.to_markdown(data, fname, ctype)
            except Exception as exc:
                text = ""
        else:
            text = ""
        saved = None
        if save_path:
            with open(save_path, "wb") as f:
                f.write(data)
            saved = save_path
        return {"status": "ok", "content_type": ctype or (ext_match.group(1) if ext_match else ctype),
                "text": text, "saved_file": saved, "method": method, "bytes": len(data),
                "converted": bool(convert)}

    # 其他二进制：只保存，不转换
    saved = None
    if save_path:
        with open(save_path, "wb") as f:
            f.write(data)
        saved = save_path
    return {"status": "ok", "content_type": ctype, "text": "", "saved_file": saved,
            "method": method, "bytes": len(data),
            "message": "二进制文件已保存，未提取文本"}


def fetch_with_fallback(urls: list[str], save_dir: Optional[str] = None,
                        referer: Optional[str] = None, mode: str = "auto") -> dict:
    """按顺序尝试多个URL，第一个成功就返回。用于"多条搜索结果挨个试"的场景。

    referer 会透传给每个请求（同站附件防盗链场景）。
    """
    errors = []
    for i, url in enumerate(urls):
        save_path = None
        if save_dir:
            import os
            os.makedirs(save_dir, exist_ok=True)
            raw_name = urllib.parse.unquote(url.rsplit("/", 1)[-1].split("?")[0])
            fname = cc.safe_filename(raw_name) or f"file_{i}"
            save_path = os.path.join(save_dir, fname)
        result = fetch_url(url, save_path=save_path, referer=referer, mode=mode)
        if result.get("status") == "ok":
            result["tried"] = i + 1
            return result
        errors.append(f"[{url}] {result.get('error', '未知错误')}")
        _sleep_rand("fetch")
    return {"status": "failed", "error": "全部URL抓取失败: " + " | ".join(errors)}


# ===========================================================================
# 免费标准获取渠道（针对国内收费但在国外/镜像站可免费获取的场景）
# ===========================================================================

# ---------------------------------------------------------------------------
# 渠道分层（项目渠道纪律：只走官方/协会/公开存档，不碰付费墙与盗版文档站）
# ---------------------------------------------------------------------------

# 付费墙/盗版文档站：只记录其搜索结果供人工判断，绝不自动抓取下载
# （这类站点正文需登录/付费，抓到的只会是预览残页或跳转HTML，且有合规风险）
_BLOCKED_DOC_SITES = [
    "doc88.com",        # 道客巴巴
    "docin.com",        # 豆丁
    "max.book118.com",  # 原创力文档
    "book118.com",
    "wenku.baidu.com",  # 百度文库
    "ishare.iask.com",  # 爱问共享资料
    "renrendoc.com",    # 人人文档
    "guifanku.com",     # 规范库（已被项目拉黑）
    "bzwx.com",
    "bzfxw.com",        # 标准分享网（水印/引流站）
    "antpedia.com",     # 蚂蚁标准网（付费/引流）
    "chem900.com",
    "bookzz.org",       # 影子图书馆，不自动抓取
]

# 可信公开渠道：官方、政府、行业协会、标准组织、公开历史存档
# 命中这些域名（含子域）的候选优先级最高
_TRUSTED_DOMAIN_KEYWORDS = (
    # 国家标准官方
    "openstd.samr.gov.cn", "std.samr.gov.cn", "samr.gov.cn", "sac.gov.cn",
    "hbba.sacinfo.org.cn", "std.sacinfo.org.cn",
    # 中央部委/政府公开
    "mohurd.gov.cn", "mem.gov.cn", "mee.gov.cn", "gov.cn",
    # 行业协会/标准化组织（本轮免费攻坚实测有效）
    "cwec.org.cn",       # 中国水利工程协会（SL 送审稿/正式稿）
    "cnass.cn",          # 中国安全生产协会
    "cesa.org.cn",       # 中国节能协会
    "cgs.gov.cn",        # 地质调查局
    "nea.gov.cn",        # 能源局
    "chinagb.com",       # 国标查询（目录/元数据）
    # 公开存档
    "web.archive.org", "archive.org",
    "gitee.com/raw",     # Gitee raw 直链
    "raw.githubusercontent.com",
)

# 次可信：开源存档仓库页面（非 raw，需再解析一层）
_SECONDARY_DOMAIN_KEYWORDS = (
    "github.com", "gitee.com",
)


def _domain_of(url: str) -> str:
    try:
        return urllib.parse.urlparse(url).netloc.lower()
    except Exception:
        return ""


def _is_blocked_url(url: str) -> bool:
    """命中付费墙/盗版文档站域名（含子域）。"""
    dom = _domain_of(url)
    if not dom:
        return False
    return any(dom == b or dom.endswith("." + b) for b in _BLOCKED_DOC_SITES)


def _domain_tier(url: str) -> int:
    """候选源可信度分层，数值越小越优先。

    0=官方/协会/政府/公开存档直链  1=开源raw直链  2=开源仓库页面  9=普通网页
    """
    dom = _domain_of(url)
    # raw 直链按精确域/子域匹配，不做宽松子串（防止 evil 域仿冒）
    if dom == "raw.githubusercontent.com" or dom.endswith(".raw.githubusercontent.com"):
        return 1
    for k in _TRUSTED_DOMAIN_KEYWORDS:
        if k in ("gitee.com/raw", "raw.githubusercontent.com"):
            continue
        if dom == k or dom.endswith("." + k):
            return 0
    for k in _SECONDARY_DOMAIN_KEYWORDS:
        if dom == k or dom.endswith("." + k):
            return 2
    return 9


# 兼容旧名（保留 Internet Archive 定向搜索能力，付费站已剔除）
_DOC_SHARE_SITES = ["archive.org"]

# 国外可免费获取中国标准的站点/镜像
_FREE_STD_SITES = [
    "openstd.samr.gov.cn",  # 国家标准全文公开系统（GB国标官方免费）
    "hbba.sacinfo.org.cn",  # 行业标准平台（部分预览免费）
    "github.com",           # GitHub 上有大量标准存档仓库
    "gitee.com",            # Gitee 镜像
]


def search_doc_share_sites(std_no: str, title: str = "", limit: int = 5) -> list[dict]:
    """对公开存档站做定向搜索（百度 site: 限定）。

    渠道纪律：道客巴巴/豆丁/百度文库/book118 等付费墙站点已从自动抓取中
    剔除（见 _BLOCKED_DOC_SITES）；当前只保留 Internet Archive 这类公开历史存档。
    返回 [{title, url, snippet, site}, ...] 列表。
    """
    query = f'"{std_no}" {title}'.strip()
    results = []
    for site in _DOC_SHARE_SITES:
        try:
            q = f"{query} site:{site}"
            with _sc_session(timeout=20) as sess:
                resp = sess.get("https://www.baidu.com/s", params={"wd": q, "rn": "20"})
                items = _parse_baidu(resp, sess=sess) if resp.status == 200 else []
            for it in items[:limit]:
                it["site"] = site
                results.append(it)
            if len(results) >= limit * 3:
                break
            _sleep_rand("search")
        except Exception as e:
            logger.debug("doc share search %s 失败：%s", site, e)
            continue
    return results[:limit * 2]


def search_github_standard(std_no: str, title: str = "", limit: int = 10) -> list[dict]:
    """在 GitHub 上搜索标准号，找 PDF/MD 存档。

    GitHub 有大量标准全文仓库（如 'Chinese-Standard-Download'、'GB-Standard' 等）。
    免认证的 Code Search API 不可用（返回 401），所以实际策略是：
    1) 用百度 site:github.com 搜索标准号，找到匹配的仓库文件
    2) 同时用 GitHub Repositories Search API（免认证可用，限速 10/分钟）
       找标准存档仓库

    返回 [{title, url, snippet, repo, file_path, raw_url}, ...]
    """
    std_clean = std_no.replace("/", " ").replace("—", "-").replace("／", " ").strip()
    q = f'"{std_clean}"'.strip()
    if title:
        # 标题取前 8 个字，避免查询过长
        q = f'{q} {title[:8]}'

    results = []

    def _ingest(items):
        for it in items[:limit]:
            url = it.get("url", "")
            if "github.com" not in url:
                continue
            # 尝试推断 raw_url（github.com/owner/repo/blob/... → raw.githubusercontent.com/owner/repo/...）
            raw_url = ""
            if "/blob/" in url:
                raw_url = url.replace("github.com", "raw.githubusercontent.com") \
                              .replace("/blob/", "/", 1)
            results.append({
                "title": it.get("title", ""),
                "url": url,
                "raw_url": raw_url,
                "snippet": it.get("snippet", ""),
                "repo": _extract_repo_from_github_url(url),
                "file_path": _extract_path_from_github_url(url),
                "site": "github.com",
            })
            if len(results) >= limit * 2:
                break

    # 1) site:github.com 搜索（需精确URL：Google 的结果链接是不透明令牌、
    #    取不到真实网址，故走 Bing/百度；百度限流时如实记录）
    items, _eng, _errs = _exact_url_search(f"{q} site:github.com", limit=20)
    _ingest(items)
    for e in _errs:
        logger.debug("GitHub 站内搜索：%s", e)

    _sleep_rand("search")

    # 2) GitHub Repositories Search API（找标准存档仓库，免认证可用）
    try:
        api = "https://api.github.com/search/repositories"
        with _sc_session(timeout=30) as sess2:
            resp = sess2.get(api, params={"q": f"{std_clean} standard OR 标准",
                                          "per_page": 5},
                             headers={"Accept": "application/vnd.github+json"})
        if resp.status == 200:
            data = resp.json()
            for item in data.get("items", [])[:5]:
                results.append({
                    "title": item.get("full_name", ""),
                    "url": item.get("html_url", ""),
                    "snippet": (item.get("description", "") or "")[:200],
                    "repo": item.get("full_name", ""),
                    "site": "github.com",
                })
    except Exception as e:
        logger.debug("GitHub repo search 失败：%s", e)

    return results[:limit]


def _extract_repo_from_github_url(url: str) -> str:
    """从 github.com/owner/repo/blob/... 提取 owner/repo。"""
    m = re.search(r"github\.com/([^/]+/[^/]+)", url)
    return m.group(1) if m else ""


def _extract_path_from_github_url(url: str) -> str:
    """从 github.com/owner/repo/blob/branch/path 提取 path。"""
    m = re.search(r"/blob/[^/]+/(.+)$", url)
    return m.group(1) if m else ""


def search_wayback(url: str, limit: int = 3) -> list[dict]:
    """查询 Wayback Machine 上的历史快照。

    用途：原URL失效（404/删除）时，从 web.archive.org 取历史版本。
    走 Wayback CDX API（免认证）。
    """
    try:
        api = "http://archive.org/wayback/available"
        with _sc_session(timeout=20) as sess:
            resp = sess.get(api, params={"url": url})
        if resp.status == 200:
            data = resp.json()
            snap = data.get("archived_snapshots", {}).get("closest", {})
            if snap.get("available"):
                return [{"title": f"Wayback 快照 {snap.get('timestamp', '')}",
                          "url": snap.get("url", ""),
                          "snippet": f"原URL的历史快照，时间 {snap.get('timestamp', '')}",
                          "site": "web.archive.org"}]
    except Exception as e:
        logger.debug("Wayback 查询失败：%s", e)
    return []


def acquire_standard_via_free_channels(std_no: str, title: str = "",
                                        save_dir: Optional[str] = None,
                                        max_attempts: int = 10,
                                        mode: str = "auto") -> dict:
    """【不花钱获取收费标准】综合搜索+抓取。

    渠道纪律（自动抓取顺序按域名可信度排序，官方/协会优先）：
    1. 官方/政府/行业协会公开页（openstd、mohurd、cwec、cnass 等）
    2. GitHub/Gitee 标准存档（raw 直链优先）
    3. Internet Archive 历史快照
    4. 其他普通网页
    付费墙/盗版文档站（道客巴巴/豆丁/百度文库/book118 等）只记录到
    blocked_candidates 供人工判断，绝不自动抓取。
    PDF 下载后强制：封面标准号身份校验（防错配）+ 魔数/体积校验 +
    页数/双页拼版/末页完整体检，不合格自动删件换源。

    mode 会透传给 fetch_url（默认 "auto"：HTTP→cloudscraper→Stealthy→Dynamic）。

    返回：{"status": "ok"/"failed", "text": 抓到的文本,
           "source_url": 实际成功的URL, "source_type": 来源类型,
           "tried": [尝试记录], "candidates": [可抓取候选URL],
           "blocked_candidates": [命中付费墙被跳过的URL]}
    """
    tried = []
    candidates = []  # [(url, source_type, save_path_hint)]
    blocked_candidates = []  # 付费墙/盗版站：只记录，不抓取

    def _add_candidate(url: str, source_type: str, hint: str = ""):
        if not url or url in [c[0] for c in candidates] \
                 or url in blocked_candidates:
            return
        if _is_blocked_url(url):
            blocked_candidates.append(url)
        else:
            candidates.append((url, source_type, hint))

    # 构建相关性检查的关键词
    # 标准号提取数字部分（如 "DL/T 596-2021" → ["596", "2021"]）
    # 标题提取关键中文词（如 "电力设备预防性试验规程" → ["预防性试验", "试验规程", "电力设备"]）
    relevance_keywords = _build_relevance_keywords(std_no, title)

    # 1) GitHub/Gitee 存档（raw 直链最干净）
    gh_results = search_github_standard(std_no, title, limit=10)
    for it in gh_results:
        # 优先 raw_url（直接下载文件），没有则用 html_url
        url = it.get("raw_url") or it.get("url")
        _add_candidate(url, "github", it.get("file_path", ""))

    # GitHub 搜索用了百度，archive 定向搜索也要用百度，中间加长延迟避免限流
    if gh_results:
        _sleep_rand("engine")

    # 2) Internet Archive 公开存档（付费文档站已按渠道纪律剔除）
    ds_results = search_doc_share_sites(std_no, title, limit=5)
    for it in ds_results:
        _add_candidate(it.get("url"), "archive", "")

    # 搜索也用了百度，web_search 之前加延迟
    if ds_results:
        _sleep_rand("engine")

    # 3) 直接搜索标准号（不带 site:，扩大范围）
    # 用引号包住完整标准号，防止前缀被误解（如 "DL/T" 匹配到 "DLsite"）。
    # 注意：**不要**再追加 "标准 规程" 之类后缀——实测会把 "GB" 误判成存储单位
    # （gigabyte），结果清一色"存储容量换算"，政府直链 PDF 全部沉底。
    std_no_quoted = f'"{std_no}"'
    web_query = f'{std_no_quoted} {title}'.strip()
    web_results = search_web(web_query, limit=10)
    for it in web_results.get("results", []):
        # Google 覆盖项（exact=False）只有域名级URL（cite 丢失 path/query），
        # 抓不到文档、只会浪费尝试预算，不进候选
        if it.get("exact") is False:
            continue
        _add_candidate(it.get("url"), "web_search", "")

    if not candidates:
        return {"status": "failed", "error": "未找到任何可抓取候选源（付费墙站点已跳过）",
                "tried": tried, "source_type": None, "candidates": [],
                "blocked_candidates": blocked_candidates}

    # 4) 按域名可信度稳定排序：官方/协会 → raw直链 → 仓库页 → 普通站，
    #    同级保持搜索结果原顺序。这样 max_attempts 的预算优先花在正规源上。
    candidates.sort(key=lambda c: _domain_tier(c[0]))

    # 5) 逐个抓取，前 max_attempts 个（已按可信度排序）
    seen_hashes = set()  # 本次抓取内容去重，同一字节流不重复落盘
    html_fallback = None  # 网页（非文档）兜底：可能只是"标准信息页/转载页"，未必是全文
    for url, source_type, hint in candidates[:max_attempts]:
        tried.append({"url": url, "type": source_type})
        try:
            save_path = None
            if save_dir:
                import os
                os.makedirs(save_dir, exist_ok=True)
                # 文件名一律"标准号 标题.扩展名"：URL尾段先 URL 解码再只取扩展名，
                # 杜绝 %E5%BB%BA... 编码名 / "XX相关标准.pdf" 这类无标准号文件名
                raw = urllib.parse.unquote(url.rsplit("/", 1)[-1].split("?")[0])
                ext = os.path.splitext(raw)[1].lower()
                if ext not in (".pdf", ".doc", ".docx", ".md", ".markdown", ".txt", ".zip", ".rar"):
                    ext = ".bin"
                base = ("%s %s" % (std_no, title)).strip() if title else std_no
                fname = cc.safe_filename(base + ext)
                save_path = os.path.join(save_dir, fname)
            r = fetch_url(url, save_path=save_path, convert=False, mode=mode)
            if r.get("status") != "ok":
                tried[-1]["error"] = r.get("error", "抓取失败")[:100]
                _sleep_rand()
                continue

            text = r.get("text", "")
            saved_file = r.get("saved_file")

            # 相关性检查：抓到的文本必须包含标准号数字 或 标题关键词之一
            # 文档共享站的 HTML 页面通常包含标准号（在标题/简介里），可以通过
            # 但纯垃圾内容（如博客、导航站）会被过滤
            is_relevant = _check_relevance(text, relevance_keywords)

            # 如果是 PDF/MD 直链下载（saved_file 有值且是文档类型），即使文本提取少也接受
            is_direct_doc = bool(saved_file and any(
                saved_file.lower().endswith(ext) for ext in
                [".pdf", ".md", ".markdown", ".doc", ".docx", ".txt"]
            ))

            if is_direct_doc:
                # 同字节流去重（同一汇编件常被多个候选URL重复返回）
                import hashlib
                try:
                    fdata = open(saved_file, "rb").read()
                except OSError:
                    fdata = b""
                digest = hashlib.md5(fdata).hexdigest() if fdata else None
                if digest and digest in seen_hashes:
                    try:
                        os.remove(saved_file)
                    except OSError:
                        pass
                    tried[-1]["error"] = "与已下载内容字节相同，去重跳过"
                    _sleep_rand()
                    continue
                # PDF 直链也要核对封面身份，防止"搜 A 标准、下载到 B 标准"
                # （教训：JGJ59-2011 扫描件曾被 8 个不同 JGJ 查询全部误收）
                is_pdf = bool(saved_file and saved_file.lower().endswith(".pdf"))
                detail = "非PDF直链"
                health = None
                if is_pdf:
                    ok_id, detail = _verify_pdf_identity(saved_file, std_no)
                    if not ok_id:
                        try:
                            os.remove(saved_file)
                        except OSError:
                            pass
                        tried[-1]["error"] = "身份校验失败: %s" % detail
                        _sleep_rand()
                        continue
                    # 封面身份通过后做轻量体检：页数过少=预览/截断件，拒收；
                    # 双页拼版/末页空白/扫描件只警告不阻塞，交人工与转换环节处理
                    health = cc.pdf_health_check(saved_file)
                    if health.get("error") is None and health["pages"] < 3:
                        try:
                            os.remove(saved_file)
                        except OSError:
                            pass
                        tried[-1]["error"] = "完整性拒收: 仅%d页（%s）" % (
                            health["pages"], "；".join(health.get("warnings", [])))
                        _sleep_rand()
                        continue
                if not is_pdf and len(text.strip()) < 200 \
                        and not _check_relevance(text, relevance_keywords):
                    # 非PDF直链（.md/.txt/.doc/.docx）没有封面身份校验兜底，
                    # 必须防住 404/错误占位页被当成正文收下
                    # （实测：Google 重建的 GitHub 链接可能已失效，返回 "404: Not Found"）
                    try:
                        os.remove(saved_file)
                    except OSError:
                        pass
                    tried[-1]["error"] = "非PDF直链内容过短且不相关（疑似404/错误页）"
                    _sleep_rand()
                    continue
                if digest:
                    seen_hashes.add(digest)
                # 身份/去重/体检全部通过后，此刻才做全文转换：
                # 扫描件整篇 OCR（分钟级）只花在通过件上，错配件最多损失封面3页OCR
                try:
                    text = cc.to_markdown(
                        fdata, os.path.basename(saved_file),
                        r.get("content_type", "")) if fdata else ""
                except Exception as exc:
                    text = ""
                    tried[-1]["convert_warning"] = str(exc)[:100]
                return {"status": "ok", "text": text,
                        "saved_file": saved_file, "is_document": True,
                        "source_url": url, "source_type": source_type,
                        "identity_check": detail,
                        "health_check": health,
                        "tried": tried, "candidates": [c[0] for c in candidates],
                        "blocked_candidates": blocked_candidates}
            if is_relevant and len(text.strip()) >= 200:
                # HTML 页面：可能只是"标准信息页/转载页/新闻页"，未必是标准全文。
                # 先记为兜底、继续往后找真正的文档（PDF/MD 直链）；只有确实找不到
                # 文档时，才退回这个网页结果（并明确标注 is_document=False）。
                if html_fallback is None:
                    html_fallback = {
                        "status": "ok", "text": text,
                        "saved_file": saved_file, "is_document": False,
                        "source_url": url, "source_type": source_type,
                        "identity_check": "网页（非标准文档，可能是信息页/转载页）",
                        "health_check": None,
                        "tried": tried, "candidates": [c[0] for c in candidates],
                        "blocked_candidates": blocked_candidates}
                tried[-1]["error"] = "已记录网页兜底，继续查找正式文档"
            else:
                tried[-1]["error"] = "内容相关性不足或文本过短"
        except Exception as e:
            tried[-1]["error"] = str(e)[:100]
        _sleep_rand()

    # 5) 最后试 Wayback（如果原URL有线索）
    if candidates:
        wb = search_wayback(candidates[0][0])
        for it in wb:
            url = it.get("url")
            if not url:
                continue
            tried.append({"url": url, "type": "wayback"})
            try:
                r = fetch_url(url, mode=mode)
                if r.get("status") == "ok" and r.get("text", "").strip():
                    text = r.get("text", "")
                    if _check_relevance(text, relevance_keywords):
                        return {"status": "ok", "text": text,
                                "source_url": url, "source_type": "wayback",
                                "tried": tried, "candidates": [c[0] for c in candidates],
                                "blocked_candidates": blocked_candidates}
                    else:
                        tried[-1]["error"] = "Wayback内容相关性不足"
            except Exception as e:
                tried[-1]["error"] = str(e)[:100]

    # 6) 没拿到正式文档，但抓到了相关网页——退回网页兜底（明确标注非文档）
    if html_fallback is not None:
        return html_fallback

    # 7) 全部失败——返回候选列表供用户手工处理
    return {"status": "failed",
            "error": "所有候选源抓取失败或内容相关性不足",
            "tried": tried, "source_type": None,
            "candidates": [c[0] for c in candidates],
            "blocked_candidates": blocked_candidates,
            "message": "付费墙/盗版文档站链接不自动抓取；blocked_candidates "
                       "仅为命中黑名单的候选URL，需人工自行判断获取方式"}


def _build_relevance_keywords(std_no: str, title: str = "") -> list[str]:
    """构建相关性检查关键词列表。

    - 从标准号提取数字部分（596、2021）
    - 从标题提取关键中文词组（2-4字）
    """
    import re
    kws = []
    # 标准号里的数字（年份和序号，3位以上保留：596、2021、408 等）
    for m in re.findall(r"\d+", std_no):
        if len(m) >= 3:
            kws.append(m)
    # 标题关键词（中文 2-4 字词组）
    if title:
        # 简单分词：取 2-4 字的连续中文片段
        for m in re.findall(r"[\u4e00-\u9fa5]{2,6}", title):
            if len(m) >= 2:
                kws.append(m)
    # 去重，保留顺序
    seen = set()
    out = []
    for k in kws:
        if k not in seen:
            seen.add(k)
            out.append(k)
    return out


def _check_relevance(text: str, keywords: list[str]) -> bool:
    """检查文本是否包含足够多的关键词。

    严格判定规则（避免单一数字误匹配，如 Steam 页面碰巧含 "2021"）：
    - 至少 1 个标准号数字 + 至少 1 个标题关键词，或
    - 至少 3 个标题关键词
    """
    if not keywords or not text:
        return False
    hit_numbers = 0
    hit_words = 0
    for kw in keywords:
        if kw.isdigit():
            if kw in text:
                hit_numbers += 1
        else:
            if kw in text:
                hit_words += 1
    # 严格规则：数字+标题词同时命中，或3个以上标题词
    return (hit_numbers >= 1 and hit_words >= 1) or hit_words >= 3


def _std_cover_pattern(std_no: str):
    """把标准号编译成"封面容错"正则，容忍 OCR/排版变体。

    例：JGJ 82-2011 能匹配封面/OCR出的 "JGJ82—2011"、"J G J 8 2-2011"；
    GB/T 150.1-2024 能匹配 "GB/T 150.1-2024"、"GBT150.1—2024"；
    GB 50058-2014 能匹配 OCR 误读的 "GB.50058-2014"。无法解析时返回 None。
    """
    s = std_no.strip().upper().replace("／", "/")
    m = re.match(r"^([A-Z]{1,4})(?:/T)?\s*([0-9]+(?:\.[0-9]+)*)\s*[-—–~]\s*(\d{4})$", s)
    if not m:
        return None
    prefix, serial, year = m.groups()
    p = r"\s*".join(list(prefix))
    # GB/T 的 T：允许 GB/T、GBT、GB T 三种写法
    p += r"(?:[\s./．/／]*T)?" if "/" in std_no.upper().replace("／", "/") else r"(?:[\s/／]*T)?"
    segs = []
    for i, seg in enumerate(serial.split(".")):
        if i:
            segs.append(r"[\s.．]*")
        segs.append(r"\s*".join(list(seg)))
    # 字母与序号之间允许句点/空格（OCR 常见 "GB.50058"）
    pat = p + r"[\s.．]*" + "".join(segs)
    pat += r"[^0-9A-Za-z]{0,6}" + r"\D{0,2}".join([year[:2], year[2:]])
    return re.compile(pat, re.IGNORECASE)


def _verify_pdf_identity(pdf_path: str, std_no: str, max_pages: int = 3):
    """核对下载到的 PDF 封面前几页印的标准号是否就是查询的标准号。

    扫描件无文字层时自动 OCR 前几页再判。
    - 明确不一致 → (False, 识别到的标准号)
    - 匹配 / 无法判断（格式不支持、OCR不可用）→ (True, 原因)，不阻塞
    """
    pat = _std_cover_pattern(std_no)
    if pat is None:
        return True, "标准号格式不在封面校验范围，跳过"
    try:
        import fitz
    except ImportError:
        return True, "pymupdf不可用，跳过封面校验"
    try:
        doc = fitz.open(pdf_path)
        text = "\n".join(doc[i].get_text() for i in range(min(max_pages, doc.page_count)))
        doc.close()
    except Exception as e:
        return False, "PDF无法打开: %s" % e

    def hit(t):
        return bool(pat.search(t) or pat.search(re.sub(r"[ \t]+", "", t)))

    if hit(text):
        return True, "封面标准号匹配"
    # 文字层几乎没有中文 → 按扫描件处理，OCR 前几页
    if sum(1 for c in text if "一" <= c <= "鿿") < 20:
        try:
            import ocr_engine as oe
            ocr_text = oe.ocr_pdf_bytes(open(pdf_path, "rb").read(),
                                       dpi=200, max_pages=max_pages)
            if hit(ocr_text):
                return True, "封面OCR标准号匹配"
            found = re.findall(r"[A-Z]{2,4}\s*T?\s*\d{2,4}(?:\.\d+)?\s*[-—–~]\s*20\d{2}",
                               ocr_text.upper())
            return False, "封面标准号不匹配（件内识别到: %s）" % (
                ",".join(sorted(set(x.replace(" ", "") for x in found))[:3]) or "无标准号")
        except Exception as e:
            return True, "OCR校验不可用，放行: %s" % e
    return False, "文字层封面未找到标准号 %s" % std_no