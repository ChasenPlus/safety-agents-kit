# 标准文档爬虫 MCP Server

## 文件说明

- `download_std.py` — 你提供的原始脚本，交互式命令行工具，保留作为存档/参照，本目录下其他文件不依赖它运行。`crawler_core.py` 是基于它重构的非交互式版本，已逐项核对过关键实现细节（分页查询上限、下载重试次数、文件名安全处理、限速延迟时机等）与原始脚本保持一致。
- `crawler_core.py` — 去交互化重构版核心库。转换逻辑与原脚本基本一致；新增了 `search_industry_records()`（只查询不下载）和 `download_one_standard()`（精确下载单条）两个函数；网络请求层用 `curl_cffi`（`impersonate="chrome"`）带真实浏览器TLS指纹，规避部分网站对"看起来不像浏览器"的请求的拦截。
- `web_fetcher.py` — 网页搜索与任意URL抓取（**Google→百度→Bing→DuckDuckGo** 搜索、HTML正文提取、免费渠道获取标准）。**2026-10-09 用 Scrapling 重做抓取/解析层**：
  - 网络层改为 Scrapling `FetcherSession`（底层即 curl_cffi，带真实 Chrome TLS/JA3 指纹 + stealthy headers）。
  - 新增「逐级升级」抓取阶梯 `_do_request`：**Scrapling HTTP → cloudscraper（Cloudflare/JS挑战）→ StealthyFetcher（反检测无头浏览器，可解 Cloudflare Turnstile）→ DynamicFetcher（Playwright 渲染JS）**；浏览器两级仅在环境装了浏览器依赖时启用，未装则自动跳过、不报错。
  - **搜索首选 Google（完整版）**：直连被墙，普通 HTTP 也拿不到结果，故走无头浏览器 + **系统代理**（自动读 Windows 注册表 `ProxyEnable/ProxyServer`，如 Clash 的 `127.0.0.1:7897`）。Google 结果卡片是 `div.MjjYud`，点击链接已换成不可解读的 `/goto?url=token`，真实URL改由 `<cite>`（`域名 › 路径面包屑`）重建（GitHub 这类纯「域+路径」站可精确还原）。Google 不可用时自动降级到 百度→Bing→DuckDuckGo。
  - HTML 正文提取改用 `Response.markdown(main_content_only=True)`（自动剔除脚本/隐藏内容/疑似提示注入），过短时回退全文。
  - 搜索结果解析改用 Scrapling `Selector`（CSS 选择器），比原先 BeautifulSoup 硬编码 class 更抗改版。
  - **全部硬校验铁律原样保留**：付费墙/盗版站拦截（`_BLOCKED_DOC_SITES`）、候选源可信度分层（`_domain_tier`）、MD5去重、`%PDF`魔数/体积校验、封面标准号身份校验、PDF体检、Referer防盗链透传、四档随机延时。
  - `fetch_url` / `fetch_with_fallback` / `acquire_standard_via_free_channels` 均新增 `mode` 参数：`http`/`auto`/`stealth`/`dynamic`。
- `gb_standard.py` — 国家标准全文公开系统（openstd.samr.gov.cn）国标下载。
- `ocr_engine.py` / `image_extract.py` — 扫描件OCR、图片抽取（rapidocr-onnxruntime，含300DPI/水印/双页拼版处理）。
- `sak_crawler_mcp.py` — 用 FastMCP 包装成MCP server，暴露 `list_industries`/`search_standards`/`download_standard`/`download_industry_bulk`/`convert_local_dir`/`compare_files`/`search_web`/`fetch_url`/`acquire_standard`/`acquire_standard_free`/`diagnose_ocr`/`check_pdf_file` 工具。
- `diagnose_crawler.py` — 综合诊断与回归测试脚本（OCR诊断、搜索结果渠道测试、免费获取全流程测试）。
- `requirements.txt` — 依赖清单。

## 安装与启动

```bash
pip install -r requirements.txt

# 一次性：下载 Scrapling 的反检测浏览器（StealthyFetcher / DynamicFetcher 用）。
# 不装也能跑——抓取阶梯会自动跳过浏览器两级，只走 HTTP + cloudscraper。
scrapling install

python sak_crawler_mcp.py
```

**证书注意**：部署路径含中文时，curl_cffi 加载 certifi 内置 CA 会报 `CURLE_SSL_CACERT(77)`，须通过环境变量 `CURL_CA_BUNDLE` 指向 ASCII 路径的 `cacert.pem`（`sak.json` 里已配置）。

## 独立测试（不经过宿主 IDE，直接验证核心逻辑）

```python
import crawler_core as cc
result = cc.search_industry_records("电力", keyword="高处作业", limit=5)
print(result)

import web_fetcher as wf
print(wf.fetch_url("https://openstd.samr.gov.cn/bzgk/std/std_list?p.p1=0&p.p2=GB%2030871", mode="auto"))
```

或用综合诊断脚本：`python diagnose_crawler.py`。

## 关于本目录不对外分发

这里的代码是你自己写的、经过验证的私有资产（尤其是对 hbba.sacinfo.org.cn 接口的调用方式），不要跟着公开的agents/skills一起分享给外部。