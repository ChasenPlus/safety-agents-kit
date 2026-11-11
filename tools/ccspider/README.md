# ccspider —— 行业标准批量下载 / 文档转 Markdown（独立小工具）

不依赖本项目其他代码，单独拷走也能用。适合没有 Python 环境、只想下几份标准的现场场景。

| 文件 | 说明 |
|:---|:---|
| `ccspider.py` | 源码，约 550 行，只依赖下列第三方库 |
| `ccspider.exe` | PyInstaller 单文件打包版（Windows x64），双击即用，无需 Python。**体积约 45 MB，不随仓库分发**，到 [Releases](../../releases/latest) 下载，或按下面的命令自己打包 |

> 仓库里只有源码，克隆不会带上 exe。自己打包一份即可，命令见「依赖与重新打包」。

## 用法

双击 `ccspider.exe`（从 Releases 下载），或在装了依赖的机器上 `python ccspider.py`：

```
============================== 标准文档处理工具 ==============================
  1) 下载行业标准并转 md
  2) 本地目录文件转 md
  3) 对比两个文件
  0) 退出
```

- **1) 下载行业标准**：输入行业中文名或代码，多个用 `=` 分隔（如 `电力=安全生产`）。脚本分页拉取该行业全部标准记录，逐条下载并转成 Markdown，输出到 `程序目录/DownloadMD/{行业名}/`，文件名形如 `DL/T 5001-2014 火力发电厂...md`。
- **2) 本地目录转 md**：把目录下 `.doc/.docx/.pdf/.xls/.xlsx/.html/.txt` 递归转成 Markdown，输出到 `程序目录/TOMD/{YYYYMMDD}/`，保持原目录相对结构。
- **3) 对比两个文件**：按字节比较内容是否完全一致，用于校验下载件与原件。

下载行为：单线程，每次请求后随机停 1~3 秒，失败重试 2 次，已存在的文件自动跳过（可断点续跑）。下载接口返回空文件时写一条占位说明，不会静默丢文件。

## 依赖与重新打包

```bash
pip install curl_cffi python-docx pypdf openpyxl xlrd markdownify
pip install pyinstaller
pyinstaller -F -n ccspider ccspider.py --hidden-import=_cffi_backend --collect-all curl_cffi
```

`--collect-all curl_cffi` 与 `--hidden-import=_cffi_backend` 不能省：curl_cffi 自带 libcurl 二进制，漏了会打出「启动即报找不到模块或 DLL」的废包。

源码用 `curl_cffi` 的浏览器指纹伪装（`impersonate="chrome"`）访问目标站点，比普通 `requests` 更稳；不要手动改 `User-Agent`，否则 UA 版本与 TLS 指纹不一致反而更容易被识别。

## 与本项目 `sak-crawler` 的关系

| | `tools/ccspider`（本目录） | `server/crawler`（MCP 服务） |
|:---|:---|:---|
| 形态 | 交互式命令行 / 双击即用 | stdio MCP 服务，供智能体调用 |
| 抓取层 | curl_cffi 指纹伪装 | Scrapling 分级（HTTP → cloudscraper → 反检测浏览器） |
| 额外能力 | 本地目录批量转 md、文件比对 | 封面校验、OCR、PDF 体检、RAG 入库 |

两者共用同一套「按扩展名选转换器」的逻辑。要在智能体里用，选 `sak-crawler`；要给不写代码的同事批量下标准，用本目录的 exe。

## 合规提示

- 目标站点是行业标准信息服务平台对外公开的检索页面，本工具只做**公开信息的批量检索与格式转换**，不绕过登录、不破解权限。
- 标准文本的著作权归发布机构所有，下载件仅供内部合规检索与学习使用，不得再分发或商用；引用条款时请以正式出版物为准。
- 请控制访问频率（脚本已内置 1~3 秒随机间隔）。若站点规则或 robots 不允许，请停止使用并改为官方渠道采购。
