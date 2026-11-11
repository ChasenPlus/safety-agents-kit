# MCP 服务清单

所有服务都是 **stdio + 换行分隔 JSON-RPC 2.0**，符合 MCP 规范，可被任何 MCP 宿主复用。

启动方式统一为：

```bash
python server/<服务名>/server.py
```

配置写进 `sak.json` 的 `mcp_servers`，或任何 MCP 宿主的 `mcpServers` 段。

---

## sak-rag —— 本地语义检索

完全离线：向量召回 + 可选 cross-encoder 精排，索引落在本地。

| 工具 | 用途 |
|:---|:---|
| `rag_search(query, layers=None, top_k=8)` | 语义检索，返回路径 + 片段 + 行号 |
| `rag_index_status()` | 查看各层文件数、chunk 数、待索引数 |
| `rag_upsert_file(path)` | 单文件入库（新增/更新） |
| `rag_delete_file(path)` | 删除索引条目 |
| `rag_reindex(full=False)` | 增量或全量重建 |

环境变量：`SAK_KB_ROOT`、`SAK_RAG_MODEL_DIR`、`SAK_RAG_INDEX_DIR`、`SAK_DIRECTORY_XLSX`。

---

## sak-private —— 核验与制度匹配

| 工具 | 用途 |
|:---|:---|
| `check_coverage(question, candidate_clause)` | R11 覆盖充分性核验：指标 / 频次 / 责任主体 / 适用范围 / 操作细节 五个维度 |
| `match_regulation(question)` | 判断问题属于哪份制度（返回候选与置信度） |

设 `ANTHROPIC_API_KEY` 时可对启发式判断做 LLM 二次精化（可选，不设也能跑）。

---

## sak-crawler —— 标准库管理

| 工具 | 用途 | 开销 |
|:---|:---|:---|
| `list_industries()` | 支持的行业分类 | 低 |
| `search_standards(industry, keyword, limit)` | 只查不下载，拿候选 | 低 |
| `download_standard(industry, pk, std_no, title, out_root)` | 精确下载单条并转 MD | 中 |
| `download_industry_bulk(industries, out_root)` | 整行业批量建库 | 高 |
| `acquire_standard_free(std_no, title, out_root)` | 官方公开渠道抓取，带身份校验与体检 | 高 |
| `check_pdf_file(path, min_pages)` | 整目录 PDF 体检（页数/双页拼版/截断/加密） | 低 |
| `convert_local_dir(src_dir, program_root)` | 本地文件批量转 Markdown | 中 |
| `compare_files(path_a, path_b)` | 两份文件差异对比 | 低 |
| `search_web(query, limit)` / `fetch_url(url, save_dir, mode)` | 网页检索与抓取 | 中 |
| `diagnose_ocr()` | OCR 环境自检 | 低 |

抓取遵循「官方 / 协会 / 政府公开页优先」，付费墙与盗版站不抓，不做验证码绕过。
浏览器依赖缺失时自动降级为 HTTP 抓取，不报错。

---

## sak-report —— 正式文档生成

| 工具 | 用途 |
|:---|:---|
| `get_data_schema()` | 拿整改通知单的字段说明 |
| `generate_rectification_notice(data)` | 生成 docx |
| `postprocess_docx(path)` | 版式后处理（字体、行距、编号） |

---

## sak-handoff —— 分支提示词分发

| 工具 | 用途 |
|:---|:---|
| `list_branches()` | 列出分支、职责、提示词字数 |
| `get_branch_prompt(branch)` | 取某分支提示词全文 |
| `get_handoff_packet(branch, task)` | 生成可直接投喂子代理的交接包 |

CLI 用户不需要这个服务，`sak handoff` 已内置同样能力。

---

## sak-fs / sak-excel —— 第三方服务

直接使用社区实现，不重复造：

```json
"fs":    { "command": "npx", "args": ["-y", "@modelcontextprotocol/server-filesystem", "./kb"] },
"excel": { "command": "npx", "args": ["-y", "@negokaz/excel-mcp-server"] }
```

需要 Node.js 18+。若不想装 Node，可只用 `skills/excel-ledger-writer` 的脚本写台账。
