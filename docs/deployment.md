# 部署与配置

## 配置文件

`sak` 按以下顺序找配置，找到第一个就用（后面的不再合并）：

1. `-c/--config` 显式指定
2. 当前目录 `sak.json`
3. `~/.sak/config.json`
4. 仓库自带 `config.example.json`（仅默认值）

环境变量优先级高于配置文件：`SAK_PROVIDER`、`SAK_MODEL`、`SAK_BASE_URL`、
`SAK_API_KEY`、`SAK_KB_ROOT`、`SAK_MAX_STEPS`。

`sak.json` 里的相对路径（`kb_root`、`prompts_dir`、MCP 的脚本参数）都**相对配置文件所在目录**解析。

```json
{
  "provider": "openai",
  "model": "deepseek-chat",
  "base_url": "https://api.deepseek.com/v1",
  "api_key_env": "DEEPSEEK_API_KEY",
  "temperature": 0.2,
  "max_tokens": 4096,
  "max_steps": 24,
  "kb_root": "./kb",
  "prompts_dir": "./prompts",
  "mcp_servers": {
    "sak-fs": { "command": "npx", "args": ["-y", "@modelcontextprotocol/server-filesystem", "./kb"] },
    "sak-rag": { "command": "python", "args": ["server/rag/server.py"] }
  }
}
```

`command` 写成 `python` / `python3` 时会自动替换为当前解释器，避免 PATH 里没有 python。

## 常见模型配置

### DeepSeek

```bash
SAK_PROVIDER=openai
SAK_MODEL=deepseek-chat
SAK_BASE_URL=https://api.deepseek.com/v1
DEEPSEEK_API_KEY=sk-xxx
```

### 通义千问（DashScope 兼容模式）

```bash
SAK_MODEL=qwen-max
SAK_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1
DASHSCOPE_API_KEY=sk-xxx
```

### 本地 Ollama（离线）

```bash
SAK_MODEL=qwen2.5:14b
SAK_BASE_URL=http://localhost:11434/v1
SAK_API_KEY=ollama
```

### Claude

```bash
SAK_PROVIDER=anthropic
SAK_MODEL=claude-sonnet-4-20250514
SAK_BASE_URL=https://api.anthropic.com
ANTHROPIC_API_KEY=sk-ant-xxx
```

### Gemini

```bash
SAK_PROVIDER=gemini
SAK_MODEL=gemini-2.0-flash
GEMINI_API_KEY=xxx
```

## MCP 服务依赖

| 服务 | 安装 |
|:---|:---|
| fs / excel | 需要 Node.js 18+，`npx` 会按需拉包 |
| private / agent_switch | `pip install fastmcp`（1.x） |
| report | `pip install mcp python-docx` |
| rag | `pip install numpy onnxruntime tokenizers` + 向量模型 |
| crawler | `pip install fastmcp scrapling pymupdf` |

RAG 的模型目录通过 `SAK_RAG_MODEL_DIR` 指定，默认找 `server/rag/models/`。
首次使用要先建索引：

```bash
python server/rag/rag_core.py reindex
python server/rag/rag_core.py status
```

## 知识库准备

```bash
sak init my-workspace       # 生成 kb/ 骨架
```

把待检索的资料放进 `kb/02-原始资料层/`（标准、制度、预案），
把团队规则放进 `kb/01-Wiki层/`，然后重建索引。

## 排错

先跑自检：

```bash
sak doctor
```

它会检查配置来源、密钥、知识库目录、8 份提示词是否可读，并逐个试启动 MCP 服务。

| 现象 | 原因 | 处理 |
|:---|:---|:---|
| `模型接口返回 401` | 密钥未设置或环境变量名与 `api_key_env` 不一致 | `sak doctor` 看 API Key 一行 |
| `MCP 服务 xxx 启动失败` | 依赖没装 / 脚本路径错 | 单独终端执行 `command args` 看真实报错 |
| 工具调用一直不收敛 | 提示词要求的工具没挂上，模型在空转 | 检查 `branches/X.json` 的 `servers` 与 `sak.json` 的键名是否一致 |
| 检索结果为空 | 索引没建 / `kb_root` 指错 | `rag_core.py status` |
| 中文乱码（Windows） | 控制台编码 | 设置 `PYTHONUTF8=1`，代码已默认给子进程注入 |
