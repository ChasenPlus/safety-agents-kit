# Safety Agents Kit（SAK）

把「八分支安全合规智能体」从某个 IDE 里解耦出来：**提示词即资产，MCP 即工具，模型随便换**。

- 任意模型：DeepSeek / 通义千问 / Kimi / 智谱 / OpenAI / Claude / Gemini / 本地 Ollama、vLLM
- 任意宿主：终端 CLI、你自己的 Python 程序、任何支持 MCP 的 IDE 或聊天客户端
- 零核心依赖：`pip install -e .` 之后即可跑，模型调用只用标准库
- 工具标准化：本地检索（RAG）、制度归属匹配、R11 覆盖核验、标准下载入库、整改通知单生成，全部是标准 MCP 服务

> 面向电力/能源行业的安全生产合规场景：条款检索与判定、巡查检查报告、制度修订、标准库管理、
> 现场工作包、专项方案、公文材料、定责考核台账。仓库内所有企业名、人名、路径均为示例占位值。

---

## 为什么要有这个项目

原来的做法是把每个功能写成一个「自定义智能体」，粘在某个商业 IDE 里。问题很直接：

| 问题 | 后果 |
|:---|:---|
| 提示词绑在平台里 | 换 IDE 就得重贴一遍，团队无法用 Git 管理 |
| 模型绑在平台里 | 想换更便宜的国产模型、或想本地离线跑，做不到 |
| 工具绑在平台里 | 检索、核验、下载能力无法被第三方复用 |
| 经验散在对话里 | 谁踩过的坑，别人下次还得再踩 |

SAK 把这四层拆开：**分支提示词是 Markdown 文件，工具是 MCP 服务，编排是 300 行 Python，模型是适配器**。
四者都可以单独替换、单独贡献。

## 与 IDE 无关的架构

```
        你（CLI / Python / 任意 MCP 客户端）
                     │
              ┌──────▼──────┐
              │  sak 编排层  │  分支提示词 + 工具循环 + 最大步数保护
              └──┬───────┬──┘
                 │       │
      ┌──────────▼─┐   ┌─▼──────────────┐
      │ 模型适配器  │   │  MCP 客户端     │  stdio / JSON-RPC 2.0
      │ OpenAI兼容 │   └─┬───┬───┬───┬──┘
      │ Anthropic  │     │   │   │   │
      │ Gemini     │   rag private fs excel crawler report
      │ mock       │     │   │   │   │
      └────────────┘   ┌─▼───▼───▼───▼─┐
                       │  你的知识库 / 文件 │
                       └────────────────┘
```

提示词只描述「怎么判断、按什么格式输出」，工具只描述「能查到什么、能生成什么」。
两者都不认识任何特定模型或 IDE。

## 八条分支

| ID | 分支 | 职责 | 典型输入 |
|:--:|:---|:---|:---|
| A | 条款检索与判定 | 标准条款原文检索、R11 覆盖充分性核验、制度缺口判定 | `临时用电配电箱未设漏电保护，合规吗？` |
| B | 巡查检查报告 | 检查文件/现场问题，出问题清单与整改通知单 | `检查 应急预案.docx` |
| C | 制度修订起草 | 按制度编号起草修订建议书，回写自查台账 | `2112` |
| D | 标准库管理 | 标准查询、下载、转 MD、归位、入检索索引 | `查询 电力 高处作业` |
| E | 现场工作包生成 | 安全/技术交底、应急处置卡、考核标准、规范卡、培训卡 | `带电作业设备检修` |
| F | 专项方案编制 | 施工组织方案（十章）、危大工程专项方案（九章） | 可研/作业描述 |
| G | 写材料发文 | 安全周报、月报、季度报、安委会材料 | `季度报` |
| H | 定责考核与台账 | 制度缺口台账、修订门禁、追责考核、Wiki 写回 | 缺口信息 / 制度编号 |

分支之间可以互相调用（B 调 A 查依据、E 调 F 出方案、A 调 D 下载标准），
跨分支调用在 CLI 里就是 `sak handoff -b A "任务"` 生成一份「交接包」交给子代理。

## 五分钟上手

### 1. 安装

```bash
git clone https://github.com/your-org/safety-agents-kit.git
cd safety-agents-kit
pip install -e .
```

只想要 CLI 而不装包，也可以直接 `python -m sak`。

### 2. 初始化工作目录

```bash
sak init my-workspace
cd my-workspace
```

会生成 `sak.json`（配置）、`kb/`（知识库骨架）、`.env.example`。

### 3. 填模型

编辑 `sak.json`，或直接用环境变量：

```bash
# DeepSeek
export SAK_PROVIDER=openai
export SAK_MODEL=deepseek-chat
export SAK_BASE_URL=https://api.deepseek.com/v1
export DEEPSEEK_API_KEY=sk-xxxx

# 本地 Ollama（完全离线）
export SAK_PROVIDER=openai
export SAK_MODEL=qwen2.5:14b
export SAK_BASE_URL=http://localhost:11434/v1
export SAK_API_KEY=ollama
```

Claude 用 `SAK_PROVIDER=anthropic`，Gemini 用 `SAK_PROVIDER=gemini`，其余见 `sak/llm/__init__.py` 的 `PRESETS`。

### 4. 先不花钱，验证链路

```bash
sak doctor
sak run -b A --mock "临时用电配电箱未设漏电保护，合规吗？"
```

`--mock` 不联网，只验证提示词加载、MCP 启动、工具往返是否正常。

### 5. 正式跑

```bash
sak run -b A "GB 50194 对配电箱漏电保护有什么要求？"
sak run -b B -f 检查说明.txt -o 报告.md
sak chat -b H          # 多轮对话
sak branches           # 看全部分支
```

## 模型适配

| provider | 适用 | 说明 |
|:---|:---|:---|
| `openai` | DeepSeek、Qwen、Kimi、智谱、OpenAI、vLLM、Ollama、任何 OpenAI 兼容网关 | 改 `base_url` 即可，无需改代码 |
| `anthropic` | Claude | Messages API + tool use |
| `gemini` | Gemini | generateContent + functionDeclarations |
| `mock` | 自测 | 零网络，用于 CI |

要接一个全新的供应商，只需实现 `sak/llm/base.py` 里的 `LLMClient.chat()`，
返回 `AssistantTurn(content, tool_calls)` 即可，30 行左右。

## 工具（MCP）

`server/` 下都是标准 stdio MCP 服务，既能被 SAK 驱动，也能挂进任何 MCP 宿主：

| 服务 | 能力 | 依赖 |
|:---|:---|:---|
| `sak-rag` | 本地语义检索（向量召回 + 可选精排，完全离线） | numpy / onnxruntime / tokenizers |
| `sak-private` | R11 覆盖充分性核验、制度归属匹配 | fastmcp |
| `sak-crawler` | 标准检索、下载、体检、转 Markdown | fastmcp / scrapling / pymupdf |
| `sak-report` | 整改通知单等正式 docx 生成 | mcp / python-docx |
| `sak-fs` | 本地文件读写 | `npx @modelcontextprotocol/server-filesystem` |
| `sak-excel` | Excel 台账读写 | `npx @negokaz/excel-mcp-server` |
| `sak-handoff` | 分支提示词分发（CLI 的 `sak handoff` 等价物） | fastmcp |

各服务单独装依赖，互不影响：

```bash
pip install -e ".[servers]"     # private / report / excel 相关
pip install -e ".[rag]"         # 本地检索
pip install -e ".[crawler]"     # 标准下载
```

详见 [docs/mcp-servers.md](docs/mcp-servers.md)。

## 知识库

`kb_template/` 是一套可直接用的知识库骨架：

```
00-导航层/      按主题、工作类型、标准层级找入口
01-Wiki层/      规则与经验（00-必读 是各分支开工必读的公共规则）
02-原始资料层/  国家标准、行业标准、企业制度、隐患图集等原文
03-输出层/      交付物落盘位置
04-脚本代码/    门禁、台账、转换等可执行脚本
```

`sak init` 会把它复制成 `kb/`。把标准原文放进 `02-原始资料层/`，再建一次索引即可被检索到。

## 文档

- [架构说明](docs/architecture.md) —— 分层、数据流、扩展点
- [部署与配置](docs/deployment.md) —— 模型、MCP、知识库、常见故障
- [分支与权限矩阵](docs/branch-permissions.md) —— 每个分支需要哪些服务
- [MCP 服务清单](docs/mcp-servers.md) —— 工具签名与依赖
- [贡献指南](CONTRIBUTING.md)

## 自测

```bash
python tests/test_smoke.py
```

不需要任何 API Key：分支加载、配置解析、MCP 启动与调用、编排循环全覆盖。

## 路线图

- [ ] 会话持久化与断点续跑
- [ ] 分支自动路由（LLM 先判断该走哪条分支）
- [ ] 检索层支持外部向量库（Chroma / Qdrant / pgvector）
- [ ] 更多行业的标准源适配（建筑、化工、矿山）
- [ ] Web UI（可选，复用同一套 MCP 服务）

## 免责声明

本项目提供的是**辅助判断工具**，输出结果不构成法律意见或合规结论。
正式引用标准条款必须回到标准原文逐字核对；涉及具体项目、具体责任认定时，
请由具备资质的专业人员复核。仓库内所有企业、项目、人名均为示例，不代表任何真实主体。

## 许可证

[MIT](LICENSE)
