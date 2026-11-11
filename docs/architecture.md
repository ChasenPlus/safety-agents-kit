# 架构说明

## 四层分离

| 层 | 位置 | 换掉它意味着 | 不该放什么 |
|:---|:---|:---|:---|
| 提示词层 | `prompts/*.md` | 换一套业务规则、换行业 | 不写工具实现细节、不写模型名称 |
| 编排层 | `sak/` | 换交互形态（CLI→服务→GUI） | 不写业务判断规则 |
| 模型层 | `sak/llm/` | 换模型供应商 | 不感知 MCP、不感知分支 |
| 工具层 | `server/` | 换检索方案、换数据源 | 不写提示词、不做业务结论 |

边界只有一条：**提示词说「查什么、怎么判、怎么输出」，工具说「我有什么」。**

## 一次调用的数据流

```
1. CLI 解析 -b A "问题"
2. 读 branches/A.json  ->  servers: [rag, private, fs]
3. 按 sak.json 的 mcp_servers 启动这三个 stdio 子进程
4. tools/list 汇总成统一工具清单（重名工具自动加 server__ 前缀）
5. system = prompts/A_clause_retrieval.md + 运行时上下文（知识库路径、日期）
6. 循环：
     messages -> LLMClient.chat(system, messages, tools)
                  ├─ 有 tool_calls：逐个 tools/call，把结果作为 role=tool 追加，回到 6
                  └─ 无 tool_calls：返回 content，结束
7. 达到 max_steps 仍未收敛则强制停止并回报最后一段输出
```

## 关键设计取舍

**为什么自己写 MCP 客户端（约 150 行）**
不依赖官方 SDK，避免 `mcp` / `fastmcp` 的版本互相钉死（现实里这两个包长期不兼容）。
协议只用 `initialize` / `notifications/initialized` / `tools/list` / `tools/call` 四个方法，
这是 stdio MCP 的最小可用子集，任何标准实现都能互通。

**为什么核心零第三方依赖**
`pip install -e .` 不装任何东西也能跑 `sak --mock`。MCP 服务端按需装，
这样「只想换个模型试试」的人不必先解决 onnxruntime 的编译问题。

**为什么分支配置是 JSON 不是 YAML**
避免为了一份 8 行的清单引入 PyYAML。分支文件的作用是「声明提示词和所需服务」，
不需要 YAML 的表达力。

**为什么工具名冲突要加前缀**
不同 MCP 服务可能都有 `search`，前缀化让模型能明确指向某个服务，也让日志可读。

**为什么工具失败不抛出**
`MCPManager.call()` 把异常转成 `[错误] ...` 文本回给模型，让模型自己决定重试还是降级。
直接抛异常会让一次工具抖动中断整个任务。

## 扩展点

| 想做的事 | 改哪里 |
|:---|:---|
| 加一个分支 | 新增 `branches/X.json` + `prompts/X_*.md` |
| 加一个模型供应商 | 新增 `sak/llm/xxx.py` 实现 `chat()`，在 `llm/__init__.py` 注册 |
| 加一个工具服务 | 新增 `server/xxx/`，在 `sak.json` 的 `mcp_servers` 注册，在分支 JSON 的 `servers` 里引用 |
| 换检索实现 | 只改 `server/rag/`，提示词与编排层不动 |
| 改输出格式 | 只改 `prompts/`，不动代码 |

## 目录

```
branches/        分支声明（id / 提示词 / 所需服务 / 可交接给谁）
prompts/         分支提示词正文
sak/             编排层（config / branches / agent / mcp_client / cli / llm）
server/          MCP 服务（rag / private / crawler / report / agent_switch）
skills/          可复用的操作 SOP 与脚本（台账写入、清单同步、标准入库等）
kb_template/     知识库骨架（导航 / Wiki / 原始资料 / 输出 / 脚本）
tests/           离线自测
docs/            文档
```
