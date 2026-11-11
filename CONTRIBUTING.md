# 贡献指南

欢迎三类贡献：**新分支**、**新工具**、**新的判断经验**。

## 环境

```bash
git clone <repo>
cd safety-agents-kit
pip install -e .
python tests/test_smoke.py       # 应全绿，且不需要任何 API Key
python tests/test_llm_wire.py    # 模型适配层报文格式
python tests/test_coverage_cases.py  # 覆盖核验回归（改 coverage_core 必跑）
```

## 提交前自检

1. `python -m compileall -q sak server scripts` 无报错
2. `python tests/test_smoke.py`、`python tests/test_llm_wire.py`、`python tests/test_coverage_cases.py` 全 PASS
3. `python scripts/desensitize.py --scan-only` 残留为 0
4. 没有提交 `.env`、`sak.json`、`kb/`、索引文件、任何标准原文 PDF

## 加一条分支

1. 新建 `prompts/X_<英文短名>.md`，写清四件事：
   - 开工预读（要读哪些笔记）
   - 执行步骤（能落到工具调用上，不要写「自行判断」）
   - 输出格式（表格列名、字段顺序、必填项）
   - 降级规则（工具不可用时怎么退回，必须显式说明）
2. 新建 `branches/X.json`：

```json
{
  "id": "X",
  "name": "分支中文名",
  "summary": "一句话职责",
  "prompt": "X_<英文短名>.md",
  "servers": ["fs", "rag"],
  "handoff_to": []
}
```

3. 在 `README.md` 与 `docs/branch-permissions.md` 的分支表里补一行。

**提示词写作要求**

- 用「步骤 + 判定条件 + 输出格式」，不写形容词式的原则
- 引用标准必须要求回原文逐字取，不允许直接复制检索片段
- 凡是「模型可能偷懒」的环节，写成硬约束（如「必须贴出校验通过原文」）
- 不要写模型名、IDE 名、本地绝对路径

## 加一个 MCP 工具服务

1. 目录 `server/<name>/`，入口固定为 `server.py`
2. 用 `fastmcp` 或裸 `mcp` 均可，stdio 传输
3. 工具函数必须有 docstring —— 它就是给模型看的工具说明
4. 依赖写进 `pyproject.toml` 的 optional-dependencies，并在 `docs/mcp-servers.md` 登记
5. 在 `sak.json` 的 `mcp_servers` 里注册（用相对路径）

## 加一条判断经验

经验不要写进代码，写进知识库：

- 放在 `kb_template/01-Wiki层/` 对应子目录
- 格式：`现象 → 原因 → 处置 → 依据`
- 如果这条经验要求某个分支改行为，在提示词的对应步骤里加一行指针，不要复制内容

## 脱敏

仓库里不得出现真实企业名、项目名、人名、手机号、内部文件路径。
`scripts/desensitize.py` 负责扫描（绝对路径、手机号、身份证号）并按词表替换：
把需要替换的词写进 `scripts/sensitive_words.txt`（一行 `原词=替换词`，该文件已 gitignore，不会提交）。

## 代码风格

- Python ≥ 3.9，只用标准库（`sak/` 核心必须零依赖）
- 类型标注齐全，`from __future__ import annotations`
- 不做「以防万一」的兜底：工具边界做校验，内部调用信任上层
- 错误信息要能直接指导排查（带服务名、命令、路径）
