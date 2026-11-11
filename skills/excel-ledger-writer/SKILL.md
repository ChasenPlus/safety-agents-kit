---
name: excel-ledger-writer
description: 向《示例电力制度自查台账.xlsx》幂等追加/更新条目的规范与可执行脚本，包含sheet新建/追加判断、汇总sheet同步更新、去重（R09）、修订阈值检测与状态回写（R12）。分支A/B/C在读写台账时应优先调用本技能自带的脚本，而不是逐个MCP工具调用手工拼写单元格，以保证格式一致性和幂等性。
---

# 示例电力制度自查台账写入规范

## 为什么用脚本而不是逐格MCP调用

台账要求：每个制度编号一个独立sheet、首次出现要建sheet并写表头、已存在要追加、汇总sheet要同步更新计数、修订起草后要原地回写状态——这是一套有状态、需要幂等保证的操作。如果每次都现场调用 MCP 工具逐个单元格写入，容易出现表头漏写、汇总没同步、重复行、状态回写变成重复新增等问题。`scripts/append_ledger_entry.py` 把这套逻辑封装成四个明确的调用模式，只需要传入结构化参数即可。

## 四种模式

### 1. append（默认，R06b用）—— 追加新条目

```bash
python skills/excel-ledger-writer/scripts/append_ledger_entry.py \
  --workbook "{工作目录}/03-输出层/制度修订编制/示例电力制度自查台账.xlsx" \
  --entry-json '{
    "记录日期": "2026-08-19",
    "触发问题": "承包商安全文化培训考核频次要求",
    "归属制度编号": "2112",
    "归属制度名称": "《安全文化建设管理规定》",
    "缺口类型": "命中但未实质覆盖",
    "命中条款说明": "2112第3.2条原文：应定期开展安全文化培训——仅笼统提及定期开展培训，未规定具体频次",
    "匹配置信度": "已确认（企业标准层有命中条款，但内容不符）",
    "缺口描述": "该问题在行业标准层已有明确的频次要求条款，但2112第3.2条表述过于原则化",
    "建议补充方向": "建议在2112第3.2条基础上，细化明确培训频次、承包商人员适用范围及考核方式",
    "参考的上位标准依据": "《规范全称》（标准编号）第X.X.X条 [原文摘录]"
  }'
```

返回：`{"status": "written", "sheet": "2112_安全文化建设管理规定", "row": 3, "summary_updated": true}`
去重命中：`{"status": "duplicate", "sheet": "...", "existing_row": N}`
异常：`{"status": "error", "message": "..."}`

### 2. check-threshold（R12自动触发用）—— 检查是否达到修订起草阈值

分支A/B每次成功追加一条R06b条目后调用：

```bash
python skills/excel-ledger-writer/scripts/append_ledger_entry.py \
  --workbook "{工作目录}/03-输出层/制度修订编制/示例电力制度自查台账.xlsx" \
  --mode check-threshold --code 2112 --threshold 2
```

返回：`{"status": "ok", "count": 2, "reached": true, "should_ask": true, ...}`

`should_ask: true` 时，应在回复中附加 `建议询问用户是否启动修订=true`，由主流程决定是否询问用户。

### 3. set-confirm-status（R12用）—— 记录用户的Y/N回应

```bash
python skills/excel-ledger-writer/scripts/append_ledger_entry.py \
  --workbook "{工作目录}/03-输出层/制度修订编制/示例电力制度自查台账.xlsx" \
  --mode set-confirm-status --code 2112 --confirm-status 待用户确认
```

`--confirm-status` 取值：`待用户确认`（发出询问时）/ `已确认Y`（用户同意）/ `已暂缓N`（用户不同意）。

### 4. update-status（R12用）—— 原地回写修订状态

```bash
python skills/excel-ledger-writer/scripts/append_ledger_entry.py \
  --workbook "{工作目录}/03-输出层/制度修订编制/示例电力制度自查台账.xlsx" \
  --mode update-status --code 2112 --name "《安全文化建设管理规定》" \
  --new-status "已生成修订建议待评审" \
  --doc-path "/mnt/user-data/outputs/制度修订建议书_2112.docx"
```

不传 `--match-list-json` 时默认更新该制度sheet内**全部**"待补充"状态的行；只想更新特定几条时传入 `--match-list-json '[["2026-08-18","问题A"],["2026-08-19","问题B"]]'`。

## 脚本行为说明（对应 R06b / R09 / R08 / R12）

1. **去重（R09）**：append模式下，查找"触发问题"与"记录日期"完全一致的已有行，命中则返回 `duplicate`，不追加。
2. **sheet新建/追加**：sheet名规则为 `{归属制度编号}_{归属制度名称去书名号}`（超长自动截断）。sheet不存在则新建并写表头；存在则追加到最后一行之后。
3. **汇总sheet同步**：append成功后同步"汇总"sheet中该制度编号的"累计发现问题数"（+1）和"最近更新日期"；不存在则新增一行。这个计数就是check-threshold判断的依据。
4. **Y/N状态机（R12核心）**：`check-threshold` 用"上次询问时条目数"（而不是"上次生成时条目数"）判断是否该重新问用户，这是两个独立字段——用户已经回答N之后，只要没有新增缺口就不会重复追问，只有等新增缺口出现时才会再问一次。这个区分是经过实测踩坑后才确定的设计，不能合并成一个字段。
5. **状态与建议书路径原地回写（R12）**：update-status模式直接修改已有行的"状态"和"修订建议书路径"两列，不新增行，避免同一批缺口在起草修订后被误当成新条目。
6. **写入失败兜底（R08）**：工作簿路径不可写、文件被占用等情况，返回 `error` 及具体原因，由调用方走R08降级内联输出流程，不得吞掉错误静默失败。

## 若平台不支持Bash执行Python

改用 `mcp__sak-excel`（`@negokaz/excel-mcp-server`）逐步操作，但必须手动复现上述六条行为，不能省略任意一条，否则台账会逐渐出现重复行、汇总失真，或修订建议重复生成/重复追问。
