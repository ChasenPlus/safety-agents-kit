# 04-脚本代码 说明

分支提示词里会调用一批脚本。仓库只带**通用、可复用**的那部分，其余属于使用方自己的业务脚本。

## 随包提供

| 脚本 | 位置 | 用途 |
|:---|:---|:---|
| `gate.py` | `01 gate/` | 交付门禁：检查 docx/xlsx 的格式、占位符、字体、目录域等，输出 `GATE PASS/FAIL` |
| `append_ledger_entry.py` | `skills/excel-ledger-writer/scripts/` | 台账幂等追加/更新（去重、阈值检测、状态回写） |
| `generate_work_package_docs.py` | `skills/official-document-formatting/scripts/` | 工作包六件套转正式 docx |
| `generate_special_plan_templates.py` | `skills/official-document-formatting/scripts/` | 专项方案模板批量生成 |

## 需要自备

以下脚本被提示词引用，但内容与使用方的标准体系、报告模板、台账结构强相关，仓库不提供。
缺失时按最后一列的方式降级，不影响其余流程。

| 脚本 | 所属分支 | 用途 | 缺失时降级方式 |
|:---|:---:|:---|:---|
| `search_standard_clause.py` | A | 条款原文定位与命中排序 | 直接用 `sak-rag` 的 `rag_search`，人工确认条款号 |
| `verify_standards_version.py` | A/D | 查本地版本跟踪表并提示联网核验 | 手工核对标准号与发布日期 |
| `mark_checklist_invalid.py` | A/B | 把清单里失效的条目打标 | 在台账里人工标注 |
| `classify_standards_three_tier.py` | D | 标准原文三层分类归位 | 按目录名手工移动 |
| `cleanup_invalid_files.py` | D | 空白/乱码/重复文件检测（只检测不删除） | 手工抽查 |
| `link_checklist_standards.py` | D | 检查清单条目与标准原文的链接 | 手工维护关联表 |
| `ocr_scanned_pdf.py` | D | 扫描件 PDF 转 Markdown | 用外部 OCR 工具转换后放入 `02-原始资料层/` |
| `generate_revision_doc.py` | C | 制度修订建议书 docx 生成 | 由模型直接输出结构化文本，再人工排版 |
| `generate_safety_report.py` | G | 安全周报/月报/季报 docx 生成 | 按 `03-输出层/安全报告/` 的模板手工套版 |
| `generate_anweihui_report.py` | G | 安委会汇报材料生成 | 同上 |
| `docx_to_md.py` | H | 既有 docx 转 Markdown 以便检索 | 用 `python-docx` 临时脚本读取 |

## 约定

- 所有脚本用 `python -X utf8` 运行，路径参数走绝对路径或 `${SAK_KB_ROOT}` 展开后的路径。
- 脚本只读写知识库与输出目录，不修改 `prompts/`、`branches/`。
- 新增脚本请写清 `--help`、返回码含义，并在本表补一行。
