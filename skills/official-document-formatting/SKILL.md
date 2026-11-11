---
name: official-document-formatting
description: "专项工作包生成智能体"产出六份文档的正式公文格式排版规范与生成脚本。生成"正式文档/Word版/能打印"时调用，默认的纯Markdown文本输出不需要本技能。已用真实数据测试过全部六份文档的排版效果（含PDF渲染核对）。
---

# 公文格式排版规范

## 设计依据

参照 GB/T 9704《党政机关公文格式》的核心排版惯例：标题黑体加粗居中、正文仿宋字、
一级标题"一、"二级标题"（一）"三级标题"1."的层级序号、落款单位+日期、页脚页码。
但这五份是内部工作文书，不是对外正式发文，所以不含密级/紧急程度/主送机关/发文字号
这类仅用于正式行文的要素——保留的是"正式、清晰、层级分明"，不追求完整套用行政公文
的全部形式要件。

**应急处置卡是例外**：现场张贴用的卡片，采用红/黄配色而不是纯黑白公文风格——应急场景
要的是"一眼看清"，不是"正式庄重"，两种设计诉求不一样，不能用同一套视觉语言。

**专项培训卡也是例外**：蓝/橙配色，嵌入二维码（扫码观看培训视频），含事故案例摘要——
培训警示场景要的是"案例+视频双重警示"，也不是纯公文风格。

## 生成方式

调用 `skills/official-document-formatting/scripts/generate_work_package_docs.py`：

```bash
python generate_work_package_docs.py --data-file <数据JSON文件路径> --output-dir <输出目录> [--only 类型1 类型2 ...]
```

- `--data-file`：包含六份文档全部内容的JSON文件（结构见下方"数据结构"）。**不要用命令行直接传JSON字符串**，中文+多层嵌套很容易在shell里转义出错，先写成文件再传路径。
- `--output-dir`：docx文件的输出目录。
- `--only`：只想生成其中几份时用（值取 `safety_briefing`/`technical_briefing`/`emergency_card`/`assessment_standard`/`regulation_card`/`training_card`），不传则六份全生成。
- 想看数据结构长什么样，跑 `python generate_work_package_docs.py --print-example` 会打印一份完整示例。

返回一段JSON，报告每份文档的生成状态（`ok`或`error`及原因）、每份文档的 `missing_images`（未找到的图片路径），以及 `_summary.missing_images` 汇总，据此判断是否成功、要不要走R08降级。**`missing_images` 非空时不得静默交付**：要么补齐图片文件重新生成，要么明确告知用户哪些条目是占位段。

## 图片怎么放进文档（图片保真管线出文档侧）

支持插图的条目字段：安全交底的 `measures`、技术交底的 `procedure_steps`/`quality_requirements`、应急处置卡的 `emergency_steps`。这些字段的列表项写成 dict 时可扩展两个键：

- `"image"`：图片路径。绝对路径直接用；相对路径依次相对 ① data-file JSON 所在目录 ② `output_dir` 解析。**标准/制度里的示意图**：从 RAG 命中结果的 `images` 字段拿相对 md 的路径，结合 `rel_path` 定位 md 同级 `{stem}_images/` 目录，转成可访问的绝对/相对路径填入；**现场照片**：约定复制到 `output_dir/现场照片/` 子目录，JSON 里写 `现场照片/{YYYYMMDD}_{地点设备}_{序号}.jpg`。
- `"image_caption"`：图注（可选），如"图1 验电部位示意图"。脚本以小五黑体居中渲染；应急卡配图集中在步骤框外，图注自动带"步骤N配图"前缀。
- 排版：图片居中、横图宽≤15cm，竖长图（高/宽>1.33）按高≤20cm 约束，等比缩放不变形。
- 图片随所属条目插入（措施图在该措施条款及其依据引用块之后、工序图在该步骤之后），不要把图集中堆到文档末尾。
- 路径不存在时插灰色居中占位段 `【图片缺失：路径】` 并登记进 `missing_images`，绝不假装生成成功。

## 数据结构（把"专项工作包生成智能体"产出的内容组织成这个结构）

关键字段说明（完整示例见脚本 `--print-example` 输出）：

- `work_name`、`date`、`doc_no`：工作名称、日期、文号后缀（编号不冲突即可，比如当天第几份用01/02区分）。
- `hazards`（安全交底用）：危险源字符串列表。
- `measures`（安全交底用）：安全措施列表，每项可以是纯字符串，也可以是 `{"text":..., "source":..., "clause_text":...}` 这种带原文引用的结构（有引用时会在该条下方生成灰底引用块，对应R00）；dict 项还可带 `image`/`image_caption`，规则见上方"图片怎么放进文档"。
- `procedure_steps`、`quality_requirements`、`key_control_points`（技术交底用）：结构同上，前两者的 dict 项支持 `image`/`image_caption`。
- `scenario`、`risks`、`emergency_steps`、`site_contact`、`emergency_phone`、`sources`（应急处置卡用）；`emergency_steps` 的 dict 项（`{"text":..., "image":..., "image_caption":...}`）支持配图。
- `regulations`（制度规范卡用）：`[{"code":..., "name":..., "clauses":[{"article":..., "text":...}, ...]}, ...]`——这里的内容对应"专项工作包生成智能体"步骤2b的企业标准全量搜索结果，**不经R11过滤，原样放进来**，脚本本身不做任何筛选。
- `assessment_regulations`（工作相关考核标准用）：结构与`regulations`完全相同，`[{"code":..., "name":..., "clauses":[{"article":..., "text":...}, ...]}, ...]`——同样来自②b的全量搜索结果，但只挑其中"考核/验收/评分/合格判定"相关的条款，**同样不经R11过滤**。和`regulations`可以有重叠（同一条款既涉及考核也可能出现在制度规范卡里），不用互斥去重。
- `training_items`（专项培训卡用）：`[{"risk_type":..., "case_summary":..., "case_source":..., "video_path":..., "video_label":...}, ...]`——内容来自步骤2f：风险类型从步骤0/2a已识别的风险出发，案例摘要从《新能源项目典型事故案例汇编》PDF按风险-章节映射表提取（每类1-2个案例，每个2-3句），视频路径按风险-视频映射表匹配（无视频的风险`video_path`留空）。脚本自动生成QR码编码`video_path`，嵌入docx。

## 执行前必须确认的事（避免把没核验过的内容当成正式文档发出去）

生成前，调用方（分支E自己）要确认：`hazards`/`measures`/`procedure_steps`/`quality_requirements`/`key_control_points` 这几个字段里的内容（对应安全交底/技术交底/应急处置卡），全部来自2a步骤R11核验通过的条款——**这个脚本本身不做核验，只负责排版**，核验责任在调用它之前就应该完成。`regulations`和`assessment_regulations`这两个字段则相反，必须是2b步骤的全量摘录结果，不能偷偷塞进只有核验通过的内容（那样就失去了"全量、用于警示"的意义）。

## 关于emoji/特殊符号的技术限制（已踩过坑，别重复踩）

生成应急处置卡这类需要图标点缀的场景，**只能用基本平面内的Unicode符号**（比如 ⚠ ⚡ ★ ● 这些），**不能用需要代理对编码的高级emoji**（比如 🚨 🎯 这类"表情符号"）——后者在LibreOffice等渲染器里和中文混排时，会导致该符号之后的中文文字整体显示为方块乱码（这是实测踩过的坑，不是理论推测）。脚本内部已经统一改用 `⚠`/`⚡` 这类安全符号，如果要扩充新的图标点缀，先用同样的方式（转PDF截图检查）验证过再用，不要凭感觉挑emoji。

## 生成后必须核对排版（不能生成完就直接交付）

按docx技能的要求，生成后转一次PDF、截图看一眼排版是否正常（页边距、表格是否错位、中文字体是否正确显示，尤其是刚才提到的emoji混排问题），确认没问题再交付给用户，不能生成完docx文件就直接当作完成——之前吃过这个亏。
