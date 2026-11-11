# -*- coding: utf-8 -*-
"""
专项工作包 —— 公文格式文档生成脚本

给分支E（专项工作包生成）的六份文档提供正式、美观的docx排版：
安全交底、技术交底、应急处置卡、工作相关考核标准、制度规范卡、专项培训卡。

六份文档分四组，视觉设计也不同：
    - A组（安全交底、技术交底）：纯公文风格，黑体标题+仿宋正文+层级序号+签字表格。
    - B组（工作相关考核标准、制度规范卡）：公文抬头框架相同，但正文部分改用"声明框+
      逐条引用块"的结构，不是层级序号叙述——因为这两份是企业标准全量摘录、不经R11过滤，
      内容本质是"罗列条款原文"而不是"叙述式的规程要求"，排版上也应该体现这个差异。
    - 应急处置卡单独一组：现场张贴用的卡片，采用醒目的红/黄配色而不是纯黑白公文风格，
      应急场景要的是"一眼看清"，不是"正式庄重"。
    - 专项培训卡：蓝/橙配色，嵌入二维码（扫码观看培训视频），含事故案例摘要，
      用真实事故案例+视频警示作业人员，定位是"培训警示"而非"交底告知"。

排版依据：参照 GB/T 9704《党政机关公文格式》的核心排版惯例（标题黑体加粗居中、
正文仿宋三号字、一级标题"一、"二级标题"（一）"三级标题"1."的层级序号、
落款单位+日期右对齐、页脚页码），但这些是内部工作文书而非对外发文，所以不含
密级/紧急程度/主送机关/发文字号这类仅用于正式行文的要素，重点保留"看起来正式、
读起来清晰、层级分明"这几个核心诉求。

用法：
    python generate_work_package_docs.py --data-file work_package_data.json --output-dir /path/to/output

data-file 的JSON结构说明见本文件末尾的 EXAMPLE_DATA，或直接运行：
    python generate_work_package_docs.py --print-example > example_data.json
生成一份示例数据文件，照着改。
"""

import argparse
import json
import os
import sys
import copy
import tempfile

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import docx_image_helpers as img_hlp


# ============================================================
# 公文风格基础设置
# ============================================================

COMPANY_NAME = "示例电力"
TITLE_FONT = "黑体"
BODY_FONT = "仿宋_GB2312"
BODY_FONT_FALLBACK = "仿宋"
SIG_FONT = "楷体_GB2312"

RED = RGBColor(0xC0, 0x00, 0x00)
AMBER_FILL = "FFF2CC"
RED_FILL = "FCE4E4"
GRAY_FILL = "F2F2F2"
BLUE_TITLE_FILL = "1F4E79"     # 专项培训卡标题条（深蓝）
BLUE_CASE_FILL = "D6E4F0"      # 案例摘要块底色（浅蓝）
BLUE_BORDER = "2E75B6"         # 案例摘要块边框（中蓝）
ORANGE_ACCENT = "ED7D31"       # 培训警示强调色（橙）

# ============================================================
# 培训资料路径与风险-视频映射
# ============================================================

VIDEO_DIR = os.path.expandvars(r"$SAK_KB_ROOT\02-原始资料层\08培训资料\视频")

# 风险类型 → 视频文件名（仅此映射表中的风险类型有培训视频）
RISK_VIDEO_MAP = {
    "有限空间": "1有限空间.mp4",
    "高处坠落": "2高坠.mp4",
    "二氧化碳致裂管爆炸": "3.mp4",
    "CO2致裂管爆炸": "3.mp4",
    "坍塌": "4坍塌.mp4",
}


def set_cjk_font(run, font_name, size=None, bold=None, color=None):
    """python-docx的font.name只设置西文字体，中文字体要单独通过oxml设置east asia字体，
    否则在Word里中文字符会用系统默认字体渲染，跟指定的仿宋/黑体对不上。"""
    run.font.name = font_name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:eastAsia'), font_name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def add_page_number_field(paragraph):
    """插入 PAGE 域代码，实现页脚"第 X 页"里的自动页码。"""
    run = paragraph.add_run()
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = "PAGE"
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    run._element.append(fld_begin)
    run._element.append(instr)
    run._element.append(fld_end)


def set_cell_shading(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def set_cell_borders(cell, color="808080", size="4"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), size)
        el.set(qn('w:color'), color)
        borders.append(el)
    tcPr.append(borders)


def _doc_no(prefix, data):
    """公文发文字号：机关代字+性质字+〔四位年份〕+顺序号（不编虚位）。
    范式：示例电力党字〔2025〕1号、示例集团集团发〔2026〕173号。"""
    year = str(data["date"])[:4]
    seq = int(str(data.get("doc_no", "1")))
    return f"示例电力{prefix}〔{year}〕{seq}号"


def new_official_document(title_lines, doc_number, subtitle=None):
    """创建一份标准公文风格文档骨架：单位抬头 + 文号 + 大标题 + 分隔线。"""
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(3.7)
    section.bottom_margin = Cm(3.5)
    section.left_margin = Cm(2.8)
    section.right_margin = Cm(2.6)

    # 单位抬头
    header_p = doc.add_paragraph()
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = header_p.add_run(COMPANY_NAME)
    set_cjk_font(r, TITLE_FONT, size=22, bold=True, color=RED)

    # 文号
    num_p = doc.add_paragraph()
    num_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = num_p.add_run(doc_number)
    set_cjk_font(r, BODY_FONT, size=12)

    # 分隔线（用下框线模拟公文抬头下的红线）
    line_p = doc.add_paragraph()
    pPr = line_p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '18')
    bottom.set(qn('w:color'), 'C00000')
    pBdr.append(bottom)
    pPr.append(pBdr)

    doc.add_paragraph()

    # 主标题
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for i, line in enumerate(title_lines):
        if i > 0:
            title_p.add_run().add_break()
        r = title_p.add_run(line)
        set_cjk_font(r, TITLE_FONT, size=20, bold=True)

    if subtitle:
        sub_p = doc.add_paragraph()
        sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = sub_p.add_run(subtitle)
        set_cjk_font(r, BODY_FONT, size=14, color=RGBColor(0x59, 0x59, 0x59))

    doc.add_paragraph()

    # 页脚页码
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = fp.add_run("第 ")
    set_cjk_font(r1, BODY_FONT, size=10)
    add_page_number_field(fp)
    r2 = fp.add_run(" 页")
    set_cjk_font(r2, BODY_FONT, size=10)

    return doc


def add_body_paragraph(doc, text, size=15, indent_first_line=True, bold=False):
    p = doc.add_paragraph()
    if indent_first_line:
        p.paragraph_format.first_line_indent = Pt(size * 2)
    p.paragraph_format.line_spacing = 1.4
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    set_cjk_font(r, BODY_FONT, size=size, bold=bold)
    return p


def add_level1_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    set_cjk_font(r, TITLE_FONT, size=15, bold=True)
    return p


def add_level2_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    set_cjk_font(r, BODY_FONT, size=15, bold=True)
    return p


def add_numbered_item(doc, index, text, size=15):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Pt(size * 2)
    p.paragraph_format.first_line_indent = Pt(-size)
    p.paragraph_format.line_spacing = 1.4
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(f"{index}. ")
    set_cjk_font(r, BODY_FONT, size=size, bold=True)
    r2 = p.add_run(text)
    set_cjk_font(r2, BODY_FONT, size=size)
    return p


def add_citation_block(doc, source, text, size=14):
    """引用块：灰底 + 悬挂缩进，用于展示条款原文，一字不差（对应R00）。"""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_shading(cell, GRAY_FILL)
    set_cell_borders(cell, color="BFBFBF", size="2")
    p_src = cell.paragraphs[0]
    r = p_src.add_run(f"依据：{source}")
    set_cjk_font(r, BODY_FONT, size=12, color=RGBColor(0x59, 0x59, 0x59))
    p_text = cell.add_paragraph()
    p_text.paragraph_format.space_before = Pt(4)
    r2 = p_text.add_run(text)
    set_cjk_font(r2, BODY_FONT, size=size)
    doc.add_paragraph()


def add_signature_table(doc, roles):
    """签字确认表格：交底人/被交底人这类落款区域。roles 是角色名称列表。"""
    doc.add_paragraph()
    tbl = doc.add_table(rows=2, cols=len(roles))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, role in enumerate(roles):
        c = tbl.cell(0, i)
        set_cell_borders(c)
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(role)
        set_cjk_font(r, BODY_FONT, size=13, bold=True)

        c2 = tbl.cell(1, i)
        set_cell_borders(c2)
        c2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r2 = p2.add_run("（签字）")
        set_cjk_font(r2, BODY_FONT, size=12, color=RGBColor(0x99, 0x99, 0x99))
    for row in tbl.rows:
        row.height = Cm(1.5)


def add_info_header_table(doc, fields):
    """工作名称/交底日期/交底人 这类文档头部信息表格。fields是[(标签,值),...]。"""
    tbl = doc.add_table(rows=len(fields), cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    for i, (label, value) in enumerate(fields):
        c1 = tbl.cell(i, 0)
        c1.width = Cm(3.2)
        set_cell_shading(c1, GRAY_FILL)
        set_cell_borders(c1)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p1.add_run(label)
        set_cjk_font(r1, BODY_FONT, size=13, bold=True)

        c2 = tbl.cell(i, 1)
        set_cell_borders(c2)
        p2 = c2.paragraphs[0]
        r2 = p2.add_run(value or "______________")
        set_cjk_font(r2, BODY_FONT, size=13)
    doc.add_paragraph()


def add_footer_note(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    r = p.add_run(text)
    set_cjk_font(r, BODY_FONT, size=11, color=RGBColor(0x80, 0x80, 0x80))


# ============================================================
# 五份文档各自的生成函数
# ============================================================

def gen_safety_briefing(data, out_path, base_dir=None):
    """安全交底。dict 型 measure 可带 image/image_caption（措施示意图、现场照片）。
    返回 missing_images 路径列表。"""
    missing = []
    out_dir = data.get("output_dir")
    disp = data.get("display_name", data["work_name"])
    doc = new_official_document(
        ["安全技术交底书"],
        _doc_no("安交", data),
        subtitle=disp,
    )
    add_info_header_table(doc, [
        ("工作名称", disp),
        ("交底日期", data["date"]),
        ("交底人", data.get("briefer", "")),
        ("被交底人", data.get("briefed", "")),
    ])

    if data.get("note"):
        add_footer_note(doc, data["note"])

    sections = []
    sections.append(("一、危险源辨识", data["hazards"]))
    if data.get("entry_requirements"):
        sections.append(("二、入场准入与培训要求", data["entry_requirements"]))
    if data.get("site_discipline"):
        sections.append(("三、外委现场管理纪律", data["site_discipline"]))
    sections.append(("四、安全技术措施", data["measures"]))
    if data.get("penalties"):
        sections.append(("五、违章后果", data["penalties"]))
    # 章节号随可选段动态重排，保证序号连续
    sections = [(f"{['一','二','三','四','五','六'][i]}、{title.split('、',1)[1]}", body)
                for i, (title, body) in enumerate(sections)]

    for title, body in sections:
        add_level1_heading(doc, title)
        if title.endswith("安全技术措施"):
            for i, measure in enumerate(body, 1):
                text = measure["text"] if isinstance(measure, dict) else measure
                add_numbered_item(doc, i, text)
                if isinstance(measure, dict):
                    if measure.get("source"):
                        add_citation_block(doc, measure["source"], measure["clause_text"])
                    img, cap = img_hlp.get_item_image(measure)
                    if img:
                        img_hlp.add_image_block(doc, img, cap, base_dir=base_dir,
                                                output_dir=out_dir, missing=missing)
        else:
            for i, item in enumerate(body, 1):
                add_numbered_item(doc, i, item)

    add_level1_heading(doc, f"{['一','二','三','四','五','六'][len(sections)]}、交底确认")
    add_body_paragraph(doc, "被交底人已知悉上述风险及防范措施，签字确认：", indent_first_line=False)
    add_signature_table(doc, ["交底人", "被交底人", "现场负责人"])

    doc.save(out_path)
    return missing


def gen_technical_briefing(data, out_path, base_dir=None):
    """技术交底。dict 型 procedure_steps/quality_requirements 可带
    image/image_caption（工艺流程图、质量标准示意图）。返回 missing_images。"""
    missing = []
    out_dir = data.get("output_dir")
    doc = new_official_document(
        ["技术交底书"],
        _doc_no("技交", data),
        subtitle=data.get("display_name", data["work_name"]),
    )
    add_info_header_table(doc, [
        ("工作名称", data.get("display_name", data["work_name"])),
        ("交底日期", data["date"]),
        ("交底人", data.get("briefer", "")),
        ("被交底人", data.get("briefed", "")),
    ])

    add_level1_heading(doc, "一、作业方法与工艺流程")
    for i, step in enumerate(data["procedure_steps"], 1):
        text = step["text"] if isinstance(step, dict) else step
        add_numbered_item(doc, i, text)
        if isinstance(step, dict):
            img, cap = img_hlp.get_item_image(step)
            if img:
                img_hlp.add_image_block(doc, img, cap, base_dir=base_dir,
                                        output_dir=out_dir, missing=missing)

    add_level1_heading(doc, "二、技术要求与质量标准")
    for item in data["quality_requirements"]:
        text = item["text"] if isinstance(item, dict) else item
        add_body_paragraph(doc, "• " + text, indent_first_line=False)
        if isinstance(item, dict):
            if item.get("source"):
                add_citation_block(doc, item["source"], item["clause_text"])
            img, cap = img_hlp.get_item_image(item)
            if img:
                img_hlp.add_image_block(doc, img, cap, base_dir=base_dir,
                                        output_dir=out_dir, missing=missing)

    add_level1_heading(doc, "三、关键工序控制点")
    for i, point in enumerate(data["key_control_points"], 1):
        add_numbered_item(doc, i, point)

    add_level1_heading(doc, "四、交底确认")
    add_body_paragraph(doc, "被交底人已知悉上述技术要求，签字确认：", indent_first_line=False)
    add_signature_table(doc, ["交底人", "被交底人", "技术负责人"])

    doc.save(out_path)
    return missing


def gen_emergency_card(data, out_path, base_dir=None):
    """应急处置卡 —— 现场张贴用，红/黄配色，不走纯公文黑白风格。
    dict 型 emergency_steps 可带 image/image_caption（疏散路线图等）。"""
    missing = []
    out_dir = data.get("output_dir")
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

    # 顶部红色标题条
    tbl_head = doc.add_table(rows=1, cols=1)
    cell = tbl_head.cell(0, 0)
    set_cell_shading(cell, "C00000")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("⚠ 应急处置卡")
    set_cjk_font(r, TITLE_FONT, size=26, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(data.get("display_name", data["work_name"]))
    set_cjk_font(r2, BODY_FONT, size=16, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

    doc.add_paragraph()

    add_level1_heading(doc, "适用场景")
    add_body_paragraph(doc, data["scenario"], indent_first_line=False)

    # 风险点：黄色底
    p_h = doc.add_paragraph()
    r_h = p_h.add_run("⚡ 主要风险点")
    set_cjk_font(r_h, TITLE_FONT, size=15, bold=True, color=RED)
    tbl_risk = doc.add_table(rows=1, cols=1)
    c = tbl_risk.cell(0, 0)
    set_cell_shading(c, AMBER_FILL)
    set_cell_borders(c, color="F5A623", size="4")
    for i, risk in enumerate(data["risks"]):
        target_p = c.paragraphs[0] if i == 0 else c.add_paragraph()
        r = target_p.add_run(f"{i+1}. {risk}")
        set_cjk_font(r, BODY_FONT, size=14)
    doc.add_paragraph()

    # 处置步骤：红色底突出
    p_s = doc.add_paragraph()
    r_s = p_s.add_run("⚠ 应急处置步骤")
    set_cjk_font(r_s, TITLE_FONT, size=15, bold=True, color=RED)
    tbl_steps = doc.add_table(rows=1, cols=1)
    c2 = tbl_steps.cell(0, 0)
    set_cell_shading(c2, RED_FILL)
    set_cell_borders(c2, color="C00000", size="4")
    step_images = []
    for i, step in enumerate(data["emergency_steps"]):
        text = step["text"] if isinstance(step, dict) else step
        target_p = c2.paragraphs[0] if i == 0 else c2.add_paragraph()
        r = target_p.add_run(f"{i+1}. {text}")
        set_cjk_font(r, BODY_FONT, size=15, bold=True)
        if isinstance(step, dict):
            img, cap = img_hlp.get_item_image(step)
            if img:
                step_images.append((i + 1, img, cap))
    doc.add_paragraph()

    # 配图（疏散/路线图）集中放红框外，图注带步骤序号保持对应
    for no, img, cap in step_images:
        full_cap = f"步骤{no}配图" + (f"：{cap}" if cap else "")
        img_hlp.add_image_block(doc, img, full_cap, base_dir=base_dir,
                                output_dir=out_dir, missing=missing)

    # 通用急救要点（可选）：灰底块，跨岗位通用的急救常识
    if data.get("first_aid_tips"):
        p_fa = doc.add_paragraph()
        r_fa = p_fa.add_run("✚ 通用急救要点")
        set_cjk_font(r_fa, TITLE_FONT, size=15, bold=True, color=RGBColor(0x2E, 0x75, 0xB6))
        tbl_fa = doc.add_table(rows=1, cols=1)
        c_fa = tbl_fa.cell(0, 0)
        set_cell_shading(c_fa, "EAF1F8")
        set_cell_borders(c_fa, color="2E75B6", size="4")
        for i, tip in enumerate(data["first_aid_tips"]):
            target_p = c_fa.paragraphs[0] if i == 0 else c_fa.add_paragraph()
            r = target_p.add_run(f"{i+1}. {tip}")
            set_cjk_font(r, BODY_FONT, size=13)
        doc.add_paragraph()

    add_level1_heading(doc, "应急联系")
    contact_fields = [
        ("现场负责人", data.get("site_contact", "")),
        ("第二负责人", data.get("second_leader", "")),
        ("应急电话", data.get("emergency_phone", "")),
    ]
    add_info_header_table(doc, contact_fields)

    add_footer_note(doc, "依据：" + "；".join(data.get("sources", [])))
    doc.save(out_path)
    return missing


def _article_label(article):
    """条款编号标签：纯数字编号包成'第X条'，'附录B B7…'这类非数字开头原样显示。"""
    a = str(article)
    return f"第{a}条" if a[:1].isdigit() else a


def gen_assessment_standard(data, out_path):
    """工作相关考核标准 —— 企业标准全量摘录中"考核/验收"相关条款，不经R11过滤，
    用来警示参加本项工作的人员"公司制度对这件事确实有考核要求"（哪怕写得不够具体）。"""
    doc = new_official_document(
        [f"{data.get('display_name', data['work_name'])} — 工作相关考核标准"],
        _doc_no("考核", data),
        subtitle="示例电力制度全量摘录（考核/验收相关）",
    )

    for reg in data["assessment_regulations"]:
        add_level1_heading(doc, f"{reg['code']}《{reg['name']}》")
        for clause in reg["clauses"]:
            add_citation_block(doc, _article_label(clause['article']), clause["text"])

    doc.save(out_path)
    return []


def gen_regulation_card(data, out_path):
    """制度规范卡 —— 企业标准全量摘录，不经R11过滤"""
    doc = new_official_document(
        [f"{data.get('display_name', data['work_name'])} — 制度规范卡"],
        _doc_no("规范", data),
        subtitle="示例电力制度全量摘录",
    )

    for reg in data["regulations"]:
        add_level1_heading(doc, f"{reg['code']}《{reg['name']}》")
        for clause in reg["clauses"]:
            add_citation_block(doc, _article_label(clause['article']), clause["text"])

    doc.save(out_path)
    return []


# ============================================================
# 专项培训卡 —— 二维码 + 案例摘要
# ============================================================

def generate_qr_code(data_str, output_dir=None):
    """生成QR码PNG图片，返回临时文件路径。
    data_str: 要编码的内容（文件路径或URL）。
    output_dir: QR码图片保存目录，默认用 tempfile.gettempdir()。
    返回 (qr_path, error) 元组：成功时 error=None，失败时 qr_path=None、error=异常信息。"""
    try:
        import qrcode
        qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_M,
                           box_size=10, border=2)
        qr.add_data(data_str)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        save_dir = output_dir or tempfile.gettempdir()
        os.makedirs(save_dir, exist_ok=True)
        # 按编码内容命名：同一视频在多份培训卡中只生成/复用一张二维码图
        import hashlib
        fname = f"qr_{hashlib.md5(data_str.encode('utf-8')).hexdigest()[:10]}.png"
        qr_path = os.path.join(save_dir, fname)
        if not os.path.isfile(qr_path):
            img.save(qr_path)
        return qr_path, None
    except Exception as e:
        return None, str(e)


def gen_training_card(data, out_path, base_dir=None):
    """专项培训卡 —— 蓝/橙配色，嵌入二维码（扫码观看培训视频），含事故案例摘要。
    内容来自 data["training_items"]，每项含 risk_type/case_summary/case_source/video_path/video_label。
    返回 missing_images 列表（QR码生成失败时登记）。"""
    missing = []
    out_dir = data.get("output_dir")
    qr_save_dir = os.path.join(out_dir, "二维码") if out_dir else None

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

    # 顶部蓝色标题条
    tbl_head = doc.add_table(rows=1, cols=1)
    cell = tbl_head.cell(0, 0)
    set_cell_shading(cell, BLUE_TITLE_FILL)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("● 专项培训卡")
    set_cjk_font(r, TITLE_FONT, size=26, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(data.get("display_name", data["work_name"]))
    set_cjk_font(r2, BODY_FONT, size=16, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

    doc.add_paragraph()

    # 培训说明
    add_level1_heading(doc, "培训要求")
    add_body_paragraph(doc,
        "本卡片根据工作涉及的风险类型，匹配了示例电力《新能源项目典型事故案例汇编》中的"
        "相关事故案例和培训视频。请认真阅读案例摘要，扫描二维码观看培训视频，"
        "汲取事故教训，切实提高安全意识。",
        indent_first_line=False)

    if data.get("note"):
        add_footer_note(doc, data["note"])
    for i, rule in enumerate(data.get("training_discipline", []), 1):
        add_numbered_item(doc, i, rule, size=13)

    # 连续相同 risk_type 的条目合并到同一个风险类型标题下
    groups = []
    for item in data.get("training_items", []):
        rt = item.get("risk_type", "")
        if groups and groups[-1][0] == rt:
            groups[-1][1].append(item)
        else:
            groups.append((rt, [item]))

    for risk_type, items in groups:
        is_key = any(it.get("is_key_risk") for it in items)

        # 风险类型标题（橙色强调，重点风险带标注）
        p_rt = doc.add_paragraph()
        p_rt.paragraph_format.space_before = Pt(10)
        heading = f"● 风险类型：{risk_type}" + ("（本岗位重点风险）" if is_key else "")
        r_rt = p_rt.add_run(heading)
        set_cjk_font(r_rt, TITLE_FONT, size=16, bold=True,
                     color=RGBColor(0xED, 0x7D, 0x31))

        for item in items:
            case_summary = item.get("case_summary", "")
            case_source = item.get("case_source", "")
            video_path = item.get("video_path", "")
            video_label = item.get("video_label", "")

            # 案例摘要块（浅蓝底）
            if case_summary:
                tbl_case = doc.add_table(rows=1, cols=1)
                c = tbl_case.cell(0, 0)
                set_cell_shading(c, BLUE_CASE_FILL)
                set_cell_borders(c, color=BLUE_BORDER, size="4")
                p_cs = c.paragraphs[0]
                case_title = item.get("case_title", "")
                if case_title:
                    r_tl = p_cs.add_run(case_title + "\n")
                    set_cjk_font(r_tl, TITLE_FONT, size=13, bold=True)
                r_lbl = p_cs.add_run("案例摘要：")
                set_cjk_font(r_lbl, BODY_FONT, size=13, bold=True)
                r_ct = p_cs.add_run(case_summary)
                set_cjk_font(r_ct, BODY_FONT, size=13)
                if case_source:
                    p_src = c.add_paragraph()
                    p_src.paragraph_format.space_before = Pt(4)
                    r_src = p_src.add_run(f"来源：{case_source}")
                    set_cjk_font(r_src, BODY_FONT, size=11,
                                 color=RGBColor(0x59, 0x59, 0x59))
                doc.add_paragraph()

            # QR码区域
            if video_path:
                # 检查视频文件是否存在
                if os.path.isfile(video_path):
                    qr_path, qr_err = generate_qr_code(video_path, output_dir=qr_save_dir)
                    if qr_path:
                        # QR码图片 + 扫码提示（用表格左右排列）
                        tbl_qr = doc.add_table(rows=1, cols=2)
                        tbl_qr.alignment = WD_TABLE_ALIGNMENT.CENTER
                        # 左列：QR码图片
                        c_qr = tbl_qr.cell(0, 0)
                        c_qr.width = Cm(5)
                        c_qr.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                        p_qr = c_qr.paragraphs[0]
                        p_qr.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_qr.add_run().add_picture(qr_path, width=Cm(3.5))
                        # 右列：扫码提示文字
                        c_txt = tbl_qr.cell(0, 1)
                        c_txt.width = Cm(11)
                        c_txt.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                        p_hint = c_txt.paragraphs[0]
                        p_hint.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        r_hint = p_hint.add_run("扫码观看培训视频")
                        set_cjk_font(r_hint, TITLE_FONT, size=14, bold=True,
                                     color=RGBColor(0xED, 0x7D, 0x31))
                        if video_label:
                            p_vl = c_txt.add_paragraph()
                            r_vl = p_vl.add_run(video_label)
                            set_cjk_font(r_vl, BODY_FONT, size=12)
                        p_path = c_txt.add_paragraph()
                        p_path.paragraph_format.space_before = Pt(4)
                        r_path = p_path.add_run(f"视频路径：{video_path}")
                        set_cjk_font(r_path, BODY_FONT, size=10,
                                     color=RGBColor(0x80, 0x80, 0x80))
                        doc.add_paragraph()
                    else:
                        p_err = doc.add_paragraph()
                        p_err.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r_err = p_err.add_run(f"【二维码生成失败：{qr_err}】")
                        set_cjk_font(r_err, BODY_FONT, size=12,
                                     color=RGBColor(0x80, 0x80, 0x80))
                        missing.append(f"{risk_type}二维码({qr_err})")
                        doc.add_paragraph()
                else:
                    p_nf = doc.add_paragraph()
                    p_nf.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_nf = p_nf.add_run(f"【培训视频文件未找到：{video_path}】")
                    set_cjk_font(r_nf, BODY_FONT, size=12,
                                 color=RGBColor(0x80, 0x80, 0x80))
                    missing.append(f"{risk_type}视频文件不存在({video_path})")
                    doc.add_paragraph()

    # 培训确认签字栏
    add_level1_heading(doc, "培训确认")
    add_body_paragraph(doc, "参加培训人员已知悉上述事故案例并观看培训视频，签字确认：",
                       indent_first_line=False)
    add_signature_table(doc, ["培训人", "参训人", "现场负责人"])

    add_footer_note(doc, "本卡片案例摘要摘自《新能源项目典型事故案例汇编》，培训视频存放于项目培训资料库。")
    doc.save(out_path)
    return missing


GENERATORS = {
    "safety_briefing": gen_safety_briefing,
    "technical_briefing": gen_technical_briefing,
    "emergency_card": gen_emergency_card,
    "assessment_standard": gen_assessment_standard,
    "regulation_card": gen_regulation_card,
    "training_card": gen_training_card,
}

FILENAME_MAP = {
    "safety_briefing": "安全交底",
    "technical_briefing": "技术交底",
    "emergency_card": "应急处置卡",
    "assessment_standard": "工作相关考核标准",
    "regulation_card": "制度规范卡",
    "training_card": "专项培训卡",
}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--data-file", help="包含六份文档内容的JSON文件路径")
    parser.add_argument("--output-dir", default=None,
                         help="生成的docx文件输出目录。Windows下若路径含中文，强烈建议改用--data-file里的output_dir字段传递，"
                              "命令行参数在中文Windows的cmd/powershell下可能因代码页不是UTF-8而被错误解码，导致最终生成的文件夹/文件名乱码")
    parser.add_argument("--only", nargs="*", choices=list(GENERATORS.keys()),
                         help="只生成指定的几份文档，不传则六份全生成")
    parser.add_argument("--print-example", action="store_true", help="打印一份示例JSON数据结构后退出")
    args = parser.parse_args()

    if args.print_example:
        print(json.dumps(EXAMPLE_DATA, ensure_ascii=False, indent=2))
        return

    if not args.data_file:
        print(json.dumps({"status": "error", "message": "必须提供 --data-file"}, ensure_ascii=False))
        sys.exit(1)

    with open(args.data_file, "r", encoding="utf-8") as f:
        data = json.load(f)
    # 图片相对路径优先相对 data 文件所在目录解析
    data_base_dir = os.path.dirname(os.path.abspath(args.data_file))

    # output_dir优先取JSON里的output_dir字段(走utf-8安全的文件读取链路)，
    # 命令行--output-dir仅作兜底，且中文路径下不保证正确解码
    output_dir = data.get("output_dir") or args.output_dir
    if not output_dir:
        print(json.dumps({"status": "error", "message": "必须在--data-file的JSON里提供output_dir字段，或传--output-dir（不推荐，中文路径有乱码风险）"}, ensure_ascii=False))
        sys.exit(1)

    os.makedirs(output_dir, exist_ok=True)
    targets = args.only or list(GENERATORS.keys())
    results = {}
    all_missing = []
    for key in targets:
        work_name = data["work_name"]
        fname = f"{FILENAME_MAP[key]}_{work_name}.docx"
        out_path = os.path.join(output_dir, fname)
        try:
            import inspect
            sig = inspect.signature(GENERATORS[key])
            if "base_dir" in sig.parameters:
                missing = GENERATORS[key](data, out_path, base_dir=data_base_dir)
            else:
                missing = GENERATORS[key](data, out_path)
            results[key] = {"status": "ok", "path": out_path,
                            "missing_images": missing}
            all_missing.extend(f"{FILENAME_MAP[key]}: {m}" for m in missing)
        except Exception as e:
            results[key] = {"status": "error", "message": str(e)}

    results["_summary"] = {
        "generated": sum(1 for k in targets if results.get(k, {}).get("status") == "ok"),
        "missing_images": all_missing,
    }
    print(json.dumps(results, ensure_ascii=False, indent=2))


EXAMPLE_DATA = {
    "work_name": "带电作业设备检修",
    "date": "2026-08-31",
    "doc_no": "01",
    "briefer": "______",
    "briefed": "______",
    "hazards": [
        "触电风险：作业过程中人体或工具误碰带电部位",
        "高处坠落风险：登高作业未按规定使用安全带",
    ],
    "measures": [
        {"text": "作业前必须验电，确认无电后方可开始检修", "source": "DL 409-1991 第X.X条", "clause_text": "（示例条款原文占位）",
         "image": "./带电作业设备检修_images/p005_01.png", "image_caption": "图1 验电部位示意图"},
        "作业全程设专人监护，监护人不得离岗",
    ],
    "procedure_steps": [
        "停电", "验电", "挂接地线",
        {"text": "设置遮栏和标示牌", "image": "现场照片/20260831_配电室_01.jpg",
         "image_caption": "遮栏与标示牌设置现场"},
        "开始检修作业"],
    "quality_requirements": [
        {"text": "接地线截面积不小于25平方毫米", "source": "GB 26860-2011 第X条", "clause_text": "（示例条款原文占位）"},
    ],
    "key_control_points": ["验电顺序不能颠倒", "接地线装拆顺序：先接地端后接导体端"],
    "scenario": "带电设备检修作业现场发生触电事故时",
    "risks": ["触电", "高处坠落", "机械伤害"],
    "emergency_steps": ["立即切断电源", "使用绝缘工具使伤者脱离电源", "拨打120并报告现场负责人",
                        {"text": "对伤者进行心肺复苏（如具备资质）",
                         "image": "现场照片/应急疏散图.jpg", "image_caption": "现场应急疏散路线"}],
    "site_contact": "______",
    "emergency_phone": "______",
    "sources": ["DL 409-1991", "GB 26860-2011"],
    "assessment_regulations": [
        {
            "code": "2087", "name": "带电作业管理规定",
            "clauses": [{"article": "4.2", "text": "（示例：企业标准中考核/验收相关条款原文全量摘录占位，实际生成时替换为真实检索结果，例如关于验电操作、接地线装设等考核评分要求的原文）"}],
        },
    ],
    "regulations": [
        {
            "code": "2087", "name": "带电作业管理规定",
            "clauses": [{"article": "3.1", "text": "（示例：企业标准原文全量摘录占位，实际生成时替换为真实检索结果）"}],
        },
    ],
    "training_items": [
        {
            "risk_type": "高处坠落",
            "case_summary": "2020年某风电建设工程，吊装单位作业人员在风机塔筒吊装作业中，"
                           "未按规定使用安全带，从高处坠落受伤。事故直接原因为作业人员安全防护意识不足、"
                           "违章作业，间接原因为现场安全监护不到位。",
            "case_source": "《新能源项目典型事故案例汇编》第一部分 高处坠落类",
            "video_path": os.path.expandvars(r"$SAK_KB_ROOT\02-原始资料层\08培训资料\视频\2高坠.mp4"),
            "video_label": "高处坠落培训视频",
        },
        {
            "risk_type": "触电",
            "case_summary": "2020年某风电建设工程，作业人员在拖拽揽风绳过程中，揽风绳与35kV集电线路C相放电，"
                           "造成3人触电倒地，1人经抢救无效死亡。事故直接原因为揽风绳与带电线路放电。",
            "case_source": "《新能源项目典型事故案例汇编》第一部分 人员触电类",
            "video_path": "",
            "video_label": "",
        },
    ],
}


if __name__ == "__main__":
    main()
