#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
整改通知单生成核心模块
====================
格式参数全部来自真实模板实测（示例电力2026年安全综合巡查报告·示例发电厂版），
不是凭经验估的：

  纸张 A4 (11910 x 16840 dxa)
  页边距 上下1440dxa(1英寸) 左右1800dxa(1.25英寸)
  正文 宋体 小四(12磅/sz=24)
  标题 16磅(sz=32) 加粗居中
  问题表格 总宽8539dxa 居中 全单线边框(sz=4)
  双栏列宽 左3986dxa(2.77in) 右4536dxa(3.15in)
  单元格内边距 左右各108dxa

【重要边界】本模块只负责"把已查证的结构化数据按模板格式排版"，
不负责检索标准、不负责核验条款、不生成任何内容文字。
所有标准号/条款号/条款原文必须由调用方（分支A检索核验后）提供，
本模块原样填入，缺失字段会显式标注而不是编造。
"""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml

import os
import re
import zipfile


# ============================================================
# 格式常量（实测自模板，勿随意改动）
# ============================================================
FONT_BODY = "宋体"
SZ_TITLE = Pt(16)        # 封面/正文大标题
SZ_BODY = Pt(12)         # 正文小四
SZ_H1 = Pt(14)           # 一级标题(一、二、三)
SZ_H2 = Pt(12)           # 二级标题((一)(二))
SZ_H3 = Pt(12)           # 三级标题(1. 2. 3.)

TABLE_TOTAL_W = Cm(15.06)     # 8539 dxa
COL_LEFT_W = Cm(7.03)         # 3986 dxa
COL_RIGHT_W = Cm(8.00)        # 4536 dxa

# 图片尺寸只允许四种（2026-10-28 学术论文式定稿，与 gate P2 一致）
CM = 360000                       # 1cm = 360000 EMU
SZ_PIC_LAND = (int(6.0 * CM), int(4.5 * CM))    # 单图横版
SZ_PIC_PORT = (int(4.5 * CM), int(6.0 * CM))    # 单图竖版

# 图片边框：0.75pt = 9525 EMU 黑色细边框（模板（三）3 强制，gate P20）
PIC_BORDER_XML = (
    '<a:ln %s w="9525" cap="flat" cmpd="sng" algn="ctr">'
    '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
    '<a:prstDash val="solid"/></a:ln>'
) % nsdecls('a')

# 目录只收章/节/问题标题（模板 TOC 域带 \u，正文段一旦带大纲级别就会被吸进目录）
RE_CH = re.compile(r'^[一二三四五六七八九十]+、')
RE_SEC = re.compile(r'^（[一二三四五六七八九十]+）')
RE_PROB = re.compile(r'^★?\d+[.．]')

# 问题描述的标准句式模板——这是模板里实际使用的固定写法，不可改动
CLAUSE_SENTENCE = '{finding}，不符合《{std_name}》（{std_no}）第{article}条\u201c{clause_text}\u201d的要求。'


def _set_run(run, size=SZ_BODY, bold=False, font=FONT_BODY):
    """统一设置run的字体（中英文都要设，不然中文会fallback到默认字体）"""
    run.font.size = size
    run.font.bold = bold
    run.font.name = font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)


def _add_para(doc, text="", size=SZ_BODY, bold=False,
              align=WD_ALIGN_PARAGRAPH.LEFT, indent_chars=0,
              space_before=0, space_after=0, line_spacing=1.5):
    """
    添加段落。indent_chars是首行缩进的字符数（中文公文习惯用字符而不是厘米），
    2字符缩进 = 2 * 字号，这样换字号时缩进比例不会错乱。
    """
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    if indent_chars:
        pf.first_line_indent = Pt(size.pt * indent_chars)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line_spacing
    if text:
        run = p.add_run(text)
        _set_run(run, size=size, bold=bold)
    return p


def _set_cell_font(cell, size=SZ_BODY, bold=False):
    """设置单元格内所有文字的字体"""
    for p in cell.paragraphs:
        p.paragraph_format.line_spacing = 1.5
        for run in p.runs:
            _set_run(run, size=size, bold=bold)


def _build_cover(doc, data):
    """封面：标题 + 公司名 + 日期，全部居中，上方留大片空白"""
    for _ in range(8):
        _add_para(doc)
    _add_para(doc, data["report_title"], size=SZ_TITLE, bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER)
    for _ in range(14):
        _add_para(doc)
    _add_para(doc, data["company"], size=SZ_BODY, bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER)
    _add_para(doc)
    _add_para(doc, data["report_date"], size=SZ_BODY, bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()


def _build_opening(doc, data):
    """正文首页：大标题 + 政策引言段落"""
    _add_para(doc, data["report_title"], size=SZ_TITLE, bold=True,
              align=WD_ALIGN_PARAGRAPH.CENTER, space_after=12)
    if data.get("opening_text"):
        _add_para(doc, data["opening_text"], indent_chars=2)


def _build_section_overall(doc, data):
    """一、整体情况"""
    _add_para(doc, "一、整体情况", size=SZ_H1, bold=True,
              indent_chars=2, space_before=6, space_after=6)
    if data.get("overall_text"):
        _add_para(doc, data["overall_text"], indent_chars=2)


def _build_section_last_year(doc, data):
    """二、上一年度问题整改情况"""
    _add_para(doc, "二、上一年度问题整改情况", size=SZ_H1, bold=True,
              indent_chars=2, space_before=6, space_after=6)
    if data.get("last_year_text"):
        _add_para(doc, data["last_year_text"], indent_chars=2)


def _format_problem_sentence(item):
    """
    把一条问题拼成模板的标准句式。
    标准依据字段缺失时显式标注，绝不编造——这是R00原文铁律的体现。
    """
    finding = item.get("finding", "").strip()
    std_name = item.get("std_name", "").strip()
    std_no = item.get("std_no", "").strip()
    article = item.get("article", "").strip()
    clause_text = item.get("clause_text", "").strip()

    # 四个依据字段必须齐全才能拼标准句式
    if all([std_name, std_no, article, clause_text]):
        return CLAUSE_SENTENCE.format(
            finding=finding, std_name=std_name, std_no=std_no,
            article=article, clause_text=clause_text)
    # 任一字段缺失：原样输出发现的问题，并显式标注依据待补
    missing = [n for n, v in [("标准名称", std_name), ("标准号", std_no),
                              ("条款号", article), ("条款原文", clause_text)] if not v]
    return f"{finding}。[标准依据待补充：缺少{'、'.join(missing)}，需检索核验后填入，本工具不代为编造]"


def _build_problem_table(doc, items):
    """
    问题表格：左栏问题描述（带编号），右栏留空给用户插图。
    最后一行是"暴露的问题："，右栏同样留空。
    """
    # 行数 = 问题条数 + 1行"暴露的问题"
    table = doc.add_table(rows=len(items) + 1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = False

    for row_idx, item in enumerate(items):
        row = table.rows[row_idx]
        left, right = row.cells[0], row.cells[1]
        left.width = COL_LEFT_W
        right.width = COL_RIGHT_W

        text = f"（{item['seq']}）{_format_problem_sentence(item)}"
        left.paragraphs[0].text = text
        left.paragraphs[0].paragraph_format.line_spacing = 1.5
        _set_cell_font(left)
        # 右栏：默认留空由用户插图；传了 images 才由本工具插入（自动带边框、统一尺寸）
        right.paragraphs[0].text = ""
        if item.get("images"):
            _insert_item_images(doc, right, item["images"])

    # 末行："暴露的问题："
    last = table.rows[len(items)]
    last.cells[0].width = COL_LEFT_W
    last.cells[1].width = COL_RIGHT_W
    last.cells[0].paragraphs[0].text = "暴露的问题："
    _set_cell_font(last.cells[0], bold=True)
    last.cells[1].paragraphs[0].text = ""
    return table


def _build_section_problems(doc, data):
    """三、本次巡查主要问题——按专业大类 → 问题小项 → 表格"""
    _add_para(doc, data.get("problems_section_title", "三、本次巡查发现的主要问题"),
              size=SZ_H1, bold=True, indent_chars=2, space_before=6, space_after=6)

    for cat in data.get("categories", []):
        # （一）安全管理方面
        _add_para(doc, f"（{cat['cat_no']}）{cat['cat_name']}",
                  size=SZ_H2, bold=True, indent_chars=2,
                  space_before=6, space_after=6)

        for prob in cat.get("problems", []):
            # 1.安全目标管理不规范。
            _add_para(doc, f"{prob['prob_no']}.{prob['prob_title']}",
                      size=SZ_H3, bold=True, indent_chars=2,
                      space_before=6, space_after=6)
            if prob.get("items"):
                _build_problem_table(doc, prob["items"])
                _add_para(doc)  # 表格后空一行

        # 每个大类末尾的整改建议
        if cat.get("suggestions"):
            _add_para(doc, "整改建议：", size=SZ_BODY, bold=True,
                      indent_chars=2, space_before=6, space_after=6)
            for s in cat["suggestions"]:
                _add_para(doc, s, indent_chars=2)
            _add_para(doc)


def _build_tail(doc, data):
    """结尾：整改要求 + 签字栏"""
    if data.get("closing_text"):
        _add_para(doc, data["closing_text"], indent_chars=2,
                  space_before=12)
    _add_para(doc)
    _add_para(doc)
    # 签字栏
    _add_para(doc, data.get("sign_left", "检查组组长（签字）："),
              indent_chars=2, space_before=12)
    _add_para(doc)
    _add_para(doc, data.get("sign_right", "受检单位负责人（签字）："),
              indent_chars=2)
    _add_para(doc)
    _add_para(doc, data.get("sign_date", "年    月    日"),
              align=WD_ALIGN_PARAGRAPH.RIGHT)


def _apply_page_setup(doc):
    """页面设置：A4 + 实测页边距"""
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.54)      # 1440 dxa
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.17)     # 1800 dxa
        section.right_margin = Cm(3.17)


# ============================================================
# 交付洁净后处理（2026-10-28 事故沉淀，出厂即合规）
#   事故1：目录里混进整段正文——模板 TOC 域是 TOC \o "1-9" \h \u，
#          \u 按段落大纲级别收录，正文段只要带 w:outlineLvl 就会被吸进目录
#   事故2：图片没边框——模板（三）3 要求 0.75pt 黑色细边框，漏加看不出来
# 所以：正文段一律不许带大纲级别；只有章/节/问题标题带；每张图必须有边框
# ============================================================
def _set_outline_level(p, lvl):
    """写 w:outlineLvl（按 CT_PPr 顺序插在 w:rPr 之前，乱放会被 Word 静默丢弃）"""
    pPr = p._p.get_or_add_pPr()
    ol = pPr.find(qn("w:outlineLvl"))
    if ol is None:
        ol = OxmlElement("w:outlineLvl")
        rPr = pPr.find(qn("w:rPr"))
        if rPr is not None:
            rPr.addprevious(ol)
        else:
            pPr.append(ol)
    ol.set(qn("w:val"), str(lvl))


def _clean_outline_levels(doc):
    """
    目录收录范围矫正：
      章  一、二、…        → lvl 0
      节  （一）（二）…     → lvl 1
      问题 ★4. / 5.         → lvl 2
      其余一切段落（正文、表格内、图注、签字栏）→ 删掉大纲级别
    返回 (设置级别数, 清除级别数)。
    """
    set_n = clear_n = 0
    for p in doc.paragraphs:
        t = "".join(x.text or "" for x in p._p.iter(qn("w:t"))).strip()
        pPr = p._p.pPr
        if RE_CH.match(t):
            _set_outline_level(p, 0)
            set_n += 1
        elif RE_SEC.match(t):
            _set_outline_level(p, 1)
            set_n += 1
        elif RE_PROB.match(t) and len(t) < 170:
            _set_outline_level(p, 2)
            set_n += 1
        elif pPr is not None and pPr.find(qn("w:outlineLvl")) is not None:
            pPr.remove(pPr.find(qn("w:outlineLvl")))
            clear_n += 1
    return set_n, clear_n


def _add_picture_borders(doc):
    """给文档内每张图补 0.75pt 黑色细边框；返回(新增数, 已有数)"""
    n = skip = 0
    for spPr in doc.element.body.iter(qn("pic:spPr")):
        if spPr.find(qn("a:ln")) is not None:
            skip += 1
            continue
        ln = parse_xml(PIC_BORDER_XML)
        geo = spPr.find(qn("a:prstGeom"))
        if geo is not None:              # CT_ShapeProperties：几何 → 填充 → 线
            geo.addnext(ln)
        else:
            xf = spPr.find(qn("a:xfrm"))
            (xf.addnext(ln) if xf is not None else spPr.insert(0, ln))
        n += 1
    return n, skip


def _patch_settings(path):
    """在 settings.xml 写入 doNotAutoCompressPictures，避免后续保存时图片被重编码"""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/settings.xml":
                s = data.decode("utf-8")
                if "doNotAutoCompressPictures" not in s:
                    s = s.replace(
                        "</w:settings>",
                        '<w:doNotAutoCompressPictures w:val="true"/></w:settings>')
                data = s.encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, path)


def _insert_item_images(doc, cell, images):
    """
    可选：把证据照片直接插进右栏（不传 images 则保持右栏留空）。
    尺寸只走两种：横版 6.0×4.5cm、竖版 4.5×6.0cm（允许变形，保证版面统一）；
    边框由 _add_picture_borders 统一补，不在此处重复。
    """
    from docx.shared import Emu
    from docx.image.image import Image
    added = 0
    for img in images or []:
        path = img.get("path") if isinstance(img, dict) else img
        if not path or not os.path.exists(path):
            continue
        try:
            meta = Image.from_file(path)
            cx, cy = SZ_PIC_PORT if meta.px_height > meta.px_width else SZ_PIC_LAND
        except Exception:
            cx, cy = SZ_PIC_LAND
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing = 1.0    # 行距必须 auto，否则会裁图
        p.add_run().add_picture(path, width=Emu(cx), height=Emu(cy))
        added += 1
    return added


def _add_page_number(doc):
    """页脚居中页码——模板里是有页码的"""
    for section in doc.sections:
        footer = section.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        fldChar1 = OxmlElement("w:fldChar")
        fldChar1.set(qn("w:fldCharType"), "begin")
        instrText = OxmlElement("w:instrText")
        instrText.set(qn("xml:space"), "preserve")
        instrText.text = "PAGE"
        fldChar2 = OxmlElement("w:fldChar")
        fldChar2.set(qn("w:fldCharType"), "end")
        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
        _set_run(run, size=SZ_BODY)


def generate_report(data, out_path):
    """
    主入口：把结构化数据渲染成整改通知单docx。

    data结构见 EXAMPLE_DATA，核心是categories → problems → items 三层嵌套。
    本函数不做任何内容生成，只做格式排版。
    出厂即做交付洁净后处理（大纲级别矫正 + 图片边框 + 禁压缩标记），
    返回处理统计，供调用方如实回报。
    """
    doc = Document()
    # 设置全局默认字体
    style = doc.styles["Normal"]
    style.font.name = FONT_BODY
    style.font.size = SZ_BODY
    style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_BODY)

    _apply_page_setup(doc)
    _add_page_number(doc)

    _build_cover(doc, data)
    _build_opening(doc, data)
    _build_section_overall(doc, data)
    _build_section_last_year(doc, data)
    _build_section_problems(doc, data)
    _build_tail(doc, data)

    # ── 交付洁净后处理（避免返工）────────────────────────────
    # 1) 大纲级别：只有章/节/问题标题带，正文一律清除（目录才不会吸进正文）
    lvl_set, lvl_clear = _clean_outline_levels(doc)
    # 2) 图片边框：每张图 0.75pt 黑色细边框
    bd_new, bd_old = _add_picture_borders(doc)

    doc.save(out_path)
    # 3) 禁压缩标记：写进 settings.xml，后续保存不重编码图片
    _patch_settings(out_path)
    return {
        "path": out_path,
        "大纲级别_已设置标题数": lvl_set,
        "大纲级别_已清除正文数": lvl_clear,
        "图片边框_新加": bd_new,
        "图片边框_原有": bd_old,
    }


def postprocess_docx(path, out_path=None):
    """
    对**任意已存在的成品 docx**做交付洁净后处理（不必是本工具生成的）：
      1) 大纲级别矫正：章 lvl0 / 节 lvl1 / 问题 lvl2，其余段落一律清除
      2) 每张图补 0.75pt 黑色细边框
      3) settings.xml 写 doNotAutoCompressPictures
    只改 XML，不动 word/media，图片字节完全不变。
    out_path 为空则原地保存；即使原地保存也不是"用 Word 保存"，不会重编码图片。
    """
    src = path
    dst = out_path or path
    d = Document(src)
    lvl_set, lvl_clear = _clean_outline_levels(d)
    bd_new, bd_old = _add_picture_borders(d)
    d.save(dst)
    _patch_settings(dst)
    return {
        "path": dst,
        "大纲级别_已设置标题数": lvl_set,
        "大纲级别_已清除正文数": lvl_clear,
        "图片边框_新加": bd_new,
        "图片边框_原有": bd_old,
    }


# ============================================================
# 示例数据结构（--print-example 会输出这个）
# ============================================================
EXAMPLE_DATA = {
    "report_title": "XX发电安全综合巡查报告",
    "company": "示例电力股份有限公司",
    "report_date": "2026年7月2日",
    "opening_text": "为深入贯彻落实……（政策引言段落，由调用方提供，本工具不代为生成）",
    "overall_text": "2025年安全综合巡查发现问题41大项69小项……本次安全综合巡查发现问题X大项X小项，重点问题X项。",
    "last_year_text": "本次巡查期间，巡查组对历年巡查问题整改落实情况进行了复查。经复查，已整改X项，整改率X%。",
    "problems_section_title": "三、2026年安全综合巡查主要问题",
    "categories": [
        {
            "cat_no": "一",
            "cat_name": "安全管理方面",
            "problems": [
                {
                    "prob_no": "1",
                    "prob_title": "安全目标管理不规范。",
                    "items": [
                        {
                            "seq": "1",
                            "finding": "查阅XX公司与员工签订的2026年度安全环保目标责任书控制措施中，存在将保证实现安全目标的控制措施当成安全目标的情况",
                            "std_name": "安全目标管理规定",
                            "std_no": "Q/DHP 2187-2024",
                            "article": "3.2",
                            "clause_text": "四级控制 一级是企业控制轻伤和一类障碍，不发生人身重伤和设备、电网和供热事故",
                            "images": ["D:/照片/证据1.jpg"]
                        }
                    ]
                }
            ],
            "suggestions": [
                "一是……（整改建议由调用方提供，可基于定责分析起草，但须标注AI生成草稿）"
            ]
        }
    ],
    "closing_text": "请受检单位针对本次巡查发现的问题，制定整改计划，明确责任人和整改期限，于X日内将整改情况报送……",
    "sign_left": "检查组组长（签字）：",
    "sign_right": "受检单位负责人（签字）：",
    "sign_date": "年    月    日"
}
