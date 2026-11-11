# -*- coding: utf-8 -*-
"""
md_to_docx.py — 通用Markdown转Word转换器（示例电力工作包/方案专用）

用途：将施工组织方案、专项施工方案等Markdown文档转换为带真实表格的正式Word文档。
背景：2026-10-06金隅兴发项目初版转换把md表格降级为文本行，不符合正式文档要求，故重写。

支持语法：
  # / ## / ### / #### 标题（黑体，居中仅一级标题）
  | a | b | 表格（转为Word真实表格，含表头加粗）
  - 无序列表 / 1. 有序列表
  **加粗**（行内解析）
  > 引用（仿宋灰色缩进，用于警示声明）
  普通段落（仿宋小四=12pt）

用法：
  python md_to_docx.py <input.md> [output.docx] [--title "文档标题"]
  省略output时默认同名.docx；--title省略时取md首个# 标题。

依赖：python-docx（.trae/python-venv已装）
"""
import os
import re
import sys
import argparse

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn


def set_cn_font(run, name, size, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    # 中文字体需同时设置 eastAsia
    r = run._element.rPr.rFonts
    r.set(qn('w:eastAsia'), name)


def add_runs_with_bold(paragraph, text, font_name='仿宋', size=12, base_bold=False):
    """解析 **加粗** 行内语法，分段加run。"""
    parts = re.split(r'(\*\*[^*]+\*\*)', text)
    for part in parts:
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            run = paragraph.add_run(part[2:-2])
            set_cn_font(run, font_name, size, bold=True)
        else:
            run = paragraph.add_run(part)
            set_cn_font(run, font_name, size, bold=base_bold)


def add_heading(doc, text, level):
    p = doc.add_paragraph()
    if level == 1:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        set_cn_font(run, '黑体', 22, bold=True)
        doc.add_paragraph()  # 标题后空行
    elif level == 2:
        run = p.add_run(text)
        set_cn_font(run, '黑体', 16, bold=True)
    elif level == 3:
        run = p.add_run(text)
        set_cn_font(run, '黑体', 14, bold=True)
    else:
        run = p.add_run(text)
        set_cn_font(run, '黑体', 12, bold=True)


def add_table(doc, rows):
    """rows: list[list[str]]，第一行为表头。"""
    if not rows:
        return
    ncols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=ncols)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j in range(ncols):
            cell_text = row[j] if j < len(row) else ''
            cell = table.cell(i, j)
            cell.text = ''
            p = cell.paragraphs[0]
            # 表头加粗；表格内不使用颜色填充底色（用户偏好）
            add_runs_with_bold(p, cell_text, font_name='仿宋', size=10.5, base_bold=(i == 0))


def md_to_docx(md_path, out_path, title=None):
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.read().splitlines()

    doc = Document()
    # 页边距：上3.7 下3.5 左2.8 右2.6（党政机关公文格式）
    for section in doc.sections:
        section.top_margin = Cm(3.7)
        section.bottom_margin = Cm(3.5)
        section.left_margin = Cm(2.8)
        section.right_margin = Cm(2.6)

    # 标题：优先用参数，否则用首个一级标题
    if title:
        add_heading(doc, title, 1)

    table_buf = []

    def flush_table():
        if table_buf:
            add_table(doc, table_buf)
            table_buf.clear()

    first_h1_skipped = False
    for line in lines:
        raw = line
        line = line.strip()

        # 表格行累积
        if line.startswith('|'):
            cells = [c.strip() for c in line.split('|')]
            # 去掉首尾空cell（markdown表格分隔符产生）
            if cells and cells[0] == '':
                cells = cells[1:]
            if cells and cells[-1] == '':
                cells = cells[:-1]
            # 跳过分隔行 |---|---|
            if cells and all(re.fullmatch(r':?-{2,}:?', c or '---') for c in cells):
                continue
            table_buf.append(cells)
            continue
        else:
            flush_table()

        if not line:
            continue
        if line == '---':
            continue
        if line.startswith('# ') and not line.startswith('## '):
            # 一级标题：若已用参数title则跳过首个，否则作为标题
            if title or first_h1_skipped:
                add_heading(doc, line[2:], 1)
            else:
                add_heading(doc, line[2:], 1)
                first_h1_skipped = True
        elif line.startswith('#### '):
            add_heading(doc, line[5:], 4)
        elif line.startswith('### '):
            add_heading(doc, line[4:], 3)
        elif line.startswith('## '):
            add_heading(doc, line[3:], 2)
        elif line.startswith('> '):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(1.0)
            add_runs_with_bold(p, line[2:], font_name='仿宋', size=10.5)
            for run in p.runs:
                run.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
        elif re.match(r'^[-*] ', line):
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.left_indent = Cm(0.75)
            add_runs_with_bold(p, line[2:])
        elif re.match(r'^\d+\. ', line):
            p = doc.add_paragraph(style='List Number')
            p.paragraph_format.left_indent = Cm(0.75)
            add_runs_with_bold(p, re.sub(r'^\d+\. ', '', line))
        else:
            p = doc.add_paragraph()
            add_runs_with_bold(p, line)

    flush_table()
    doc.save(out_path)
    return out_path


def main():
    ap = argparse.ArgumentParser(description='Markdown转Word（示例电力工作包/方案专用）')
    ap.add_argument('input', help='输入md文件路径')
    ap.add_argument('output', nargs='?', help='输出docx路径（默认同名.docx）')
    ap.add_argument('--title', help='文档标题（默认取md首个# 标题）')
    args = ap.parse_args()

    out = args.output or os.path.splitext(args.input)[0] + '.docx'
    md_to_docx(args.input, out, title=args.title)
    print('OK: ' + out)


if __name__ == '__main__':
    main()
