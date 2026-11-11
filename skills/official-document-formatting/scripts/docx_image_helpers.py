# -*- coding: utf-8 -*-
"""公文 docx 插图公共助手。

图片保真管线的出文档侧：工作包 JSON 里 dict 项可带 image / image_caption 字段，
- 路径为绝对路径直接使用；相对路径依次相对 data 文件目录、output_dir 解析
  （现场照片约定放输出目录下的 现场照片/ 子目录）。
- 图片居中、宽≤15cm（竖长图改为高≤20cm），图注小五黑体居中。
- 路径不存在不静默：插灰色占位段【图片缺失：path】并登记进 missing 列表，
  由调用方汇总进结果 JSON 的 missing_images。
"""

import os

from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

TITLE_FONT = "黑体"
BODY_FONT = "仿宋_GB2312"

IMG_MAX_WIDTH_CM = 15.0
IMG_MAX_HEIGHT_CM = 20.0


def set_cjk_font(run, font_name, size=None, bold=None, color=None):
    """中文字体需通过 oxml 设 eastAsia，与 generate_work_package_docs 中同名函数一致。"""
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


def resolve_image_path(path, base_dir=None, output_dir=None):
    """按 绝对路径 → base_dir → output_dir → 当前工作目录 顺序解析，返回
    存在的绝对路径；都不存在返回 None（调用方按缺失处理，不猜测）。"""
    if not path:
        return None
    norm = path.replace("/", os.sep).strip()
    cands = []
    if os.path.isabs(norm):
        cands.append(norm)
    else:
        rel = norm.lstrip("." + os.sep)
        for d in (base_dir, output_dir):
            if d:
                cands.append(os.path.join(d, rel))
        cands.append(os.path.abspath(norm))
    for c in cands:
        if os.path.isfile(c):
            return c
    return None


def _image_display_size(abs_path):
    """按像素宽高比决定 docx 约束边：横图限宽 15cm，竖长图限高 20cm。
    读不到尺寸时保守按限宽处理。"""
    try:
        from PIL import Image
        with Image.open(abs_path) as im:
            w, h = im.size
        if h > w * 1.33:  # 竖长图：按高约束，避免单图超出一页
            return None, Cm(IMG_MAX_HEIGHT_CM)
    except Exception:
        pass
    return Cm(IMG_MAX_WIDTH_CM), None


def add_image_block(doc, path, caption="", base_dir=None, output_dir=None,
                    missing=None):
    """在文档当前位置插入图片+图注；缺图插灰色占位并登记。返回 True 插图成功。"""
    if missing is None:
        missing = []
    abs_path = resolve_image_path(path, base_dir=base_dir, output_dir=output_dir)
    if not abs_path:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(f"【图片缺失：{path}】")
        set_cjk_font(r, BODY_FONT, size=12, color=RGBColor(0x80, 0x80, 0x80))
        missing.append(path)
        return False

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    width, height = _image_display_size(abs_path)
    p.add_run().add_picture(abs_path, width=width, height=height)

    if caption:
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_after = Pt(8)
        cr = cp.add_run(caption)
        set_cjk_font(cr, TITLE_FONT, size=9, bold=False,
                     color=RGBColor(0x40, 0x40, 0x40))
    return True


def get_item_image(item):
    """dict 型条目取图片字段 (image, image_caption)，非 dict 或无图返回 ("","")。"""
    if isinstance(item, dict):
        return item.get("image", ""), item.get("image_caption", "")
    return "", ""
