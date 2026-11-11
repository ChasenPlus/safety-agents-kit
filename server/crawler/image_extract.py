# -*- coding: utf-8 -*-
"""文档图片抽取与资产落盘的原子能力（映射表驱动，禁止正则猜文件名）。

上层约定（见 docs/architecture.md）：
- 图片目录：与 md 同目录的 ``{stem}_images/``
- 命名：docx 图 d01.png；pdf 图 p001_01.png
- md 引用一律正斜杠相对路径：``![图注](./{stem}_images/d01.png)``

本模块只做三件事：
1. docx：从 python-docx 对象取 rels 图片字节、按 XML 顺序枚举段落/单元格内的 rId；
2. pdf：按页枚举图片块（含 bbox/像素尺寸/xref）并做图标/水印/整页扫描过滤；
3. 落盘：Pillow 长边压缩/超限转 JPEG，返回最终文件名（gif/svg 原样透传）。
"""
from __future__ import annotations

import io
import os
from typing import Dict, List, Optional, Tuple

# Office Open XML 命名空间
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_V = "urn:schemas-microsoft-com:vml"

# PDF 内嵌图降噪阈值
PDF_MIN_PX = 120            # 图片像素宽或高低于此值视为图标/装饰
PDF_MIN_AREA_RATIO = 0.02   # bbox 占页面面积低于此值视为水印/小标
PDF_SCAN_PAGE_RATIO = 0.50  # bbox 占面超过此值且页面几乎无文字 → 整页扫描，不抽
IMG_MAX_LONG_SIDE = 1600
IMG_MAX_BYTES = 2 * 1024 * 1024


# ============================================================
# docx
# ============================================================

def docx_rels_images(doc) -> Dict[str, Tuple[str, bytes]]:
    """从 python-docx Document 取全部图片关系。

    返回 {rId: (扩展名(小写,无点), 字节)}。
    header/footer 等部件的图片关系也在 part 包中，但只有 document.xml
    正文引用到的 rId 才会被上层取用，多余关系无副作用。
    """
    out: Dict[str, Tuple[str, bytes]] = {}
    rels = getattr(doc.part, "rels", None)
    if not rels:
        return out
    for rid, rel in rels.items():
        try:
            target = rel.target_part
        except Exception:
            continue
        ref = getattr(target, "partname", None)
        blob = getattr(target, "blob", None)
        if blob is None or ref is None:
            continue
        ext = os.path.splitext(str(ref))[1].lower().lstrip(".")
        if ext in {"jpg", "jpeg", "png", "gif", "bmp", "webp", "tif", "tiff", "svg"}:
            out[rid] = (ext, bytes(blob))
    return out


def iter_element_rids(element) -> List[str]:
    """按 XML 出现顺序枚举一个段落/单元格内引用的图片 rId（含重复引用）。

    同时覆盖现代 DrawingML（<a:blip r:embed>）与老版 VML（<v:imagedata r:id>）。
    """
    rids: List[str] = []
    for el in element.iter():
        tag = el.tag
        if tag == "{%s}blip" % _NS_A:
            rid = el.get("{%s}embed" % _NS_R) or el.get("{%s}link" % _NS_R)
            if rid:
                rids.append(rid)
        elif tag == "{%s}imagedata" % _NS_V:
            rid = el.get("{%s}id" % _NS_R)
            if rid:
                rids.append(rid)
    return rids


# ============================================================
# pdf
# ============================================================

def page_image_blocks(page,
                      min_px: int = PDF_MIN_PX,
                      min_area_ratio: float = PDF_MIN_AREA_RATIO,
                      scan_ratio: float = PDF_SCAN_PAGE_RATIO
                      ) -> List[dict]:
    """枚举页面图片块，按阅读顺序（y0, x0）排序并过滤。

    基于 ``get_image_info(xrefs=True)`` 而非 ``get_text("dict")``——后者的图片
    块 xref 为 None，无法回取字节。返回：
    [{xref, bbox(x0,y0,x1,y1), wpx, hpx, area_ratio, is_fullpage}]
    - 像素尺寸 <min_px> 或面积占比 <min_area_ratio> → 图标/水印/蒙版碎片，剔除；
    - 同 xref 在一页内出现多次（蒙版拼贴）只保留面积最大的一块；
    - is_fullpage=True 表示该图几乎铺满整页（扫描附图页），上层不抽对象。
    """
    page_area = max(page.rect.width * page.rect.height, 1.0)
    merged: Dict[int, dict] = {}
    try:
        infos = page.get_image_info(xrefs=True)
    except Exception:
        return []
    for info in infos or []:
        bbox = info.get("bbox")
        if not bbox or len(bbox) != 4:
            continue
        wpx = int(info.get("width", 0) or 0)
        hpx = int(info.get("height", 0) or 0)
        if wpx < min_px or hpx < min_px:
            continue
        bw, bh = bbox[2] - bbox[0], bbox[3] - bbox[1]
        ratio = (bw * bh) / page_area
        if ratio < min_area_ratio:
            continue
        xref = int(info.get("xref", 0) or 0)
        if not xref:
            continue  # 无 xref 无法回取字节，放弃（理论上 xrefs=True 都有）
        item = {
            "xref": xref,
            "bbox": tuple(float(v) for v in bbox),
            "wpx": wpx,
            "hpx": hpx,
            "area_ratio": ratio,
            "is_fullpage": ratio >= scan_ratio,
        }
        if xref in merged:
            if ratio > merged[xref]["area_ratio"]:
                merged[xref] = item
        else:
            merged[xref] = item
    out = list(merged.values())
    out.sort(key=lambda it: (round(it["bbox"][1], 1), it["bbox"][0]))
    return out


def extract_xref_image(fitz_doc, xref: int) -> Optional[Tuple[str, bytes]]:
    """抽 xref 对应图片字节，返回 (扩展名, bytes)；失败返回 None。"""
    try:
        info = fitz_doc.extract_image(xref)
    except Exception:
        return None
    if not info:
        return None
    ext = (info.get("ext") or "png").lower()
    img = info.get("image")
    if not img:
        return None
    return ext, bytes(img)


# ============================================================
# 落盘（压缩/映射）
# ============================================================

def save_normalized(raw: bytes,
                    ext: str,
                    out_path_no_ext: str,
                    max_long_side: int = IMG_MAX_LONG_SIDE,
                    max_bytes: int = IMG_MAX_BYTES
                    ) -> Optional[str]:
    """把图片字节落盘并做体积归一，返回最终文件名（含扩展名，不含目录）。

    - gif/svg 原样透传（动图/矢量不压扁）；
    - 长边 >max_long_side 等比缩小；
    - 位图处理后仍 >max_bytes → JPEG q85（透明通道合成白底）；
    - Pillow 无法识别时原样写盘。
    """
    ext = (ext or "png").lower().lstrip(".")
    if ext in ("gif", "svg"):
        name = os.path.basename(out_path_no_ext) + "." + ext
        with open(out_path_no_ext + "." + ext, "wb") as f:
            f.write(raw)
        return name

    final_ext = ext if ext in ("png", "jpg", "jpeg", "bmp", "webp", "tif", "tiff") else "png"
    try:
        from PIL import Image
        img = Image.open(io.BytesIO(raw))
        img.load()
        if max(img.size) > max_long_side:
            img.thumbnail((max_long_side, max_long_side), Image.LANCZOS)
        save_fmt = "JPEG" if final_ext in ("jpg", "jpeg") else (
            "PNG" if final_ext == "png" else final_ext.upper().replace("TIF", "TIFF"))
        buf = io.BytesIO()
        if save_fmt == "JPEG" and img.mode in ("RGBA", "LA", "P"):
            bg = Image.new("RGB", img.size, (255, 255, 255))
            rgba = img.convert("RGBA")
            bg.paste(rgba, mask=rgba.split()[-1])
            img = bg
        elif save_fmt != "JPEG" and img.mode not in ("RGB", "RGBA", "L", "LA"):
            img = img.convert("RGB")
        img.save(buf, format=save_fmt, quality=85 if save_fmt == "JPEG" else None)
        out_bytes = buf.getvalue()
        if save_fmt != "JPEG" and len(out_bytes) > max_bytes:
            # PNG 仍超限（通常是照片型大图）→ 转 JPEG
            bg = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode in ("RGBA", "LA", "P"):
                rgba = img.convert("RGBA")
                bg.paste(rgba, mask=rgba.split()[-1])
                img2 = bg
            else:
                img2 = img.convert("RGB")
            buf = io.BytesIO()
            img2.save(buf, format="JPEG", quality=85)
            out_bytes = buf.getvalue()
            final_ext = "jpg"
    except Exception:
        out_bytes = raw
        final_ext = "png" if final_ext not in ("jpg", "jpeg") else final_ext

    fname = os.path.basename(out_path_no_ext) + "." + final_ext
    with open(out_path_no_ext + "." + final_ext, "wb") as f:
        f.write(out_bytes)
    return fname
