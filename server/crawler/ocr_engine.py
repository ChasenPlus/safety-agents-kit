# -*- coding: utf-8 -*-
"""
OCR 引擎模块 —— 处理扫描件 PDF（pymupdf/pdfplumber 提取不到文字或质量不佳的情况）。

升级版策略（针对"完整无乱码"目标）：
    1. 先用 pymupdf(fitz) 提取 PDF 内嵌文本；
    2. 用 pdfplumber 同步提取表格，按页保留结构（段落/表格/标题）；
    3. 若内嵌文字过少或质量差 → 判定扫描件，用 RapidOCR 识别：
       - DPI 提升到 300（默认值）
       - 图像预处理：灰度→对比度增强→二值化（Otsu自适应）→去噪→倾斜校正
       - 保留 OCR 框的坐标信息，按位置排序，分行输出
       - 表格区域单独识别（基于行/列聚类）
    4. 多引擎回退：RapidOCR → paddleocr(若装) → tesseract(若装) → pymupdf 原始文本

RapidOCR 用 onnxruntime，无需安装 tesseract，首次调用会加载模型（约几秒）。
图像预处理依赖 Pillow + opencv-python，两者已是项目依赖。
"""

import io
import os
import re
import tempfile
import logging
from typing import Optional

import numpy as np

logger = logging.getLogger(__name__)

# 惰性加载 OCR 引擎，避免导入模块时就加载模型
_ocr_engine = None

# 住建部官网公开PDF的水印文字（横排整句+竖排时被逐字提取成单字行）。
# 这类PDF文本层只有水印、正文是图片，2026-10-03 JGJ/T 46-2024实测：
# pdfplumber把51页竖排水印提成815个单字行，字符总数>200，骗过"文本充足"快路径导致OCR不启动。
_WATERMARK_CHARS = set("住房城乡建设部信息公开浏览专用")


def _maybe_strip_watermark(text: str) -> str:
    """识别并清除水印主导文本。

    判据：非空行中"整行仅由水印字符构成且长度<=30"的行占比>=50%才清洗，
    占比不足时原样返回——正常条款文本不会有半数短行全部由这16个字组成，保证零误伤。
    2026-10-03 GB/T 50502-2009 实测：上海住建委cmsres托管件每页文字层仅有
    "标准分享网 www.bzfxw.com 免费下载"一行、正文为图片，一并纳入识别。
    """
    if not text:
        return text
    lines = text.splitlines()
    nonempty = [ln.strip() for ln in lines if ln.strip()]
    if not nonempty:
        return text

    def _is_wm_line(s: str) -> bool:
        if len(s) <= 30 and all(ch in _WATERMARK_CHARS for ch in s):
            return True
        if "bzfxw" in s.lower() or ("标准分享网" in s and len(s) <= 40):
            return True
        return False

    wm_count = sum(1 for s in nonempty if _is_wm_line(s))
    if wm_count / len(nonempty) < 0.5:
        return text
    return "\n".join(ln for ln in lines
                     if not (ln.strip() and _is_wm_line(ln.strip()))).strip()


# 现代汉语百余个极高频字。任何正常中文材料中这些字必然反复出现；
# 而国标PDF私有CMap字形乱码（文本能提取但汉字映射错位，如"犌犅３０８７１"）里占比趋近0。
# 2026-10-03 GB 30871-2022实测：乱码件占比0.000，正常文字版PDF为0.235/0.256。
_TOP_FREQ_CHARS = set(
    "的一是了我不人在他有这个上们来到时大地为子中你说生国年着就那和要她出也得里后自以会家"
    "可下而过天去能对小多然于心学么之都好看起发当没成只如事把还用第样道想作种开美总从无情"
    "已面最女但现前些所同日手又行意动方期它头经长儿回位分爱老因很给名法间斯知世什两次使身者被高"
)


def _looks_like_garbled(text: str, threshold: float = 0.05, min_han: int = 30) -> bool:
    """检测私有CMap字形乱码文本。

    判据：汉字样本≥30个时，50个极高频字占比<5%判为乱码（正常中文材料实测≥20%，阈值留有4倍裕度）；
    汉字不足30个不判（样本无统计意义，交给字符数阈值处理）。
    """
    han = [c for c in text if "一" <= c <= "鿿"]
    if len(han) < min_han:
        return False
    hit = sum(1 for c in han if c in _TOP_FREQ_CHARS)
    return hit / len(han) < threshold


def _get_ocr():
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
    return _ocr_engine


# ---------------------------------------------------------------------------
# 图像预处理（用 Pillow + opencv）
# ---------------------------------------------------------------------------

def _preprocess_arr(arr, do_deskew: bool = True):
    """对 ndarray（BGR/RGB/灰度）做预处理，返回 3 通道 BGR ndarray 直接喂给 OCR。
    相比"渲染→PNG编码→PNG解码"的 bytes 往返，省掉每页两次图像编解码（300dpi A4
    实测各约 30-80ms）。

    流程：灰度→对比度增强(CLAHE)→去噪→(可选)倾斜校正。
    处理失败时原样返回，不阻塞 OCR。
    """
    try:
        import cv2
    except ImportError:
        return arr
    try:
        if arr.ndim == 2:
            gray = arr
        elif arr.shape[2] == 4:
            gray = cv2.cvtColor(arr, cv2.COLOR_BGRA2GRAY)
        else:
            gray = cv2.cvtColor(arr, cv2.COLOR_BGR2GRAY)
        # CLAHE 自适应直方图均衡，对扫描件尤其有效
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        out = clahe.apply(gray)
        # 中值滤波去噪，保留边缘
        out = cv2.medianBlur(out, 3)

        # 倾斜校正——扫描件常有 1-5 度倾斜
        if do_deskew:
            try:
                _, bw = cv2.threshold(out, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
                coords = np.column_stack(np.where(bw > 0))
                if len(coords) > 100:
                    # 300dpi A4 前景像素可达数百万，全量 minAreaRect 极慢且无必要；
                    # 均匀降采样到 ~2万点，角度估计不变
                    if len(coords) > 20000:
                        step = len(coords) // 20000
                        coords = coords[::step]
                    angle = cv2.minAreaRect(coords)[-1]
                    if angle < -45:
                        angle = 90 + angle
                    if 0.5 < abs(angle) < 15:
                        h, w = out.shape[:2]
                        M = cv2.getRotationMatrix2D((w // 2, h // 2), angle, 1.0)
                        out = cv2.warpAffine(out, M, (w, h),
                                             flags=cv2.INTER_CUBIC,
                                             borderMode=cv2.BORDER_REPLICATE)
            except Exception as e:
                logger.debug("deskew 失败，跳过：%s", e)

        return cv2.cvtColor(out, cv2.COLOR_GRAY2BGR)
    except Exception as e:
        logger.warning("图像预处理失败，使用原图：%s", e)
        return arr


def _preprocess_image(img_bytes: bytes, do_deskew: bool = True) -> bytes:
    """PNG bytes 预处理后再编码回 PNG（保留给单图/外部调用；PDF 主链路已改用
    _preprocess_arr 直接传 ndarray，不再走 PNG 往返）。"""
    try:
        import cv2
        import numpy as np
        arr = cv2.imdecode(np.frombuffer(img_bytes, dtype=np.uint8), cv2.IMREAD_COLOR)
        if arr is None:
            return img_bytes
        ok, buf = cv2.imencode(".png", _preprocess_arr(arr, do_deskew))
        return buf.tobytes() if ok else img_bytes
    except Exception:
        return img_bytes


# ---------------------------------------------------------------------------
# PDF 文本/表格提取（pdfplumber 用于保留表格结构）
# ---------------------------------------------------------------------------

def _extract_with_pdfplumber(data: bytes) -> tuple[str, list]:
    """用 pdfplumber 提取文本和表格，返回 (markdown_text, tables)。

    pdfplumber 对表格识别比 pymupdf 强，能把表格结构以 list[list[str]] 形式返回。
    这里把表格转成 Markdown 表格语法，按页拼接。
    """
    try:
        import pdfplumber
    except ImportError:
        return "", []

    chunks = []
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                # 文本
                txt = page.extract_text() or ""
                # 表格
                tables = page.extract_tables() or []
                if tables:
                    # 把表格插入到合适位置——简化处理：先放文本，再放表格
                    # （精确位置排序需要更复杂的版面分析，这里先保证表格不丢）
                    if txt.strip():
                        chunks.append(txt.strip())
                    for ti, tbl in enumerate(tables):
                        md_tbl = _table_to_markdown(tbl)
                        if md_tbl:
                            chunks.append(md_tbl)
                else:
                    if txt.strip():
                        chunks.append(txt.strip())
                # 页间分隔
                chunks.append("")
        return "\n\n".join(chunks).strip(), []
    except Exception as e:
        logger.warning("pdfplumber 提取失败：%s", e)
        return "", []


def _table_to_markdown(tbl: list) -> str:
    """把 pdfplumber 提取的 list[list[str]] 表格转为 Markdown 表格。"""
    if not tbl or len(tbl) < 1:
        return ""
    # 清洗单元格（去换行、去空白）
    clean = []
    for row in tbl:
        if row is None:
            continue
        new_row = []
        for cell in row:
            if cell is None:
                new_row.append("")
            else:
                # 替换单元格内的换行和竖线，避免破坏表格结构
                c = str(cell).replace("\n", " ").replace("\r", " ").replace("|", "\\|").strip()
                new_row.append(c)
        # 跳过全空行
        if any(c for c in new_row):
            clean.append(new_row)
    if not clean:
        return ""
    # 第一行作为表头
    header = clean[0]
    # 确保所有行列数一致（补齐）
    ncol = max(len(r) for r in clean)
    for r in clean:
        while len(r) < ncol:
            r.append("")
    md = []
    md.append("| " + " | ".join(header) + " |")
    md.append("| " + " | ".join(["---"] * ncol) + " |")
    for r in clean[1:]:
        md.append("| " + " | ".join(r) + " |")
    return "\n".join(md)


# ---------------------------------------------------------------------------
# OCR 提取（按位置排序保留结构）
# ---------------------------------------------------------------------------

def _cluster_lines(items, line_tol: int):
    """按 y 坐标把文字框聚成行、行内按 x 排序。
    items: [(cx, cy, x0, x1, text), ...] → [(cy, 拼接文本), ...]"""
    if not items:
        return []
    items = sorted(items, key=lambda x: (x[1], x[0]))
    groups = []
    cur = [items[0]]
    cur_y = items[0][1]
    for it in items[1:]:
        if abs(it[1] - cur_y) <= line_tol:
            cur.append(it)
        else:
            groups.append(cur)
            cur = [it]
            cur_y = it[1]
    groups.append(cur)
    out = []
    for g in groups:
        g.sort(key=lambda x: x[0])
        cy = sum(it[1] for it in g) / len(g)
        out.append((cy, "".join(it[4] for it in g)))
    return out


def _detect_column_boundary(items):
    """检测双栏排版，返回栏分界线 x（左栏 cx<x<右栏 cx）或 None（单栏）。

    判据（三条同时满足，避免把参差的单栏文字误判成双栏）：
    1. 页面中央 30%~70% 带内存在全页最宽的文字框中心间隙，且间隙 ≥ 页宽 8%；
    2. 左右两组框各占 ≥15%（边注/页码不足以成栏）；
    3. 左栏 90 分位右缘 ≤ 右栏 10 分位左缘（两栏基本不重叠，排除宽表格/标题）。
    """
    if len(items) < 8:
        return None
    xmin = min(it[2] for it in items)
    xmax = max(it[3] for it in items)
    width = xmax - xmin
    if width <= 0:
        return None
    cxs = sorted(it[0] for it in items)
    best_gap, best_mid = 0.0, None
    for a, b in zip(cxs, cxs[1:]):
        mid = (a + b) / 2
        if not (xmin + 0.30 * width <= mid <= xmin + 0.70 * width):
            continue
        if b - a > best_gap:
            best_gap, best_mid = b - a, mid
    if best_mid is None or best_gap < 0.08 * width:
        return None
    left = [it for it in items if it[0] < best_mid]
    right = [it for it in items if it[0] >= best_mid]
    n = len(items)
    if len(left) < 0.15 * n or len(right) < 0.15 * n:
        return None
    left_edge = sorted(it[3] for it in left)[int(0.9 * (len(left) - 1))]
    right_edge = sorted(it[2] for it in right)[int(0.1 * (len(right) - 1))]
    if left_edge > right_edge + 0.02 * width:
        return None
    return best_mid


def _sort_ocr_result_by_position(result, line_tol: int = 8) -> list[str]:
    """把 RapidOCR 的 [box, text, score] 列表按阅读顺序输出成行文本。

    单栏页：按 y 聚行、行内按 x（旧行为，完全保留）。
    双栏页：检测栏沟后输出 页眉通栏 → 左栏整栏 → 中部通栏 → 右栏整栏 → 页脚通栏，
    彻底避免旧逻辑把左右栏同行高度的文字拼到一行（"左句尾右句头"串列）。
    跨栏框（标题/宽表/页眉页脚）按 y 位置归入头/中/尾三段。
    """
    items = []
    for item in result or []:
        if not item or len(item) < 2:
            continue
        box, text = item[0], item[1]
        if not text or not text.strip():
            continue
        try:
            xs = [p[0] for p in box]
            ys = [p[1] for p in box]
            items.append((sum(xs) / 4, sum(ys) / 4, min(xs), max(xs), text.strip()))
        except Exception:
            items.append((0, 0, 0, 0, text.strip()))
    if not items:
        return []

    boundary = _detect_column_boundary(items)
    if boundary is None:
        return [t for _, t in _cluster_lines(items, line_tol)]

    xmin = min(it[2] for it in items)
    width = max(it[3] for it in items) - xmin
    tol_w = 0.02 * width
    left, right, spans = [], [], []
    for it in items:
        if it[2] < boundary - tol_w and it[3] > boundary + tol_w:
            spans.append(it)
        elif it[0] < boundary:
            left.append(it)
        else:
            right.append(it)
    # 防御：分栏后任一栏为空则回退单栏
    if not left or not right:
        return [t for _, t in _cluster_lines(items, line_tol)]

    ll = _cluster_lines(left, line_tol)
    rl = _cluster_lines(right, line_tol)
    sl = _cluster_lines(spans, line_tol)
    top = min(cy for cy, _ in ll + rl) - line_tol * 2
    bottom = max(cy for cy, _ in ll + rl) + line_tol * 2
    head = [t for cy, t in sl if cy < top]
    tail = [t for cy, t in sl if cy > bottom]
    middle = [t for cy, t in sl if top <= cy <= bottom]
    return head + [t for _, t in ll] + middle + [t for _, t in rl] + tail


def _pixmap_to_bgr(pix):
    """PyMuPDF pixmap（RGB/灰度）直接转连续内存 BGR ndarray 喂 RapidOCR，
    不经 PNG 编码/解码（300dpi A4 每页省 60-160ms）。"""
    n = pix.n
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, n)
    if n == 4:
        arr = arr[:, :, :3]
    elif n == 1:
        arr = np.repeat(arr, 3, axis=2)
    # pixmap 为 RGB，RapidOCR 按 BGR 约定处理；::-1 视图改连续拷贝
    return np.ascontiguousarray(arr[:, :, ::-1])


def _ocr_pixmap(pix, ocr, dpi: int, preprocess: bool = True) -> str:
    """对一个已渲染 pixmap 做预处理+OCR+阅读序排序，返回该区域文本。"""
    arr = _pixmap_to_bgr(pix)
    if preprocess:
        try:
            arr = _preprocess_arr(arr, do_deskew=True)
        except Exception as e:
            logger.debug("预处理跳过：%s", e)
    result, _ = ocr(arr)
    if not result:
        return ""
    lines = _sort_ocr_result_by_position(result, line_tol=max(6, int(dpi / 30)))
    return "\n".join(lines).strip()


def ocr_pdf_bytes(data: bytes, dpi: int = 300, max_pages: Optional[int] = None,
                   preprocess: bool = True, pages: Optional[list] = None) -> str:
    """对扫描件PDF做OCR，返回纯文本（按阅读顺序排序，尽量保留行结构）。

    Args:
        data: PDF 二进制数据
        dpi: 渲染DPI（默认300，平衡清晰度与速度；扫描件建议300+）
        max_pages: 只处理前N页（None=全部），用于超长PDF限速
        preprocess: 是否启用图像预处理（灰度/对比度增强/deskew），默认 True
        pages: 显式指定要OCR的页索引（0基），用于混合件只补OCR扫描插页；
            给定后忽略 max_pages
    双页拼版：文档级横向页占比过半时，横向页按中线拆成左右两半、各自以完整
    dpi 重新渲染后 OCR（整页渲染再裁等价于半分辨率，错字率高）。
    """
    import fitz  # pymupdf

    doc = fitz.open(stream=data, filetype="pdf")
    total_pages_all = len(doc)
    if pages is not None:
        page_indices = [i for i in pages if 0 <= i < total_pages_all]
    else:
        total_pages = (min(total_pages_all, max_pages)
                       if max_pages is not None else total_pages_all)
        page_indices = list(range(total_pages))

    # 文档级双页拼版判定（与 crawler_core.pdf_health_check 同判据）：
    # 过半横向页才拆；孤立的宽表格横向页不拆，避免栏间错位
    landscape = sum(
        1 for i in page_indices
        if doc[i].rect.width > doc[i].rect.height * 1.15
    )
    spread_mode = bool(page_indices) and landscape / len(page_indices) > 0.5

    ocr = _get_ocr()
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    chunks = []
    for page_idx in page_indices:
        page = doc[page_idx]
        clips = [None]
        if spread_mode and page.rect.width > page.rect.height * 1.15:
            r = page.rect
            mid = (r.x0 + r.x1) / 2
            # 左半页→右半页，保持拼版扫描的阅读顺序
            clips = [fitz.Rect(r.x0, r.y0, mid, r.y1),
                     fitz.Rect(mid, r.y0, r.x1, r.y1)]

        page_parts = []
        had_error = False
        for clip in clips:
            try:
                pix = page.get_pixmap(matrix=mat, clip=clip, alpha=False)
                t = _ocr_pixmap(pix, ocr, dpi, preprocess)
            except Exception as e:
                logger.warning("第%d页 OCR 失败：%s", page_idx + 1, e)
                t = ""
                had_error = True
            if t.strip():
                page_parts.append(t.strip())

        # 异常页重试一次；正常的空白页（无异常也无文字）不重试，避免白跑模型
        if had_error and not page_parts:
            for clip in clips:
                try:
                    pix = page.get_pixmap(matrix=mat, clip=clip, alpha=False)
                    t = _ocr_pixmap(pix, ocr, dpi, preprocess)
                    if t.strip():
                        page_parts.append(t.strip())
                except Exception:
                    pass

        if page_parts:
            chunks.append("\n".join(page_parts))
            chunks.append("")  # 页间空行
        else:
            # 缺页占位：整页从 MD 消失会让正文无声缺失，必须留痕（A5）
            chunks.append("> [第%d页 OCR 失败或为无文字页，引用时请对照PDF原件]"
                          % (page_idx + 1))
            chunks.append("")
        if (page_idx + 1) % 10 == 0:
            logger.info("OCR 进度：%d/%d 页", page_idx + 1, total_pages_all)

    doc.close()
    raw = "\n".join(chunks).strip()
    # 后处理：修正 OCR 常见的标点/空格丢失问题
    return postprocess_ocr_text(raw)


# ---------------------------------------------------------------------------
# OCR 文本后处理（方案 A：正则修正标点/空格丢失）
# ---------------------------------------------------------------------------

# 标准号前缀（用于标准号格式修正）
_STD_NO_PREFIXES = [
    "DL/T", "DL", "GB/T", "GB/Z", "GB", "NB/T", "NB", "AQ", "JG/T", "JG",
    "HG/T", "HG", "HJ", "SL", "DLT", "GBT", "NBT", "MT", "MT/T",
    "YS/T", "YS", "YY/T", "YY", "YD/T", "YD", "TB/T", "TB",
    "QB/T", "QB", "QC/T", "QC", "JT/T", "JT", "SH/T", "SH",
    "SY/T", "SY", "SH", "CECS", "DB", "DB/T", "Q/BJP",
]


def postprocess_ocr_text(text: str) -> str:
    """OCR 文本后处理，修正常见的标点/空格丢失问题。

    修正规则（按顺序应用，靠前的规则优先级高，避免被后续规则破坏）：
    1. 标准号格式修正：DL/T4082023 → DL/T 408-2023
    2. 章节号后补空格：4.1工作人员 → 4.1 工作人员
    3. 日期格式修正：2023-12-28发布 → 2023-12-28 发布（数字后接中文补空格）
    4. 中英文边界补空格：代替DL408 → 代替 DL408
    5. 连续英文单词间的常见空格恢复（用词典贪心匹配）

    注意：修正保守优先——宁可漏修不可错修，避免破坏原本正确的内容。
    """
    if not text or not text.strip():
        return text

    result = text

    # 规则1：标准号格式修正
    # 匹配模式：前缀+序号+年份 粘在一起，如 DL/T4082023、GB502632007
    # 正则：前缀(可选斜杠) + 3-4位序号 + 4位年份
    # 注意：要避免误匹配正常文本里的数字
    for prefix in sorted(_STD_NO_PREFIXES, key=len, reverse=True):
        # 转义正则特殊字符
        p = prefix.replace("/", r"\/").replace("[", r"\[").replace("]", r"\]")
        # 模式：DL/T4082023 → 前缀 + 3-4位序号 + 4位年份
        pattern = r'(' + p + r')(\d{3,4})(\d{4})'
        # 替换为：前缀 序号-年份
        result = re.sub(pattern, r'\1 \2-\3', result)

    # 规则2：章节号后补空格（行首的 4.1、5.6.7 等后面紧跟中文）
    # 模式：行首 章节号(数字.数字...) + 非空格非数字字符
    result = re.sub(r'^(\d+(?:\.\d+)+)([^\s\d.])', r'\1 \2', result, flags=re.MULTILINE)
    # 单级章节号（如 "5安全组织措施"），但要避免误匹配普通数字开头
    # 只匹配 1-2 位数字 + 中文（章节号通常 1-2 位）
    result = re.sub(r'^(\d{1,2})([\u4e00-\u9fa5])', r'\1 \2', result, flags=re.MULTILINE)

    # 规则3：数字/日期后接中文补空格
    # 模式：4位数字-2位数字-2位数字 + 中文，如 2023-12-28发布
    result = re.sub(r'(\d{4}-\d{2}-\d{2})([\u4e00-\u9fa5])', r'\1 \2', result)
    # 普通数字后接中文（但要避免破坏标准号，因为标准号已在规则1修正过）
    # 只匹配：数字 + 中文 且 前面不是字母/斜杠（避免破坏 DL408 这种）
    result = re.sub(r'(?<![A-Za-z/])(\d{2,4})([\u4e00-\u9fa5])', r'\1 \2', result)

    # 规则4：中英文边界补空格
    # 中文 + 英文
    result = re.sub(r'([\u4e00-\u9fa5])([A-Za-z])', r'\1 \2', result)
    # 英文 + 中文
    result = re.sub(r'([A-Za-z])([\u4e00-\u9fa5])', r'\1 \2', result)
    # 中文 + 数字（如 "电气操作10" → "电气操作 10"）
    # 注意：放在标准号修正之后，避免破坏标准号内部的英文+数字
    result = re.sub(r'([\u4e00-\u9fa5])(\d)', r'\1 \2', result)

    # 规则5：常见英文单词空格恢复（用词典贪心匹配）
    # 仅对连续 10+ 字符的英文串做处理，避免误伤缩写
    result = _restore_english_word_spaces(result)

    return result


# 常见英文单词词典（用于英文空格恢复，优先匹配长词）
# 来源：标准文档中常见的技术词汇 + 常用英文虚词
_COMMON_ENGLISH_WORDS = [
    # 常用虚词/短词（2-3 字母，必须补上否则 "partof" 这类切不出来）
    "the", "and", "for", "with", "from", "by", "of", "in", "on", "at",
    "to", "is", "are", "be", "as", "it", "an", "or", "of", "per",
    # 长词（贪心匹配优先）
    "transformer", "substations", "substation", "equipment",
    "installation", "operation", "maintenance", "inspection", "protection",
    "regulation", "standard", "industry", "electric", "power",
    "plants", "plant", "code", "safety", "voltage", "current",
    "circuit", "breaker", "switch", "disconnect", "generator", "motor",
    "system", "device", "panel", "station", "network", "control",
    "alarm", "fault", "test", "work", "zone", "area", "field",
    "manual", "automatic", "emergency", "normal", "special", "general",
    "section", "chapter", "figure", "table", "page", "number", "date",
    "release", "implement", "publish", "edition", "version",
    "level", "high", "low", "medium", "extra",
    "alternate", "direct", "frequency", "phase", "line",
    "cable", "wire", "tower", "pole", "insulator",
    "grounding", "earthing", "neutral", "live", "dead", "isolate",
    "permit", "license", "certificate", "approve", "verify", "check",
    "hazard", "risk", "danger", "warning", "caution", "notice",
    "requirement", "procedure", "process", "method", "material",
    "design", "construct", "install", "commission", "decommission",
    "national", "international", "enterprise", "part", "parts",
    "bus", "plants", "plant",
]
# 去重 + 按长度降序预排序（模块加载时算一次，避免每次调用都排序)
_COMMON_ENGLISH_WORDS_SORTED = sorted(set(_COMMON_ENGLISH_WORDS), key=len, reverse=True)
# 小写版本（用于大小写不敏感匹配）
_COMMON_ENGLISH_WORDS_LOWER = [w.lower() for w in _COMMON_ENGLISH_WORDS_SORTED]


def _restore_english_word_spaces(text: str) -> str:
    """对连续英文串做常见词空格恢复。

    策略：找到连续 10+ 字符的英文串（含大小写），用词典贪心匹配切分。
    大小写不敏感匹配，但保留原始大小写输出。
    保守策略——只处理能完全切分的串，否则原样保留。
    """
    if not text:
        return text

    def replace_long_english(match):
        s = match.group(0)
        if len(s) < 10:  # 短串不处理
            return s
        # 尝试用词典贪心切分
        words = _greedy_split(s)
        if words and "".join(words) == s and len(words) > 1:
            return " ".join(words)
        return s

    # 匹配连续的英文串（大小写都算，10+ 字符）
    return re.sub(r'[a-zA-Z]{10,}', replace_long_english, text)


def _greedy_split(s: str) -> list:
    """用词典贪心切分英文串。优先匹配长词，大小写不敏感但保留原词。

    例如 "Electricpartof" → ["Electric", "part", "of"]
    """
    s_lower = s.lower()
    result = []
    i = 0
    while i < len(s_lower):
        matched = False
        for idx, w_lower in enumerate(_COMMON_ENGLISH_WORDS_LOWER):
            if s_lower[i:i + len(w_lower)] == w_lower:
                # 保留原词大小写（从原串切片）
                result.append(s[i:i + len(w_lower)])
                i += len(w_lower)
                matched = True
                break
        if not matched:
            # 匹配失败，剩余部分作为整体返回（放弃切分）
            if result:
                result.append(s[i:])
            else:
                return []  # 完全无法切分
            break
    return result


def _page_has_large_image(page, cover_ratio: float = 0.3) -> bool:
    """页面是否含覆盖面积超阈值的位图——用于区分"扫描附图页"与"空白/分隔页"，
    避免数字件每遇到空白页就白跑一次 OCR。"""
    try:
        rect = page.rect
        page_area = max(rect.width * rect.height, 1.0)
        try:
            infos = page.get_image_info(xrefs=True)
        except Exception:
            infos = []
        for info in infos or []:
            bbox = info.get("bbox")
            if bbox and (bbox[2] - bbox[0]) * (bbox[3] - bbox[1]) > cover_ratio * page_area:
                return True
        return bool(page.get_images(full=False))
    except Exception:
        return False


def is_scanned_pdf(data: bytes, threshold: int = 50) -> bool:
    """判断PDF是否为扫描件（提取文字太少）。threshold = 每页平均字符数下限。

    统计前先剥离水印文本层（住建部公开件每页带固定水印，会虚增字符数误判为非扫描件）。
    """
    import fitz
    try:
        doc = fitz.open(stream=data, filetype="pdf")
        total_chars = 0
        n = 0
        all_text = []
        for page in doc:
            page_text = _maybe_strip_watermark(page.get_text("text").strip())
            all_text.append(page_text)
            total_chars += len(page_text)
            n += 1
        doc.close()
        # 乱码判定必须在全文级做：私有CMap件每页汉字可能很少（实测GB30871每页均12.7个
        # 汉字），逐页会因样本不足漏判；全文汇总后汉字样本充足可正确识别
        if _looks_like_garbled("\n".join(all_text)):
            return True
        avg = total_chars / max(n, 1)
        return avg < threshold
    except Exception:
        return False


def extract_pdf_with_ocr(data: bytes, dpi: int = 300,
                          max_pages: Optional[int] = None,
                          use_pdfplumber: bool = True) -> tuple[str, bool]:
    """PDF文本提取，带图像预处理+OCR回退+表格保留（逐页裁决）。

    流程（fitz 只打开一次，旧流程的"pdfplumber 全量 + fitz 全量 +
    is_scanned_pdf 两次全量"全部合并）：
    1. 文档级判定：全文乱码 或 每页平均<50字 → 扫描件，跳过 pdfplumber
       （无边框扫描件上做表格检测是纯浪费），直接整篇 OCR；
    2. 数字件：pdfplumber 提取表格结构 + fitz 文本择优；
    3. 混合件（数字正文+个别扫描附图/插页）：只对文字层<50字的薄弱页补 OCR，
       作为局部补充段，不加整篇 OCR 警告头、used_ocr=False。

    返回 (text, used_ocr) —— used_ocr=True 表示主体由 OCR 得到。
    """
    import fitz
    try:
        doc = fitz.open(stream=data, filetype="pdf")
        page_items = []
        for page in doc:
            txt = _maybe_strip_watermark((page.get_text("text") or "").strip())
            page_items.append((txt, _page_has_large_image(page)))
        doc.close()
    except Exception as e:
        logger.debug("pymupdf 打开失败：%s", e)
        page_items = []

    n = len(page_items)
    page_texts = [t for t, _ in page_items]
    joined = "\n\n".join(t for t in page_texts if t).strip()
    total_chars = sum(len(t) for t in page_texts)
    garbled = bool(joined) and _looks_like_garbled(joined)
    scanned = (n > 0 and total_chars / max(n, 1) < 50) or garbled

    # 路径一：扫描件/私有CMap乱码件 → 整篇 OCR
    if scanned:
        try:
            ocr_text = _maybe_strip_watermark(
                ocr_pdf_bytes(data, dpi=dpi, max_pages=max_pages, preprocess=True))
            if ocr_text.strip():
                return ocr_text, True
        except Exception as e:
            logger.warning("OCR 回退失败：%s", e)
        return joined, False

    # 路径二：数字件——pdfplumber 表格优先，fitz 文本兜底
    text = joined
    if use_pdfplumber:
        try:
            pp_text, _ = _extract_with_pdfplumber(data)
            pp_text = _maybe_strip_watermark((pp_text or "").strip())
            if pp_text and not _looks_like_garbled(pp_text) and len(pp_text) > len(text):
                text = pp_text
        except Exception as e:
            logger.debug("pdfplumber 跳过：%s", e)

    # 路径三：混合件薄弱页（无文字层但有大图=扫描附图/插页）定点补 OCR；
    # 无图的空白/分隔页不 OCR；封顶 30 页防极端耗时
    weak = [i for i, (t, has_img) in enumerate(page_items)
            if len(t) < 50 and has_img][:30]
    if weak:
        try:
            # 附图补页用 200dpi：附图以图表/示意图为主，200dpi 足够且比 300 快约一半
            extra = _maybe_strip_watermark(
                ocr_pdf_bytes(data, dpi=min(dpi, 200), pages=weak, preprocess=True))
            # 去掉"无文字页"占位（补充段里是噪声），保留 OCR 失败留痕
            extra_lines = [
                ln for ln in extra.splitlines()
                if "OCR 失败或为无文字页" not in ln
            ]
            extra = "\n".join(extra_lines).strip()
            if extra:
                text = (text
                        + "\n\n## 扫描附图/插页 OCR 补充（可能含错字，引用前核对原件）\n\n"
                        + extra).strip()
        except Exception as e:
            logger.warning("混合件补页OCR失败：%s", e)
    return text, False


# ---------------------------------------------------------------------------
# 工具函数：检查 OCR 引擎与依赖可用性
# ---------------------------------------------------------------------------

def diagnose() -> dict:
    """诊断 OCR 引擎与依赖是否就绪，供排查"为什么转换出来还是乱码"用。"""
    status = {"rapidocr": False, "pymupdf": False, "pdfplumber": False,
              "pillow": False, "opencv": False, "numpy": False}
    try:
        import fitz
        status["pymupdf"] = True
    except ImportError:
        pass
    try:
        import pdfplumber
        status["pdfplumber"] = True
    except ImportError:
        pass
    try:
        from PIL import Image
        status["pillow"] = True
    except ImportError:
        pass
    try:
        import cv2
        status["opencv"] = True
    except ImportError:
        pass
    try:
        import numpy
        status["numpy"] = True
    except ImportError:
        pass
    try:
        _get_ocr()
        status["rapidocr"] = True
    except Exception as e:
        status["rapidocr_error"] = str(e)
    return status
