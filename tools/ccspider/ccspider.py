# -*- coding: utf-8 -*-
"""
标准文档处理工具（行业标准批量下载转 Markdown / 本地目录文件转 Markdown）

输出位置：
    - 行业下载：程序目录/DownloadMD/{行业中文名}/*.md
    - 本地转换：程序目录/TOMD/{日期YYYYMMDD}/（当天的所有 md 集中存放，目录存在则直接使用）

说明：
    - 单线程，每次请求后随机延迟 1~3 秒，失败自动重试 2 次，无死循环。
    - 下载接口 /portal/download/{pk} 的响应头 content-disposition / content-type
      提供真实文件名与格式，脚本据此调用对应转换器，最终统一输出 .md。

依赖：
    pip install curl_cffi python-docx pypdf openpyxl xlrd markdownify

打包为无界面 exe：
    pip install pyinstaller
    pyinstaller -F -n ccspider ccspider.py --hidden-import=_cffi_backend --collect-all curl_cffi
    （curl_cffi 自带 libcurl 二进制，--collect-all / --hidden-import 不能少，
     否则打包出的 exe 启动即报找不到模块或 DLL）
    （程序目录以 exe 所在目录为准）
"""

import io
import os
import re
import sys
import time
import random
import datetime
import urllib.parse

from curl_cffi import requests

BASE = "https://hbba.sacinfo.org.cn"

# 浏览器指纹伪装目标（curl_cffi 专有；设成 None 即关闭伪装）
IMPERSONATE = "chrome"

# 行业代码表：代码 -> 名称（从 stdList 页面汇总）
INDUSTRY_MAP = {
    "AQ": "安全生产", "BB": "包装", "CB": "船舶", "CH": "测绘", "CJ": "城镇建设",
    "CY": "新闻出版", "DA": "档案", "DB": "地震", "DL": "电力", "DY": "电影",
    "DZ": "地质矿产", "EJ": "核工业", "FZ": "纺织", "GA": "公共安全", "GC": "国家物资储备",
    "GF": "国防工业", "GH": "供销合作", "GM": "国密", "GY": "广播电视和网络视听",
    "HB": "航空", "HG": "化工", "HJ": "环境保护", "HS": "海关", "HY": "海洋",
    "JB": "机械", "JC": "建材", "JG": "建筑工程", "JR": "金融", "JS": "机关事务",
    "JT": "交通", "JY": "教育", "KA": "矿山安全", "LB": "旅游", "LD": "劳动和劳动安全",
    "LS": "粮食", "LY": "林业", "MH": "民用航空", "MR": "市场监管", "MT": "煤炭",
    "MZ": "民政", "NB": "能源", "NY": "农业", "QB": "轻工", "QC": "汽车",
    "QJ": "航天", "QX": "气象", "RB": "认证认可", "RF": "人民防空", "SB": "国内贸易",
    "SC": "水产", "SF": "司法", "SH": "石油化工", "SJ": "电子", "SL": "水利",
    "SN": "出入境检验检疫", "SW": "税务", "SY": "石油天然气", "TB": "铁路", "TD": "土地管理",
    "TY": "体育", "WB": "物资管理", "WH": "文化", "WJ": "兵工民品", "WM": "外经贸",
    "WS": "卫生", "WW": "文物保护", "XB": "稀土", "XF": "消防救援", "YB": "黑色冶金",
    "YC": "烟草", "YD": "通信", "YJ": "减灾救灾与综合性应急管理", "YS": "有色金属",
    "YY": "医药", "YZ": "邮政", "ZY": "中医药",
}
NAME2CODE = {v: k for k, v in INDUSTRY_MAP.items()}

# 注意：不要手动设置 User-Agent。impersonate 会自动带上与 TLS 指纹配套的完整浏览器头部，
# 手动覆盖会造成 UA 版本与指纹版本不一致，反而更容易被识别。
HEADERS = {
    "X-Requested-With": "XMLHttpRequest",
    "Referer": BASE + "/stdList",
}


def program_dir():
    """返回程序所在目录（支持 PyInstaller 打包后的 exe）。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ---------------- 通用工具 ----------------

def sleep_rand():
    time.sleep(random.uniform(1, 3))


def normalize_md_text(raw):
    """折叠多余空行与行尾空白，保持正文不乱。"""
    lines = [ln.rstrip() for ln in raw.splitlines()]
    out, blank = [], False
    for ln in lines:
        if not ln.strip():
            if out and not blank:
                out.append("")
            blank = True
            continue
        blank = False
        out.append(ln.strip())
    while out and out[-1] == "":
        out.pop()
    return "\n".join(out)


def safe_filename(s):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", s)
    return s.strip(" .") or "unnamed"


# ---------------- PDF / Word / Excel / HTML 转换 ----------------

def docx_to_md(data):
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    document = docx.Document(io.BytesIO(data))
    out = []

    def iter_block(parent):
        for child in parent.element.body.iterchildren():
            if child.tag.endswith("}p"):
                yield Paragraph(child, parent)
            elif child.tag.endswith("}tbl"):
                yield Table(child, parent)

    for block in iter_block(document):
        if isinstance(block, Paragraph):
            if not block.text.strip():
                continue
            name = (block.style.name if block.style else "") or ""
            txt = block.text.strip()
            if "Heading 1" in name:
                out.append("# " + txt)
            elif "Heading 2" in name:
                out.append("## " + txt)
            elif "Heading 3" in name:
                out.append("### " + txt)
            else:
                out.append(txt)
        elif isinstance(block, Table):
            md_rows = []
            for row in block.rows:
                cells = [c.text.strip().replace("|", "\\|") for c in row.cells]
                md_rows.append("| " + " | ".join(cells) + " |")
            if md_rows:
                n = len(md_rows[0].split("|")) - 2
                out.append(md_rows[0])
                out.append("|" + " --- |" * n)
                out.extend(md_rows[1:])
                out.append("")
    return normalize_md_text("\n".join(out))


def doc_to_md(data):
    """老版 .doc 二进制；尽力提取可见文本，失败时给出提示。"""
    try:
        text = data.decode("utf-8", errors="ignore")
    except Exception:
        text = ""
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    printable = re.findall(r"[\u4e00-\u9fa5a-zA-Z0-9\s，。；：、（）()《》\"\'\-.%/]", text)
    cleaned = "".join(printable)
    cleaned = re.sub(r"\s{2,}", "\n", cleaned)
    if len(cleaned.strip()) < 20:
        return ("> 该文件为旧版 .doc 二进制文档，当前环境缺少转换工具，"
                "无法可靠提取文本，请使用 Word 打开原文。")
    return normalize_md_text(cleaned)


def pdf_to_md(data):
    import pypdf
    reader = pypdf.PdfReader(io.BytesIO(data))
    chunks = []
    for page in reader.pages:
        try:
            txt = page.extract_text() or ""
        except Exception:
            txt = ""
        if txt.strip():
            chunks.append(txt.strip())
    return normalize_md_text("\n\n".join(chunks))


def xlsx_to_md(data):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    parts = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        parts.append("## " + ws.title)
        ncols = max(len(r) for r in rows)
        lines = []
        for r in rows:
            cells = ["" if v is None else str(v).replace("|", "\\|").replace("\n", " ")
                     for v in r]
            cells += [""] * (ncols - len(cells))
            lines.append("| " + " | ".join(cells) + " |")
        lines.insert(1, "|" + " --- |" * ncols)
        parts.append("\n".join(lines))
    return normalize_md_text("\n\n".join(parts))


def xls_to_md(data):
    import xlrd
    wb = xlrd.open_workbook(file_contents=data)
    parts = []
    for ws in wb.sheets():
        if ws.nrows == 0:
            continue
        parts.append("## " + ws.name)
        ncols = ws.ncols
        lines = []
        for ri in range(ws.nrows):
            cells = []
            for ci in range(ncols):
                v = ws.cell_value(ri, ci)
                if isinstance(v, float) and v == int(v):
                    v = int(v)
                cells.append("" if v is None else str(v).replace("|", "\\|").replace("\n", " "))
            lines.append("| " + " | ".join(cells) + " |")
        lines.insert(1, "|" + " --- |" * ncols)
        parts.append("\n".join(lines))
    return normalize_md_text("\n\n".join(parts))


def html_to_md(data):
    from markdownify import markdownify as mdify
    text = data.decode("utf-8", errors="ignore")
    return normalize_md_text(mdify(text, heading_style="ATX"))


def raw_to_md(data):
    text = data.decode("utf-8", errors="ignore")
    return normalize_md_text(text)


def to_markdown(data, fname, content_type=""):
    """按扩展名 / content-type 选择转换器，保证输出结构化的 Markdown。"""
    ext = (os.path.splitext(fname)[1] or "").lower()
    ctype = (content_type or "").lower()
    if "wordprocessingml" in ctype:
        return docx_to_md(data)
    if ext == ".docx":
        return docx_to_md(data)
    if ext == ".doc":
        return doc_to_md(data)
    if ext in (".xlsx", ".xlsm") or "spreadsheetml" in ctype:
        return xlsx_to_md(data)
    if ext == ".xls":
        return xls_to_md(data)
    if ext in (".pdf",) or "pdf" in ctype:
        return pdf_to_md(data)
    if ext in (".html", ".htm") or "text/html" in ctype:
        return html_to_md(data)
    if ext in (".md", ".markdown", ".txt", ".text"):
        return raw_to_md(data)
    return raw_to_md(data)


# ---------------- 行业标准下载 ----------------

def resolve_industry(inp):
    """支持输入中文名或行业代码，返回 (中文名, 代码)。未知输入原样使用。"""
    inp = inp.strip()
    if inp in INDUSTRY_MAP:
        return INDUSTRY_MAP[inp], inp
    if inp in NAME2CODE:
        return inp, NAME2CODE[inp]
    return inp, NAME2CODE.get(inp, inp)


def query_all_records(sess, industry_name):
    """分页查询某行业全部标准记录，返回记录列表。有页数上限防死循环。"""
    records = []
    current = 1
    size = 100
    max_pages = 10000
    while current <= max_pages:
        data = {
            "current": str(current),
            "size": str(size),
            "key": "",
            "ministry": "",
            "industry": industry_name,
            "pubdate": "",
            "date": "",
        }
        resp = sess.post(BASE + "/stdQueryList", data=data, timeout=60)
        resp.raise_for_status()
        body = resp.json()
        recs = body.get("records") or []
        records.extend(recs)
        pages = int(body.get("pages") or 0)
        if current >= pages or not recs:
            break
        current += 1
        sleep_rand()
    return records


def get_download(sess, pk, retries=2):
    """下载文件，返回 (bytes, filename, content_type)。失败返回 (None, None, None)。"""
    url = BASE + "/portal/download/" + pk
    for attempt in range(retries + 1):
        try:
            resp = sess.get(url, timeout=120)
            resp.raise_for_status()
            cd = resp.headers.get("content-disposition", "")
            fname = ""
            if "filename" in cd:
                part = cd.split("filename=", 1)[1].split(";")[0].strip().strip('"')
                fname = urllib.parse.unquote(part).replace("+", " ")
            return resp.content, fname, resp.headers.get("content-type", "")
        except Exception as exc:
            if attempt >= retries:
                print("    下载失败: %s" % exc)
                return None, None, None
            sleep_rand()
    return None, None, None


def process_industry(sess, industry_inp, out_root):
    name, code = resolve_industry(industry_inp)
    print("==> 行业: %s (代码: %s)" % (name, code))
    try:
        records = query_all_records(sess, name)
    except Exception as exc:
        print("   查询列表失败: %s" % exc)
        return
    if not records:
        print("    无记录")
        return
    print("    共 %d 条记录" % len(records))

    folder = os.path.join(out_root, name)
    os.makedirs(folder, exist_ok=True)

    ok = skipped = failed = empty = 0
    for idx, rec in enumerate(records, 1):
        std_no = (rec.get("code") or "unknown").strip()
        title = (rec.get("chName") or "").strip()
        pk = rec.get("pk") or ""
        if not pk:
            failed += 1
            continue

        md_file = os.path.join(folder, safe_filename("%s %s.md" % (std_no, title)))
        if os.path.exists(md_file):
            skipped += 1
            print("    [%d/%d] 已存在，跳过: %s" % (idx, len(records), std_no))
            continue

        print("    [%d/%d] 下载中: %s %s" % (idx, len(records), std_no, title))
        data, fname, ctype = get_download(sess, pk)
        sleep_rand()
        if data is None:
            failed += 1
            print("        下载失败: %s" % std_no)
            continue
        if len(data) == 0:
            empty += 1
            note = ("# %s %s\n\n> 该标准未提供可下载的电子文本（下载接口返回空文件）。\n\n"
                    "- 标准号: %s\n- 标准状态: %s\n"
                    % (std_no, title, std_no, rec.get("status", "")))
            with open(md_file, "w", encoding="utf-8") as f:
                f.write(note)
            print("        下载为空文件，已写入占位说明")
            continue

        if fname:
            print("    (文件: %s)" % fname)
        try:
            md_text = to_markdown(data, fname or "", ctype)
            header = "# %s %s\n\n" % (std_no, title)
            if not md_text.strip():
                md_text = "> 未提取到可用的文本内容。"
            with open(md_file, "w", encoding="utf-8") as f:
                f.write(header + md_text)
            ok += 1
        except Exception as exc:
            failed += 1
            print("        转换失败: %s" % exc)

    print("    本行业完成：成功 %d，跳过 %d，空文件 %d，失败 %d"
          % (ok, skipped, empty, failed))


# ---------------- 本地目录转换 ----------------

SCAN_EXTS = {".doc", ".docx", ".pdf", ".xls", ".xlsx", ".xlsm", ".html", ".htm", ".txt", ".md", ".markdown"}


def collect_files(root):
    """递归收集待转换文件，避免死循环（探测环路目录）。"""
    found = []
    seen = set()
    for base, dirs, files in os.walk(root):
        real = os.path.realpath(base)
        if real in seen:
            dirs[:] = []
            continue
        seen.add(real)
        dirs[:] = [d for d in dirs if os.path.realpath(os.path.join(base, d)) not in seen]
        for fn in files:
            if os.path.splitext(fn)[1].lower() in SCAN_EXTS:
                found.append(os.path.join(base, fn))
    return found


def process_local_dir(src_dir, program_root):
    if not os.path.isdir(src_dir):
        print("目录不存在: %s" % src_dir)
        return
    today = datetime.datetime.now().strftime("%Y%m%d")
    out_root = os.path.join(program_root, "TOMD", today)
    os.makedirs(out_root, exist_ok=True)
    print("输入目录: %s" % os.path.abspath(src_dir))
    print("输出目录: %s" % os.path.abspath(out_root))

    files = collect_files(src_dir)
    if not files:
        print("未找到 .docx/.pdf/.xls/.xlsx 等文档文件")
        return
    print("共 %d 个文件待转换" % len(files))

    ok = failed = 0
    for idx, fp in enumerate(files, 1):
        rel = os.path.relpath(fp, src_dir)
        name = os.path.basename(rel)
        print("    [%d/%d] 转换: %s" % (idx, len(files), rel))
        try:
            with open(fp, "rb") as f:
                data = f.read()
            if not data:
                print("        空文件，跳过")
                continue
            md_text = to_markdown(data, name)
            ext = os.path.splitext(name)[1].lower()
            if ext in (".md", ".markdown"):
                md_text = raw_to_md(data)
            if not md_text.strip():
                md_text = "> 未提取到可用的文本内容。"
            rel_md = os.path.splitext(rel)[0] + ".md"
            out_file = os.path.join(out_root, rel_md)
            os.makedirs(os.path.dirname(out_file), exist_ok=True)
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(md_text)
            ok += 1
            print("        -> %s" % os.path.relpath(out_file, out_root))
        except Exception as exc:
            failed += 1
            print("        转换失败: %s" % exc)

    print("本次转换完成：成功 %d，失败 %d" % (ok, failed))


# ---------------- 入口 ----------------

def runduibi():
    """交互式比较两个文件内容是否完全一致。"""
    from pathlib import Path
    import filecmp
    a = input("第一个文件路径: ").strip()
    b = input("第二个文件路径: ").strip()
    path_a = Path(a)
    path_b = Path(b)
    if not path_a.exists() or not path_b.exists():
        print("⚠️ 错误：输入的文件路径不存在，请检查！")
        return
    is_same = filecmp.cmp(str(path_a), str(path_b), shallow=False)
    if is_same:
        print("文件内容相同")
    else:
        print("文件内容不相同")


def print_menu():
    print("")
    print("=" * 30 + " 标准文档处理工具 " + "=" * 30)
    print("  1) 下载行业标准并转 md")
    print("  2) 本地目录文件转 md")
    print("  3) 对比两个文件")
    print("  0) 退出")
    print("=" * 77)


def menu_download(program_root):
    """菜单项 1：下载行业标准转 md。"""
    raw = input("请输入行业名/行业代码（多个用 = 分隔，例如 外经贸=黑色冶金）：\n").strip()
    if not raw:
        print("未输入行业，返回菜单。")
        return
    industries = [a.strip() for a in raw.split("=") if a.strip()]
    if not industries:
        print("未识别到有效行业，返回菜单。")
        return

    out_root = os.path.join(program_root, "DownloadMD")
    os.makedirs(out_root, exist_ok=True)
    print("输出目录: %s" % out_root)

    sess = requests.Session(impersonate=IMPERSONATE)
    sess.headers.update(HEADERS)
    for ind in industries:
        try:
            process_industry(sess, ind, out_root)
        except Exception as exc:
            print("处理行业 %s 时出错: %s" % (ind, exc))
        print("")


def menu_local_dir(program_root):
    """菜单项 2：本地目录文件转 md。"""
    src = input("请输入目录路径（例如 docs）：\n").strip()
    if not src:
        print("未输入目录，返回菜单。")
        return
    if not os.path.isdir(src):
        print("目录不存在: %s" % src)
        return
    process_local_dir(src, program_root)


def menu_compare():
    """菜单项 3：对比两个文件内容是否一致。"""
    runduibi()


def main():
    program_root = program_dir()
    print("程序目录: %s" % program_root)
    while True:
        print_menu()
        choice = input("请选择 [0-3]: ").strip()
        try:
            if choice == "0":
                print("已退出。")
                break
            elif choice == "1":
                menu_download(program_root)
            elif choice == "2":
                menu_local_dir(program_root)
            elif choice == "3":
                menu_compare()
            else:
                print("无效选项，请重新选择。")
        except KeyboardInterrupt:
            print("\n已中断当前操作，返回菜单。")
        except Exception as exc:
            print("发生异常: %s" % exc)


if __name__ == "__main__":
    main()