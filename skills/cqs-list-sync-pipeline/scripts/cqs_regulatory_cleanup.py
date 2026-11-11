import os
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cqs_regulatory_cleanup.py — 安全生产基建全周期标准化实施参考清单（cqs）通用治理工具

沉淀自 2026-10-06 (1006) 版全量治理（监管列企业制度清零/依据规范化/全簿美化/三表链接重建）。

列结构（务必先核实，2026-10-07 实测）：
  19个8列主表： A类目/分目 B主要文件 C参考依据 D依据来源 E文件来源 F文件管理 G监管要求来源 H备注（r1标题r2表头，数据r3起，冻结A3）
  17设备表(9列)：A检查内容 B检查方法 C执行周期 D异常情况 E处理措施 F参考依据 G依据来源 H监管要求来源 I备注
  s20打印表：同8列，数据从 r3 起（无空行）；不含9列设备表
  s21索引表：r1标题(合并A1:H1) r3表头 r4起数据(冻结A4)；A序号 B来源清单 C类目 D分目 E依据(挂链接) F来源
             G文件状态（在库/在线全文/官方公告/在线书目/待补） H权威本地路径（只放本地相对路径，不放URL）；
             数据区后为"待补文件清单"标题合并行+明细行
  s22 已于 2026-10-07 删除（内容已融入主表），现共 22 sheets。

功能：
  --book xlsx --scan                 只扫描报告（监管列残留/依据格式/死链/错链/http异常），不写盘
  --book xlsx --audit-mislinks       错链专项审计：依据文本标准号 vs 链接文件名标准号（只读，异常则退出码1）
  --book xlsx --clean-gov            监管列清洗（删企业片段+国标行标归口规范化）
  --book xlsx --normalize-ent-year   企业标准编号补年份（YEAR_MAP，只补已核实编号，禁止编造）
  --book xlsx --beautify             全簿美化（标题/表头/边框/字体/行高/冻结/列宽，保留D列分类色）
  --book xlsx --rebuild-links        链接治理：s21按H重建→主表C/F按s21权威路径回填→死链检测
  --book xlsx --rebuild-print        s20打印表整表重建（19个8列主表非空行镜像）
  可串联多项，按 scan→clean→year→links→print→beautify 顺序执行。

2026-10-07 补链/换版轮新增工具函数（import 复用）：
  delete_row_migrated(ws, delrow, ncol)  删行并手动迁移 hyperlink/合并区/行高（openpyxl 不迁这些）
  code_set(s)/is_mislinked(text,target)  标准号提取与错链判定（GB 与 GB/T 按同号族兼容）
  audit_mislinks(wb, kb_root)            全簿错链/死http/禁用域名审计，返回问题列表

铁律：
  1. 监管要求来源列只允许政府及以上机关表述；Q/BJP（示例电力）、QBEH（示例集团集团）、
     清洁能源/集团/公司内控制度一律只能出现在 C/F 依据列，且写全名《制度名》（Q/BJP XXXX-YYYY）。
  2. 依据标准格式：《标准名称》（标准号-年份），全角书名号+全角括号；
     企业标准年份核不到就留无年份编号（如 Q/BJP 2146），绝不编造。
  3. s21 的 H 列是唯一权威路径列；WPS 另存会丢光全簿 hyperlink 对象，以 s21!H 重建即可。
  4. openpyxl insert_rows 不迁移 hyperlink/合并区：合并区迁移必须
     “insert 前 unmerge(min_row>=锚点) → insert_rows → +n 位置 remerge”。
"""
import argparse, math, os, re, sys, zipfile
from copy import copy
import openpyxl
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

# ---------- 常量 ----------
ENT_SEG = ['集团', 'Q/BJP', 'QBEH', 'Q/BEH', '示例电力', '示例集团集团', '清洁能源', '公司制度', '企业内部']
# 已从库内文件正文核实的企业标准“编号-年份”（新增核实项只许往里加，不许臆造）
YEAR_MAP = {'2136':'2026','2137':'2026','2140':'2026','2142':'2022','2143':'2025','2144':'2025',
            '2145':'2025','2152':'2025','2158':'2026','2247':'2022','2264':'2024','2282':'2023',
            '2304':'2025','2344':'2026'}
KIND_FILL = {'国家标准':'FFDDEBF7','行业标准':'FFE2EFDA','企业标准':'FFFCE4D6','法律法规':'FFF8CBAD',
             '团体标准':'FFEDE7F6','部门规章':'FFEAD1DC','规范性文件':'FFDDEBF7','地方标准':'FFFFF2CC',
             '待补文件':'FFF2F2F2'}
NAVY, BLUE = 'FF1F4E79', 'FF2E75B6'
SPECIAL_GOV = [
 ('国家发展改革委招投标监管；集团物资管理制度；国家能源局技术改造标准',
  '国家发展改革委招投标与合同监管（招标投标法及其实施条例）；国家能源局技术改造行业标准'),
 ('国家发展改革委招投标监管；集团物资管理制度',
  '国家发展改革委招投标与合同监管（招标投标法及其实施条例）'),
 ('国家能源局技术改造标准；集团物资管理制度', '国家能源局技术改造行业标准'),
]

def main_sheets(wb):
    return [n for n in wb.sheetnames if not n.startswith(('20.', '21.', '22.'))]
def is_dev17(name): return name.startswith('17.项目运行维护（设备')
def gov_col(name): return 8 if is_dev17(name) else 7
def ref_col(name): return 6 if is_dev17(name) else 3
def src_col(name): return 7 if is_dev17(name) else 4

# ---------- 监管列清洗 ----------
def clean_gov_text(s):
    orig = s
    for a, b in SPECIAL_GOV:
        if s == a: s = b
    s = s.replace('财政部及集团基本建设财务、竣工决算审计要求', '财政部基本建设财务管理与竣工决算审计要求')
    kept = []
    for seg in [x.strip() for x in re.split(r'[；;]', s) if x.strip()]:
        if any(k in seg for k in ENT_SEG):
            continue
        seg = seg.replace('；行业协会创优评定', '').replace('行业协会创优评定', '').strip()
        if seg:
            kept.append(seg)
    s = '；'.join(kept).strip()
    s = re.sub(r'[；;]+\s*$', '', s).strip()
    if s == '国家标准':
        s = '国家市场监督管理总局、国家标准化管理委员会发布的国家标准'
    elif s == '行业标准':
        s = '国务院行业主管部门（国家能源局、住房和城乡建设部等）发布的行业标准'
    return s, (s != orig)

def add_ent_year(s):
    for num, yr in YEAR_MAP.items():
        s = re.sub(r'Q/BJP\s*' + num + r'(?!\s*[-—–0-9])', f'Q/BJP {num}-{yr}', s)
    return s

# ---------- 监管列内容审计（2026-10-08 用户铁律：监管列只写政府发文+强制性标准，不写推荐性标准） ----------
# 推荐性/行业指导性标准前缀（带 /T 或 DA/T 档案行标）；企业标准一律禁止
GOV_REC_PAT = re.compile(r"(GB/T|DL/T|NB/T|JB/T|JGJ/T|DB11/T|DB[0-9]{2}/T|CJJ/T|AQ/T|DA/T|SL/T|Q/BJP|Q/BEH|CGC/GF)")
# 允许出现在监管列的强制性标准前缀（GB/DL/JGJ/AQ 不带T、TSG特设规范、JJG计量检定规程、GA公安强标、XF消防强标）
GOV_FORBIDDEN_HINT = ('建标函',)  # 已核实的假文号/禁用文号
def gov_content_scan(wb):
    """返回监管列违规格列表 (sheet, row, 原因, 文本)。"""
    bad = []
    for sn in main_sheets(wb):
        ws = wb[sn]; gc = gov_col(sn)
        for r in range(3, ws.max_row + 1):
            v = ws.cell(r, gc).value
            if not isinstance(v, str) or not v.strip():
                continue
            m = GOV_REC_PAT.search(v)
            if m:
                bad.append((sn, r, f'推荐性/企业标准 {m.group(1)}', v[:100]))
            for h in GOV_FORBIDDEN_HINT:
                if h in v:
                    bad.append((sn, r, f'禁用文号/表述 {h}', v[:100]))
    return bad

# ---------- 链接 ----------
def rebuild_links(wb, kb_root, dead_out):
    s21 = wb['21.依据链接索引']
    s21._hyperlinks = []
    n21 = 0
    for r in range(4, s21.max_row + 1):
        p = s21.cell(r, 8).value
        if p and str(p).strip():
            s21.cell(r, 5).hyperlink = '..\\' + str(p).strip(); n21 += 1
    m = {}
    for r in range(4, s21.max_row + 1):
        e, h = s21.cell(r, 5).value, s21.cell(r, 8).value
        if e and h and str(h).strip():
            m.setdefault(str(e).strip(), str(h).strip())
    nmain = ndead = nwait = 0
    for sn in main_sheets(wb):
        ws = wb[sn]; rc = ref_col(sn)
        for r in range(3, ws.max_row + 1):
            cell = ws.cell(r, rc); v = cell.value
            if not v or not str(v).strip().startswith('《'):
                continue
            p = m.get(str(v).strip())
            if cell.hyperlink:
                nmain += 1
            elif p:
                if os.path.exists(os.path.join(kb_root, p)):
                    cell.hyperlink = '..\\' + p; nmain += 1
                else:
                    ndead += 1; dead_out.append((sn, r, p))
            else:
                nwait += 1
    return n21, nmain, ndead, nwait

def dead_scan(wb, kb_root):
    bad = []
    for sn in wb.sheetnames:
        if sn.startswith('21.') or sn.startswith('22.'):
            col = 5
        elif sn.startswith('20.'):
            col = 3
        else:
            col = ref_col(sn)
        ws = wb[sn]
        for r in range(1, ws.max_row + 1):
            t = ws.cell(r, col).hyperlink
            if t and t.target.startswith('..\\') and not os.path.exists(os.path.join(kb_root, t.target[3:])):
                bad.append((sn, r, t.target[:80]))
    return bad

# ---------- 删行手动迁移（2026-10-07 实测：openpyxl delete_rows 不迁 hyperlink/合并区/行高） ----------
def delete_row_migrated(ws, delrow, ncol):
    """删除第 delrow 行：链接/合并区(收缩或平移)/行高 全部手动迁移。仅支持单行删除。"""
    maxr = ws.max_row
    links = [(c.row, c.column, c.hyperlink.target)
             for row in ws.iter_rows(min_row=1, max_row=maxr, min_col=1, max_col=ncol)
             for c in row if c.hyperlink is not None]
    shrinks, shifts = [], []
    for rng in list(ws.merged_cells.ranges):
        b = (rng.min_col, rng.min_row, rng.max_col, rng.max_row)
        if rng.min_row <= delrow <= rng.max_row:
            shrinks.append(b)
        elif rng.min_row > delrow:
            shifts.append(b)
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row > delrow or (rng.min_row <= delrow <= rng.max_row):
            ws.unmerge_cells(str(rng))
    heights = {r - 1: ws.row_dimensions[r].height for r in range(delrow + 1, maxr + 1)
               if ws.row_dimensions[r].height}
    ws.delete_rows(delrow, 1)
    for c1, r1, c2, r2 in shrinks:
        if r1 <= delrow - 1:
            ws.merge_cells(start_row=r1, start_column=c1, end_row=delrow - 1, end_column=c2)
    for c1, r1, c2, r2 in shifts:
        ws.merge_cells(start_row=r1 - 1, start_column=c1, end_row=r2 - 1, end_column=c2)
    # 本版本保存时从 cell._hyperlink 重建链接集合；删后逐格重挂并刷新 ref
    for r, c, t in links:
        if r == delrow:
            continue
        nr = r - 1 if r > delrow else r
        cell = ws.cell(nr, c)
        cell._hyperlink = Hyperlink(ref=cell.coordinate, target=t)
    for r in range(delrow, maxr):
        ws.row_dimensions[r].height = heights.get(r)

# ---------- 错链审计（文本标准号 vs 链接文件名标准号，2026-10-07 实测抓出36处历史错链） ----------
CODE_PAT = re.compile(
    r"(GB/T|GB|DL/T|DL|JGJ/T|JGJ|NB/T|NB|JB/T|JB|DA/T|JJG|Q/BJP|DB11/T|CGC/GF|DLGJ|XF|AQ/T|AQ|CJJ/T|CJJ)"
    r"[\s—\-]*(\d{2,5})(?:\s*[.\-/]\s*(\d{1,2}))?")
BAD_HOSTS = ('antpedia', 'book118', 'docin.com', 'doc88', 'renrendoc', 'taodocs', 'infoeach')
URL_RE = re.compile(r"^https?://[^\s\"<>]+$")

def code_set(s):
    """标准号集合（前缀去/T差并归一：GBT 与 GB 视为同号族，文件名常省略/T）。"""
    out = set()
    if not s:
        return out
    for m in CODE_PAT.finditer(s):
        pre = m.group(1).replace('/', '').replace(' ', '')
        if pre.endswith('T') and pre != 'DAT':
            pre = pre[:-1]
        out.add((pre, int(m.group(2)), int(m.group(3)) if m.group(3) else None))
    return out

def is_mislinked(text, target):
    a, b = code_set(text), code_set(os.path.basename(target))
    return bool(a and b and not (a & b))

def s21_data_end(s21):
    """s21 正表区末行（'待补文件清单'标题行的上一行；找不到退回 max_row）。"""
    for r in range(4, s21.max_row + 1):
        v = s21.cell(r, 1).value
        if isinstance(v, str) and '待补文件清单' in v:
            return r - 1
    return s21.max_row

def audit_mislinks(wb, kb_root):
    """返回问题列表：(级别, sheet, 坐标/行, 问题描述)。覆盖主表C/F、s21 E 的本地/网络链接。"""
    issues = []
    for sn in main_sheets(wb):
        ws = wb[sn]
        col = ref_col(sn)
        tag = 'dev17错链' if is_dev17(sn) else '主表错链'
        for r in range(3, ws.max_row + 1):
            cell = ws.cell(r, col)
            v = cell.value
            if not isinstance(v, str) or cell.hyperlink is None:
                continue
            t = str(cell.hyperlink.target)
            if t.startswith('http'):
                if not URL_RE.match(t) or any(b in t for b in BAD_HOSTS):
                    issues.append((tag, sn, cell.coordinate, f'异常URL: {t}'))
            else:
                rel = t.replace('..\\', '').replace('../', '')
                fp = os.path.join(kb_root, rel)
                if not os.path.exists(fp):
                    issues.append(('死链', sn, cell.coordinate, rel))
                elif is_mislinked(v.strip(), rel):
                    issues.append((tag, sn, cell.coordinate, f'{v[:36]} -> {rel}'))
    if '21.依据链接索引' in wb.sheetnames:
        s21 = wb['21.依据链接索引']; end = s21_data_end(s21)
        for r in range(4, end + 1):
            cell = s21.cell(r, 5); v = cell.value
            if not isinstance(v, str) or cell.hyperlink is None:
                continue
            t = str(cell.hyperlink.target)
            if t.startswith('http'):
                if not URL_RE.match(t) or any(b in t for b in BAD_HOSTS):
                    issues.append(('s21异常URL', '21.依据链接索引', f'E{r}', t))
            else:
                rel = t.replace('..\\', '').replace('../', '')
                if not os.path.exists(os.path.join(kb_root, rel)):
                    issues.append(('s21死链', '21.依据链接索引', f'E{r}', rel))
                elif is_mislinked(v.strip(), rel):
                    issues.append(('s21错链', '21.依据链接索引', f'E{r}', f'{v[:36]} -> {rel}'))
    return issues

# ---------- s20 重建 ----------
def rebuild_print(wb):
    s20 = wb['20.打印（全）']
    srcs = [n for n in main_sheets(wb) if not is_dev17(n)]
    for rng in list(s20.merged_cells.ranges):
        if rng.min_row > 2:
            s20.unmerge_cells(str(rng))
    if s20.max_row > 2:
        s20.delete_rows(3, s20.max_row - 2)
    src0 = wb[srcs[0]]
    for c in range(1, 9):
        L = get_column_letter(c)
        if src0.column_dimensions[L].width:
            s20.column_dimensions[L].width = src0.column_dimensions[L].width
    out = 3
    for sn in srcs:
        ws = wb[sn]
        rmap = {}
        for r in range(3, ws.max_row + 1):
            if any(ws.cell(r, c).value is not None for c in range(1, 9)):
                rmap[r] = out; out += 1
        for r, tr in rmap.items():
            for c in range(1, 9):
                sc, dc = ws.cell(r, c), s20.cell(tr, c)
                dc.value = sc.value
                if sc.has_style:
                    dc.font, dc.fill, dc.border = copy(sc.font), copy(sc.fill), copy(sc.border)
                    dc.alignment, dc.number_format = copy(sc.alignment), sc.number_format
                if sc.hyperlink is not None:
                    dc.hyperlink = sc.hyperlink.target
            if ws.row_dimensions[r].height:
                s20.row_dimensions[tr].height = ws.row_dimensions[r].height
        for rng in list(ws.merged_cells.ranges):
            if rng.min_row > 2 and rng.min_row in rmap and rng.max_row in rmap and rng.max_col <= 8:
                s20.merge_cells(start_row=rmap[rng.min_row], start_column=rng.min_col,
                                end_row=rmap[rng.max_row], end_column=rng.max_col)
    return out - 1

# ---------- 美化（2026-10-08 全面精修轮：分块斑马纹 + 内容自适应行高 + 打印设置 + 保链字体） ----------
W8 = {'A':18,'B':40,'C':52,'D':10,'E':15,'F':15,'G':38,'H':24}
W9 = {'A':18,'B':13,'C':11,'D':23,'E':25,'F':40,'G':10,'H':34,'I':12}
W21 = {'A':6,'B':28,'C':22,'D':32,'E':46,'F':11,'G':10,'H':46}
BAND1, BAND2, SEC_FILL = 'FFF2F6FB', 'FFFFFFFF', 'FFD6E4F0'
HYPER_FONT = Font(name='微软雅黑', size=10, underline='single', color='FF0563C1')

def _fonts():
    thin = Side(style='thin', color='FF8EAADB')
    return (Font(name='微软雅黑', size=16, bold=True, color='FFFFFFFF'),
            Font(name='微软雅黑', size=11, bold=True, color='FFFFFFFF'),
            Font(name='微软雅黑', size=10),
            Border(left=thin, right=thin, top=thin, bottom=thin),
            Alignment(horizontal='center', vertical='center', wrap_text=True),
            Alignment(horizontal='left', vertical='center', wrap_text=True))

def last_data_row(ws, ncol):
    last = 2
    for r in range(3, ws.max_row + 1):
        if any(ws.cell(r, c).value is not None for c in range(1, ncol + 1)):
            last = r
    return last

def _vislen(s):
    """显示宽度：CJK/全角计 2，西文计 1。"""
    return sum(2 if ord(ch) > 127 else 1 for ch in str(s))

def _autoheight(ws, hr, lastr, ncol, widths, min_h=23.0, per_line=14.5, pad=7.0, cap=140.0):
    """按内容估算行高（避免文字被截断）；竖向合并区按跨度分摊需求，多推高不压缩。"""
    merges = [m for m in ws.merged_cells.ranges if m.max_row > hr and m.min_col <= ncol]
    ml = {}
    for mr in merges:
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                ml[(r, c)] = mr
    minh = {r: min_h for r in range(hr + 1, lastr + 1)}

    def need(v, w):
        return max(1, math.ceil(_vislen(v) / max(1.0, w - 1))) * per_line + pad

    for r in range(hr + 1, lastr + 1):
        for c in range(1, ncol + 1):
            v = ws.cell(r, c).value
            if v is None or str(v) == '' or (r, c) in ml:
                continue
            minh[r] = max(minh[r], need(v, widths.get(get_column_letter(c), 10)))
    for mr in merges:
        v = ws.cell(mr.min_row, mr.min_col).value
        if v is None or str(v) == '':
            continue
        w = sum(widths.get(get_column_letter(c), 10) for c in range(mr.min_col, mr.max_col + 1))
        nd = need(v, w)
        rows = list(range(mr.min_row, mr.max_row + 1))
        cur = sum(minh[r] for r in rows)
        if cur < nd:
            add = (nd - cur) / len(rows)
            for r in rows:
                minh[r] += add
    for r, h in minh.items():
        ws.row_dimensions[r].height = round(min(h, cap), 1)

def beautify(wb):
    FT, FH, FD, BD, AC, AL = _fonts()

    def head(ws, row, ncol, h):
        ws.row_dimensions[row].height = h
        for c in range(1, ncol + 1):
            x = ws.cell(row, c)
            x.font = FH; x.fill = PatternFill('solid', fgColor=BLUE); x.alignment = AC; x.border = BD

    def body(ws, hr, ncol, widths, keycol, srccol, dev=False):
        """分块斑马纹 + 字体/边框/对齐统一（保链字体）；返回数据末行。"""
        lastr = last_data_row(ws, ncol)
        band = False
        for r in range(hr + 1, lastr + 1):
            if ws.cell(r, keycol).value is not None:
                band = not band
            sec = dev and ws.cell(r, keycol).value is not None and \
                all(ws.cell(r, c).value in (None, '') for c in range(2, ncol + 1))
            for c in range(1, ncol + 1):
                x = ws.cell(r, c)
                x.font = HYPER_FONT if x.hyperlink is not None else FD
                x.border = BD
                x.alignment = AC if c == srccol else AL
                x.fill = PatternFill('solid', fgColor=(SEC_FILL if sec else (BAND1 if band else BAND2)))
            if sec:
                ws.cell(r, keycol).font = Font(name='微软雅黑', size=10, bold=True)
            sv = ws.cell(r, srccol).value
            if sv and str(sv).strip() in KIND_FILL:
                ws.cell(r, srccol).fill = PatternFill('solid', fgColor=KIND_FILL[str(sv).strip()])
                ws.cell(r, srccol).alignment = AC
        _autoheight(ws, hr, lastr, ncol, widths)
        ws.freeze_panes = f'A{hr + 1}'
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.print_title_rows = f'{hr}:{hr}'

    def titlerow(ws):
        ws.row_dimensions[1].height = 36
        t = ws.cell(1, 1); t.font = FT; t.fill = PatternFill('solid', fgColor=NAVY); t.alignment = AC

    for sn in main_sheets(wb):
        ws = wb[sn]; dev = is_dev17(sn); ncol = 9 if dev else 8
        widths = W9 if dev else W8
        for L, w in widths.items():
            ws.column_dimensions[L].width = w
        titlerow(ws)
        head(ws, 2, ncol, 26)
        body(ws, 2, ncol, widths, 1 if dev else 2, src_col(sn), dev)
    ws = wb['20.打印（全）']
    for L, w in W8.items():
        ws.column_dimensions[L].width = w
    titlerow(ws)
    head(ws, 2, 8, 26)
    body(ws, 2, 8, W8, 2, 4)
    ws = wb['21.依据链接索引']
    for L, w in W21.items():
        ws.column_dimensions[L].width = w
    titlerow(ws)
    head(ws, 3, 8, 26)
    body(ws, 3, 8, W21, 4, 6)
    if '22.NBT32037新增归档内容' in wb.sheetnames:
        ws = wb['22.NBT32037新增归档内容']
        hr = 2 if ws.cell(2, 1).value else 3
        titlerow(ws)
        head(ws, hr, 8, 22)
        body(ws, hr, 8, W8, 2, 4)

# ---------- 扫描 ----------
def scan(wb, kb_root):
    bad_gov = bad_fmt = 0
    for sn in main_sheets(wb):
        ws = wb[sn]; gc, rc = gov_col(sn), ref_col(sn)
        for r in range(3, ws.max_row + 1):
            g = ws.cell(r, gc).value
            if g and any(k in str(g) for k in ENT_SEG):
                bad_gov += 1
            v = ws.cell(r, rc).value
            if v and str(v).strip().startswith('《') and not re.search(r'[（(][^（）()]{2,}[)）]\s*$', str(v)):
                bad_fmt += 1
    dead = dead_scan(wb, kb_root)
    issues = audit_mislinks(wb, kb_root)
    gov_bad = gov_content_scan(wb)
    print(f'[scan] 监管列企业制度残留={bad_gov}  监管列推荐性标准/假文号={len(gov_bad)}  依据缺编号={bad_fmt}  死链={len(dead)}  错链/异常URL={len(issues)}')
    for g in gov_bad[:40]:
        print('   监管违规:', g)
    for d in dead[:20]:
        print('   死链:', d)
    for it in issues[:40]:
        print('   ', it)
    return bad_gov + bad_fmt + len(dead) + len(issues) + len(gov_bad)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--book', required=True)
    ap.add_argument('--kb-root', default=os.path.expandvars(r'$SAK_KB_ROOT\02-原始资料层'))
    ap.add_argument('--scan', action='store_true')
    ap.add_argument('--audit-mislinks', action='store_true',
                    help='错链/死链/异常URL专项审计（只读），有问题退出码1')
    ap.add_argument('--clean-gov', action='store_true')
    ap.add_argument('--normalize-ent-year', action='store_true')
    ap.add_argument('--rebuild-links', action='store_true')
    ap.add_argument('--rebuild-print', action='store_true')
    ap.add_argument('--beautify', action='store_true')
    a = ap.parse_args()
    wb = openpyxl.load_workbook(a.book)
    if a.scan:
        sys.exit(1 if scan(wb, a.kb_root) else 0)
    if a.audit_mislinks:
        issues = audit_mislinks(wb, a.kb_root)
        for it in issues:
            print('   ', it)
        print(f'[audit-mislinks] 问题 {len(issues)} 处')
        sys.exit(1 if issues else 0)
    if a.clean_gov:
        n = 0
        for sn in main_sheets(wb):
            ws = wb[sn]; gc = gov_col(sn)
            for r in range(3, ws.max_row + 1):
                x = ws.cell(r, gc)
                if x.value:
                    nv, ch = clean_gov_text(str(x.value))
                    if ch:
                        x.value = nv; n += 1
        print('清洗监管单元格:', n)
    if a.normalize_ent_year:
        n = 0
        for sn in main_sheets(wb):
            ws = wb[sn]; rc = ref_col(sn)
            for r in range(3, ws.max_row + 1):
                x = ws.cell(r, rc)
                if x.value and 'Q/BJP' in str(x.value):
                    nv = add_ent_year(str(x.value))
                    if nv != str(x.value):
                        x.value = nv; n += 1
        print('企业标准补年份:', n)
    if a.rebuild_links:
        dead = []
        n21, nmain, ndead, nwait = rebuild_links(wb, a.kb_root, dead)
        print(f'链接重建: s21={n21} 主表={nmain} 待补无链={nwait} 死链(库内缺件,需s21改待补)={ndead}')
        for d in dead:
            print('   ', d)
    if a.rebuild_print:
        print('s20重建末行:', rebuild_print(wb))
    if a.beautify:
        beautify(wb)
        print('美化完成')
    wb.save(a.book)
    with zipfile.ZipFile(a.book) as z:
        print('zip完整性:', 'OK' if z.testzip() is None else '损坏')

if __name__ == '__main__':
    main()
