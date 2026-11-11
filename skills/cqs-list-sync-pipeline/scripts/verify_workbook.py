#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
verify_workbook.py — cqs清单工作簿全量验证（替换原件前的放行门禁）

6 项检查（任一 FAIL 退出码非 0，禁止替换原件）：
  1. zip 完整性 + openpyxl 重载
  2. 死链检查：全部超链接 '..\' 相对路径解析后必须存在（相对 xlsx 所在目录）
  3. s21 三方一致：E 名称行有 H 路径 ⇔ 有链接，且链接 == '..\'+H
  4. s20 与主表逐行逐列值 diff（打印表必须是主表的忠实拼接）
  5. D 列依据来源色标一致：D 值 → FILL 映射（国家标准蓝/行业标准绿/企业标准橙/法律法规橙/团体标准紫）
  6. 合并区计数（信息项，与 --expect-merges 给定时参与判定）

2026-10-09 实测基准（22 sheets）：全簿 9765 链接 0 死链、s20 与主表 0 差异、合并区 8996 个。
布局：8列主表/打印表数据自 r3 起（MAIN_DATA0=3），s21 数据自 r4 起（INDEX_DATA0=4）。

用法：
  python verify_workbook.py --book <xlsx> --kb-root <02-原始资料层目录> \
      --print-sheet "20.打印（全）" --index-sheet "21.依据链接索引" \
      --main-sheets "1.项目开发及技术更新" ... [--expect-merges 3942]
  python verify_workbook.py --selftest
"""
import argparse
import os
import sys
import zipfile

try:
    from openpyxl import load_workbook, Workbook
except ImportError:
    sys.exit("缺少 openpyxl，请用 .trae/python-venv/Scripts/python.exe 运行")

COLS = 8
# 本簿实际布局（2026-10-09 复核）：8列主表/打印表 r1大标题(A1:H1合并)、r2表头，数据自 r3 起；
# s21索引表 r1标题、r2空、r3表头，数据自 r4 起。旧版统一用 HEADER_ROWS=3 会导致
# 打印表比对/D列色标整体错一行（假 FAIL），已拆分为两个起始行常量。
MAIN_DATA0 = 3
INDEX_DATA0 = 4
FILL_BY_KIND = {
    "国家标准": "FFDDEBF7",
    "行业标准": "FFE2EFDA",
    "企业标准": "FFFCE4D6",
    "法律法规": "FFF8CBAD",
    "团体标准": "FFEDE7F6",
}


def check_zip_reload(path):
    try:
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
            if bad:
                return False, f"zip 损坏条目: {bad}"
        load_workbook(path)
        return True, "zip/重载正常"
    except Exception as e:
        return False, f"重载失败: {e}"


def check_dead_links(wb, book_dir):
    """全簿超链接：'..\\' 开头的一律解析并验存在。返回 (ok, 总数, 死链列表)。"""
    total, dead = 0, []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.hyperlink is not None:
                    t = c.hyperlink.target
                    if not t:
                        continue
                    total += 1
                    if t.startswith("..\\"):
                        p = os.path.normpath(os.path.join(book_dir, t))
                        if not os.path.exists(p):
                            dead.append(f"{ws.title}!{c.coordinate} -> {t}")
    return (not dead), total, dead


def check_index_consistency(ws):
    """s21 三方一致：H 有路径 ⇔ E 有链接且 == '..\\'+H；G=待补 行不得有链接。

    网络行约定（铁律）：G=在线全文/在线书目/官方公告 时 H 留空、E 挂 http(s) URL；
    故 H 空 + http 链接是合法状态，不再判失败；但 H 空不得挂本地 '..\\' 链接。
    """
    bad = []
    for r in range(INDEX_DATA0, ws.max_row + 1):
        h = ws.cell(r, 8).value
        e = ws.cell(r, 5)
        status = ws.cell(r, 7).value
        has_h = bool(h and str(h).strip())
        link = e.hyperlink.target if e.hyperlink else None
        if has_h:
            expect = "..\\" + str(h).strip()
            if link != expect:
                bad.append(f"r{r}: 链接={link!r} 期望={expect!r}")
        elif link:
            if not link.startswith("http"):
                bad.append(f"r{r}: H 空却挂非 http 链接 {link!r}")
            elif not e.value:
                bad.append(f"r{r}: H 空挂链接但 E 为空")
        elif e.value:
            bad.append(f"r{r}: E 有值却无链接 {str(e.value)[:30]!r}")
        if status == "待补" and link:
            bad.append(f"r{r}: 待补行挂链接 {link!r}")
    return (not bad), bad


def last_data_row(ws, ncol=COLS):
    """数据区最后非空行（跳过尾部空格式行）。"""
    last = 0
    for r in range(MAIN_DATA0, ws.max_row + 1):
        for c in range(1, ncol + 1):
            v = ws.cell(r, c).value
            if v is not None and str(v).strip():
                last = r
                break
    return last


def check_print_vs_mains(wb, print_sheet, main_sheets, max_show=10):
    """s20 数据区 == 各主表数据区顺序拼接（逐行逐列值比对）。"""
    dst = wb[print_sheet]
    diffs, out_row = [], MAIN_DATA0
    for name in main_sheets:
        src = wb[name]
        for r in range(MAIN_DATA0, last_data_row(src) + 1):
            tr = out_row + (r - MAIN_DATA0)
            for col in range(1, COLS + 1):
                sv, dv = src.cell(r, col).value, dst.cell(tr, col).value
                if (sv or "") != (dv or ""):
                    diffs.append(f"{name}!r{r} vs 打印!r{tr} col{col}: {sv!r} != {dv!r}")
            if len(diffs) >= max_show:
                return False, diffs
        out_row += last_data_row(src) - MAIN_DATA0 + 1
    if dst.max_row > out_row - 1:
        # 尾部允许多出的全空行
        for r in range(out_row, dst.max_row + 1):
            if any(dst.cell(r, c).value for c in range(1, COLS + 1)):
                diffs.append(f"打印表尾部多出非空行 r{r}")
                break
    return (not diffs), diffs


def check_d_fill(wb, sheet_names, max_show=10):
    """主表+s20：D 列值若命中 FILL 映射，底色必须一致。"""
    bad = []
    for name in sheet_names:
        ws = wb[name]
        for r in range(MAIN_DATA0, last_data_row(ws) + 1):
            v = ws.cell(r, 4).value
            if v in FILL_BY_KIND:
                rgb = ws.cell(r, 4).fill.fgColor.rgb
                if rgb != FILL_BY_KIND[v]:
                    bad.append(f"{name}!D{r}: {v} 底色={rgb} 期望={FILL_BY_KIND[v]}")
                    if len(bad) >= max_show:
                        return False, bad
    return (not bad), bad


def count_merges(wb):
    return {ws.title: len(ws.merged_cells.ranges) for ws in wb.worksheets}


def verify(book, kb_root, print_sheet, index_sheet, main_sheets, expect_merges=None):
    book_dir = kb_root or os.path.dirname(os.path.abspath(book))
    results = []
    ok, msg = check_zip_reload(book)
    results.append(("1.zip完整性+重载", ok, [msg]))
    if not ok:
        return results

    wb = load_workbook(book)

    ok, total, dead = check_dead_links(wb, book_dir)
    results.append((f"2.死链检查（共{total}条）", ok, dead[:10]))

    if index_sheet in wb.sheetnames:
        ok, bad = check_index_consistency(wb[index_sheet])
        results.append(("3.索引E/H/链接三方一致", ok, bad[:10]))

    if print_sheet in wb.sheetnames and main_sheets:
        ok, diffs = check_print_vs_mains(wb, print_sheet, main_sheets)
        results.append(("4.打印表与主表逐行比对", ok, diffs[:10]))

    fill_sheets = list(main_sheets or []) + ([print_sheet] if print_sheet in wb.sheetnames else [])
    if fill_sheets:
        ok, bad = check_d_fill(wb, fill_sheets)
        results.append(("5.D列依据来源色标", ok, bad[:10]))

    merges = count_merges(wb)
    total_m = sum(merges.values())
    info = [f"{k}: {v}" for k, v in merges.items()]
    ok = (expect_merges is None) or (total_m == expect_merges)
    results.append((f"6.合并区计数（共{total_m}）", ok, info if not ok else info[:5]))

    return results


def selftest(tmp=None):
    import tempfile
    from openpyxl.styles import PatternFill
    # 模拟真实布局：主表/打印表 r1大标题(合并)+r2表头，数据 r3 起；s21 r1标题+r2空+r3表头，数据 r4 起
    d = tmp or tempfile.mkdtemp(prefix="cqs_verify_")
    list_dir = os.path.join(d, "清单")
    os.makedirs(list_dir, exist_ok=True)
    os.makedirs(os.path.join(d, "sub"), exist_ok=True)
    with open(os.path.join(d, "sub", "x.pdf"), "w") as f:
        f.write("x")
    wb = Workbook()
    m = wb.active
    m.title = "1.甲"
    p = wb.create_sheet("20.打印（全）")
    idx = wb.create_sheet("21.依据链接索引")
    for ws in (m, p):
        ws.cell(1, 1).value = "大标题"
        ws.merge_cells("A1:H1")
        for c in range(1, COLS + 1):
            ws.cell(2, c).value = f"head{c}"
    idx.cell(1, 1).value = "依据链接索引"
    idx.merge_cells("A1:H1")
    for c in range(1, COLS + 1):
        idx.cell(3, c).value = f"ihead{c}"
    m.cell(3, 1).value = "数据"
    m.cell(3, 3).value = "依据甲"
    m.cell(3, 3).hyperlink = "..\\sub\\x.pdf"
    m.cell(3, 4).value = "行业标准"
    m.cell(3, 4).fill = PatternFill("solid", fgColor="FFE2EFDA")
    p.cell(3, 1).value = "数据"
    p.cell(3, 3).value = "依据甲"
    p.cell(3, 3).hyperlink = "..\\sub\\x.pdf"
    p.cell(3, 4).value = "行业标准"
    p.cell(3, 4).fill = PatternFill("solid", fgColor="FFE2EFDA")
    # s21：行1 本地在库（H 有路径），行2 网络行（H 空 + http，合法）
    idx.cell(4, 1).value = 1
    idx.cell(4, 5).value = "依据甲"
    idx.cell(4, 7).value = "在库"
    idx.cell(4, 8).value = "sub\\x.pdf"
    idx.cell(4, 5).hyperlink = "..\\sub\\x.pdf"
    idx.cell(5, 1).value = 2
    idx.cell(5, 5).value = "依据乙"
    idx.cell(5, 7).value = "在线全文"
    idx.cell(5, 5).hyperlink = "https://example.org/x"
    book = os.path.join(list_dir, "t.xlsx")
    wb.save(book)
    res = verify(book, None, "20.打印（全）", "21.依据链接索引", ["1.甲"])
    for name, ok, detail in res:
        assert ok, (name, detail)
    print("[selftest] OK（%d 项全过）" % len(res))


def main():
    ap = argparse.ArgumentParser(description="cqs清单工作簿全量验证")
    ap.add_argument("--book")
    ap.add_argument("--kb-root", help="02-原始资料层目录；缺省取 xlsx 所在目录")
    ap.add_argument("--print-sheet", default="20.打印（全）")
    ap.add_argument("--index-sheet", default="21.依据链接索引")
    ap.add_argument("--main-sheets", nargs="*")
    ap.add_argument("--expect-merges", type=int)
    ap.add_argument("--selftest", action="store_true")
    args = ap.parse_args()
    if args.selftest:
        selftest()
        return
    if not args.book:
        ap.error("需要 --book 或 --selftest")
    res = verify(args.book, args.kb_root, args.print_sheet,
                 args.index_sheet, args.main_sheets, args.expect_merges)
    all_ok = True
    for name, ok, detail in res:
        print(f"[{'PASS' if ok else 'FAIL'}] {name}")
        if not ok:
            all_ok = False
            for line in detail:
                print(f"    {line}")
    print("=" * 40)
    print("全量验证通过，可替换原件" if all_ok else "存在 FAIL 项，禁止替换原件")
    sys.exit(0 if all_ok else 1)


if __name__ == "__main__":
    main()
