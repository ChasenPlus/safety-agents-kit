#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rebuild_print_sheet.py — s20 打印（全）表整表重建

【策略】永远不用 insert_rows：清空目标表数据区 → 按主表顺序整段复制
（值/样式/超链接/行高/合并区加行偏移）。2026-10-03 实测重建 2918 行、
合并区与主表逐行比对 0 差异。

表结构约定：主表与打印表均为 1=大标题、2=列头、3=空行、4 起为数据。

用法：
  python rebuild_print_sheet.py --book <xlsx> --print-sheet "20.打印（全）" \
      --main-sheets "1.项目开发及技术更新" "2.项目建设前期（可研）" ...
  python rebuild_print_sheet.py --book <xlsx> --config sheets.json   # {"print_sheet":..., "main_sheets":[...]}
  python rebuild_print_sheet.py --selftest
"""
import argparse
import json
import sys
from copy import copy

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("缺少 openpyxl，请用 python 运行")

COLS = 8          # A..H
HEADER_ROWS = 3   # 1标题 2列头 3空行，数据从第4行起


def rebuild(wb, print_sheet, main_sheets, dry_run=False):
    dst = wb[print_sheet]
    if not dry_run:
        # 清目标表数据区：删行最干净（连同行高），合并区先全解再按新区重建
        for rng in list(dst.merged_cells.ranges):
            if rng.min_row > HEADER_ROWS:
                dst.unmerge_cells(str(rng))
        if dst.max_row > HEADER_ROWS:
            dst.delete_rows(HEADER_ROWS + 1, dst.max_row - HEADER_ROWS)

    # 列宽以第一个主表为准
    src0 = wb[main_sheets[0]]
    if not dry_run:
        for col in range(1, COLS + 1):
            letter = get_column_letter(col)
            w = src0.column_dimensions[letter].width
            if w:
                dst.column_dimensions[letter].width = w

    out_row = HEADER_ROWS + 1
    total_links = 0
    for name in main_sheets:
        src = wb[name]
        base = out_row - HEADER_ROWS - 1  # 合并区行偏移 = 目标起始 - 源起始(4)
        n_rows = src.max_row - HEADER_ROWS
        if dry_run:
            print(f"  [dry-run] {name}: {n_rows} 行 → r{out_row}-{out_row+n_rows-1}")
            out_row += n_rows
            continue
        for r in range(HEADER_ROWS + 1, src.max_row + 1):
            tr = r + base
            for col in range(1, COLS + 1):
                sc, dc = src.cell(r, col), dst.cell(tr, col)
                dc.value = sc.value
                if sc.has_style:
                    dc.font, dc.fill, dc.border = copy(sc.font), copy(sc.fill), copy(sc.border)
                    dc.alignment, dc.number_format = copy(sc.alignment), sc.number_format
                if sc.hyperlink is not None:
                    dc.hyperlink = sc.hyperlink.target
                    total_links += 1
            h = src.row_dimensions[r].height
            if h:
                dst.row_dimensions[tr].height = h
        # 合并区加行偏移复制
        for rng in list(src.merged_cells.ranges):
            if rng.min_row > HEADER_ROWS:
                dst.merge_cells(start_row=rng.min_row + base, start_column=rng.min_col,
                                end_row=rng.max_row + base, end_column=rng.max_col)
        out_row += n_rows
        print(f"[ok] {name}: {n_rows} 行")
    print(f"[done] {print_sheet} 重建至 {out_row - 1} 行，链接 {total_links} 条")
    return out_row - 1


def selftest():
    wb = Workbook()
    m1 = wb.active
    m1.title = "1.甲"
    m2 = wb.create_sheet("2.乙")
    p = wb.create_sheet("20.打印（全）")
    for ws in (m1, m2, p):
        ws.cell(1, 1).value = "标题"
        ws.cell(2, 1).value = "列头"
    for i, v in enumerate(("a", "b", "c")):
        m1.cell(4 + i, 1).value = v
    m1.cell(4, 3).value = "依据"
    m1.cell(4, 3).hyperlink = "..\\x.pdf"
    m1.merge_cells("A4:A5")
    m2.cell(4, 1).value = "z"
    last = rebuild(wb, "20.打印（全）", ["1.甲", "2.乙"])
    assert last == 7, last
    assert p.cell(4, 1).value == "a" and p.cell(7, 1).value == "z"
    assert p.cell(4, 3).hyperlink.target == "..\\x.pdf"
    assert any(r.min_row == 4 and r.max_row == 5 for r in p.merged_cells.ranges), \
        [str(r) for r in p.merged_cells.ranges]
    print("[selftest] OK")


def main():
    ap = argparse.ArgumentParser(description="s20打印表整表重建")
    ap.add_argument("--book", help="工作簿路径")
    ap.add_argument("--print-sheet", default="20.打印（全）")
    ap.add_argument("--main-sheets", nargs="*", help="主表名（按顺序）")
    ap.add_argument("--config", help='{"print_sheet":..., "main_sheets":[...]}')
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--selftest", action="store_true")
    args = ap.parse_args()
    if args.selftest:
        selftest()
        return
    if args.config:
        with open(args.config, encoding="utf-8") as f:
            cfg = json.load(f)
        print_sheet, main_sheets = cfg["print_sheet"], cfg["main_sheets"]
    else:
        if not args.book or not args.main_sheets:
            ap.error("需要 --book + --main-sheets，或 --config")
        print_sheet, main_sheets = args.print_sheet, args.main_sheets
    book = args.book or cfg["workbook"]
    wb = load_workbook(book)
    rebuild(wb, print_sheet, main_sheets, dry_run=args.dry_run)
    if not args.dry_run:
        wb.save(book)
        print(f"[saved] {book}")


if __name__ == "__main__":
    main()
