import os
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sync_index_sheet.py — s21 依据链接索引表同步（插入→重排序号→重建链接→标题合并复位）

【教训来源】2026-10-03 s21 五段锚点插入 59 条索引：
  - insert_rows 导致 1993 处超链接错位（值移了、hyperlink 对象没移）；
    修复 = 清空 ws._hyperlinks 后按 H 列（权威路径列）逐行重建，重建后 E/H/链接不一致=0。
  - insert_rows 导致标题合并区 A3172:H3172 错位，数据行被压成 MergedCell 无法赋值；
    修复 = unmerge 旧区 → merge 新区 → 恢复被清数据行 → 重排序号。
  - 多锚点必须按行号倒序插入，前面的锚点才不失效。

s21 列约定：A序号 B来源清单 C类目 D分目/主要文件 E依据名称及编号(挂链接)
           F依据来源 G文件状态(在库/待补) H本地文件路径(相对02-原始资料层,权威列)
表头 3 行，数据自第 4 行起。行模板交替：灰底 FFF7F9FC / 白底 FFFFFFFF。

用法：
  python sync_index_sheet.py --config index.json [--dry-run]
  python sync_index_sheet.py --book <xlsx> --sheet "21.依据链接索引" --rebuild-links-only
  python sync_index_sheet.py --print-sample-config
  python sync_index_sheet.py --selftest
"""
import argparse
import json
import sys
from copy import copy

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import PatternFill
except ImportError:
    sys.exit("缺少 openpyxl，请用 .trae/python-venv/Scripts/python.exe 运行")

HEADER_ROWS = 3
COL_PATH = 8   # H 权威路径列
COL_NAME = 5   # E 依据名称列（挂链接）
COL_STATUS = 7  # G 在库/待补
FILL_GRAY = "FFF7F9FC"
FILL_WHITE = "FFFFFFFF"


def snapshot_row_style(ws, row, ncols=8):
    snap = []
    for col in range(1, ncols + 1):
        c = ws.cell(row, col)
        snap.append({"font": copy(c.font), "fill": copy(c.fill),
                     "border": copy(c.border), "alignment": copy(c.alignment),
                     "number_format": c.number_format})
    return snap


def apply_row_style(ws, row, snap, fill_rgb=None):
    for col, st in enumerate(snap, start=1):
        c = ws.cell(row, col)
        c.font, c.border = copy(st["font"]), copy(st["border"])
        c.alignment, c.number_format = copy(st["alignment"]), st["number_format"]
        c.fill = PatternFill("solid", fgColor=fill_rgb) if fill_rgb else copy(st["fill"])


def rebuild_hyperlinks(ws, first_row=HEADER_ROWS + 1, last_row=None):
    """清空全表链接对象，按 H 列权威路径逐行重建。G=待补 或 H 空 → 不挂链接。"""
    last_row = last_row or ws.max_row
    ws._hyperlinks = []  # 关键：insert_rows 后旧链接对象已错位，必须整体作废
    n = 0
    for r in range(first_row, last_row + 1):
        path = ws.cell(r, COL_PATH).value
        if path and str(path).strip():
            ws.cell(r, COL_NAME).hyperlink = "..\\" + str(path).strip()
            n += 1
    return n


def renumber(ws, first_row=HEADER_ROWS + 1, last_row=None):
    """主区重排序号：E 列非空的行才编号。返回最大序号。"""
    last_row = last_row or ws.max_row
    n = 0
    for r in range(first_row, last_row + 1):
        if ws.cell(r, COL_NAME).value:
            n += 1
            ws.cell(r, 1).value = n
        else:
            ws.cell(r, 1).value = None
    return n


def fix_title_merge(ws, title_text, correct_row, ncols=8):
    """标题合并区错位修复：unmerge 所有含标题文本的合并区 → 在正确行重合并。"""
    fixed = False
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == 1 and rng.max_col >= ncols:
            v = ws.cell(rng.min_row, 1).value
            if v and title_text in str(v) and rng.min_row != correct_row:
                ws.unmerge_cells(str(rng))
                fixed = True
    if fixed or not any(r.min_row == correct_row and r.min_col == 1
                        for r in ws.merged_cells.ranges):
        ws.merge_cells(start_row=correct_row, start_column=1,
                       end_row=correct_row, end_column=ncols)
    return fixed


def insert_entries(ws, anchor_row, entries, tpl_gray, tpl_white, dry_run=False):
    """
    在 anchor_row（段首行）前插入一组索引条目。entries: [{B,C,D,E,F,G,H}]（键为列字母）。
    奇偶行交替灰/白底。
    """
    n = len(entries)
    if dry_run:
        print(f"  [dry-run] r{anchor_row} 前插入 {n} 条")
        return
    ws.insert_rows(anchor_row, n)
    for i, e in enumerate(entries):
        r = anchor_row + i
        fill = FILL_GRAY if i % 2 == 0 else FILL_WHITE
        apply_row_style(ws, r, tpl_gray if i % 2 == 0 else tpl_white, fill_rgb=fill)
        for col_letter, val in e.items():
            if col_letter in ("A",):
                continue  # 序号最后统一重排
            from openpyxl.utils import column_index_from_string
            ws.cell(r, column_index_from_string(col_letter)).value = val


def run(config, dry_run=False):
    wb = load_workbook(config["workbook"])
    ws = wb[config.get("sheet", "21.依据链接索引")]
    tpl_gray = snapshot_row_style(ws, config.get("tpl_gray_row", 5))
    tpl_white = snapshot_row_style(ws, config.get("tpl_white_row", 6))

    # 1) 多锚点倒序插入（样式快照已在 insert_rows 之前完成）
    anchors = sorted(config.get("anchors", []), key=lambda a: a["row"], reverse=True)
    for a in anchors:
        insert_entries(ws, a["row"], a["entries"], tpl_gray, tpl_white, dry_run=dry_run)

    if dry_run:
        return

    # 2) 标题合并区复位（先于重排/重建，避免 MergedCell 只读挡路）
    tm = config.get("title_merge")
    if tm:
        fix_title_merge(ws, tm["text"], tm["row"])

    # 3) 主区重排序号
    last = config.get("renumber_last_row", ws.max_row)
    n_seq = renumber(ws, last_row=last)

    # 4) 全表链接按 H 列重建（必须在所有 insert_rows 完成之后）
    n_link = rebuild_hyperlinks(ws)

    wb.save(config["workbook"])
    print(f"[saved] {config['workbook']}  序号至 {n_seq}，重建链接 {n_link} 条")


SAMPLE_CONFIG = {
    "workbook": os.path.expandvars(r"$SAK_KB_ROOT\02-原始资料层\05问题分析参考清单\安全生产基建全周期标准化实施参考清单（cqs）.xlsx"),
    "sheet": "21.依据链接索引",
    "tpl_gray_row": 5,
    "tpl_white_row": 6,
    "title_merge": {"text": "待补", "row": 3231},
    "renumber_last_row": 3226,
    "anchors": [{
        "row": 541,
        "entries": [{
            "B": "7.项目施工阶段（管理）", "C": "综合管理 / 配套储能·综合",
            "D": "施工单位资质及……", "E": "《光伏发电建设项目文件归档与档案整理规范》（NB/T 32037-2026）",
            "F": "行业标准", "G": "在库",
            "H": r"行业标准\能源\NB_T32037—2026 光伏发电建设项目文件归档与档案整理规范.pdf"
        }]
    }]
}


def selftest():
    wb = Workbook()
    ws = wb.active
    for col, w in (("A", 6), ("B", 26), ("E", 48), ("H", 56)):
        ws.column_dimensions[col].width = w
    ws.cell(3, 1).value = "序号"
    for i in range(6):
        r = 4 + i
        ws.cell(r, 5).value = f"依据{i+1}"
        ws.cell(r, 7).value = "在库" if i % 2 == 0 else "待补"
        ws.cell(r, 8).value = rf"行业标准\能源\f{i}.pdf" if i % 2 == 0 else None
    # 模拟 insert_rows 后链接错位：故意挂错
    ws.cell(4, 5).hyperlink = "..\\wrong.pdf"
    n = rebuild_hyperlinks(ws)
    assert n == 3, n
    assert ws.cell(4, 5).hyperlink.target.endswith("f0.pdf")  # 错位已纠正
    assert ws.cell(5, 5).hyperlink is None  # 待补无链接
    m = renumber(ws)
    assert m == 6 and ws.cell(9, 1).value == 6
    ws.merge_cells("A10:H10")
    ws.cell(10, 1).value = "待补文件清单"
    ws.insert_rows(10, 2)  # 标题被挤到 12，合并区还在 10 → 模拟错位
    fix_title_merge(ws, "待补", 12)
    assert any(r.min_row == 12 for r in ws.merged_cells.ranges)
    print("[selftest] OK")


def main():
    ap = argparse.ArgumentParser(description="s21索引表同步")
    ap.add_argument("--config")
    ap.add_argument("--book", help="仅重建链接时直接给工作簿路径")
    ap.add_argument("--sheet", default="21.依据链接索引")
    ap.add_argument("--rebuild-links-only", action="store_true",
                    help="不插入，只按H列重建全表链接+重排序号")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--print-sample-config", action="store_true")
    ap.add_argument("--selftest", action="store_true")
    args = ap.parse_args()
    if args.selftest:
        selftest()
        return
    if args.print_sample_config:
        print(json.dumps(SAMPLE_CONFIG, ensure_ascii=False, indent=2))
        return
    if args.rebuild_links_only:
        if not args.book:
            ap.error("--rebuild-links-only 需要 --book")
        wb = load_workbook(args.book)
        ws = wb[args.sheet]
        n_seq = renumber(ws)
        n_link = rebuild_hyperlinks(ws)
        wb.save(args.book)
        print(f"[saved] {args.book}  序号至 {n_seq}，重建链接 {n_link} 条")
        return
    if not args.config:
        ap.error("需要 --config 或 --rebuild-links-only / --selftest")
    with open(args.config, encoding="utf-8") as f:
        run(json.load(f), dry_run=args.dry_run)


if __name__ == "__main__":
    main()
