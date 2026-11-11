import os
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
insert_group_rows.py — cqs基建清单主表"分组依据行"结构化写入

【教训来源】2026-10-03 NB/T 32037 新增归档 13 组 59 行落位（s7/s8/s9/s11 插行 + s13 覆盖占位行）
【核心坑】openpyxl insert_rows 只移动单元格值，不移动 hyperlink 对象与 merged_cells 定义。
  本脚本对策：
  - mode=overwrite（优先）：在既有占位行上直接写，完全绕开 insert_rows；
  - mode=insert：先快照模板行样式再 insert_rows（否则模板行被挤走后抄到错误样式），
    写完后【调用方负责】对被挤动的区域重建超链接并复核合并区
    （见 sync_index_sheet.py 的 rebuild_hyperlinks）。
  - MergedCell 只读：写合并区覆盖的单元格前必须先 unmerge。
  - D 列底色：抄完模板样式后按 D 值用 FILL 映射逐行校正（勿照抄模板行底色）。

用法：
  python insert_group_rows.py --config groups.json [--dry-run]
  python insert_group_rows.py --print-sample-config
  python insert_group_rows.py --selftest      # 内存自检，不写盘

解释器：.trae/python-venv/Scripts/python.exe（已装 openpyxl）
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

# 依据来源分类 → D列底色（2026-10-03 全表统一映射，勿改）
FILL_BY_KIND = {
    "国家标准": "FFDDEBF7",
    "行业标准": "FFE2EFDA",
    "企业标准": "FFFCE4D6",
    "法律法规": "FFF8CBAD",
    "团体标准": "FFEDE7F6",
}

COLS = 8  # A..H


def est_lines(text, width):
    """按列宽估算折行数：中文记2.1、英文记1，换行符重置。"""
    if not text:
        return 1
    total = 0
    for para in str(text).split("\n"):
        cur = 0.0
        lines = 1
        for ch in para:
            cur += 2.1 if ord(ch) > 127 else 1.0
            if cur > width:
                lines += 1
                cur = 0.0
        total += lines
    return max(total, 1)


def first_row_height(group, widths):
    """组首行高 = max(43.2, min(lines*13.4+3, 200))，lines 取 B/G/H 三列估算最大值。"""
    b = est_lines(group.get("doc", ""), widths.get("B", 48))
    g = est_lines(group.get("note", ""), widths.get("G", 20))
    h = est_lines(group.get("hreg", ""), widths.get("H", 40))
    lines = max(b, g, h)
    return max(43.2, min(lines * 13.4 + 3, 200))


def snapshot_row_style(ws, row):
    """把模板行 8 列样式复制成独立对象列表（insert_rows 前调用，防模板被挤走）。"""
    snap = []
    for col in range(1, COLS + 1):
        c = ws.cell(row, col)
        snap.append({
            "font": copy(c.font), "fill": copy(c.fill), "border": copy(c.border),
            "alignment": copy(c.alignment), "number_format": c.number_format,
        })
    return snap


def apply_style(ws, row, snap):
    for col, st in enumerate(snap, start=1):
        c = ws.cell(row, col)
        c.font, c.fill, c.border = copy(st["font"]), copy(st["fill"]), copy(st["border"])
        c.alignment, c.number_format = copy(st["alignment"]), st["number_format"]


def unmerge_covering(ws, r0, r1):
    """解开所有覆盖 [r0, r1] 的合并区（MergedCell 只读，必须先解）。"""
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row <= r1 and rng.max_row >= r0:
            ws.unmerge_cells(str(rng))


def write_group(ws, r0, group, tpl_first_snap, tpl_next_snap, widths, dry_run=False):
    """
    在 r0 起写入一组：组首行（A/B/E/F/G/H 有值）+ 后续依据行（仅 C/D 有值），
    组内 A/B/E/F/G/H 纵向合并。返回下一空行行号。
    group = {cat, doc, esrc, fman, note, hreg, cites:[{text, kind, path|None}]}
    path 为相对 02-原始资料层 的路径；None 表示待补（不挂链接）。
    """
    cites = group["cites"]
    n = len(cites)
    if dry_run:
        print(f"  [dry-run] {ws.title} r{r0}-{r0+n-1} 《{cites[0]['text'][:20]}…》等{n}条")
        return r0 + n

    unmerge_covering(ws, r0, r0 + n - 1)
    for i, cite in enumerate(cites):
        r = r0 + i
        apply_style(ws, r, tpl_first_snap if i == 0 else tpl_next_snap)
        c3 = ws.cell(r, 3)
        c3.value = cite["text"]
        c3.hyperlink = None
        if cite.get("path"):
            c3.hyperlink = "..\\" + cite["path"]
        c4 = ws.cell(r, 4)
        c4.value = cite["kind"]
        # D列底色按分类值校正（不照抄模板行底色——s7 曾因此误绿 59 行）
        rgb = FILL_BY_KIND.get(cite["kind"])
        if rgb:
            c4.fill = PatternFill("solid", fgColor=rgb)
        if i == 0:
            ws.cell(r, 1).value = group.get("cat", "")
            ws.cell(r, 2).value = group.get("doc", "")
            ws.cell(r, 5).value = group.get("esrc", "")
            ws.cell(r, 6).value = group.get("fman", "")
            ws.cell(r, 7).value = group.get("note", "")
            ws.cell(r, 8).value = group.get("hreg", "")
    if n > 1:
        for col_letter in ("A", "B", "E", "F", "G", "H"):
            ws.merge_cells(f"{col_letter}{r0}:{col_letter}{r0 + n - 1}")
    # 行高：组首按内容估算，依据行固定 29.8
    ws.row_dimensions[r0].height = first_row_height(group, widths)
    for i in range(1, n):
        ws.row_dimensions[r0 + i].height = 29.8
    return r0 + n


def run(config, dry_run=False):
    wb = load_workbook(config["workbook"])
    for sheet_cfg in config["sheets"]:
        ws = wb[sheet_cfg["sheet"]]
        widths = {k: ws.column_dimensions[k].width or 12
                  for k in ("B", "G", "H")}
        # 模板样式快照必须在任何 insert_rows 之前完成
        tpl_first = snapshot_row_style(ws, sheet_cfg["template_first"])
        tpl_next = snapshot_row_style(ws, sheet_cfg["template_next"])
        groups = sheet_cfg["groups"]
        # 同一表多组插行：按锚点行号倒序处理，前面的锚点才不会失效
        groups = sorted(groups, key=lambda g: g["anchor_row"], reverse=True)
        for g in groups:
            r0 = g["anchor_row"]
            mode = g.get("mode", "insert")
            if mode == "insert" and not dry_run:
                ws.insert_rows(r0, len(g["cites"]))
            elif mode == "overwrite":
                pass  # 直接覆盖占位行，不动行结构
            else:
                raise ValueError(f"未知 mode: {mode}")
            write_group(ws, r0, g, tpl_first, tpl_next, widths, dry_run=dry_run)
        print(f"[ok] {sheet_cfg['sheet']}: {len(groups)} 组处理完毕")
    if not dry_run:
        wb.save(config["workbook"])
        print(f"[saved] {config['workbook']}")
    print("[提醒] insert 模式下被挤动区域的既有超链接/合并区可能已错位，"
          "请继续跑 sync_index_sheet.py 重建链接、verify_workbook.py 全量验证")


SAMPLE_CONFIG = {
    "workbook": os.path.expandvars(r"$SAK_KB_ROOT\02-原始资料层\05问题分析参考清单\安全生产基建全周期标准化实施参考清单（cqs）.xlsx"),
    "sheets": [{
        "sheet": "7.项目施工阶段（管理）",
        "template_first": 4,
        "template_next": 5,
        "groups": [{
            "anchor_row": 88,
            "mode": "insert",
            "cat": "综合管理 / 配套储能·综合",
            "doc": "施工单位资质及……（按NB/T 32037-2026附录B 8370）",
            "esrc": "项目所辖部门", "fman": "项目所辖部门",
            "note": "", "hreg": "国家能源局要求",
            "cites": [
                {"text": "《光伏发电建设项目文件归档与档案整理规范》（NB/T 32037-2026）",
                 "kind": "行业标准",
                 "path": r"行业标准\能源\NB_T32037—2026 光伏发电建设项目文件归档与档案整理规范.pdf"},
                {"text": "《某待补标准》（GB 00000-2026）", "kind": "国家标准", "path": None}
            ]
        }]
    }]
}


def selftest():
    wb = Workbook()
    ws = wb.active
    ws.title = "t"
    for col, w in (("B", 48), ("G", 20), ("H", 40)):
        ws.column_dimensions[col].width = w
    ws.cell(4, 3).value = "tpl1"
    ws.cell(5, 3).value = "tpl2"
    g = {"cat": "类目A", "doc": "主要文件" * 30, "esrc": "E", "fman": "F",
         "note": "备注", "hreg": "监管",
         "cites": [{"text": "依据一", "kind": "国家标准",
                    "path": r"行业标准\能源\x.pdf"},
                   {"text": "依据二", "kind": "行业标准", "path": None}]}
    t1 = snapshot_row_style(ws, 4)
    t2 = snapshot_row_style(ws, 5)
    nxt = write_group(ws, 10, g, t1, t2, {"B": 48, "G": 20, "H": 40})
    assert nxt == 12, nxt
    assert ws.cell(10, 1).value == "类目A"
    assert ws.cell(11, 3).value == "依据二"
    assert ws.cell(10, 3).hyperlink.target.endswith("x.pdf")
    assert ws.cell(11, 3).hyperlink is None  # 待补不挂链接
    assert ws.cell(11, 4).fill.fgColor.rgb == "FFE2EFDA"
    assert "A10:A11" in [str(r) for r in ws.merged_cells.ranges]
    assert est_lines("中文中文abc", 10) >= 2
    print("[selftest] OK")


def main():
    ap = argparse.ArgumentParser(description="cqs清单主表分组插行")
    ap.add_argument("--config", help="groups.json 路径")
    ap.add_argument("--dry-run", action="store_true", help="只打印计划不写盘")
    ap.add_argument("--print-sample-config", action="store_true")
    ap.add_argument("--selftest", action="store_true")
    args = ap.parse_args()
    if args.selftest:
        selftest()
        return
    if args.print_sample_config:
        print(json.dumps(SAMPLE_CONFIG, ensure_ascii=False, indent=2))
        return
    if not args.config:
        ap.error("需要 --config 或 --selftest / --print-sample-config")
    with open(args.config, encoding="utf-8") as f:
        run(json.load(f), dry_run=args.dry_run)


if __name__ == "__main__":
    main()
