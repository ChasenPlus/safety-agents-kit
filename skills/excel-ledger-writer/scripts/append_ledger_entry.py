#!/usr/bin/env python3
"""
示例电力制度自查台账 —— 幂等追加/状态管理脚本
用法：
    append 模式（默认，R06b用）：
        python3 append_ledger_entry.py --workbook <台账.xlsx> --entry-json '<JSON>'
    check-threshold 模式（R12自动检测用，是否该询问用户是否启动修订）：
        python3 append_ledger_entry.py --workbook <台账.xlsx> --mode check-threshold --code 2112 --threshold 2
    set-confirm-status 模式（R12，记录用户对"是否启动修订"的Y/N回应）：
        python3 append_ledger_entry.py --workbook <台账.xlsx> --mode set-confirm-status --code 2112 --confirm-status 待用户确认|已确认Y|已暂缓N
    update-status 模式（R12，起草/更新修订建议书后回写条目状态与"一本台账"追踪信息）：
        python3 append_ledger_entry.py --workbook <台账.xlsx> --mode update-status --code 2112 --name "《...》" \
            --new-status 已生成修订建议待评审 --doc-path /path/to.docx

行为要点：
1. 去重（R09）：append模式下"触发问题"+"记录日期"完全一致的行已存在则跳过。
2. sheet规则：每个"归属制度编号"对应一个sheet，sheet名 "{编号}_{制度名称去书名号}"。
3. 汇总sheet同步：append成功后同步"累计发现问题数"+1、"最近更新日期"。
4. R12"一本建议书"追踪：汇总sheet额外维护 修订建议书路径 / 建议书版本号 / 上次生成时条目数 /
   上次询问时条目数 / 修订确认状态 五列，用于支持"同一制度只有一份建议书、每次触发是重新
   生成覆盖而不是新建文件"和"阈值到达后不直接生成、而是先问用户Y/N，且不重复追问"这两条要求。
5. 写入失败不静默吞掉，返回 status=error 及原因。

始终以单行JSON打印结果到stdout。
"""

import argparse
import json
import sys
import os

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print(json.dumps({"status": "error", "message": "缺少依赖：请先执行 pip install openpyxl --break-system-packages"}, ensure_ascii=False))
    sys.exit(1)

ARIAL = "Arial"
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name=ARIAL, bold=True, color="FFFFFF", size=11)
NORMAL_FONT = Font(name=ARIAL, size=10)
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

ENTRY_HEADERS = [
    "记录日期", "触发问题", "归属制度编号", "归属制度名称",
    "缺口类型(完全空白/命中但未实质覆盖)",
    "命中条款说明(仅缺口类型=命中但未实质覆盖时填写)",
    "匹配置信度", "缺口描述", "建议补充方向", "参考的上位标准依据", "状态",
    "修订建议书路径(仅状态=已生成修订建议待评审时填写)",
]
ENTRY_KEYS = [
    "记录日期", "触发问题", "归属制度编号", "归属制度名称",
    "缺口类型", "命中条款说明", "匹配置信度", "缺口描述",
    "建议补充方向", "参考的上位标准依据", "状态", "修订建议书路径",
]

SUMMARY_HEADERS = [
    "制度编号", "制度名称", "累计发现问题数", "最近更新日期", "对应工作表",
    "修订建议书路径", "建议书版本号", "上次生成时条目数", "上次询问时条目数", "修订确认状态",
]


def sanitize_sheet_name(code: str, name: str) -> str:
    clean_name = name.replace("《", "").replace("》", "")
    sheet_name = f"{code}_{clean_name}"
    for ch in r':\/?*[]':
        sheet_name = sheet_name.replace(ch, "")
    return sheet_name[:31]


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def ensure_entry_sheet(wb, sheet_name):
    if sheet_name in wb.sheetnames:
        return wb[sheet_name], False
    ws = wb.create_sheet(sheet_name)
    for c, h in enumerate(ENTRY_HEADERS, start=1):
        ws.cell(row=1, column=c, value=h)
    style_header(ws, 1, len(ENTRY_HEADERS))
    ws.freeze_panes = "A2"
    widths = [12, 26, 12, 26, 18, 40, 22, 34, 30, 30, 12, 30]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws, True


def find_duplicate(ws, entry):
    for row in range(2, ws.max_row + 1):
        date_val = ws.cell(row=row, column=1).value
        question_val = ws.cell(row=row, column=2).value
        if date_val == entry["记录日期"] and question_val == entry["触发问题"]:
            return row
    return None


def append_entry_row(ws, entry):
    row = ws.max_row + 1
    if ws.max_row == 1 and ws.cell(row=1, column=1).value == ENTRY_HEADERS[0]:
        row = 2
    for c, key in enumerate(ENTRY_KEYS, start=1):
        val = entry.get(key, "" if key != "状态" else "待补充")
        cell = ws.cell(row=row, column=c, value=val)
        cell.font = NORMAL_FONT
        cell.border = BORDER
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    return row


def ensure_summary_sheet(wb):
    if "汇总" in wb.sheetnames:
        return wb["汇总"]
    ws = wb.create_sheet("汇总", 0)
    for c, h in enumerate(SUMMARY_HEADERS, start=1):
        ws.cell(row=1, column=c, value=h)
    style_header(ws, 1, len(SUMMARY_HEADERS))
    ws.freeze_panes = "A2"
    widths = [12, 30, 16, 16, 34, 44, 12, 16, 16, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


def find_summary_row(ws_sum, code):
    for row in range(2, ws_sum.max_row + 1):
        if ws_sum.cell(row=row, column=1).value == code:
            return row
    return None


def update_summary_on_append(ws_sum, code, name, sheet_name, record_date):
    row = find_summary_row(ws_sum, code)
    if row:
        count_cell = ws_sum.cell(row=row, column=3)
        count_cell.value = (count_cell.value or 0) + 1
        ws_sum.cell(row=row, column=4, value=record_date)
        return
    new_row = ws_sum.max_row + 1
    if ws_sum.max_row == 1 and ws_sum.cell(row=1, column=1).value == SUMMARY_HEADERS[0]:
        new_row = 2
    values = [code, name, 1, record_date, sheet_name, "", 0, 0, 0, ""]
    for c, v in enumerate(values, start=1):
        cell = ws_sum.cell(row=new_row, column=c, value=v)
        cell.font = NORMAL_FONT
        cell.border = BORDER


def update_entry_status(wb, code, name, matched_dates_questions, new_status, doc_path=""):
    sheet_name = sanitize_sheet_name(code, name)
    if sheet_name not in wb.sheetnames:
        return 0
    ws = wb[sheet_name]
    updated = 0
    for row in range(2, ws.max_row + 1):
        row_date = ws.cell(row=row, column=1).value
        row_question = ws.cell(row=row, column=2).value
        if row_date is None and row_question is None:
            continue
        match = False
        if matched_dates_questions:
            if (row_date, row_question) in matched_dates_questions:
                match = True
        else:
            if ws.cell(row=row, column=11).value in ("待补充", None, ""):
                match = True
        if match:
            ws.cell(row=row, column=11, value=new_status)
            if doc_path:
                ws.cell(row=row, column=12, value=doc_path)
            updated += 1
    return updated


def sync_summary_after_amendment(ws_sum, code, doc_path):
    """一本建议书生成/更新完成后，同步汇总sheet的追踪字段——
    版本号+1、记下这次生成时的累计条目数（供下次diff用）、重置确认状态。
    不重置"上次询问时条目数"——那是用来防止重复提问的独立维度。"""
    row = find_summary_row(ws_sum, code)
    if not row:
        return None
    current_count = ws_sum.cell(row=row, column=3).value or 0
    version_cell = ws_sum.cell(row=row, column=7)
    new_version = (version_cell.value or 0) + 1
    ws_sum.cell(row=row, column=6, value=doc_path)
    ws_sum.cell(row=row, column=7, value=new_version)
    ws_sum.cell(row=row, column=8, value=current_count)
    ws_sum.cell(row=row, column=9, value=current_count)  # 生成动作本身也算"已经处理过这批"，同步归零询问计数器
    ws_sum.cell(row=row, column=10, value="")  # 重置确认状态，等下一次新增缺口再触发
    return {"version": new_version, "count_at_generation": current_count}


def check_threshold(wb, code, threshold):
    """R12自动检测：是否应该去问用户"要不要启动/更新修订建议书"。

    两个独立维度，不能合并成一个：
    - reached：单纯反映"是否有值得关注的新增缺口"（count相对上次生成建议书的时点是否有增长），仅供参考展示。
    - should_ask：真正决定"这次要不要去问用户"，条件是 count 相对"上次问过用户的时点"(count_at_last_ask)
      有新增，且当前不是"提问已发出、还没收到回应"的中间状态。无论用户上次回答Y还是N，
      只要之后没有新增缺口，就不再重复提问。
    """
    empty_result = {
        "count": 0, "reached": False, "should_ask": False, "already_pending": False,
        "existing_doc_path": "", "existing_version": 0, "count_at_last_generation": 0,
        "count_at_last_ask": 0, "confirm_status": "",
    }
    if "汇总" not in wb.sheetnames:
        return dict(empty_result)
    ws = wb["汇总"]
    row = find_summary_row(ws, code)
    if not row:
        return dict(empty_result)
    count = ws.cell(row=row, column=3).value or 0
    doc_path = ws.cell(row=row, column=6).value or ""
    version = ws.cell(row=row, column=7).value or 0
    count_at_gen = ws.cell(row=row, column=8).value or 0
    count_at_ask = ws.cell(row=row, column=9).value or 0
    confirm_status = ws.cell(row=row, column=10).value or ""

    reached = count >= threshold and count > count_at_gen
    already_pending = confirm_status == "待用户确认"
    should_ask = count >= threshold and count > count_at_ask and not already_pending

    return {
        "count": count, "reached": reached, "should_ask": should_ask,
        "already_pending": already_pending, "existing_doc_path": doc_path,
        "existing_version": version, "count_at_last_generation": count_at_gen,
        "count_at_last_ask": count_at_ask, "confirm_status": confirm_status,
    }


def set_confirm_status(wb, code, status):
    if "汇总" not in wb.sheetnames:
        return False
    ws = wb["汇总"]
    row = find_summary_row(ws, code)
    if not row:
        return False
    if status == "待用户确认":
        current_count = ws.cell(row=row, column=3).value or 0
        ws.cell(row=row, column=9, value=current_count)  # 记下"这次是在累计多少条时问的"，防止重复追问
    ws.cell(row=row, column=10, value=status)
    return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook", default=None, help="台账xlsx文件路径（也可通过--config-file的workbook字段提供，二者必须提供其一）。"
                                                           "Windows下路径含中文时建议改用--config-file传递，"
                                                           "命令行参数可能因终端代码页不是UTF-8被错误解码导致乱码")
    parser.add_argument("--mode", choices=["append", "check-threshold", "set-confirm-status", "update-status"],
                         default="append")
    parser.add_argument("--entry-json", help="条目JSON字符串(append模式用)")
    parser.add_argument("--entry-file", help="条目JSON文件路径(append模式用，与--entry-json二选一)")
    parser.add_argument("--code", help="归属制度编号")
    parser.add_argument("--name", help="归属制度名称(update-status模式用，需与sheet命名规则一致)")
    parser.add_argument("--new-status", help="新状态值(update-status模式用)")
    parser.add_argument("--doc-path", default="", help="修订建议书文件路径(update-status模式)")
    parser.add_argument("--match-list-json", default="[]", help="update-status模式：指定要更新哪些行，空数组则更新全部'待补充'行")
    parser.add_argument("--threshold", type=int, default=2, help="check-threshold模式：触发询问的累计条目数阈值")
    parser.add_argument("--confirm-status", choices=["待用户确认", "已确认Y", "已暂缓N", ""],
                         help="set-confirm-status模式：用户对是否启动修订的回应")
    parser.add_argument("--config-file", help="可选：JSON文件，可提供workbook/doc_path/code/name字段覆盖同名命令行参数，"
                                               "所有字段可选、按需提供。路径含中文时强烈建议用这种方式而不是命令行直传，"
                                               "这里走utf-8安全的文件读取，不经过命令行编码链路")
    args = parser.parse_args()

    if args.config_file:
        with open(args.config_file, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        if "workbook" in cfg:
            args.workbook = cfg["workbook"]
        if "doc_path" in cfg:
            args.doc_path = cfg["doc_path"]
        if "code" in cfg:
            args.code = cfg["code"]
        if "name" in cfg:
            args.name = cfg["name"]

    if not args.workbook:
        print(json.dumps({"status": "error", "message": "必须提供 --workbook，或通过 --config-file 的 workbook 字段提供"}, ensure_ascii=False))
        sys.exit(1)

    if args.mode == "check-threshold":
        if not args.code:
            print(json.dumps({"status": "error", "message": "check-threshold模式需要--code"}, ensure_ascii=False)); sys.exit(1)
        if os.path.exists(args.workbook):
            wb = openpyxl.load_workbook(args.workbook)
        else:
            # 台账文件还不存在，等同于"这个制度目前是0条记录"，不是错误——
            # 用一个空白内存工作簿走正常的check_threshold()逻辑，自然落入"汇总sheet不存在"分支，
            # 保证返回字段和文件存在但该制度尚无记录时完全一致，不用在这里手写重复一份。
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
        result = check_threshold(wb, args.code, args.threshold)
        print(json.dumps({"status": "ok", **result}, ensure_ascii=False))
        return

    if args.mode == "set-confirm-status":
        if not (args.code and args.confirm_status is not None):
            print(json.dumps({"status": "error", "message": "set-confirm-status模式需要--code --confirm-status"}, ensure_ascii=False)); sys.exit(1)
        if not os.path.exists(args.workbook):
            print(json.dumps({"status": "error", "message": "台账文件不存在"}, ensure_ascii=False)); sys.exit(1)
        try:
            wb = openpyxl.load_workbook(args.workbook)
            ok = set_confirm_status(wb, args.code, args.confirm_status)
            wb.save(args.workbook)
            print(json.dumps({"status": "ok" if ok else "error", "message": None if ok else "未找到该制度编号的汇总记录"}, ensure_ascii=False))
        except Exception as e:
            print(json.dumps({"status": "error", "message": str(e)}, ensure_ascii=False)); sys.exit(1)
        return

    if args.mode == "update-status":
        if not (args.code and args.name and args.new_status):
            print(json.dumps({"status": "error", "message": "update-status模式需要--code --name --new-status"}, ensure_ascii=False)); sys.exit(1)
        if not os.path.exists(args.workbook):
            print(json.dumps({"status": "error", "message": "台账文件不存在"}, ensure_ascii=False)); sys.exit(1)
        try:
            matched = [tuple(x) for x in json.loads(args.match_list_json)]
            wb = openpyxl.load_workbook(args.workbook)
            updated = update_entry_status(wb, args.code, args.name, matched, args.new_status, args.doc_path)
            summary_result = None
            if args.doc_path:
                ws_sum = ensure_summary_sheet(wb)
                summary_result = sync_summary_after_amendment(ws_sum, args.code, args.doc_path)
            wb.save(args.workbook)
            print(json.dumps({"status": "ok", "updated_rows": updated, "summary": summary_result}, ensure_ascii=False))
        except Exception as e:
            print(json.dumps({"status": "error", "message": str(e)}, ensure_ascii=False)); sys.exit(1)
        return

    # ---- append 模式 ----
    if not args.entry_json and not args.entry_file:
        print(json.dumps({"status": "error", "message": "必须提供 --entry-json 或 --entry-file 之一"}, ensure_ascii=False)); sys.exit(1)

    try:
        if args.entry_json:
            entry = json.loads(args.entry_json)
        else:
            with open(args.entry_file, "r", encoding="utf-8") as f:
                entry = json.load(f)
    except Exception as e:
        print(json.dumps({"status": "error", "message": f"条目JSON解析失败: {e}"}, ensure_ascii=False)); sys.exit(1)

    required = ["记录日期", "触发问题", "归属制度编号", "归属制度名称"]
    missing = [k for k in required if not entry.get(k)]
    if missing:
        print(json.dumps({"status": "error", "message": f"缺少必填字段: {missing}"}, ensure_ascii=False)); sys.exit(1)

    entry.setdefault("状态", "待补充")
    entry.setdefault("缺口类型", "完全空白")
    entry.setdefault("命中条款说明", "（不适用）")
    entry.setdefault("修订建议书路径", "")

    try:
        if os.path.exists(args.workbook):
            wb = openpyxl.load_workbook(args.workbook)
        else:
            os.makedirs(os.path.dirname(args.workbook) or ".", exist_ok=True)
            wb = openpyxl.Workbook()
            wb.remove(wb.active)

        sheet_name = sanitize_sheet_name(entry["归属制度编号"], entry["归属制度名称"])
        ws, created = ensure_entry_sheet(wb, sheet_name)

        dup_row = find_duplicate(ws, entry)
        if dup_row:
            wb.save(args.workbook)
            print(json.dumps({"status": "duplicate", "sheet": sheet_name, "existing_row": dup_row}, ensure_ascii=False))
            return

        row = append_entry_row(ws, entry)

        ws_sum = ensure_summary_sheet(wb)
        update_summary_on_append(ws_sum, entry["归属制度编号"], entry["归属制度名称"], sheet_name, entry["记录日期"])

        wb.save(args.workbook)
        print(json.dumps({"status": "written", "sheet": sheet_name, "row": row, "summary_updated": True}, ensure_ascii=False))

    except Exception as e:
        print(json.dumps({"status": "error", "message": str(e)}, ensure_ascii=False)); sys.exit(1)


if __name__ == "__main__":
    main()
