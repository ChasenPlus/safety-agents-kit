# -*- coding: utf-8 -*-
"""sak-excel：xlsx 读写 MCP 服务器（openpyxl，不依赖 Node）。

替代原先 `npx @negokaz/excel-mcp-server`。
用途：读台账、写回状态、批量追加行。不做复杂公式计算，也不改格式模板。

环境变量：
    SAK_FS_ROOTS   允许读写的根目录，多个用 os.pathsep 分隔；默认 SAK_KB_ROOT。
"""
import os
from pathlib import Path

from fastmcp import FastMCP
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

mcp = FastMCP("sak-excel")


def _roots() -> list[Path]:
    raw = os.environ.get("SAK_FS_ROOTS") or os.environ.get("SAK_KB_ROOT") or ""
    items = [p for p in raw.split(os.pathsep) if p.strip()]
    if not items:
        items = [str(Path(__file__).resolve().parents[2] / "kb_template")]
    return [Path(p).expanduser().resolve() for p in items]


def _resolve(path: str) -> Path:
    p = Path(path).expanduser()
    if not p.is_absolute():
        p = _roots()[0] / p
    p = p.resolve()
    for root in _roots():
        if p == root or p.is_relative_to(root):
            return p
    raise PermissionError(f"路径不在允许范围内：{p}\n允许的根目录：{[str(r) for r in _roots()]}")


@mcp.tool()
def describe_sheets(path: str) -> dict:
    """列出 xlsx 的所有工作表：名称、行列数、表头首行。写表格前先调这个确认 sheet 名。"""
    p = _resolve(path)
    if not p.is_file():
        return {"status": "error", "message": f"文件不存在：{p}"}
    wb = load_workbook(p, read_only=True, data_only=True)
    try:
        sheets = []
        for ws in wb.worksheets:
            header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1), [])]
            sheets.append({
                "name": ws.title,
                "rows": ws.max_row,
                "columns": ws.max_column,
                "header": [str(h) if h is not None else "" for h in header],
            })
        return {"status": "ok", "path": str(p), "sheets": sheets}
    finally:
        wb.close()


@mcp.tool()
def read_sheet(path: str, sheet: str = "", max_rows: int = 200) -> dict:
    """读取工作表内容（默认第一张表）。返回二维数组，空单元格为 null。
    max_rows 限制返回行数，避免整表撑爆上下文。"""
    p = _resolve(path)
    if not p.is_file():
        return {"status": "error", "message": f"文件不存在：{p}"}
    wb = load_workbook(p, read_only=True, data_only=True)
    try:
        if sheet:
            if sheet not in wb.sheetnames:
                return {"status": "error", "message": f"工作表不存在：{sheet}",
                        "sheets": wb.sheetnames}
            ws = wb[sheet]
        else:
            ws = wb.worksheets[0]
        rows = []
        for row in ws.iter_rows(max_row=max_rows, values_only=True):
            rows.append(list(row))
        return {
            "status": "ok",
            "path": str(p),
            "sheet": ws.title,
            "total_rows": ws.max_row,
            "returned_rows": len(rows),
            "truncated": ws.max_row > len(rows),
            "rows": rows,
        }
    finally:
        wb.close()


@mcp.tool()
def write_sheet(path: str, rows: list[list], sheet: str = "", mode: str = "append",
                start_row: int = 0) -> dict:
    """写工作表。mode="append" 追加到已有数据之后；mode="overwrite" 从 start_row（1 起）开始覆盖写。
    sheet 为空时用第一张表；文件不存在会新建。返回写入的起止行号，便于幂等核对。"""
    p = _resolve(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    if p.is_file():
        wb = load_workbook(p)
    else:
        from openpyxl import Workbook
        wb = Workbook()
    try:
        if sheet:
            if sheet in wb.sheetnames:
                ws = wb[sheet]
            else:
                ws = wb.create_sheet(sheet)
                # 新建工作簿默认带一张空 Sheet，指定了 sheet 名就把它去掉
                default = wb["Sheet"] if "Sheet" in wb.sheetnames else None
                if (default is not None and len(wb.sheetnames) > 1
                        and default.max_row == 1
                        and all(c.value is None for c in default[1])):
                    wb.remove(default)
        else:
            ws = wb.worksheets[0]
        if mode == "overwrite":
            begin = start_row or 1
            if begin == 1 and ws.max_row:
                ws.delete_rows(1, ws.max_row)
                begin = 1
        elif mode == "append":
            # 空表（只有一行且全空）从第 1 行开始写，避免留一行空白
            first_row_empty = ws.max_row == 1 and all(c.value is None for c in ws[1])
            begin = 1 if first_row_empty else ws.max_row + 1
        else:
            return {"status": "error", "message": f"未知 mode：{mode}（append / overwrite）"}

        width = 0
        for i, row in enumerate(rows):
            for j, value in enumerate(row, 1):
                ws.cell(row=begin + i, column=j, value=value)
            width = max(width, len(row))
        wb.save(p)
        return {
            "status": "ok",
            "path": str(p),
            "sheet": ws.title,
            "mode": mode,
            "start_row": begin,
            "end_row": begin + len(rows) - 1 if rows else begin - 1,
            "rows_written": len(rows),
            "columns": width,
            "last_column": get_column_letter(width) if width else "",
        }
    finally:
        wb.close()


if __name__ == "__main__":
    mcp.run()
