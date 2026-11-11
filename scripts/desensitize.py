#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""脱敏脚本：把仓库内残留的企业名、人名、绝对路径、内部项目代号替换为示例值。

用法：
    python scripts/desensitize.py --dry-run   # 只统计，不落盘
    python scripts/desensitize.py             # 执行替换

对文本类文件（.md/.py/.json/.yaml/.yml/.txt/.toml/.example）生效；
文件名命中的会一并改名。图片/PDF/docx/xlsx 不处理，需人工确认后不要提交。
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TEXT_SUFFIXES = {".md", ".py", ".json", ".yaml", ".yml", ".txt", ".toml", ".example", ".cfg", ".ini"}
SKIP_DIRS = {".git", "__pycache__", "node_modules", ".venv", "venv"}

# 顺序敏感：长串在前，避免"京能电力"被先替换成"示例集团电力"
REPLACEMENTS = [
    ("北京京能电力股份有限公司", "示例电力股份有限公司"),
    ("京能电力全能小助手", "示例电力全能小助手"),
    ("京能电力", "示例电力"),
    ("京能", "示例集团"),
    ("岱海发电", "示例发电厂"),
    ("岱海", "示例电厂"),
    ("石热", "示例电厂A"),
    ("京玉", "示例电厂B"),
    ("京泰", "示例电厂C"),
    ("盛乐", "示例电厂D"),
    ("漳山", "示例电厂E"),
    ("金隅兴发科技园三期分布式光伏", "示例光伏项目"),
    ("内蒙古", "某省"),
    # 绝对路径
    (r"D:\工作\AI中心\我的AI大脑", "${SAK_KB_ROOT}"),
    (r"D:\工作\AI中心\TRAE用简化版", "${SAK_PROJECT_ROOT}"),
    (r"D:\工作\AI中心", "${SAK_WORK_ROOT}"),
    (r"C:\Users\CC", "${HOME}"),
    # 项目内代号
    ("jingneng", "sak"),
    ("Jingneng", "Sak"),
    ("JINGNENG", "SAK"),
]

# 扫描用：替换后仍应归零的模式
RESIDUAL_PATTERNS = [
    ("京能", re.compile("京能")),
    ("岱海", re.compile("岱海")),
    ("绝对路径 D:\\工作", re.compile(r"D:\\工作")),
    ("用户目录 C:\\Users", re.compile(r"C:\\Users")),
    ("jingneng", re.compile("jingneng", re.IGNORECASE)),
]


def iter_text_files() -> list[Path]:
    out = []
    for p in ROOT.rglob("*"):
        if not p.is_file():
            continue
        if any(part in SKIP_DIRS for part in p.parts):
            continue
        if p.resolve() == Path(__file__).resolve():
            continue  # 脚本自己的替换表不能被替换
        if p.suffix.lower() in TEXT_SUFFIXES or p.name.endswith(".example"):
            out.append(p)
    return out


def desensitize(text: str) -> str:
    for old, new in REPLACEMENTS:
        text = text.replace(old, new)
    return text


def scan() -> int:
    """返回残留命中总数。"""
    total = 0
    for path in iter_text_files():
        try:
            text = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue
        for name, pat in RESIDUAL_PATTERNS:
            hits = pat.findall(text)
            if hits:
                total += len(hits)
                print(f"[残留] {path.relative_to(ROOT)}  {name} x{len(hits)}")
    print(f"残留命中合计：{total}")
    return total


def apply(dry_run: bool) -> None:
    changed = 0
    for path in iter_text_files():
        try:
            text = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue
        new_text = desensitize(text)
        if new_text != text:
            changed += 1
            print(f"[改写] {path.relative_to(ROOT)}")
            if not dry_run:
                path.write_text(new_text, encoding="utf-8")
    print(f"共 {changed} 个文件需要改写（dry_run={dry_run}）")


def rename_files(dry_run: bool) -> None:
    for path in sorted(ROOT.rglob("*"), key=lambda p: len(p.parts), reverse=True):
        if not path.exists():
            continue
        new_name = desensitize(path.name)
        if new_name == path.name:
            continue
        target = path.with_name(new_name)
        print(f"[改名] {path.relative_to(ROOT)} -> {new_name}")
        if not dry_run:
            path.rename(target)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--scan-only", action="store_true")
    args = ap.parse_args()

    if args.scan_only:
        return 1 if scan() else 0

    apply(args.dry_run)
    rename_files(args.dry_run)
    # 改名后需要重新遍历再扫一遍
    return 1 if (not args.dry_run and scan()) else 0


if __name__ == "__main__":
    sys.exit(main())
