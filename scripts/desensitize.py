#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""脱敏检查：扫描仓库里残留的绝对路径、手机号、身份证号，并按本地词表替换。

用法：
    python scripts/desensitize.py --scan-only   # 只扫描，命中即返回非 0
    python scripts/desensitize.py --dry-run     # 预览会改哪些文件
    python scripts/desensitize.py               # 按词表执行替换

仓库本身已脱敏，脚本里不内置任何真实企业名。需要替换自己环境里的词时，
在 `scripts/sensitive_words.txt` 里一行写一条 `原词=替换词`（该文件已被 gitignore）。
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WORDS_FILE = Path(__file__).resolve().with_name("sensitive_words.txt")
TEXT_SUFFIXES = {".md", ".py", ".json", ".yaml", ".yml", ".txt", ".toml", ".example", ".cfg", ".ini"}
SKIP_DIRS = {".git", "__pycache__", "node_modules", ".venv", "venv"}

# 通用模式：与具体企业无关，任何人跑都应归零
RESIDUAL_PATTERNS = [
    ("Windows 绝对路径", re.compile(r"[A-Za-z]:\\+[^\s\"'|<>）)]{2,}")),
    ("macOS/Linux 用户目录", re.compile(r"/(?:Users|home)/[^\s\"'|<>）)]+")),
    ("手机号", re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)")),
    ("身份证号", re.compile(r"(?<!\d)\d{17}[\dXx](?!\d)")),
]


def iter_text_files() -> list[Path]:
    out = []
    for p in ROOT.rglob("*"):
        if not p.is_file():
            continue
        if any(part in SKIP_DIRS for part in p.parts):
            continue
        if p.resolve() == Path(__file__).resolve():
            continue
        if p.suffix.lower() in TEXT_SUFFIXES or p.name.endswith(".example"):
            out.append(p)
    return out


def load_words() -> list[tuple[str, str]]:
    if not WORDS_FILE.is_file():
        return []
    pairs: list[tuple[str, str]] = []
    for line in WORDS_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        old, new = line.split("=", 1)
        if old.strip():
            pairs.append((old.strip(), new.strip()))
    return pairs


def desensitize(text: str, pairs: list[tuple[str, str]]) -> str:
    for old, new in pairs:
        text = text.replace(old, new)
    return text


def scan() -> int:
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
                print(f"[残留] {path.relative_to(ROOT)}  {name} x{len(hits)}: {hits[:3]}")
    print(f"残留命中合计：{total}")
    return total


def apply(pairs: list[tuple[str, str]], dry_run: bool) -> None:
    changed = 0
    for path in iter_text_files():
        try:
            text = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue
        new_text = desensitize(text, pairs)
        if new_text != text:
            changed += 1
            print(f"[改写] {path.relative_to(ROOT)}")
            if not dry_run:
                path.write_text(new_text, encoding="utf-8")
    print(f"共 {changed} 个文件需要改写（dry_run={dry_run}）")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--scan-only", action="store_true")
    args = ap.parse_args()

    if args.scan_only:
        return 1 if scan() else 0

    pairs = load_words()
    if not pairs:
        print(f"未找到词表 {WORDS_FILE}，只做通用模式扫描")
    apply(pairs, args.dry_run)
    return 1 if (not args.dry_run and scan()) else 0


if __name__ == "__main__":
    sys.exit(main())
