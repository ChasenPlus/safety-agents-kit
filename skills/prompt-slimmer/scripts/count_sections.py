# -*- coding: utf-8 -*-
"""按 markdown 标题切节并统计各节字数，用于定位提示词中需要瘦身的大块。

用法:
    python -X utf8 count_sections.py <markdown文件路径>

字数不含标题本身、不计空白字符；超过 800 字的节会被标记为大块。
"""
import re
import sys


def main(path):
    with open(path, encoding='utf-8') as f:
        lines = f.read().split('\n')

    sections = []
    cur_title = '(文件开头)'
    cur_buf = 0
    for line in lines:
        if re.match(r'^#{1,6}\s', line):
            sections.append((cur_title, cur_buf))
            cur_title = re.sub(r'^#+\s*', '', line).strip()
            cur_buf = 0
        else:
            cur_buf += len(line.strip())
    sections.append((cur_title, cur_buf))

    total = sum(count for _, count in sections)
    print('文件: %s' % path)
    print('总字数(不含标题/空白): %d   章节数: %d' % (total, len(sections)))
    print('-' * 72)
    for title, count in sections:
        flag = '  <<< 大块(>800字)' if count > 800 else ''
        print('%6d  %s%s' % (count, title[:48], flag))


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python -X utf8 count_sections.py <markdown文件路径>')
        sys.exit(1)
    main(sys.argv[1])
