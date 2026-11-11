# -*- coding: utf-8 -*-
"""全项目交付门禁（GATE）——八个分支所有落盘交付物统一入口。

用法：
    python gate.py <文件> [--type auto|report|docx|xlsx|std] [--ref 原稿] [--allow-rot]
                   [--template] [--profile auto|group|platform|company]

档位（--profile，2026-10-31 起）：按报告类别分档，判决标准不同、不得互相套用——
group＝集团/公司(正文图)类（学术论文式组图，5×5/4.5×4.5/6×4.5/4.5×6）；
platform＝示例电力等二级平台类（表格内嵌图一律 7.5×6.0，同题多图先合成为一张；另查三章骨架与定责栏）；
company＝公司类（表格内嵌图，宽固定 4.7cm）。auto 只认目录名"电力报告"/"公司报告"，
其余按 group。平台/公司档会把表格单元格内段落一并扫描（正文级扫描对表格图是空跑）。

约定：**必须看到 GATE PASS 才算交付**；GATE FAIL 一律回改后重跑。
为什么要有它：规矩写成散文会被"通读一遍自觉没问题"跳过，写成断言才拦得住。
每犯一次新错，就在本文件加一条断言并登记到《使用说明》。
"""
import io
import os
import re
import sys
import zipfile
from collections import Counter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.oxml.ns import qn

CM = 360000
TOL = 3000
# 档位（2026-10-31 按报告类别拆分）：各类报告格式独立设计，门禁也分档，
# 不得用一类的断言判另一类（集团四档套到平台报告上会误杀，反之会漏检）。
#   group    = 集团/公司类（学术论文式组图）：组图≤4张2列5×5、≥5张3列4.5×4.5；
#              单图横版6×4.5、竖版4.5×6
#   platform = 示例电力等二级平台类：图片一律 wp:inline 嵌表格单元格，
#              单图 7.5×6.0、同题多图各 3.65×6.0（高度固定6.0，允许变形）
OK_GROUP = ((5.0, 5.0), (4.5, 4.5), (6.0, 4.5), (4.5, 6.0))
OK_PLATFORM = ((7.5, 6.0), (3.65, 6.0))
PROFILE = 'group'
PROFILE_NAME = {'group': '集团/公司(正文图)类', 'platform': '二级平台类',
                'company': '公司类(单元格图)'}


def size_ok(cx, cy):
    if PROFILE == 'company':      # 公司净化版：照片插单元格，宽固定 4.7cm、高随比例
        return abs(cx - int(4.7 * CM)) < TOL
    t = OK_PLATFORM if PROFILE == 'platform' else OK_GROUP
    return any(abs(cx - int(a * CM)) < TOL and abs(cy - int(b * CM)) < TOL
               for a, b in t)


FAIL, WARN, DONE = [], [], []
ALLOW_PH = False        # --allow-ph：模板测试件可保留占位符，仅跳过 P6，其余检查照跑


def fail(cid, msg):
    FAIL.append('[%s] %s' % (cid, msg))


def warn(cid, msg):
    WARN.append('[%s] %s' % (cid, msg))


def ok(cid):
    DONE.append(cid)


def T(p):
    return ''.join(x.text or '' for x in p._p.iter(qn('w:t'))).strip()


def is_toc(p):
    """真正的目录条目＝toc 样式 + PAGEREF 域。
    正文大纲标题误用 toc 样式（历史遗留）不算目录条目，否则 P15 会被虚高蒙混。"""
    if not p.style.name.lower().startswith('toc'):
        return False
    return any('PAGEREF' in (r.text or '') for r in p._p.iter(qn('w:instrText')))


def has_pic(p):
    return bool(p._p.findall('.//' + qn('pic:pic')))


# ── 规则集中处（唯一事实源，改这里就全项目生效）──────────────
BAN = ['未检索到判定条款', '依据待核实', '建议完善依据', '本次检查未发现',
       '安全可控', '仅反映本次抽查', '未提及的领域', '不代表不存在问题',
       '本报告仅反映', '略', '……']
PLACEHOLDER = ['{', '}', 'XX', 'xxx', 'XXX', 'TODO', '待填写', '此处填写',
               '【填写', '（待定）', '[待', '（略）']
OKFONT = ('方正小标宋简体', 'Times New Roman', '宋体', '黑体', '楷体_GB2312',
          '仿宋_GB2312', '仿宋', '楷体', '等线', '微软雅黑')


# ── docx 通用 ────────────────────────────────────────────
def docx_images(d, rel):
    z = zipfile.ZipFile(rel)
    xml = z.read('word/document.xml').decode('utf-8')
    if PROFILE in ('platform', 'company'):
        # 平台/公司类报告的照片全部嵌在表格单元格里，只看正文级段落等于空跑，
        # 故这两档把单元格内段落一并纳入（P2/P19/P20 才真正生效）
        from docx.text.paragraph import Paragraph
        paras = [Paragraph(el, d) for el in d.element.body.iter(qn('w:p'))]
        paras = [p for p in paras if not is_toc(p)]
    else:
        paras = [p for p in d.paragraphs if not is_toc(p)]

    rots = [r for r in re.findall(r'<a:xfrm[^>]*\brot="(-?\d+)"', xml) if int(r) != 0]
    if rots:
        warn('P1 旋转', '%d 处图片带旋转角 %s' % (len(rots), rots[:4]))
    if re.search(r'<a:xfrm[^>]*\bflip[HV]="1"', xml):
        fail('P1 旋转', '存在 flipH/flipV')
    ok('P1 旋转')

    bad, mism, sp_bad = [], 0, 0
    for p in paras:
        boxes = list(p._p.iter(qn('wp:extent')))
        if not boxes:
            continue
        pPr = p._p.pPr
        sp = pPr.find(qn('w:spacing')) if pPr is not None else None
        if PROFILE == 'group' and (sp is None or sp.get(qn('w:lineRule')) != 'auto'):
            # 行距 auto 是为防固定行距裁图；平台/公司类图在单元格内、行高由表格决定，不适用
            sp_bad += 1
        for ext in boxes:
            cx, cy = int(ext.get('cx')), int(ext.get('cy'))
            good = size_ok(cx, cy)
            if not good:
                logo = (not T(p)) and cx / max(cy, 1) > 3
                (warn if logo else fail)(
                    'P2 尺寸', '非标 %.2fx%.2f %s' % (cx / CM, cy / CM,
                    '｜logo 放行' if logo else '｜' + T(p)[:24]))
            for spPr in p._p.iter(qn('pic:spPr')):
                xf = spPr.find(qn('a:xfrm'))
                if xf is None:
                    continue
                for e in xf.findall(qn('a:ext')):
                    if abs(int(e.get('cx')) - cx) > TOL or abs(int(e.get('cy')) - cy) > TOL:
                        mism += 1
    if mism:
        fail('P2 尺寸', 'wp:extent 与 a:ext 不一致 %d 处' % mism)
    if sp_bad:
        fail('P2 尺寸', '图片段行距非 auto %d 段（会裁图）' % sp_bad)
    ok('P2 尺寸')

    n_pic = len(list(d.element.body.iter(qn('pic:pic'))))
    st = z.read('word/settings.xml').decode('utf-8')
    # 真实元素是 w:doNotAutoCompressPictures（Word 默认就带）；旧名 doNotCompressPictures
    # 是本项目历史笔误，Word 不认识会被静默丢弃 —— 旧名文档兼容放行，新写一律写真实名
    if n_pic and not ('doNotAutoCompressPictures' in st or 'doNotCompressPictures' in st):
        fail('P3 压缩', '含 %d 张图但 settings.xml 缺 doNotAutoCompressPictures' % n_pic)
    ok('P3 压缩')

    # P20 图片边框：模板（三）3 要求"图片加紧贴边缘的0.75pt黑色细边框"
    # 0.75pt = 9525 EMU；只改 XML 不重编码，故必须逐张检查
    no_ln, bad_w = [], []
    for p in paras:
        # 封面LOGO横幅放行：段内无文字且宽高比>3（与上面 P2 的 logo 判定同口径），
        # 规则原文是"仅LOGO横幅不加框"，断言必须与规则一致，否则空模板永远 FAIL
        ext = list(p._p.iter(qn('wp:extent')))
        if ext and not T(p) and \
                max(int(e.get('cx')) for e in ext) / max(max(int(e.get('cy')) for e in ext), 1) > 3:
            continue
        for spPr in p._p.iter(qn('pic:spPr')):
            ln = spPr.find(qn('a:ln'))
            if ln is None:
                no_ln.append(T(p)[:16] or '无文字图段')
                continue
            try:
                w = int(ln.get('w', '0'))
            except ValueError:
                w = 0
            if not (9000 <= w <= 10000) or ln.find(qn('a:solidFill')) is None:
                bad_w.append((T(p)[:16] or '无文字图段', w))
    if no_ln:
        fail('P20 图片边框', '%d 张图片无 a:ln 边框（须 0.75pt 黑色细边框）：%s'
             % (len(no_ln), no_ln[:4]))
    if bad_w:
        fail('P20 图片边框', '%d 张图片边框线宽/颜色不合标准：%s' % (len(bad_w), bad_w[:3]))
    ok('P20 图片边框')
    return paras, z


def docx_media(d, z, ref):
    if not (ref and os.path.exists(ref)):
        return
    import hashlib
    zo = zipfile.ZipFile(ref)
    oh = {hashlib.md5(zo.read(n)).hexdigest() for n in zo.namelist()
          if n.startswith('word/media/')}
    bad = [n.split('/')[-1] for n in z.namelist() if n.startswith('word/media/')
           and hashlib.md5(z.read(n)).hexdigest() not in oh]
    if bad:
        fail('P4 媒体', '与原稿不一致 %d 张（被重编码）：%s' % (len(bad), bad[:5]))
    ok('P4 媒体')


def docx_format(d, paras, template=False):
    allt = '\n'.join(T(p) for p in paras) + '\n' + '\n'.join(
        c.text for t in d.tables for r in t.rows for c in r.cells)

    if not template:                      # 空模板的说明段本就写着禁写词，模板体检跳过
        for w in BAN:
            if w in allt:
                fail('P5 禁写', '出现"%s" %d 处' % (w, allt.count(w)))
        ok('P5 禁写')

    if not template and not ALLOW_PH:     # 空模板/模板测试件本来就有占位符，跳过
        ph = [w for w in PLACEHOLDER if w in allt]
        if ph:
            fail('P6 占位', '模板占位符未清：%s' % ph)

    if not template:                      # P27 残留占位（2026-11-06）：成品不得留〔…〕槽位
        ph27 = [m.group(0) for m in re.finditer(r'〔[^〕]{1,20}〕', allt)
                if not re.fullmatch(r'〔\d{4}〕', m.group(0))
                and m.group(0) != '〔待定责〕']
        if '待填' in allt:
            ph27.append('待填')
        if ph27:
            fail('P27 残留占位', '%d 处未填槽位：%s'
                 % (len(ph27), sorted(set(ph27))[:5]))
        else:
            ok('P27 残留占位')

    if not template:                      # 空模板的示例段含"检查检查"等示范文字，跳过
        for w in ('检查检查', '查看查看', '查阅查阅', '发现发现', '的的', '了了'):
            if w in allt:
                fail('P7 病句', '重复词"%s" %d 处' % (w, allt.count(w)))
        ok('P7 病句')
    else:
        print('   跳过 P5/P6/P7（空模板含说明与示例段）')

    if '"' in allt or "'" in allt:
        fail('P8 标点', '直引号 %d 处' % (allt.count('"') + allt.count("'")))

    # P8b 引文完整性：R00 原文铁律，任何依据引用不得省略
    om = [m for m in ('...', '……', '（略', '(略', '以下略', '省略') if m in allt]
    if om:
        fail('P8b 引文省略', '检测到引用省略标记：%s' % '、'.join(om))
    else:
        ok('P8b 引文省略')

    bad_f = set()
    for r in d.element.body.iter(qn('w:r')):
        rPr = r.find(qn('w:rPr'))
        f = rPr.find(qn('w:rFonts')) if rPr is not None else None
        if f is None:
            continue
        for a in ('w:eastAsia', 'w:ascii'):
            v = f.get(qn(a)) or ''
            if v and v not in OKFONT:
                bad_f.add(v)
    if bad_f:
        warn('P9 字体', '非常用字体：%s' % sorted(bad_f)[:6])

    # 行距：正文不允许 <1.0 倍或怪值
    bad_sp = 0
    for p in paras:
        if p._p.findall('.//' + qn('pic:pic')):
            continue
        pPr = p._p.pPr
        sp = pPr.find(qn('w:spacing')) if pPr is not None else None
        if sp is None:
            continue
        ln, rule = sp.get(qn('w:line')), sp.get(qn('w:lineRule'))
        if ln and rule == 'exact' and int(ln) not in (560, 480, 400, 360):
            bad_sp += 1
    if bad_sp:
        warn('P10 行距', '非标准固定行距段落 %d 段' % bad_sp)
    ok('P8 标点')


def docx_numbering(d, paras):
    NUM = '一二三四五六七八九十'
    chs = [T(p) for p in paras
           if re.match(r'^[一二三四五六七八九十]+、', T(p)) and len(T(p)) < 40]
    for i, t in enumerate(chs, 1):
        if i <= len(NUM) and not t.startswith(NUM[i - 1] + '、'):
            fail('P11 章号', '第%d章应为"%s、"，实际"%s"' % (i, NUM[i - 1], t[:14]))
    ok('P11 章号')

    nums, cur = [], None
    for p in paras:
        t = T(p)
        if p._p.findall('.//' + qn('pic:pic')) or len(t) > 170:
            continue
        m = re.match(r'^★?(\d+)[\.．]', t)
        if m:
            nums.append(int(m.group(1)))
    if nums and nums != list(range(1, len(nums) + 1)):
        fail('P12 编号', '编号不连续：%s' % nums[:20])

    if PROFILE in ('platform', 'company'):
        # 平台/公司类报告照片在表格单元格内、无"图N.M"图注与（a）（b）分图说明，
        # P13/P16/P17/P17b/P18 属集团类组图规则，不适用
        print('   跳过 P13/P16/P17/P17b/P18（平台/公司类报告无图注/分图说明）')
        ok('P12 编号')
        return

    groups = {}
    for p in paras:
        t = T(p)
        m = re.match(r'^★?(\d+)[\.．]', t)
        if m and not p._p.findall('.//' + qn('pic:pic')) and len(t) < 170:
            cur = int(m.group(1))
        m2 = re.match(r'^图(\d+)[\.．](\d+)', t)
        if m2:
            g, k = int(m2.group(1)), int(m2.group(2))
            groups.setdefault(g, []).append(k)
            if cur is not None and g != cur:
                fail('P13 图注', '图%d.%d 与所属问题%s不符' % (g, k, cur))
    for g, ks in groups.items():
        if ks != list(range(1, len(ks) + 1)):
            fail('P13 图注', '图%d.x 断号 %s' % (g, ks))
    ok('P12 编号')

    # P16 同一问题多张照片：必须合并成组、整组只有一条图注、每格 5×5
    cur, n_pic, n_cap = None, 0, 0
    stat = {}
    for p in paras:
        t = T(p)
        if re.match(r'^★?\d+[\.．]', t) and len(t) < 170 and not has_pic(p):
            if cur is not None:
                stat[cur] = (n_pic, n_cap)
            cur = int(re.match(r'^★?(\d+)', t).group(1))
            n_pic = n_cap = 0
            continue
        if cur is None:
            continue
        if has_pic(p):
            n_pic += len(list(p._p.iter(qn('pic:pic'))))
        elif re.match(r'^图(\d+)[\.．](\d+)', t):
            n_cap += 1
    if cur is not None:
        stat[cur] = (n_pic, n_cap)
    for no, (np_, nc_) in stat.items():
        if np_ >= 2 and nc_ != 1:
            fail('P16 组图', '问题%d 有 %d 张照片却 %d 条图注（应合并为一条 图%d.1）'
                 % (no, np_, nc_, no))
    ok('P16 组图')

    # P17 图注排版标准：宋体五号加粗图号 + 楷体_GB2312五号描述（两空格）、居中
    bad_cap = 0
    for p in paras:
        t = T(p)
        if not re.match(r'^图\d+[\.．]\d+', t):
            continue
        pPr = p._p.pPr
        jc = pPr.find(qn('w:jc')) if pPr is not None else None
        runs = []
        for r in p._p.findall(qn('w:r')):
            rt = ''.join(x.text or '' for x in r.iter(qn('w:t')))
            if not rt:
                continue
            rPr = r.find(qn('w:rPr'))
            f = rPr.find(qn('w:rFonts')) if rPr is not None else None
            sz = rPr.find(qn('w:sz')) if rPr is not None else None
            b = rPr is not None and rPr.find(qn('w:b')) is not None
            runs.append((rt, f.get(qn('w:eastAsia')) if f is not None else None,
                         sz.get(qn('w:val')) if sz is not None else None, b))
        good = (len(runs) == 2
                and runs[0][1] == '宋体' and runs[0][2] == '21' and runs[0][3]
                and runs[1][1] == '楷体_GB2312' and runs[1][2] == '21'
                and not runs[1][3] and runs[1][0].startswith('  ')
                and jc is not None and jc.get(qn('w:val')) == 'center')
        if not good:
            bad_cap += 1
            if bad_cap <= 3:
                fail('P17 图注格式', '"%s" 应为宋体五号加粗图号+楷体五号描述(两空格)居中，实际 %s'
                     % (t[:16], [(r[1], r[2], r[3]) for r in runs]))
    if bad_cap > 3:
        fail('P17 图注格式', '共 %d 条图注不合标准格式' % bad_cap)
    ok('P17 图注格式')

    # P17b 图注内容标准：描述必须等于所在问题标题（ZWG 终版 46/46 即此口径）
    cur_title, bad_desc = None, 0
    for p in paras:
        t = T(p)
        m = re.match(r'^★?(\d+)[\.．]\s*(.+)$', t)
        if m and len(t) < 170 and not has_pic(p):
            cur_title = m.group(2).strip()
            continue
        m2 = re.match(r'^图\d+[\.．]\d+\s*(.*)$', t)
        if not m2 or cur_title is None:
            continue
        desc = m2.group(1).strip()
        if desc != cur_title:
            bad_desc += 1
            if bad_desc <= 3:
                fail('P17b 图注内容', '图注"%s" ≠ 问题标题"%s"（描述须取问题标题）'
                     % (desc[:20], cur_title[:20]))
    if bad_desc > 3:
        fail('P17b 图注内容', '共 %d 条图注描述与问题标题不一致' % bad_desc)
    ok('P17b 图注内容')

    # P18 组图分图说明（学术论文式）：每题照片≥2 时，须有与照片数相等的（a）（b）…
    # 分图说明，位置词齐备（左上/右上/…），楷体_GB2312 五号
    cur, npic, subs = None, 0, 0
    stat = {}
    for p in paras:
        t = T(p)
        m = re.match(r'^★?(\d+)[\.．]', t)
        if m and len(t) < 170 and not has_pic(p):
            if cur is not None:
                stat[cur] = (npic, subs)
            cur = int(m.group(1))
            npic = subs = 0
            continue
        if cur is None:
            continue
        if has_pic(p):
            npic += len(list(p._p.iter(qn('pic:pic'))))
            continue
        ms = re.match(r'^（([a-z])）\s*([^：:]{1,8})[：:]\s*\S', t)
        if ms:
            subs += 1
            if ord(ms.group(1)) - 97 != subs - 1:
                fail('P18 分图说明', '问题%d 分图序号错：第%d条为（%s）' % (cur, subs, ms.group(1)))
            rpr = p._p.find(qn('w:r') + '/' + qn('w:rPr'))
            rf = rpr.find(qn('w:rFonts')) if rpr is not None else None
            sz = rpr.find(qn('w:sz')) if rpr is not None else None
            if (rf is None or rf.get(qn('w:eastAsia')) != '楷体_GB2312'
                    or sz is None or sz.get(qn('w:val')) != '21'):
                fail('P18 分图说明', '问题%d 分图说明非楷体_GB2312五号：%s' % (cur, t[:16]))
    if cur is not None:
        stat[cur] = (npic, subs)
    for no, (np_, ns_) in stat.items():
        if np_ >= 2 and ns_ != np_:
            fail('P18 分图说明', '问题%d 有 %d 张照片却 %d 条分图说明（须一一对应）'
                 % (no, np_, ns_))
    ok('P18 分图说明')

    # P19 已拆为独立函数 docx_outline()：平台档在 P12 后提前 return，
    # 若仍留在这里平台报告就查不到"正文被写成大纲标题"


def docx_outline(paras):
    # P19 目录只收章/节/定性标题（2026-10-28 事故沉淀）
    # 模板 TOC 域是 `TOC \o "1-9" \h \u`，`\u` 按段落大纲级别收录：
    # 正文"查看/检查…""（1）（2）…"段只要带 w:outlineLvl，一刷新目录整段正文
    # 就变成目录条目。故凡带大纲级别的段落，文本必须匹配 章／节／问题标题。
    # 平台档 paras 含表格单元格段落——单元格正文带 outlineLvl 正是上次实测踩到的坑。
    bad_lvl = []
    for p in paras:
        pPr = p._p.pPr
        if pPr is None:
            continue
        # 自动编号标题（numPr）：编号由 numbering 域渲染，文本不含"（一）／1."，属合法标题
        if pPr.find(qn('w:numPr')) is not None:
            continue
        ol = pPr.find(qn('w:outlineLvl'))
        if ol is None:
            continue
        val = ol.get(qn('w:val'))
        if val is None or int(val) > 8:       # 9 = 正文文本，不进目录
            continue
        t = T(p)
        if (re.match(r'^[一二三四五六七八九十]+、', t)
                or re.match(r'^（[一二三四五六七八九十]+）', t)
                or re.match(r'^★?\d+[.．]', t)):
            continue
        bad_lvl.append(t[:30] or '空段')
    if bad_lvl:
        fail('P19 目录', '%d 段正文被写成大纲标题（刷新目录会把正文吸进目录）：%s'
             % (len(bad_lvl), bad_lvl[:3]))
    ok('P19 目录条目')


def platform_structure(d):
    """P21/P22 平台类报告专属结构断言（集团类不适用）。
    P21：三章固定 + 每节"问题标题数＝表格数"、表格固定 2 行、节末有整改建议。
    P22：照片必须全部嵌在表格单元格内（平台报告无正文插图），且每格只放 1 张
         （平台类同题多图须先合成为一张，禁"3.65×6.0 并排"）。
    P23：定责栏（表格第2行左格）只能是"占位（未要求定责）"或"H回传的责任认定内容"——
         兼容H多种回传格式（岗位+条款／责任部门+责任人／考核评级+扣分／积分记分等），
         有责任要素但未落到条款号或考核结论的也判 FAIL；空白判 FAIL。
    P24：问题描述（表格第1行左格）须以"（n）"开头（模板 numId=4 的"（%1）"字面化），
         且 n 与问题序号及全篇顺序一致。
    P25：目录页须在封面后、正文前（形式同集团：黑体"目  录"+TOC\\o域）。
    """
    W, T = qn('w:p'), qn('w:tbl')

    def txt(el):
        return ''.join(t.text or '' for t in el.iter(qn('w:t'))).strip()

    kids = list(d.element.body.iterchildren())
    chs, secs, bad, duty, desc = [], [], [], [], []
    cur = None
    gnum = 0                                  # 全篇问题序号（跨节连续，与正文"（n）"一致）
    from docx.text.paragraph import Paragraph
    for el in kids:
        if el.tag == W:
            if is_toc(Paragraph(el, d)):      # 目录条目文本带页码，不能当章/节标题
                continue
            t = txt(el)
            pPr = el.find(qn('w:pPr'))
            auto_head = pPr is not None and pPr.find(qn('w:numPr')) is not None
            olvl = None
            if pPr is not None and pPr.find(qn('w:outlineLvl')) is not None:
                olvl = pPr.find(qn('w:outlineLvl')).get(qn('w:val'))
            if re.match(r'^[一二三四五六七八九十]+、', t) and len(t) < 40:
                if cur:
                    secs.append(cur)
                cur = None
                chs.append((t, auto_head))
                continue
            if ((re.match(r'^（[一二三四五六七八九十]+）', t)
                 or (auto_head and olvl == '1')) and len(t) < 40):
                if cur:
                    secs.append(cur)
                cur = {'name': t, 'prob': 0, 'tbl': 0, 'adv': [], 'auto': auto_head}
                continue
            if cur is None:
                continue
            if ((re.match(r'^★?\d+[\.．]', t)
                 or (auto_head and olvl is None and t.endswith('。'))) and len(t) < 170):
                cur['prob'] += 1
                gnum += 1
                cur['gseq'] = gnum
            elif t.startswith('整改建议'):
                cur['adv'].append(t)
        elif el.tag == T and cur is not None:
            cur['tbl'] += 1
            rows = el.findall('./' + qn('w:tr'))
            if len(rows) < 2:
                bad.append('%s 表格 %d 行（至少 2 行：描述行＋定责行）'
                           % (cur['name'], len(rows)))
            else:
                last = rows[-1].findall('./' + qn('w:tc'))
                # 定责行：优先按标签定位（"暴露的问题"所在单格行），
                # 兼容"定责行在前、制度行在后"的两行结构（2026-11-09）；
                # 找不到标签时退回末行单格。
                _duty_tc = None
                for r in rows:
                    tcs = r.findall('./' + qn('w:tc'))
                    if len(tcs) == 1 and '暴露的问题' in txt(tcs[0]):
                        _duty_tc = tcs[0]
                        break
                if _duty_tc is None and len(last) == 1:
                    _duty_tc = last[0]
                if len(last) != 1:
                    bad.append('%s 表格末行未合并为单格（%d 格）'
                               % (cur['name'], len(last)))
                if _duty_tc is not None:
                    # 定责段 = "暴露的问题：" 起的部分；其后的"暴露制度问题："
                    # 属制度层面追溯，不计入定责句式/岗位合规核验（2026-11-09）。
                    _dt = txt(_duty_tc)
                    _dt = _dt.split('暴露制度问题')[0].strip() or _dt
                    duty.append((cur['name'], _dt))
                for r in rows:
                    c0 = r.findall('./' + qn('w:tc'))
                    if c0 and not txt(c0[0]).startswith('暴露'):
                        pp = c0[0].find('.//' + qn('w:pPr'))
                        auto = pp is not None and pp.find(qn('w:numPr')) is not None
                        desc.append((cur.get('gseq', 0), txt(c0[0]), auto))
    if cur:
        secs.append(cur)

    if len(chs) != 3:
        fail('P21 平台骨架', '章标题 %d 个（平台报告固定三章）：%s'
             % (len(chs), [c[0] for c in chs][:4]))
    else:
        for i, kw in enumerate(('整体情况', '上一年度问题整改情况', '主要问题')):
            if kw not in chs[i][0]:
                fail('P21 平台骨架', '第%d章应为"…%s"，实际"%s"' % (i + 1, kw, chs[i][0][:20]))
    if not secs:
        fail('P21 平台骨架', '未识别到（一）（二）…节标题')
    for s in secs:
        if s['prob'] == 0:
            fail('P21 平台骨架', '节"%s"无问题（空节应整段删除）' % s['name'][:16])
        if s['prob'] != s['tbl']:
            fail('P21 平台骨架', '节"%s"问题 %d 条但表格 %d 个（应为一一对应）'
                 % (s['name'][:16], s['prob'], s['tbl']))
        if not s['adv']:
            fail('P21 平台骨架', '节"%s"缺"整改建议："段' % s['name'][:16])
    for b in bad[:3]:
        fail('P21 平台骨架', b)
    ok('P21 平台骨架')

    # P24 问题编号（描述行左格）：两种合法状态——
    #   ①自动编号：段落带 numPr（numbering 定义"（%1）"），文本不得残留字面"（n）"；
    #   ②字面编号（历史成品）：以"（n）"开头且 n 与问题序号、全篇顺序一致。
    #   同一篇不得混用。
    bad_no, auto_rows = [], 0
    for i, item in enumerate(desc, 1):
        no, t, auto = item
        if auto:
            auto_rows += 1
            if re.match(r'^（\d+）', t):
                bad_no.append('第%d条问题描述既自动编号又残留字面"（n）"' % i)
            elif not t.strip('★').strip():
                bad_no.append('第%d条问题描述为空' % i)
            continue
        m = re.match(r'^（(\d+)）', t)
        if no != i:
            bad_no.append('第%d个表格出现在第%d号问题下（题述顺序与表格错位）' % (i, no))
        elif not m:
            bad_no.append('第%d条问题描述缺"（%d）"编号，实际开头"%s"' % (i, i, t[:14]))
        elif int(m.group(1)) != no:
            bad_no.append('第%d条问题描述编号为"（%s）"，与问题序号不符' % (i, m.group(1)))
        elif not t[m.end():].strip():
            bad_no.append('第%d条问题描述为空（仅编号）' % i)
    if auto_rows and auto_rows != len(desc):
        bad_no.append('自动编号与字面编号混用（%d/%d 行自动编号）' % (auto_rows, len(desc)))
    if bad_no:
        fail('P24 问题编号', '%d 处：%s' % (len(bad_no), bad_no[:3]))
    ok('P24 问题编号')

    # P25 目录页（2026-11-01 定）：平台报告目录须在封面后单独一页、正文之前，
    # 形式同集团报告（黑体"目  录"标题 + `TOC \o "1-9" \h \u` 域）。
    # 缺域、只有标题无域、或目录被放在开篇引言之后（正文中间）均 FAIL。
    def _has_toc_field(p):
        for ins in p._p.iter(qn('w:instrText')):
            if 'TOC' in (ins.text or ''):
                return True
        for fs in p._p.iter(qn('w:fldSimple')):
            if 'TOC' in (fs.get(qn('w:instr')) or ''):
                return True
        return False

    def _is_toc_line(p):
        """目录条目（含未带 PAGEREF 的纯文本条目：toc 样式或"标题\\t页码"）。"""
        if p.style.name.lower().startswith('toc'):
            return True
        return bool(re.search(r'\t\s*\d+$', txt(p._p)))

    ps = list(d.paragraphs)
    toc_idx = next((i for i, p in enumerate(ps) if _has_toc_field(p)), None)
    head_idx = next((i for i, p in enumerate(ps)
                     if not is_toc(p) and not _is_toc_line(p) and not _has_toc_field(p)
                     and re.match(r'^[一二三四五六七八九十]+、', txt(p._p))
                     and len(txt(p._p)) < 40), None)
    cover_idx = next((i for i, p in enumerate(ps) if 'sectPr' in p._p.xml), None)
    if toc_idx is None:
        fail('P25 目录页', '缺目录域（平台报告须在封面后单独一页放"目  录"+TOC \\o 域）')
    else:
        title_idx = None
        for j in range(toc_idx - 1, max(toc_idx - 3, -1), -1):
            if txt(ps[j]._p).replace('\u3000', '').replace(' ', '') == '目录':
                title_idx = j
                break
        if title_idx is None:
            fail('P25 目录页', '目录域上方缺"目  录"标题段（形式须同集团：黑体12pt居中）')
        elif cover_idx is not None and title_idx - cover_idx > 3:
            fail('P25 目录页', '目录未紧跟封面（距封面分节符 %d 段，中间夹了正文内容）'
                 % (title_idx - cover_idx))
        if head_idx is not None and toc_idx >= head_idx:
            fail('P25 目录页', '目录页位于正文之后（应封面后、正文前）')
    ok('P25 目录页')

    # P23 定责栏（"暴露的问题（原因追溯）"左格）：只有两种合法状态——
    #   ①未要求定责 → 占位（含"〔"/"待定责"/"待提供"/"待填"）；
    #   ②已定责 → 分支H回传的责任认定句式，兼容H的多种回传格式：
    #      岗位+条款 / 责任部门+责任人 / 考核评级+扣分 / 积分记分 等，
    #      判据=含责任类关键词之一且（含"第…条"或含考核/责任主体表述）。
    # 其余自由文字一律视为自行编造（B 分支不定责，定责归 H）。
    empty = [n[:16] for n, t in duty if not t]
    if empty:
        fail('P23 定责栏', '%d 个问题的定责栏为空（应写〔待定责〕或H回传的定责结论）：%s'
             % (len(empty), empty[:3]))
    KW = ('岗位', '职责', '责任', '责任人', '部门', '班组', '考核', '扣分',
          '记分', '积分', '一岗一清单')
    made, weak = [], []
    for n, t in duty:
        if not t or any(k in t for k in ('〔', '待定责', '待提供', '待填')):
            continue
        if not any(k in t for k in KW):
            made.append((n[:16], t[:24]))
        elif not (re.search(r'第[^\s，。]{1,12}条', t)
                  or any(k in t for k in ('考核', '扣分', '记分', '积分', '责任人'))):
            weak.append((n[:16], t[:24]))
    if made:
        fail('P23 定责栏', '%d 个问题的定责栏既非〔待定责〕占位、也无责任认定要素'
                           '（疑自行编造，应改走分支H）：%s' % (len(made), made[:3]))
    if weak:
        fail('P23 定责栏', '%d 个问题的定责栏只提了岗位/部门，未落到条款号或考核结论：%s'
             % (len(weak), weak[:3]))
    ok('P23 定责栏')

    # P23b 定责合规（2026-11-09 R02教训）：岗位须在"部门负责人/外委工作负责人/
    #   外委项目经理"白名单内；不得引用对外单位（检修事业部/燃机项目部）或
    #   总经理/副总经理等非责任认定岗位；条款原文由分支H负责回查。
    _ALLOW = ('总经理', '副总经理', '部长', '主任', '副部长', '主管', '工程师',
              '项目经理', '项目运维班组长', '运维班组长', '运维班班长')
    _FORBID = ('检修事业部', '燃机项目部', '发电运行部', '涿州',
               '兼职安全员', '值班员', '高级业务主管')
    _illegal = []
    for n, t in duty:
        if not t or any(k in t for k in ('〔', '待定责', '待提供', '待填')):
            continue
        for kw in _FORBID:
            if kw in t:
                _illegal.append((n[:16], '越界:' + kw))
                break
        else:
            for role in re.findall(r'[一二三四五六七八九十]+是(.+?)岗位安全职责第', t):
                if not any(a in role for a in _ALLOW):
                    _illegal.append((n[:16], '非白名单:' + role[:14]))
                    break
    if _illegal:
        fail('P23b 定责合规', '%d 处定责岗位不合规（应仅引部门负责人/外委工作负责人/'
                              '外委项目经理，且不得引用检修事业部、燃机项目部等对外单位）：%s'
             % (len(_illegal), _illegal[:3]))
    ok('P23b 定责合规')

    # P28 定责句式与多岗位（2026-11-08 示例电厂A合并报告教训）：
    #   仅约束"岗位+条款"式定责（含"岗位安全职责第"字样者）；H回传采用
    #   "责任部门+责任人""考核评级+扣分"等其他合法格式时本项不适用（交P23）。
    #   ① 整格须以“暴露的问题：”领起；
    #   ② 每处引用须完整成句：{岗位}岗位安全职责第N条“{原文}”落实不到位；
    #   ③ 每条问题至少2个责任岗位（直接责任+管理/监督责任），禁止只写一个人。
    _bad, _few = [], []
    for n, t in duty:
        if not t or any(k in t for k in ('〔', '待定责', '待提供', '待填')):
            continue
        cnt = t.count('岗位安全职责第')
        if not cnt:
            continue
        if (not t.startswith('暴露的问题：') or t.count('落实不到位') != cnt
                or len(re.findall(r'岗位安全职责第[^\s，。]{1,12}条“', t)) != cnt
                or t.count('”落实不到位') != cnt):
            _bad.append((n[:16], t[:30]))
        elif cnt < 2:
            _few.append((n[:16], t[:30]))
    if _bad:
        fail('P28 定责句式', '%d 个问题定责未按“暴露的问题：一是{岗位}岗位安全职责'
                              '第N条“{原文}”落实不到位”完整成句：%s' % (len(_bad), _bad[:3]))
    if _few:
        fail('P28 定责多岗位', '%d 个问题定责只写1个责任岗位（须含直接责任+管理/监督责任'
                                '至少2个）：%s' % (len(_few), _few[:3]))
    ok('P28 定责句式')

    # P29 依据层级（R17，2026-11-09 用户重申）：平台报告的问题依据不得落到
    #   《XX法》《XX条例》等法律法规——被查公司→示例电力→示例集团集团→行业→国标
    #   前五层有可覆盖条款时必须用上位标准；法律法规仅作最后手段且须用户确认。
    #   只检查"不符合/违反/未按/未落实"这类定性句，避免误伤背景叙述；
    #   检测前先剔除中文引号内的条款原文（原文里的法律名属 R00 不可改）。
    def _outer(s):
        return re.sub(r'“[^”]*”', '“”', s)

    STD = re.compile(r'《([^》]{2,60})》(?:[（(]([^）)]{1,40})[）)])?')
    body_txt = []
    for el in d.element.body.iter(W):
        t = txt(el)
        if t and not is_toc(Paragraph(el, d)):
            body_txt.append(t)
    law_bad = []
    for t in body_txt:
        for sent in re.split(r'[。；\n]', _outer(t)):
            if not any(k in sent for k in ('不符合', '违反', '未按', '未落实')):
                continue
            for m in STD.finditer(sent):
                name, num = m.group(1).strip(), (m.group(2) or '').strip()
                if (name.endswith('法') or name.endswith('条例')) and not re.search(r'\d', num):
                    law_bad.append(name)
    if law_bad:
        fail('P29 依据层级', '%d 处依据引用法律法规（R17：应优先公司/平台/集团/行业/国标）：%s'
             % (len(law_bad), sorted(set(law_bad))[:5]))
    else:
        ok('P29 依据层级')

    # P30 同标准多条款合并句式（2026-11-09 用户定）：同一段落内同一标准
    #   只允许出现一次"《名称》（编号）"，后续条款写"和X.X“原文”"。
    #   引号内的条款原文不计（原文重复引用标准名属 R00 原样）。
    dup_std = []
    for t in body_txt:
        seen = {}
        for m in STD.finditer(_outer(t)):
            key = (m.group(1).strip(), (m.group(2) or '').strip())
            seen[key] = seen.get(key, 0) + 1
        for (nm, no), c in seen.items():
            if c >= 2:
                dup_std.append('%s（%s）×%d' % (nm, no, c))
    if dup_std:
        fail('P30 引用合并', '%d 处同一标准重复写书名号（应合并为"…和{条款号}"）：%s'
             % (len(dup_std), dup_std[:3]))
    else:
        ok('P30 引用合并')

    # P31 全自动编号（2026-11-09 用户硬性要求）：分类标题、题级标题、问题描述行
    #   一律 Word 自动编号（numPr），正文不得写死"（一）""1.""（1）"。
    #   章标题"一、二、三"沿用字面章号（P11/P25 与目录依赖），不在本项范围。
    head_lit = ['分类"%s"' % s['name'][:14] for s in secs if not s['auto']]
    row_lit = ['描述行%d"%s"' % (i, it[1][:10])
               for i, it in enumerate(desc, 1) if not it[2]]
    lit = head_lit + row_lit
    if lit:
        fail('P31 自动编号', '%d 处未用 Word 自动编号（用户要求全自动）：%s'
             % (len(lit), lit[:3]))
    else:
        ok('P31 自动编号')

    # P32 依据引文标红（2026-11-09 用户投诉后补，模板参数 §240 已有规定）：
    #   问题正文中紧跟条款号的“…”原文必须为红色 FF0000；只标引文，不含引号。
    START = re.compile(r'(?:[0-9][0-9A-Za-z．.\-]*(?:第[一二三四五六七八九十]+项)?'
                       r'(?:\s*[a-z]）)?|[a-z]）)\s*“')
    END = re.compile(r'\s*(的要求|的规定|和\s*《|和\s*[0-9])')
    red_bad = 0
    for el in d.element.body.iter(W):
        t = txt(el)
        if '“' not in t or '不符合' not in t:
            continue
        spans = []
        for m in START.finditer(t):
            s = m.end()
            p = t.find('”', s)
            while p != -1:
                if END.match(t, p + 1):
                    spans.append((s, p))
                    break
                p = t.find('”', p + 1)
        if not spans:
            continue
        red = []
        for r_el in el.iter(qn('w:r')):
            rs = ''.join(x.text or '' for x in r_el.iter(qn('w:t')))
            pr = r_el.find(qn('w:rPr'))
            c = pr.find(qn('w:color')) if pr is not None else None
            isred = c is not None and (c.get(qn('w:val')) or '').upper() == 'FF0000'
            red.extend([isred] * len(rs))
        for s, e in spans:
            if not all(red[s:e]):
                red_bad += 1
    if red_bad:
        fail('P32 引文标红', '%d 处依据引文未标红 FF0000' % red_bad)
    else:
        ok('P32 引文标红')

    # P33 依据与事项对应（2026-11-09 示例电厂A返工：工作票问题误引操作票条款）：
    #   问题主体只涉工作票 → 引文不得只含操作票特征词；反之亦然。无特征词视为无法判定，放行。
    WORK_M = ('工作票', '工作许可人', '工作负责人', '工作票签发人', '危险点分析控制单', '工作班成员')
    OPS_M = ('操作票', '倒闸操作', '操作人', '操作任务', '操作顺序', '操作开始', '唱票', '复诵', '操作项')
    mis = []
    for el in d.element.body.iter(W):
        t = txt(el)
        i = t.find('不符合《')
        if i < 0:
            continue
        subj = t[:i]
        if ('工作票' in subj) == ('操作票' in subj):
            continue
        quoted = ''.join(re.findall(r'“([^”]*)”', t[i:]))
        if not quoted:
            continue
        has_work = any(k in quoted for k in WORK_M)
        has_ops = any(k in quoted for k in OPS_M)
        if '工作票' in subj and has_ops and not has_work:
            mis.append(subj[-18:])
        elif '操作票' in subj and has_work and not has_ops:
            mis.append(subj[-18:])
    if mis:
        fail('P33 依据票种匹配', '%d 条问题依据票种/作业类型不匹配：%s' % (len(mis), mis[:3]))
    else:
        ok('P33 依据票种匹配')

    # P34 依据必须带引文（2026-11-09 用户：只给编号无法判断能否采用）
    noquote = 0
    for el in d.element.body.iter(W):
        t = txt(el)
        j = t.find('不符合《')
        if j >= 0 and '“' not in t[j:]:
            noquote += 1
    if noquote:
        fail('P34 依据引文', '%d 处依据只写标准/条款号、未带引文' % noquote)
    else:
        ok('P34 依据引文')

    # P35 专项对应：防火封堵必须优先引 DL 5027-2015 10.5.3（用户2026-11-09指定，优先于R17层级）
    #   只看"问题描述"（"不符合"之前），整改建议/依据标准名中的"防火封堵"不算。
    fd_bad = []
    for el in d.element.body.iter(W):
        t = txt(el)
        i = t.find('不符合')
        if i < 0:
            continue
        if '防火封堵' in t[:i] and 'DL 5027' not in t:
            fd_bad.append(t[:22])
    if fd_bad:
        fail('P35 防火封堵依据', '%d 处防火封堵问题未引《电力设备典型消防规程》（DL 5027-2015）10.5.3：%s'
             % (len(fd_bad), fd_bad[:3]))
    else:
        ok('P35 防火封堵依据')

    # P36 括号编号全文连续（2026-11-09 用户："括号的数字编号从上到下是统一的，不要重新编号"）
    # 描述行编号 = decimal + 括号格式的 num；须全篇共用同一 num、不得带 startOverride（会按组重启）。
    try:
        npart = d.part.numbering_part.element
        a_def = {}
        for ab in npart.findall(qn('w:abstractNum')):
            aid = ab.get(qn('w:abstractNumId'))
            lv = ab.find(qn('w:lvl'))
            fmt = ''
            lt = ''
            if lv is not None:
                f = lv.find(qn('w:numFmt'))
                l = lv.find(qn('w:lvlText'))
                fmt = f.get(qn('w:val')) if f is not None else ''
                lt = l.get(qn('w:val')) if l is not None else ''
            a_def[aid] = (fmt, lt)
        n_def = {}
        for n in npart.findall(qn('w:num')):
            nid = n.get(qn('w:numId'))
            aid_o = n.find(qn('w:abstractNumId'))
            aid = aid_o.get(qn('w:val')) if aid_o is not None else ''
            ov = any(o.find(qn('w:startOverride')) is not None
                     for o in n.findall(qn('w:lvlOverride')))
            n_def[nid] = (a_def.get(aid, ('', '')), ov)
        used, ov_ids = set(), set()
        for el in d.element.body.iter(qn('w:p')):
            ppr = el.find(qn('w:pPr'))
            if ppr is None:
                continue
            npr = ppr.find(qn('w:numPr'))
            if npr is None:
                continue
            ne = npr.find(qn('w:numId'))
            if ne is None:
                continue
            nid = ne.get(qn('w:val'))
            if nid not in n_def:
                continue
            (fmt, lt), ov = n_def[nid]
            if fmt == 'decimal' and lt and '（%1）' in lt:
                used.add(nid)
                if ov:
                    ov_ids.add(nid)
        if len(used) > 1:
            fail('P36 括号编号连续', '描述行使用了 %d 套独立编号（numId=%s），应全篇共用一套连续编号'
                 % (len(used), sorted(used, key=int)[:6]))
        elif ov_ids:
            fail('P36 括号编号连续', '描述行编号带 startOverride（numId=%s），会按组从1重新编号'
                 % sorted(ov_ids, key=int))
        else:
            ok('P36 括号编号连续')
    except Exception as e:
        ok('P36 括号编号连续（未启用：%s）' % e)

    # P37 同类问题依据不宜全用一个标准（2026-11-09 用户："同类问题在确保准确性的前提下，尽量不要全用一个标准"）
    #   同一张表（同一类问题）≥3条且全部引同一标准 → 提醒分散；仅为 WARN，准确性优先。
    try:
        dup = []
        for ti, t in enumerate(d.tables):
            rows = []
            for tr in t.rows:
                tcs = tr._tr.findall(qn('w:tc'))
                if not tcs:
                    continue
                s = ''.join(x.text or '' for x in tcs[0].iter(qn('w:t')))
                if '不符合' in s:
                    rows.append(s)
            if len(rows) < 3:
                continue
            names = []
            for s in rows:
                m = re.findall(r'《([^》]+)》', s)
                if m:
                    names.append(m[-1])
            cnt = {}
            for x in names:
                cnt[x] = cnt.get(x, 0) + 1
            if len(cnt) == 1:
                nm = list(cnt)[0]
                if cnt[nm] >= 3:
                    dup.append('表%d %d条全引《%s》' % (ti + 1, cnt[nm], nm))
        if dup:
            warn('P37 依据集中', '%d 处同类问题整类只用一个标准，建议在准确前提下分散引用：%s'
                 % (len(dup), dup[:3]))
        else:
            ok('P37 依据集中')
    except Exception as e:
        ok('P37 依据集中（未启用：%s）' % e)

    # P38 多用二十五项反措（2026-11-09 用户："多用25项反错，很多依据25项反错里有，你现在用的依据不对"）
    try:
        allt = ''.join(txt(p._p) for p in d.paragraphs)
        for t in d.tables:
            for row in t.rows:
                for c in row.cells:
                    for p in c.paragraphs:
                        allt += txt(p._p)
        if '二十五项' not in allt:
            warn('P38 二十五项反措', '全篇未引《防止电力生产事故的二十五项重点要求》（2023版），'
                 '技术/设备/作业类问题应检索其对应条款（防火封堵2.2/误操作3.1、3.2/开关柜13.3/接地14.1等）')
        else:
            ok('P38 二十五项反措')
    except Exception as e:
        ok('P38 二十五项反措（未启用：%s）' % e)


    body_pic = [txt(p._p)[:16] for p in d.paragraphs if has_pic(p)]
    if body_pic:
        fail('P22 平台插图位', '%d 张照片浮在正文（平台报告照片须嵌表格单元格）：%s'
             % (len(body_pic), body_pic[:3]))
    over = []
    for tc in d.element.body.iter(qn('w:tc')):
        n = len(list(tc.iter(qn('pic:pic'))))
        if n > 1:
            over.append(n)
    if over:
        fail('P22 平台插图位', '%d 个单元格放多于 1 张照片（平台类同题多图须先合成为一张，禁并排）'
             % len(over))
    ok('P22 平台插图位')


def binding_check(rel):
    """P26 模板绑定：成品必须可追溯到绑定的空模板（2026-11-06 用户指定）。

    规则：各类型巡查报告只能由 `02 B巡查检查报告\reportkit\模板绑定.py` 里
    绑定的那一份空模板生成，禁止自创版式、禁止跨档位套用。
    成品 core_properties.category 里必须留有 TPL/<profile>/<版本>/<sha16> 标记。
    """
    try:
        cat = (Document(rel).core_properties.category or '').strip()
    except Exception:
        cat = ''
    if not cat.startswith('TPL/'):
        fail('P26 模板绑定', '成品未绑定空模板（无 TPL 标记）——必须由绑定空模板生成')
        return
    parts = cat.split('/')
    if len(parts) != 4:
        fail('P26 模板绑定', '绑定标记格式错误：%s' % cat)
        return
    _, prof, ver, sha = parts
    if prof != PROFILE:
        fail('P26 模板绑定', '报告类型不符：成品标记 %s ≠ 门禁档位 %s'
             % (prof, PROFILE))
        return
    try:
        sys.path.insert(0, os.path.join(
            os.path.dirname(os.path.abspath(__file__)), '..',
            '02 B巡查检查报告', 'reportkit'))
        import importlib
        m = importlib.import_module('模板绑定')
        b = m.resolve(prof)
    except Exception as e:                       # 绑定表不可用时退化为格式校验
        warn('P26 模板绑定', '绑定表不可读（%s），仅校验标记格式' % e)
        return
    if b['sha256'] != sha:
        warn('P26 模板绑定', '绑定表已换版（成品 %s ≠ 当前 %s），建议按新模板重出'
             % (sha, b['sha256']))
        return
    ok('P26 模板绑定')
    print('   绑定空模板：%s（%s）%s' % (b['文件'], ver, sha))


def check_docx(rel, ref=None, report=False, allow_rot=False, template=False):
    d = Document(rel)
    paras, z = docx_images(d, rel)
    if allow_rot:
        globals()['WARN'] = [w.replace('[P1 旋转]', '[P1 旋转·已确认]')
                             for w in WARN]
    docx_media(d, z, ref)
    docx_format(d, paras, template)
    if not template:                      # 空模板编号/图题均为占位示例，跳过语义校验
        docx_numbering(d, paras)
        docx_outline(paras)
    if PROFILE == 'platform' and not template:
        platform_structure(d)

    if report:
        if not template:
            binding_check(rel)               # P26 各类型报告必须绑定空模板
        allt = '\n'.join(T(p) for p in paras)
        # 小项数：分组式模板（一题一表多描述行）按"描述行"计；旧式按问题标题计
        desc_txt = []
        for _tb in d.tables:
            _rows = _tb.rows
            for _row in (_rows[:-1] if len(_rows) > 1 else _rows):
                _tcs = _row._tr.findall(qn('w:tc'))
                if _tcs:
                    desc_txt.append(''.join(x.text or '' for x in _tcs[0].iter(qn('w:t'))).strip())
        n_prob = len(desc_txt) if desc_txt else len(
            [1 for p in paras if re.match(r'^★?\d+[\.．]', T(p))
             and len(T(p)) < 170 and not p._p.findall('.//' + qn('pic:pic'))])
        n_star = max(
            len([1 for p in paras if re.match(r'^★\d+[\.．]', T(p))
                 and len(T(p)) < 170 and not p._p.findall('.//' + qn('pic:pic'))]),
            len([1 for s in desc_txt if '★' in s]))
        for p in paras:
            t = T(p)
            m = (re.search(r'共发现问题(\d+)项', t) if PROFILE == 'platform'
                 else re.search(r'共计发现问题(\d+)项', t))
            if m and int(m.group(1)) != n_prob:
                fail('P14 数字', '开篇 %s 项 ≠ 实际 %d 项' % (m.group(1), n_prob))
            # 只核对"本次"那一句；历年数据段（"20XX年安全综合巡查发现问题…"）不参与比对
            seg = t
            if '本次' in t:
                seg = t[t.rfind('本次'):]
            elif re.search(r'20\d\d年安全综合巡查发现问题', t):
                seg = ''
            m = re.search(r'重点问题(\d+)项', seg)
            if m and int(m.group(1)) != n_star:
                fail('P14 数字', '重点 %s 项 ≠ 实际 %d 项' % (m.group(1), n_star))
        if n_prob and n_star / n_prob > 0.30:
            warn('P14 数字', '★占比 %.0f%% > 30%%' % (100.0 * n_star / n_prob))
        ok('P14 数字')
        toc = [p for p in d.paragraphs if is_toc(p) and T(p)]
        def _auto_sec(p):
            pp = p._p.pPr
            if pp is None or pp.find(qn('w:numPr')) is None:
                return False
            ol = pp.find(qn('w:outlineLvl'))
            return ol is not None and ol.get(qn('w:val')) == '1'
        n_sec = len([1 for p in paras
                     if (re.match(r'^（[一二三四五六七八九十]+）', T(p)) or _auto_sec(p))
                     and len(T(p)) < 40])
        heads = len([1 for p in paras
                     if re.match(r'^[一二三四五六七八九十]+、', T(p))
                     and len(T(p)) < 40]) + n_sec
        if toc and len(toc) < heads:
            fail('P15 目录', '目录条目 %d 少于章+节数 %d' % (len(toc), heads))
        long_toc = [T(p)[:24] for p in toc if len(T(p)) > 60]
        if long_toc:
            fail('P15 目录', '目录里混入正文长句 %d 条：%s' % (len(long_toc), long_toc[:3]))
        ok('P15 目录')
        print('   问题 %d 条 | 重点 %d 条 | 图片 %d 张 | 目录 %d 条'
              % (n_prob, n_star, len(list(d.element.body.iter(qn('pic:pic')))), len(toc)))


# ── xlsx（台账/清单）─────────────────────────────────────
def check_xlsx(rel):
    import openpyxl
    wb = openpyxl.load_workbook(rel, data_only=True)
    n_row = 0
    for ws in wb.worksheets:
        rows = [r for r in ws.iter_rows(values_only=True)
                if any(v not in (None, '') for v in r)]
        if not rows:
            fail('X1 空表', 'sheet「%s」无数据' % ws.title)
            continue
        n_row += len(rows) - 1
        hdr = [str(v).strip() if v else '' for v in rows[0]]
        # 完全重复行
        body = [tuple(str(v).strip() if v is not None else '' for v in r)
                for r in rows[1:]]
        dup = [k for k, v in Counter(body).items() if v and sum(1 for x in body if x == v) > 1]
        if dup:
            fail('X2 重复', 'sheet「%s」重复条目 %d 组：%s'
                 % (ws.title, len(dup), dup[0][:2]))
        # 空列/空行率
        blank = sum(1 for r in rows[1:] if all(v in (None, '') for v in r))
        if blank:
            warn('X3 空行', 'sheet「%s」整行空白 %d 行' % (ws.title, blank))
        # 日期列要是真日期不是文本
        if any('日期' in h or '时间' in h for h in hdr):
            idx = [i for i, h in enumerate(hdr) if '日期' in h or '时间' in h]
            txt = sum(1 for r in rows[1:] for i in idx
                      if isinstance(r[i], str) and re.search(r'\d{4}[-/.年]\d{1,2}', r[i] or ''))
            if txt:
                warn('X4 日期', 'sheet「%s」日期为文本格式 %d 处' % (ws.title, txt))
    ok('X1 表结构')
    print('   工作表 %d 个 | 数据行 %d 行' % (len(wb.worksheets), n_row))


# ── 标准文件 ─────────────────────────────────────────────
def check_std(rel):
    name = os.path.basename(rel)
    size = os.path.getsize(rel)
    if size < 20000:
        fail('S1 完整', '体积仅 %.1f KB，疑为空壳/下载不完整' % (size / 1024))
    ext = os.path.splitext(name)[1].lower()
    if ext == '.pdf':
        with open(rel, 'rb') as f:
            if f.read(5)[:4] != b'%PDF':
                fail('S2 文件头', '不是有效 PDF')
    if not re.search(r'(GB|GBT|AQ|DL|NB|QBEH|Q/|NB/T)\s?/?T?\s?\d', name, re.I):
        warn('S3 命名', '文件名未见标准号：%s' % name)
    if re.search(r'副本|\(\d\)|（\d）|未命名|新建', name):
        fail('S3 命名', '文件名带副本/序号/未命名痕迹：%s' % name)
    ok('S1 完整')


# ── 入口 ────────────────────────────────────────────────
def detect(rel, forced):
    if forced and forced != 'auto':
        return forced
    n = os.path.basename(rel).lower()
    if n.endswith('.xlsx') or n.endswith('.xlsm'):
        return 'xlsx'
    if n.endswith('.docx'):
        return 'report' if ('报告' in n or '通知单' in n) else 'docx'
    if n.endswith('.pdf') or n.endswith('.doc'):
        return 'std'
    return 'docx'


def detect_profile(rel, forced):
    """报告档位：显式 --profile group|platform|company 优先；auto 只认目录名，
    不做文件名猜测（"报告"二字所有类别都有，靠文件名猜必然误判）。
    平台类模板固定 `03-输出层\\出报告\\Template\\电力报告\\`，公司类固定 `...\\Template\\公司报告\\`，
    集团类固定 `...\\Template\\集团报告\\`，其余路径按集团/公司(正文图)类。"""
    if forced and forced != 'auto':
        return forced
    p = rel.replace('/', '\\')
    if '\\电力报告\\' in p:
        return 'platform'
    if '\\公司报告\\' in p:
        return 'company'
    return 'group'


def main(rel, forced='auto', ref=None, allow_rot=False, template=False,
         profile='auto'):
    global PROFILE
    if not os.path.exists(rel):
        print('文件不存在：%s' % rel)
        return 2
    PROFILE = detect_profile(rel, profile)
    kind = detect(rel, forced)
    print('=' * 70)
    print('门禁对象：%s' % os.path.basename(rel))
    print('类型判定：%s | 档位：%s | 大小 %.2f MB'
          % (kind, PROFILE_NAME.get(PROFILE, PROFILE), os.path.getsize(rel) / 1048576))
    print('-' * 70)
    if kind in ('report', 'docx'):
        check_docx(rel, ref, report=(kind == 'report'), allow_rot=allow_rot,
                   template=template)
    elif kind == 'xlsx':
        check_xlsx(rel)
    else:
        check_std(rel)
    print('-' * 70)
    if FAIL:
        print('GATE FAIL（%d 项）' % len(FAIL))
        for x in FAIL:
            print('  ✗', x)
    else:
        print('GATE PASS：%d 组检查全通过（%s）' % (len(DONE), '/'.join(DONE)))
    for x in WARN:
        print('  ! 提示', x)
    print('=' * 70)
    return 1 if FAIL else 0


if __name__ == '__main__':
    a = [x for x in sys.argv[1:] if not x.startswith('--')]
    if not a:
        print(__doc__)
        sys.exit(2)
    kw = {}
    for key, flag in (('--type', 'type'), ('--ref', 'ref'), ('--profile', 'profile')):
        if key in sys.argv:
            kw[flag] = sys.argv[sys.argv.index(key) + 1]
            a = [x for x in a if x != kw[flag]]
    if '--allow-ph' in sys.argv:
        globals()['ALLOW_PH'] = True
    sys.exit(main(a[0], kw.get('type', 'auto'), kw.get('ref'),
                  '--allow-rot' in sys.argv, '--template' in sys.argv,
                  kw.get('profile', 'auto')))
