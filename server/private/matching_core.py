# -*- coding: utf-8 -*-
"""
示例电力制度归属匹配 —— 核心判断逻辑

判断某个问题理论上应归属示例电力哪个企业标准/制度编号，读取
《示例电力企业标准体系目录.xlsx》做关键词匹配，命中则"已确认"，
未命中则退化为按"所属大类"的推断模式。

同义词扩展库是这份代码里最值得持续积累、调优的部分——随着实际使用中出现的
"关键词没对上但其实该匹配"的案例增多，往这个词典里加同义词对，匹配准确度
会持续提升，这是外部拿不到你实际使用数据就无法复现的部分。
"""

import os
import re

try:
    import openpyxl
except ImportError:
    openpyxl = None

# ============================================================
# 同义词扩展库（可持续积累）
# 格式：{标准写法: [同义/近义表达, ...]}
# ============================================================
SYNONYM_MAP = {
    "培训考核": ["培训", "考核", "教育培训", "培训频次", "考核频次", "培训要求"],
    "安全文化": ["安全意识", "安全氛围", "安全宣传"],
    "承包商": ["外委", "外包", "施工单位", "承包单位"],
    "应急处置": ["应急预案", "应急演练", "事故处置", "应急响应"],
    "设备检维修": ["设备维护", "设备检修", "设备保养", "检维修"],
    "风险辨识": ["风险评估", "危险源辨识", "隐患排查"],
    "高处作业": ["登高作业", "高空作业"],
    "带电作业": ["带电检修", "带电操作"],
    "职业健康": ["职业病防护", "劳动防护"],
}


def _shared_form_len(keyword: str, question: str) -> int:
    """返回关键词与问题在同一同义词组内共同出现的字面词的最大长度，无交集返回 0。

    必须比较「字面词」而不是同义词组名：否则「绩效考核」与「考核」会被整组判为命中，
    把只重合了两个字的匹配算成长词命中。
    """
    best = 0
    for standard, synonyms in SYNONYM_MAP.items():
        forms = [standard] + synonyms
        hit_forms = [f for f in forms if f in keyword and f in question]
        if hit_forms:
            best = max(best, max(len(f) for f in hit_forms))
    return best


# 返回条数上限与弱命中阈值：目录关键词里「考核」「管理」这类两字泛词命中面很宽，
# 全部返回会让调用方在十几条候选里自己筛。按关键词长度加权后只返回最相关的前几条。
TOP_N = 5
WEAK_SCORE = 1.5


def _kw_weight(keyword: str, exact: bool) -> float:
    """关键词打分：越长越具体，权重越高；经同义词间接命中的权重减半。"""
    return _weight_of_len(len(keyword), exact)


def _weight_of_len(n: int, exact: bool) -> float:
    if n >= 5:
        w = 3.0
    elif n == 4:
        w = 2.5
    elif n == 3:
        w = 1.6
    elif n == 2:
        w = 0.7
    else:
        w = 0.3
    return w if exact else w * 0.5


def _read_directory(xlsx_path: str):
    """读取企业标准体系目录，返回记录列表。文件不存在或openpyxl未安装时返回None。"""
    if openpyxl is None:
        return None
    if not os.path.exists(xlsx_path):
        return None
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        ws = wb["制度目录"] if "制度目录" in wb.sheetnames else wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
        header = [str(h).strip() if h else "" for h in rows[0]]

        def col_idx(name_contains):
            for i, h in enumerate(header):
                if name_contains in h:
                    return i
            return None

        idx_code = col_idx("制度编号")
        idx_name = col_idx("制度名称")
        idx_category = col_idx("所属大类")
        idx_keywords = col_idx("覆盖主题关键词")

        records = []
        for row in rows[1:]:
            if not row or not row[0]:
                continue
            code = str(row[idx_code]).strip() if idx_code is not None and row[idx_code] else ""
            name = str(row[idx_name]).strip() if idx_name is not None and row[idx_name] else ""
            category = str(row[idx_category]).strip() if idx_category is not None and row[idx_category] else ""
            keywords_raw = str(row[idx_keywords]).strip() if idx_keywords is not None and row[idx_keywords] else ""
            if not code or not name:
                continue
            keywords = [k.strip() for k in re.split(r"[,，、]", keywords_raw) if k.strip()]
            records.append({
                "code": code, "name": name, "category": category, "keywords": keywords,
            })
        return records
    except Exception:
        return None


def match_regulation(question: str, xlsx_path: str) -> dict:
    """制度归属匹配主入口。

    参数：
        question: 用户原始问题（企业标准层"完全空白"场景下才调用本函数做推断）
        xlsx_path: 《示例电力企业标准体系目录.xlsx》的完整路径

    返回：{"matches": [{"code", "name", "confidence", "matched_keyword", "score"}],
           "directory_found": bool}
        按关键词具体程度加权排序，最多返回 TOP_N 条；命中条数多于 TOP_N 时
        note 会说明总命中数，不静默丢弃信息。
    """
    records = _read_directory(xlsx_path)

    if records is None:
        return {
            "matches": [],
            "directory_found": False,
            "note": "未找到制度体系目录文件，匹配准确度将持续偏低，建议尽快补充该目录",
        }
    if not records:
        return {"matches": [], "directory_found": True, "note": "目录文件存在但没有有效记录"}

    question = question or ""

    scored = []
    for rec in records:
        matched_kw = []
        exact_score = 0.0
        indirect_best = 0.0
        for kw in rec["keywords"]:
            if not kw:
                continue
            if kw in question:
                matched_kw.append(kw)
                exact_score += _kw_weight(kw, True)
                continue
            shared = _shared_form_len(kw, question)
            if shared >= 2 and shared < len(kw):
                matched_kw.append(kw)
                # 同义词是「旁证」：只取最强的一条计分，且按实际重合词的长度算，
                # 避免「考核」「培训」这类两字泛词通过大量同义关键词把分数堆高。
                indirect_best = max(indirect_best, _weight_of_len(shared, False))
        score = exact_score + indirect_best
        if matched_kw:
            weak = score < WEAK_SCORE
            scored.append({
                "code": rec["code"], "name": rec["name"],
                "confidence": "已确认（弱匹配，建议人工核实）" if weak else "已确认",
                "matched_keyword": "、".join(matched_kw),
                "score": round(score, 1),
            })

    if scored:
        scored.sort(key=lambda m: (m["score"], len(m["matched_keyword"].split("、"))), reverse=True)
        if scored[0]["score"] < WEAK_SCORE:
            # 全部是「考核」「管理」这类泛词命中：只给 3 条并明说不可靠，避免凑数
            return {
                "matches": scored[:3],
                "directory_found": True,
                "note": "仅关键词弱命中，不足以确认归属，建议结合制度目录人工核实",
            }
        result = {"matches": scored[:TOP_N], "directory_found": True}
        if len(scored) > TOP_N:
            result["total_hits"] = len(scored)
            result["note"] = f"共 {len(scored)} 条制度命中关键词，已按相关度返回前 {TOP_N} 条"
        return result

    inferred = []
    for rec in records:
        if rec["category"] and rec["category"] in question:
            inferred.append({
                "code": rec["code"], "name": rec["name"],
                "confidence": "推断，建议人工核实所属制度编号",
                "matched_keyword": f"（大类匹配：{rec['category']}）",
            })
    if inferred:
        return {"matches": inferred[:3], "directory_found": True}

    return {
        "matches": [],
        "directory_found": True,
        "note": "关键词与所属大类均未命中任何制度，建议人工核实或补充目录关键词列",
    }
