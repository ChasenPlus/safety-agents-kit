# -*- coding: utf-8 -*-
"""
R11 覆盖充分性核验 —— 核心判断逻辑

判定原则：候选条款要判定为"通过核验"，必须针对【问题实际在问的那个维度】给出明确、
具体、可比对的规定；只是同一大类下的泛泛提及（"应加强""应重视""应定期"这类），
即使关键词命中，也判定为"未通过核验"。

五个核验维度：具体指标 / 频次 / 责任主体 / 适用范围 / 操作细节

实现分两层：
1. 启发式规则层（默认，无需任何外部依赖，离线可用）：用正则+关键词库判断问题问的是
   哪个/哪些维度、条款是否针对该维度给出了具体规定。
2. 可选LLM精化层（设置了 ANTHROPIC_API_KEY 环境变量时启用）：启发式判断处于边界情况
   时，调用大模型做更细致的语义判断，结果与启发式结果不一致时以LLM判断为准，并记录在
   judgment log里供后续复核。

这是本系统真正的核心资产：判断准确度会随着你在实际使用中积累的judgment log增长而
可以持续调优（调整下面的关键词库、正则模式），这是别人拿不到你的数据就复现不了的部分。
"""

import re
import os
import json
import datetime
from pathlib import Path

# ============================================================
# 维度关键词库（可持续扩充，这是随使用积累会变得更准的部分）
# ============================================================

QUESTION_DIMENSION_KEYWORDS = {
    "频次": ["频次", "多久", "周期", "间隔", "几次", "每隔", "多长时间", "间隔多久",
             "多少次", "周期性", "定期的具体", "多久一次"],
    "具体指标": ["标准是多少", "指标", "数值", "限值", "阈值", "多少", "达到什么标准",
                "参数", "规格", "尺寸", "温度", "压力", "浓度", "距离"],
    "责任主体": ["谁来", "责任人", "适用于", "哪些人", "承包商", "外委", "谁负责",
                "由谁", "哪个部门", "岗位职责", "谁承担"],
    "适用范围": ["范围", "哪些场景", "是否适用", "适不适用", "什么情况下", "哪些情形",
                "涵盖", "包括哪些"],
    "操作细节": ["怎么做", "如何", "步骤", "方法", "流程", "具体操作", "怎么操作",
                "如何执行", "操作规程"],
}

FREQ_CONCRETE_PATTERN = re.compile(
    r"(每[一二三四五六七八九十\d]*[个]?(年|季度|月|周|星期|日|天|小时))"
    r"|(\d+\s*(次/年|次/月|次/季|次/周|次/日))"
    r"|(至少\s*\d+\s*次)"
    r"|(\d+\s*(年|月|周|日|天)\s*(一次|一回))"
)
FREQ_VAGUE_ONLY_PATTERN = re.compile(r"(定期|适时|及时|经常|不定期|按需)")

# 2026-10-29 修正：原单位白名单缺 % 与 倍，"不应低于公称工作压力的90%""不超过2倍"
# 这类百分比/倍数阈值条款被误判为"未给出具体数值"（AQ 6110-2025 7.1/7.4 实测失败）。
# 补 % ％ 倍，以及比较词后接百分比的写法（"低于…90%""超过…50%"）。
METRIC_CONCRETE_PATTERN = re.compile(
    r"(不低于|不超过|不小于|不大于|至少|最多|≥|≤|大于|小于|低于|高于|超过|达到|不得超过|不得低于)"
    r"\s*[\d.]+\s*[%％倍]?"
    r"|[\d.]+\s*(米|毫米|厘米|千米|伏|千伏|安|千安|瓦|千瓦|度|摄氏度|℃|兆帕|帕|吨|公斤|千克|%|％|倍)"
)

SUBJECT_CONCRETE_KEYWORDS = [
    "承包商", "外委", "外包", "新员工", "新入职", "班组长", "专责人", "监护人",
    "特种作业人员", "持证人员", "作业负责人", "安全员", "项目经理", "部门负责人",
]
SUBJECT_VAGUE_ONLY_PATTERN = re.compile(r"^(全员|人员|员工|相关人员|有关人员)$")

DETAIL_CONCRETE_PATTERN = re.compile(
    r"(第[一二三四五六七八九十\d]+步)"
    r"|(首先.*?然后|首先.*?其次)"
    r"|(应当?按下列(步骤|顺序|程序))"
    r"|([①②③④⑤⑥⑦⑧⑨])"
    r"|(\d+[\.\、]\s*\S+.*?\d+[\.\、])"
)

VAGUE_VERBS_ONLY = re.compile(r"^(应加强|应重视|应做好|应保证|应确保)")


def _detect_question_dimensions(question: str) -> list:
    """判断问题实际问的是哪个/哪些维度。命中多个维度全部返回；一个都没命中则返回全部5个
    （表示这是一个宽泛问题，条款只要有一个维度写具体了就算通过）。"""
    hit = []
    for dim, kws in QUESTION_DIMENSION_KEYWORDS.items():
        if any(kw in question for kw in kws):
            hit.append(dim)
    return hit if hit else list(QUESTION_DIMENSION_KEYWORDS.keys())


def _check_dimension(dim: str, clause: str) -> bool:
    """判断条款是否对某个维度给出了具体规定（而不是泛泛提及）。"""
    if dim == "频次":
        if FREQ_CONCRETE_PATTERN.search(clause):
            return True
        return False
    if dim == "具体指标":
        return bool(METRIC_CONCRETE_PATTERN.search(clause))
    if dim == "责任主体":
        for kw in SUBJECT_CONCRETE_KEYWORDS:
            if kw in clause:
                return True
        return False
    if dim == "适用范围":
        return False  # 保守：启发式层不对"适用范围"单独判通过，避免误判，宁可漏判也不能错判通过
    if dim == "操作细节":
        return bool(DETAIL_CONCRETE_PATTERN.search(clause))
    return False


def _heuristic_judge(question: str, clause: str):
    """启发式核验主逻辑。返回 (covered: bool, reason_summary: str, dims_checked: list, dims_passed: list)"""
    dims = _detect_question_dimensions(question)
    passed = [d for d in dims if _check_dimension(d, clause)]

    if passed:
        reason = f"条款对问题所问的「{ '、'.join(passed) }」维度给出了具体规定，判定为已实质覆盖。"
        return True, reason, dims, passed

    reasons = []
    for d in dims:
        if d == "频次" and FREQ_VAGUE_ONLY_PATTERN.search(clause) and not FREQ_CONCRETE_PATTERN.search(clause):
            reasons.append("仅有「定期/适时/及时」等模糊表述，未规定具体周期")
        elif d == "责任主体" and SUBJECT_VAGUE_ONLY_PATTERN.match(clause.strip()):
            reasons.append("仅笼统提及「全员/人员」，未明确具体适用主体")
        elif d == "具体指标":
            reasons.append("未给出可量化、可比对的具体数值")
        elif d == "操作细节":
            reasons.append("未说明具体操作步骤")
        elif d == "适用范围":
            reasons.append("适用范围是否覆盖本问题情境无法仅凭启发式规则判断，建议人工核实")
    if not reasons and VAGUE_VERBS_ONLY.match(clause.strip()):
        reasons.append("仅为「应加强/应重视」类原则性表述，无具体要求")
    if not reasons:
        reasons.append("未在问题所问维度上找到具体、可执行的规定")

    reason = "命中但未实质覆盖：" + "；".join(reasons) + "。"
    return False, reason, dims, passed


def _llm_refine(question: str, clause: str, heuristic_result: dict):
    """可选：调用LLM对启发式判断做二次确认/精化。仅在设置了ANTHROPIC_API_KEY时启用。
    没有配置API key时直接返回None，调用方应回退到纯启发式结果。"""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        return None
    try:
        import anthropic
    except ImportError:
        return None

    prompt = f"""你在做安全生产合规标准的"覆盖充分性核验"。判断标准：候选条款是否针对问题
实际问的具体要求（指标/频次/责任主体/适用范围/操作细节）给出了明确、可执行的规定，
而不是泛泛提及。

问题：{question}
候选条款原文：{clause}

启发式规则给出的初步判断：{"已实质覆盖" if heuristic_result["covered"] else "命中但未实质覆盖"}
理由：{heuristic_result["reason_summary"]}

请判断这个初步判断是否准确。只输出JSON，不要输出其他内容：
{{"covered": true或false, "reason_summary": "一句话理由", "agree_with_heuristic": true或false}}
"""
    try:
        client = anthropic.Anthropic(api_key=api_key)
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=300,
            messages=[{"role": "user", "content": prompt}],
        )
        text = resp.content[0].text.strip()
        text = re.sub(r"^```json\s*|\s*```$", "", text.strip())
        parsed = json.loads(text)
        return parsed
    except Exception:
        return None


def _log_judgment(log_path: str, record: dict):
    """把每次判断记入本地judgment log（jsonl格式），供后续人工复核和调优关键词库使用。
    这不是Wiki层（Wiki层是agent可见的知识积累），这是私有服务自己的内部调优日志。"""
    try:
        Path(os.path.dirname(log_path)).mkdir(parents=True, exist_ok=True)
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
    except Exception:
        pass  # 日志写入失败不应该影响主流程返回结果


def check_coverage(question: str, candidate_clause: str, log_path: str = None) -> dict:
    """R11核验主入口。

    参数：
        question: 用户原始问题
        candidate_clause: 候选条款原文
        log_path: 可选，judgment log文件路径，默认写到本目录下的 judgment_log.jsonl

    返回：{"covered": bool, "reason_summary": str}
    """
    covered, reason, dims, passed = _heuristic_judge(question, candidate_clause)
    heuristic_result = {"covered": covered, "reason_summary": reason}

    final = dict(heuristic_result)
    llm_result = _llm_refine(question, candidate_clause, heuristic_result)
    source = "heuristic"
    if llm_result is not None:
        source = "llm_refined"
        final = {
            "covered": bool(llm_result.get("covered", covered)),
            "reason_summary": llm_result.get("reason_summary", reason),
        }

    if log_path is None:
        log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "judgment_log.jsonl")
    _log_judgment(log_path, {
        "time": datetime.datetime.now().isoformat(timespec="seconds"),
        "question": question,
        "clause": candidate_clause,
        "dimensions_checked": dims,
        "dimensions_passed": passed,
        "heuristic_result": heuristic_result,
        "llm_result": llm_result,
        "final_result": final,
        "source": source,
    })

    return final
