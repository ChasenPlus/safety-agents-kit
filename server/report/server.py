#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
sak-report MCP 服务器
========================
把"整改通知单生成"能力暴露成智能体可调用的 MCP 工具。

【职责边界 —— 重要】
本 MCP 只做一件事：把**已经查证完毕的结构化数据**，按真实模板格式排版成 docx。

它不做：
  - 不检索知识库、不查标准
  - 不核验条款是否实质覆盖
  - 不生成任何内容文字（问题描述、整改建议、暴露的问题都由调用方提供）

标准号/条款号/条款原文这些依据字段，必须由分支 A 检索核验后提供。
字段缺失时本工具会在文档里显式标注"标准依据待补充"，绝不编造——
这是安全生产合规文件，编造依据的后果比留空严重得多。

交付洁净后处理（事故沉淀，出厂即做，不靠人记）：
  - 大纲级别只给章/节/问题标题，正文段一律清除
    （模板 TOC 域带 \\u 开关，正文段一旦带大纲级别，刷新目录就会把整段正文吸进目录）
  - 每张图自动加 0.75pt 黑色细边框（a:ln w="9525"）
  - settings.xml 写 doNotCompressPictures，后续保存不重编码图片

工具清单：
  generate_rectification_notice  生成整改通知单 docx（可选直接插证据照片，自动带边框）
  postprocess_docx               对任意成品 docx 做同样的大纲级别/边框洁净处理
  get_data_schema                返回输入数据结构说明和示例（调用方自检用）
"""

import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "scripts"))

from fastmcp import FastMCP

from report_core import generate_report, postprocess_docx as _postprocess_docx, EXAMPLE_DATA

app = FastMCP("sak-report")


def _validate(data):
    """
    输入校验：返回(是否通过, 问题清单)。
    宁可提前报错说清楚缺什么，也不要生成一份半成品让人以为可以直接用。
    """
    problems = []
    for field in ("report_title", "company", "report_date"):
        if not data.get(field):
            problems.append(f"缺少必填字段：{field}")

    cats = data.get("categories")
    if not cats:
        problems.append("categories为空——没有任何问题内容，生成出来会是一份空报告")
        return False, problems

    for ci, cat in enumerate(cats, 1):
        if not cat.get("cat_name"):
            problems.append(f"第{ci}个专业大类缺少cat_name")
        for pi, prob in enumerate(cat.get("problems", []), 1):
            if not prob.get("prob_title"):
                problems.append(f"第{ci}类第{pi}个问题缺少prob_title")
            items = prob.get("items", [])
            if not items:
                problems.append(f"第{ci}类第{pi}个问题（{prob.get('prob_title','')}）下没有任何具体条目")
            for ii, item in enumerate(items, 1):
                if not item.get("finding"):
                    problems.append(f"第{ci}类第{pi}问题第{ii}条缺少finding（发现的问题描述）")
    return (len(problems) == 0), problems


def _count_incomplete_evidence(data):
    """统计有多少条问题的标准依据不完整——生成后如实告知调用方，不藏着"""
    incomplete = []
    for cat in data.get("categories", []):
        for prob in cat.get("problems", []):
            for item in prob.get("items", []):
                missing = [n for n, k in [("标准名称", "std_name"), ("标准号", "std_no"),
                                          ("条款号", "article"), ("条款原文", "clause_text")]
                           if not str(item.get(k, "")).strip()]
                if missing:
                    incomplete.append({
                        "seq": item.get("seq", "?"),
                        "finding_head": str(item.get("finding", ""))[:30] + "...",
                        "missing": missing
                    })
    return incomplete


@app.tool()
def get_data_schema() -> dict:
    """返回 generate_rectification_notice 所需的完整数据结构说明和可运行示例。
    调用方（分支 A）在组装数据前应先调用此工具确认字段含义，避免字段名写错导致内容丢失。"""
    return {
        "字段说明": {
            "report_title": "报告标题，如'示例发电厂安全综合巡查报告'（必填）",
            "company": "出具单位，如'示例电力股份有限公司'（必填）",
            "report_date": "日期，如'2026年7月2日'（必填）",
            "opening_text": "开篇政策引言段落（可选，不填则跳过这段）",
            "overall_text": "一、整体情况的正文（可选）",
            "last_year_text": "二、上一年度问题整改情况的正文（可选）",
            "problems_section_title": "三、xxx主要问题 的标题文字（可选，有默认值）",
            "categories": "专业大类数组（必填），每类含cat_no/cat_name/problems/suggestions",
            "categories[].cat_no": "大类序号，中文数字如'一'",
            "categories[].cat_name": "大类名称，如'安全管理方面'",
            "categories[].problems": "该类下的问题小项数组",
            "categories[].problems[].prob_no": "小项序号，阿拉伯数字如'1'（全篇连续编号）",
            "categories[].problems[].prob_title": "小项标题，如'安全目标管理不规范。'",
            "categories[].problems[].items": "该小项下的具体问题条目数组",
            "categories[].problems[].items[].seq": "条目序号，如'1'（对应模板里的（1）（2））",
            "categories[].problems[].items[].finding": "发现的问题描述（必填），如'查阅XX公司…存在…的情况'",
            "categories[].problems[].items[].std_name": "标准名称（需检索核验后填，缺失会标注待补）",
            "categories[].problems[].items[].std_no": "标准号含年份（需检索核验后填）",
            "categories[].problems[].items[].article": "条款号（需检索核验后填）",
            "categories[].problems[].items[].clause_text": "条款原文一字不差摘录（需检索核验后填，R00铁律）",
            "categories[].suggestions": "该大类的整改建议数组，每项一条（可选）",
            "closing_text": "结尾整改要求（可选）",
            "sign_left/sign_right/sign_date": "签字栏文字（可选，有默认值）",
            "output_path": "输出docx的完整路径（可选，也可通过工具参数传）"
        },
        "生成规则": [
            "每条问题会拼成模板固定句式：{finding}，不符合《{std_name}》（{std_no}）第{article}条“{clause_text}”的要求。",
            "四个依据字段任一缺失，该条会输出为：{finding}。[标准依据待补充：缺少XXX，需检索核验后填入，本工具不代为编造]",
            "表格右栏默认留空，供用户自行插入证据截图；条目带images数组时由本工具自动插图并加边框",
            "每个问题小项的表格末行固定是'暴露的问题：'，内容需调用方自行填入表格（本工具只建空行）",
            "生成后自动做交付洁净处理：正文段清大纲级别（目录不会吸进正文）、每张图加0.75pt黑色细边框、写doNotCompressPictures",
            "图片尺寸只允许四种：单图横版6.0×4.5cm、单图竖版4.5×6.0cm、组图≤4张5.0×5.0cm、组图≥5张4.5×4.5cm"
        ],
        "示例数据": EXAMPLE_DATA
    }


@app.tool()
def postprocess_docx(path: str, output_path: str = "") -> dict:
    """对任意已存在的成品 docx 做交付洁净后处理（不必由本工具生成）：
    ①大纲级别矫正——章 lvl0/节 lvl1/问题 lvl2，其余正文段一律清除
    （模板 TOC 域带 \\u，正文段一旦带大纲级别，刷新目录就把整段正文吸进目录）；
    ②每张图补 0.75pt 黑色细边框；③settings.xml 写 doNotCompressPictures。
    只改 XML 不动 word/media，图片字节完全不变。
    用户报'目录乱''图没边框'时先跑这个，再刷新目录。

    参数：
        path: 待处理的 docx 完整路径
        output_path: 输出路径；不填则原地保存（不是 Word 保存，不会重编码图片）
    """
    if not path or not os.path.exists(path):
        return {"status": "error", "message": f"文件不存在：{path}"}
    stats = _postprocess_docx(path, output_path or None)
    stats["status"] = "ok"
    stats["提示"] = (
        "已完成洁净处理。若之前目录有误，请先在 WPS 里刷新目录域，"
        "再核对条目只含章/节/问题标题；页码可手工核对。"
        "严禁用 Word 保存 docx（会重编码图片）。"
    )
    return stats


@app.tool()
def generate_rectification_notice(data_file: str, output_path: str = "") -> dict:
    """把已查证的结构化巡查问题数据，按整改通知单模板格式生成 docx 文件。

    【重要】本工具不检索标准、不核验条款、不生成内容文字，
    std_name/std_no/article/clause_text 四个依据字段必须由调用方检索核验后提供，
    缺失时文档中会显式标注'标准依据待补充'而不是编造。
    可选：条目里带 images 数组（照片绝对路径）即直接插图，
    尺寸自动取横版 6.0×4.5cm 或竖版 4.5×6.0cm 并加 0.75pt 黑色细边框；不传则右栏留空。
    出厂自动做交付洁净处理：正文段清大纲级别（目录不会吸进正文）+ 图片边框 + 禁压缩标记。
    调用前可先用 get_data_schema 查看完整数据结构。

    参数：
        data_file: 结构化数据 JSON 文件的完整路径（不是 JSON 字符串本身）
        output_path: 输出 docx 路径；JSON 里已有 output_path 时以 JSON 为准
    """
    if not data_file or not os.path.exists(data_file):
        return {"status": "error", "message": f"数据文件不存在：{data_file}"}

    with open(data_file, "r", encoding="utf-8") as f:
        data = json.load(f)

    out_path = data.get("output_path") or output_path
    if not out_path:
        return {"status": "error",
                "message": "必须提供 output_path（JSON 里的 output_path 字段，或工具参数 output_path）"}

    ok, problems = _validate(data)
    if not ok:
        return {"status": "error",
                "message": "输入数据校验未通过，未生成文件",
                "problems": problems}

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    stats = generate_report(data, out_path)

    incomplete = _count_incomplete_evidence(data)
    result = {
        "status": "ok",
        "path": out_path,
        "已生成问题条目数": sum(len(p.get("items", []))
                          for c in data.get("categories", [])
                          for p in c.get("problems", [])),
        "交付洁净后处理": stats,
        "证据截图": ("条目带images的已自动插图并加边框" if stats.get("图片边框_新加")
                  else "表格右栏已留空，需用户自行插入"),
        "暴露的问题": "各表格末行已建空行，需调用方或用户填入"
    }
    if incomplete:
        result["⚠️标准依据不完整的条目"] = incomplete
        result["提示"] = (f"有{len(incomplete)}条问题的标准依据字段不完整，"
                        "文档中已显式标注'标准依据待补充'，未编造。"
                        "建议回到分支A检索核验补全后重新生成。")
    return result


if __name__ == "__main__":
    app.run()
