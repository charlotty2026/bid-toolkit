#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
擎标（QingBiao）—— 投标可行性分析引擎
=========================================

整合自三个模块：
  1. 标书加速出品流水线 —— 拆解招标文件、废标项、评分项、时间节点
  2. bid_feasibility_engine.py（众测产物）—— 评分分析、条款映射、承诺链审计、资质检查、报告生成
  3. 标书封神秘籍/全生命周期管理 —— 7阶段倒推法、时间节点倒推表

定位：不是帮用户写标书，是帮用户「管好标书」——拆解→分析→审计→报告，一键出。

用法：
    bid analyze 招标文件.md --profile company_profile/ --output 分析报告.md
    bid analyze 招标文件.md --list-only          # 只拆解，不做深度分析
    bid analyze 招标文件.md --gen-config         # 生成默认配置
"""

import argparse
import json
import os
import re
import sys
from collections import OrderedDict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


# ======================================================================
# 配置常量
# ======================================================================

DEFAULT_CONFIG = {
    "company": {
        "name": "企业名称",
        "profile_dir": "company_profile",
    },
    "report": {
        "title": "投标可行性分析报告",
        "include_overview": True,
        "include_scores": True,
        "include_disqualifications": True,
        "include_commitments": True,
        "include_timeline": True,
        "include_risk": True,
        "include_recommendation": True,
    },
    "timeline": {
        "bid_open_day": "T",
        "days_before_open": 7,
        "warning_days": 3,
    },
    "scoring": {
        "high_score_threshold": 0.8,
        "medium_score_threshold": 0.6,
    },
}

# 评分项 → 技术方案章节映射
CHAPTER_RULES: List[Dict[str, Any]] = [
    {
        "keywords": ["价格", "报价", "成本"],
        "chapter_no": "第二章",
        "chapter_title": "报价方案与成本测算",
        "points": [
            "报价构成与测算依据（人工、物料、设备、管理费、税金、利润）",
            "评标基准价应对策略与最低价/合理价区间测算",
            "报价不超预算的合规性承诺与让利率说明",
        ],
    },
    {
        "keywords": ["技术", "方案", "服务方案", "实施"],
        "chapter_no": "第一章",
        "chapter_title": "项目理解与整体服务技术方案",
        "points": [
            "项目特点、重难点分析与需求理解",
            "核心服务流程设计与质量管控体系",
            "专项应急预案与创新服务措施",
        ],
    },
    {
        "keywords": ["人员", "团队", "配置", "组织"],
        "chapter_no": "第三章",
        "chapter_title": "项目团队配置与人员管理",
        "points": [
            "项目经理与核心岗位简历、职称、证书及到岗承诺",
            "人员配置表与排班方案",
            "人员培训、考核、轮岗与替补机制",
        ],
    },
    {
        "keywords": ["业绩", "经验", "案例"],
        "chapter_no": "第四章",
        "chapter_title": "同类项目业绩与经验复用",
        "points": [
            "近3年同类项目业绩清单与合同关键页",
            "典型项目的服务指标、甲方评价与经验提炼",
            "本项目可复用的标准化作业模板与改进措施",
        ],
    },
    {
        "keywords": ["承诺", "服务承诺", "保障"],
        "chapter_no": "第五章",
        "chapter_title": "服务承诺与保障措施",
        "points": [
            "服务质量指标承诺（满意度、合格率、响应时效等量化值）",
            "投诉处理机制、响应时效与违约赔偿条款",
            "持续改进、廉洁服务与保密承诺",
        ],
    },
    {
        "keywords": ["资质", "资格", "认证"],
        "chapter_no": "第六章",
        "chapter_title": "企业资质与资格证明",
        "points": [
            "营业执照、资质证书、体系认证等资格文件",
            "信用等级证书、获奖荣誉等加分材料",
            "联合体协议（如有）",
        ],
    },
]


# ======================================================================
# 通用解析工具
# ======================================================================

def _split_sections(text: str) -> OrderedDict:
    """按 Markdown 二级标题（## 标题）切分文本。"""
    sections = OrderedDict()
    current = "_preamble"
    buf: List[str] = []
    for line in text.splitlines():
        m = re.match(r"^##\s+(.+?)\s*$", line)
        if m:
            sections[current] = "\n".join(buf).strip()
            current = m.group(1).strip()
            buf = []
        else:
            buf.append(line)
    sections[current] = "\n".join(buf).strip()
    return sections


def _parse_numbered_list(text: str) -> List[str]:
    """提取有序列表项（支持 "1." 与 "1、" 两种风格）。"""
    items = []
    for line in text.splitlines():
        m = re.match(r"^\s*\d+[\.、]\s*(.+?)\s*$", line)
        if m:
            items.append(m.group(1).strip())
    return items


def _parse_bold_fields(text: str) -> OrderedDict:
    """提取形如 "- **字段名**：值" 的键值对。"""
    fields = OrderedDict()
    for line in text.splitlines():
        m = re.match(r"^\s*[-*]\s*\*\*(.+?)\*\*[：:]\s*(.+?)\s*$", line)
        if m:
            fields[m.group(1).strip()] = m.group(2).strip()
    return fields


def _parse_markdown_table(text: str) -> Tuple[Optional[List[str]], List[Dict[str, str]]]:
    """解析 Markdown 表格，返回 (表头, 行字典列表)。"""
    header = None
    rows = []
    for line in text.splitlines():
        line = line.strip()
        if not line.startswith("|"):
            continue
        cells = [c.strip() for c in line.strip("|").split("|")]
        if all(re.match(r"^:?-{2,}:?$", c) for c in cells):
            continue
        if header is None:
            header = cells
        else:
            if len(cells) == len(header):
                rows.append(dict(zip(header, cells)))
    return header, rows


def _parse_date(s: str) -> Optional[date]:
    """尝试解析日期，支持 YYYY-MM-DD / YYYY年M月D日 / YYYY/M/D。"""
    s = s.strip()
    for fmt in ("%Y-%m-%d", "%Y年%m月%d日", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


# ======================================================================
# 组件1：文件读取器
# ======================================================================

def read_files(bid_file: str, profile_dir: str = "") -> Dict[str, Any]:
    """读取招标文件和企业资料库。"""
    bid_path = Path(bid_file)
    if not bid_path.is_file():
        raise FileNotFoundError(f"招标文件不存在：{bid_path}")
    bid_text = bid_path.read_text(encoding="utf-8")

    profile_files: Dict[str, str] = {}
    if profile_dir:
        profile_path = Path(profile_dir)
        if profile_path.is_dir():
            for md in sorted(profile_path.glob("*.md")):
                profile_files[md.stem] = md.read_text(encoding="utf-8")

    profile_text = "\n\n".join(profile_files.values())

    return {
        "bid_text": bid_text,
        "profile_files": profile_files,
        "profile_text": profile_text,
        "meta": {
            "read_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "bid_file": str(bid_path.resolve()),
            "profile_dir": str(Path(profile_dir).resolve()) if profile_dir else "",
            "file_size": bid_path.stat().st_size,
        },
    }


# ======================================================================
# 组件2：招标文件拆解器
# ======================================================================

def parse_bid_document(bid_text: str) -> Dict[str, Any]:
    """拆解招标文件，提取结构化信息。

    输出字段：
        project:         项目概况
        scores:          评分标准表
        invalid_clauses: 废标条款
        qualifications:  投标人资格要求
        documents:       投标文件清单
        timeline:        时间节点
        pricing_rule:    报价规则
        other_requirements: 其他要求
        sections:        原始章节切分
    """
    sections = _split_sections(bid_text)

    # 项目概况
    project = {}
    for key_section in ["项目概况", "项目基本情况", "项目概述", "项目简介"]:
        block = sections.get(key_section, "")
        if block:
            project = dict(_parse_bold_fields(block))
            break

    # 评分标准表
    scores = []
    for key_section in ["评分标准表", "评分标准", "评分办法", "评审标准"]:
        _, score_rows = _parse_markdown_table(sections.get(key_section, ""))
        if score_rows:
            for row in score_rows:
                name = row.get("评分项", row.get("评审项目", "")).strip()
                score_raw = row.get("分值", "0").strip()
                m = re.search(r"(\d+(?:\.\d+)?)", score_raw)
                score_val = float(m.group(1)) if m else 0.0
                scores.append({
                    "序号": row.get("序号", row.get("序", "")).strip(),
                    "评分项": name,
                    "分值": score_raw,
                    "分值数": score_val,
                    "评分标准": row.get("评分标准", row.get("评审细则", row.get("内容", ""))).strip(),
                })
            break

    # 废标条款
    invalid_clauses = []
    for key_section in ["废标条款", "无效投标情形", "废标条件", "否决投标条款"]:
        items = _parse_numbered_list(sections.get(key_section, ""))
        if items:
            invalid_clauses = items
            break

    # 资格要求
    qualifications = []
    for key_section in ["投标人资格要求", "资质要求", "投标人资质", "资格条件"]:
        items = _parse_numbered_list(sections.get(key_section, ""))
        if items:
            qualifications = items
            break

    # 文件清单
    documents = []
    for key_section in ["投标文件组成", "文件清单", "投标文件清单", "投标文件目录"]:
        items = _parse_numbered_list(sections.get(key_section, ""))
        if items:
            documents = items
            break

    # 时间节点
    timeline = {}
    for key_section in ["时间节点", "时间安排", "招标日程", "投标日程"]:
        fields = dict(_parse_bold_fields(sections.get(key_section, "")))
        if fields:
            timeline = fields
            break

    # 其他要求
    other_requirements = []
    for key_section in ["其他要求", "其他说明", "注意事项", "特别说明"]:
        items = _parse_numbered_list(sections.get(key_section, ""))
        if items:
            other_requirements = items
            break

    # 报价规则
    pricing_rule = ""
    for s in scores:
        if "价格" in s["评分项"]:
            pricing_rule = s["评分标准"]
            break
    extra_pricing = []
    for item in other_requirements:
        if any(k in item for k in ("保证金", "付款", "报价", "预算")):
            extra_pricing.append(item)
    if extra_pricing:
        pricing_rule = (
            (pricing_rule + "\n" if pricing_rule else "")
            + "商务附属规则：" + "；".join(extra_pricing)
        )

    return {
        "project": project,
        "scores": scores,
        "invalid_clauses": invalid_clauses,
        "qualifications": qualifications,
        "documents": documents,
        "timeline": timeline,
        "pricing_rule": pricing_rule,
        "other_requirements": other_requirements,
        "sections": sections,
    }


# ======================================================================
# 组件3：企业资料库解析
# ======================================================================

def parse_company_profile(profile_files: Dict[str, str]) -> Dict[str, Any]:
    """解析企业资料库为结构化数据。

    输入：read_files() 返回的 profile_files（{stem: text}）
    输出：{team, performance, qualifications}
    """
    today = date.today()
    result: Dict[str, Any] = {"team": [], "performance": [], "qualifications": []}

    # ---------- 人员 ----------
    team_text = profile_files.get("team", "")
    if team_text:
        sections = _split_sections(team_text)
        for role, body in sections.items():
            if role == "_preamble" or not body.strip():
                continue
            fields = _parse_bold_fields(body)
            years_str = fields.get("从业年限", "0")
            ym = re.search(r"(\d+)", years_str)
            result["team"].append({
                "角色": role,
                "姓名": fields.get("姓名", ""),
                "职称": fields.get("职称", ""),
                "从业年限": int(ym.group(1)) if ym else 0,
                "相关经验": fields.get("相关经验", ""),
                "证书": fields.get("证书", ""),
                "当前状态": fields.get("当前状态", ""),
                "原文": f"{role} {body}".replace("\n", " "),
            })

    # ---------- 业绩 ----------
    perf_text = profile_files.get("performance", "")
    if perf_text:
        sections = _split_sections(perf_text)
        for title, body in sections.items():
            if title == "_preamble" or not body.strip():
                continue
            fields = _parse_bold_fields(body)
            amount_str = fields.get("合同金额", "")
            am = re.search(r"([\d.]+)\s*万", amount_str)
            amount = float(am.group(1)) if am else 0.0
            period = fields.get("服务期限", "")
            pm = re.search(
                r"(\d{4})年(\d{1,2})月?\s*[-–~至到]\s*(\d{4})年(\d{1,2})月?", period
            )
            start_d = end_d = None
            if pm:
                start_d = date(int(pm.group(1)), int(pm.group(2)), 1)
                end_d = date(int(pm.group(3)), int(pm.group(4)), 1)
            ongoing = ("正在执行" in period) or (
                end_d is not None and end_d >= today
            )
            full_text = f"{title} {body}"
            result["performance"].append({
                "名称": re.sub(r"^业绩\d+[：:]\s*", "", title),
                "甲方": fields.get("甲方单位", ""),
                "合同金额(万)": amount,
                "起始": start_d,
                "截止": end_d,
                "进行中": ongoing,
                "服务内容": fields.get("服务内容", ""),
                "原文": full_text.replace("\n", " "),
            })

    # ---------- 资质 ----------
    qual_text = profile_files.get("qualifications", "")
    if qual_text:
        _, rows = _parse_markdown_table(qual_text)
        for r in rows:
            name = r.get("证书名称", "").strip()
            expiry_str = r.get("有效期至", "").strip()
            expiry = _parse_date(expiry_str)
            if expiry is None:
                calc_status = "❓ 未知"
            elif expiry < today:
                calc_status = "❌ 已过期"
            elif (expiry - today).days <= 90:
                calc_status = "⚠️ 即将到期"
            else:
                calc_status = "✅ 有效"
            result["qualifications"].append({
                "证书名称": name,
                "证书编号": r.get("证书编号", "").strip(),
                "发证机构": r.get("发证机构", "").strip(),
                "有效期至": expiry_str,
                "有效期date": expiry,
                "原始状态": r.get("状态", "").strip(),
                "计算状态": calc_status,
                "原文": f"{name} {r.get('证书编号','')} {r.get('发证机构','')} 有效期至{expiry_str}",
            })

    return result


# ======================================================================
# 组件4：条款映射器（评分项 → 技术方案章节）
# ======================================================================

def map_clauses_to_chapters(scoring_items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """将评分条款映射到技术方案章节。"""
    mapping = []
    for item in scoring_items:
        name = item["评分项"]
        matched = None
        for rule in CHAPTER_RULES:
            if any(k in name for k in rule["keywords"]):
                matched = rule
                break
        if matched is None:
            matched = {
                "chapter_no": "第N章",
                "chapter_title": f"{name}专项方案",
                "points": [f"围绕「{name}」编制专项响应内容"],
            }
        mapping.append({
            "评分项": name,
            "分值": item["分值"],
            "分值数": item["分值数"],
            "章节号": matched["chapter_no"],
            "章节标题": matched["chapter_title"],
            "应覆盖要点": matched["points"],
        })
    return mapping


# ======================================================================
# 组件5：评分分析器
# ======================================================================

def analyze_scores(
    scoring_items: List[Dict[str, Any]],
    profile_data: Dict[str, Any],
    chapter_mapping: List[Dict[str, Any]],
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """对评分项做分值分布、优劣势分析，结合企业资料库给出预估得分。"""
    team = profile_data.get("team", [])
    perfs = profile_data.get("performance", [])
    quals = profile_data.get("qualifications", [])

    total_score = sum(it["分值数"] for it in scoring_items) or 100
    pm = next((t for t in team if "项目经理" in t["角色"]), None)
    hospital_perfs = [p for p in perfs if "医院" in p.get("原文", "")]
    qual_3rd = [
        p
        for p in perfs
        if p.get("合同金额(万)", 0) >= 300
        and p.get("截止") is not None
        and p.get("截止", date.min) >= date(2023, 1, 1)
    ]
    iso_count = sum(
        1
        for q in quals
        if re.search(r"ISO\s*9|ISO\s*14|ISO\s*45", q["证书名称"])
    )

    chapter_by_name = {m["评分项"]: m for m in chapter_mapping}

    analysis = []
    for item in scoring_items:
        name = item["评分项"]
        std = item["评分标准"]
        mp = chapter_by_name.get(name, {})
        chapter = f'{mp.get("章节号","")} {mp.get("章节标题","")}'.strip()
        advantage = ""
        disadvantage = ""
        estimate = ""
        estimate_note = ""

        if "价格" in name:
            advantage = "可采用成本倒推+竞争性报价策略，报价空间可控"
            disadvantage = "低价竞争可能压缩利润；需准确测算成本，避免低于成本价被认定为无效投标"
            estimate = "待定"
            estimate_note = "价格分取决于最终报价与竞争对手报价，建议报价控制在合理区间"

        elif "技术" in name or "方案" in name:
            advantage = (
                f"具备{len(hospital_perfs)}个同类项目实操经验，可复用成熟流程与应急预案；"
                f"ISO三体系认证（{iso_count}项）为方案规范性提供支撑"
            )
            disadvantage = "需针对本项目实际情况定制，不能照搬通用模板"
            estimate = "中上（预估良～优）"
            estimate_note = "若项目理解深入、流程与应急预案针对性强，有望冲击高分"

        elif "人员" in name or "团队" in name:
            adv_parts = []
            dis_parts = []
            if pm:
                pj_title = pm["职称"]
                if "高级" in pj_title:
                    adv_parts.append(f"项目经理{pm['姓名']}为{pj_title}，可得职称满分")
                elif "中级" in pj_title:
                    adv_parts.append(f"项目经理{pm['姓名']}为{pj_title}，可得3分")
                adv_parts.append(f"从业{pm['从业年限']}年，含相关项目经验")
            if not any(
                p for p in team if "技术负责人" in p["角色"] and "高级" in p.get("职称", "")
            ):
                dis_parts.append("技术负责人为中级职称，团队职称梯队存在短板")
            if not dis_parts:
                dis_parts.append("需确保人员花名册与社保证明齐全")
            advantage = "；".join(adv_parts) + "；核心团队证书齐全、全员可到岗" if adv_parts else "核心团队证书齐全"
            disadvantage = "；".join(dis_parts)
            estimate = "中上"
            estimate_note = "项目经理职称可锁定高分；团队配置若补齐佐证材料，可达优档"

        elif "业绩" in name or "经验" in name:
            max_amount = max((p["合同金额(万)"] for p in qual_3rd), default=0)
            advantage = (
                f"近3年满足条件的同类业绩共{len(qual_3rd)}个，"
                f"最高单项{max_amount:.0f}万元，合同复印件可加盖公章提供"
            )
            gap = max(0, 5 - len(qual_3rd))
            if len(qual_3rd) >= 5:
                disadvantage = "业绩充分，但需注意合同关键页清晰可辨"
                estimate = "满分"
                estimate_note = "5个合格业绩即可拿满，按最高5个申报即可"
            else:
                disadvantage = f"距满分尚差{gap}个合格业绩"
                estimate = f"可拿{len(qual_3rd)*3}分"
                estimate_note = f"现有{len(qual_3rd)}个合格业绩×3分；应全部申报并确保均在有效期内"

        elif "承诺" in name or "保障" in name:
            advantage = "公司已有标准化服务承诺模板"
            disadvantage = "部分承诺需提供可核查证据（信息化平台、HIS对接等）"
            estimate = "中上"
            estimate_note = "承诺模板齐全，但需补充量化指标的支撑材料"

        else:
            advantage = "有相关项目经验可参考"
            disadvantage = "需根据招标文件具体要求定制"
            estimate = "待定"
            estimate_note = "根据实际响应内容评分"

        analysis.append({
            "评分项": name,
            "分值": item["分值"],
            "分值数": item["分值数"],
            "占比": f"{item['分值数'] / total_score * 100:.1f}%",
            "评分标准摘要": std[:80] + "..." if len(std) > 80 else std,
            "建议章节": chapter,
            "优势": advantage,
            "劣势": disadvantage,
            "预估得分": estimate,
            "得分说明": estimate_note,
        })

    # 分值分布
    distribution = {
        "总分": total_score,
        "技术分": sum(it["分值数"] for it in scoring_items if "价格" not in it["评分项"]),
        "价格分": sum(it["分值数"] for it in scoring_items if "价格" in it["评分项"]),
        "项目数": len(scoring_items),
    }

    return analysis, distribution


# ======================================================================
# 组件6：承诺链审计器
# ======================================================================

def audit_commitments(
    bid_data: Dict[str, Any], profile_data: Dict[str, Any]
) -> List[Dict[str, Any]]:
    """审计招标文件中的承诺项，在企业资料库中查找证据。"""
    # 从拆解结果中提取承诺项
    commitments = []
    for section_name in ["service_commitments", "服务承诺", "承诺要求", "投标承诺"]:
        if section_name in bid_data.get("sections", {}):
            items = _parse_numbered_list(bid_data["sections"][section_name])
            for i, text in enumerate(items, 1):
                commitments.append({"no": str(i), "text": text})

    # 如果没找到专门的承诺章节，从评分项中提取含"承诺"的
    if not commitments:
        for item in bid_data.get("scores", []):
            if "承诺" in item["评分项"]:
                commitments.append({"no": str(len(commitments) + 1), "text": item["评分标准"]})

    # 如果没有承诺项，直接返回
    if not commitments:
        return []

    # 构建可检索语料
    team_hay = "\n".join(t.get("原文", "") for t in profile_data.get("team", []))
    perf_hay = "\n".join(p.get("原文", "") for p in profile_data.get("performance", []))
    qual_hay = "\n".join(q.get("原文", "") for q in profile_data.get("qualifications", []))

    # 承诺关键词 → 证据库匹配规则
    rulebook = [
        (r"项目经理.*到岗|项目经理.*兼任", "team", "team.md：项目经理到岗承诺字段",
         "direct", ["到岗", "项目经理"]),
        (r"社保|劳动合同", "team", "team.md：人员社保单位字段",
         "direct", ["社保"]),
        (r"响应.*分钟|到场|响应时效", "team", "需在技术方案中明确响应流程并附SLA配置截图",
         "weak", []),
        (r"满意度|90%|95%|考核", "performance", "performance.md：历史项目考核优秀/合格记录",
         "direct", ["考核", "优秀", "合格"]),
        (r"安全.*事故|安全责任", "qualifications", "qualifications.md：安全生产许可证、ISO 45001",
         "direct", ["安全生产许可证", "ISO 45001"]),
        (r"核心岗位.*更换|人员稳定", "team", "team.md：核心岗位到岗承诺（需在技术方案中补人员稳定承诺函）",
         "weak", ["到岗"]),
        (r"医废|消毒|台账", "performance", "performance.md：服务内容含医废暂存/消毒（需在技术方案中附台账模板）",
         "weak", ["医废", "消毒"]),
        (r"信息化|HIS|平台|工单|巡检系统", "qualifications", "需提供信息化平台截图、功能清单及HIS对接方案",
         "weak", []),
        (r"保密|廉洁", "qualifications", "qualifications.md：保密协议、廉洁承诺制度",
         "weak", ["保密"]),
        (r"质保|保修|售后", "performance", "performance.md：历史项目质保记录",
         "weak", ["质保", "保修"]),
        (r"培训|交底", "team", "team.md：人员培训记录",
         "weak", ["培训"]),
    ]

    results = []
    for c in commitments:
        text = c["text"]
        has_direct = False
        has_weak = False
        evidence_parts = []
        lib_hit = set()

        for pattern, lib, desc, strength, verify_kws in rulebook:
            if re.search(pattern, text, re.IGNORECASE):
                source_hay = {"team": team_hay, "performance": perf_hay, "qualifications": qual_hay}[lib]
                if strength == "direct" and source_hay and all(k in source_hay for k in verify_kws if k):
                    has_direct = True
                    lib_hit.add(lib)
                    evidence_parts.append(f"[{lib}.md] {desc}")
                else:
                    has_weak = True
                    evidence_parts.append(f"[{lib}.md] {desc}")

        if not has_direct and not has_weak:
            if re.search(r"医院|服务|项目", text) and perf_hay:
                has_weak = True
                evidence_parts.append("[performance.md] 存在类似业绩，可间接支撑")

        if has_direct and not has_weak:
            status = "✅ 已匹配"
        elif has_direct and has_weak:
            status = "⚠️ 部分匹配"
        elif has_weak:
            status = "⚠️ 部分匹配"
        else:
            status = "❌ 缺失"

        results.append({
            "序号": c["no"],
            "承诺条款": text,
            "状态": status,
            "证据来源": "；".join(evidence_parts) if evidence_parts else "未在企业资料库中检索到支撑材料",
            "需补充": "是" if "weak" in status or status == "❌ 缺失" else "否",
        })

    return results


# ======================================================================
# 组件7：废标/资质检查器
# ======================================================================

def check_disqualifications(
    bid_data: Dict[str, Any], profile_data: Dict[str, Any]
) -> List[Dict[str, Any]]:
    """检查废标条款和资质要求，对照企业资料库给出状态。"""
    results = []
    today = date.today()

    # 检查废标条款
    for i, clause in enumerate(bid_data.get("invalid_clauses", []), 1):
        risk = "低"
        # 启发式判断风险等级
        if any(k in clause for k in ("盖章", "签字", "原件", "密封")):
            risk = "高"
        elif any(k in clause for k in ("时间", "截止", "逾期")):
            risk = "高"
        elif any(k in clause for k in ("资质", "证书", "许可")):
            risk = "中"
        results.append({
            "类型": "废标条款",
            "序号": str(i),
            "内容": clause[:100] + "..." if len(clause) > 100 else clause,
            "风险等级": risk,
            "状态": "⚠️ 需人工确认",
            "建议": "逐条核对，确保投标文件完全满足" if risk == "高" else "一般不会触发，但需确认",
        })

    # 检查资质要求
    for i, qual in enumerate(bid_data.get("qualifications", []), 1):
        # 在企业资质库中查找
        matched = False
        for q in profile_data.get("qualifications", []):
            qual_text = qual[:20]  # 取前20字匹配
            if any(k in q["证书名称"] for k in qual_text.split()):
                matched = True
                status = q["计算状态"]
                break
            else:
                status = "⚠️ 未匹配"

        results.append({
            "类型": "资质要求",
            "序号": str(i),
            "内容": qual[:100] + "..." if len(qual) > 100 else qual,
            "风险等级": "高" if not matched else "中",
            "状态": "✅ 已匹配" if matched else "❌ 未匹配",
            "建议": "资质证书在有效期内，可正常使用" if matched else "需补充或确认资质证书",
        })

    return results


# ======================================================================
# 组件8：时间节点分析（7阶段倒推法）
# ======================================================================

def analyze_timeline(
    bid_data: Dict[str, Any], open_date: Optional[str] = None
) -> Dict[str, Any]:
    """根据招标文件时间节点和7阶段倒推法，生成时间规划。

    7阶段倒推法口诀：
        T-1: 打印装订敲章
        T-2~T-3: 标书终稿
        T-3~T-4: 报价测算
        T-5~T-7: 技术方案交稿
        立项后一周: 资料收集截止
        拿到招标文件当天: 银行资信
    """
    today = date.today()

    # 尝试从招标文件中提取开标日期
    bid_open = None
    for key, val in bid_data.get("timeline", {}).items():
        if "开标" in key or "投标截止" in key or "递交" in key:
            bid_open = _parse_date(val)
            break

    if bid_open is None and open_date:
        bid_open = _parse_date(open_date)

    if bid_open is None:
        return {
            "status": "❓ 未找到开标日期",
            "today": today.isoformat(),
            "bid_open": "未知",
            "remaining_days": "未知",
            "timeline": [],
            "mode": "未知",
            "warning": "请在招标文件中补充时间节点信息",
        }

    remaining = (bid_open - today).days
    if remaining < 0:
        mode = "⏰ 已过期"
    elif remaining <= 3:
        mode = "🚀 极速模式"
    elif remaining <= 7:
        mode = "🐇 加速模式"
    else:
        mode = "🐢 标准模式"

    # 7阶段倒推表
    stages = [
        ("T-1", bid_open - timedelta(days=1), "打印装订敲章"),
        ("T-2~T-3", bid_open - timedelta(days=3), "标书终稿（含审核）"),
        ("T-3~T-4", bid_open - timedelta(days=4), "报价测算"),
        ("T-5~T-7", bid_open - timedelta(days=7), "技术方案交稿"),
        ("立项后1周", None, "资料收集截止"),
        ("拿到文件当天", None, "银行资信证明"),
    ]

    timeline = []
    for stage_name, stage_date, description in stages:
        if stage_date:
            days_left = (stage_date - today).days
            overdue = days_left < 0
            timeline.append({
                "阶段": stage_name,
                "截止日期": stage_date.isoformat(),
                "任务": description,
                "剩余天数": f"{days_left}天" if not overdue else f"已超{abs(days_left)}天",
                "状态": "⚠️ 已超期" if overdue else "✅ 进行中" if days_left <= 3 else "⏳ 时间充裕",
            })

    return {
        "status": f"距开标还有 {remaining} 天",
        "today": today.isoformat(),
        "bid_open": bid_open.isoformat(),
        "remaining_days": remaining,
        "mode": mode,
        "timeline": timeline,
        "warning": "时间紧张，建议压缩审核环节" if remaining <= 3 else "",
    }


# ======================================================================
# 组件9：报告生成器
# ======================================================================

def generate_report(
    bid_data: Dict[str, Any],
    profile_data: Dict[str, Any],
    score_analysis: List[Dict[str, Any]],
    score_distribution: Dict[str, Any],
    clause_mapping: List[Dict[str, Any]],
    commitment_results: List[Dict[str, Any]],
    disqualification_results: List[Dict[str, Any]],
    timeline_analysis: Dict[str, Any],
    bid_file: str,
    config: Optional[Dict[str, Any]] = None,
) -> str:
    """生成完整的 Markdown 投标可行性分析报告。"""
    cfg = config or DEFAULT_CONFIG
    project = bid_data.get("project", {})
    lines = []

    # ====== 标题 ======
    lines.append(f"# {cfg['report']['title']}")
    lines.append("")
    lines.append(f"> **生成时间**：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append(f"> **招标文件**：`{bid_file}`")
    lines.append(f"> **分析模式**：{timeline_analysis.get('mode', '标准模式')}")
    lines.append("")

    # ====== 紧急程度 ======
    lines.append("## 一、紧急程度评估")
    lines.append("")
    status = timeline_analysis.get("status", "未知")
    mode = timeline_analysis.get("mode", "未知")
    warning = timeline_analysis.get("warning", "")
    lines.append(f"- **{status}** | 建议模式：**{mode}**")
    if warning:
        lines.append(f"- ⚠️ **{warning}**")
    lines.append("")

    # 时间倒推表
    if cfg["report"]["include_timeline"] and timeline_analysis.get("timeline"):
        lines.append("### 时间倒推表（7阶段倒推法）")
        lines.append("")
        lines.append("| 阶段 | 截止日期 | 任务 | 剩余天数 | 状态 |")
        lines.append("|------|---------|------|---------|------|")
        for t in timeline_analysis["timeline"]:
            lines.append(
                f"| {t['阶段']} | {t['截止日期']} | {t['任务']} | {t['剩余天数']} | {t['状态']} |"
            )
        lines.append("")

    # ====== 项目概况 ======
    if cfg["report"]["include_overview"]:
        lines.append("## 二、项目概况")
        lines.append("")
        if project:
            for key, val in project.items():
                lines.append(f"- **{key}**：{val}")
        else:
            lines.append("（未从招标文件中提取到项目概况字段）")
        lines.append("")

    # ====== 评分分析 ======
    if cfg["report"]["include_scores"]:
        lines.append("## 三、评分项分析与策略")
        lines.append("")
        lines.append(
            f"总分 **{score_distribution.get('总分', '?')}分** | "
            f"技术分 {score_distribution.get('技术分', '?')}分 | "
            f"价格分 {score_distribution.get('价格分', '?')}分 | "
            f"共{score_distribution.get('项目数', '?')}项"
        )
        lines.append("")

        if score_analysis:
            lines.append("| 评分项 | 分值 | 占比 | 建议章节 | 优势 | 劣势 | 预估 |")
            lines.append("|--------|------|------|---------|------|------|------|")
            for a in score_analysis:
                lines.append(
                    f"| {a['评分项']} | {a['分值']} | {a['占比']} |"
                    f" {a['建议章节']} | {a['优势'][:30]}... | {a['劣势'][:30]}... | {a['预估得分']} |"
                )
            lines.append("")

            # 详细分析
            lines.append("### 详细分析")
            lines.append("")
            for a in score_analysis:
                lines.append(f"#### {a['评分项']}（{a['分值']}，占比{a['占比']}）")
                lines.append("")
                lines.append(f"- **建议章节**：{a['建议章节']}")
                lines.append(f"- **评分标准**：{a['评分标准摘要']}")
                lines.append(f"- **优势**：{a['优势']}")
                lines.append(f"- **劣势**：{a['劣势']}")
                lines.append(f"- **预估得分**：{a['预估得分']}")
                lines.append(f"- **说明**：{a['得分说明']}")
                lines.append("")

    # ====== 废标条款与资质检查 ======
    if cfg["report"]["include_disqualifications"] and disqualification_results:
        lines.append("## 四、废标条款与资质检查")
        lines.append("")

        # 高危项汇总
        high_risk = [d for d in disqualification_results if d["风险等级"] == "高"]
        if high_risk:
            lines.append("### 🔴 高危项（不处理=废标风险）")
            lines.append("")
            for d in high_risk:
                lines.append(f"- **{d['类型']} #{d['序号']}**：{d['内容'][:60]}...")
                lines.append(f"  - 状态：{d['状态']} | 建议：{d['建议']}")
            lines.append("")

        lines.append("### 完整检查清单")
        lines.append("")
        lines.append("| 类型 | 序号 | 内容 | 风险 | 状态 | 建议 |")
        lines.append("|------|------|------|------|------|------|")
        for d in disqualification_results:
            icon = "🔴" if d["风险等级"] == "高" else "🟡" if d["风险等级"] == "中" else "🟢"
            lines.append(
                f"| {d['类型']} | {d['序号']} | {d['内容'][:40]}... |"
                f" {icon} {d['风险等级']} | {d['状态']} | {d['建议'][:40]}... |"
            )
        lines.append("")

    # ====== 承诺链审计 ======
    if cfg["report"]["include_commitments"] and commitment_results:
        lines.append("## 五、承诺链审计")
        lines.append("")

        # 缺失项汇总
        missing = [c for c in commitment_results if c["状态"] == "❌ 缺失"]
        if missing:
            lines.append("### ❌ 缺失承诺（需立即补充）")
            lines.append("")
            for c in missing:
                lines.append(f"- **#{c['序号']}**：{c['承诺条款'][:60]}...")
            lines.append("")

        lines.append("### 完整承诺清单")
        lines.append("")
        lines.append("| 序号 | 承诺条款 | 状态 | 证据来源 | 需补充 |")
        lines.append("|------|---------|------|---------|-------|")
        for c in commitment_results:
            lines.append(
                f"| {c['序号']} | {c['承诺条款'][:50]}... |"
                f" {c['状态']} | {c['证据来源'][:40]}... | {c['需补充']} |"
            )
        lines.append("")

    # ====== 条款映射 ======
    if clause_mapping:
        lines.append("## 六、评分项→方案章节映射")
        lines.append("")
        lines.append("| 评分项 | 分值 | 应放章节 | 应覆盖要点 |")
        lines.append("|--------|------|---------|-----------|")
        for m in clause_mapping:
            points = "；".join(m["应覆盖要点"][:2])
            lines.append(
                f"| {m['评分项']} | {m['分值']} |"
                f" {m['章节号']} {m['章节标题']} | {points}... |"
            )
        lines.append("")

    # ====== 风险提示 ======
    if cfg["report"]["include_risk"]:
        lines.append("## 七、风险提示")
        lines.append("")

        risks = []
        # 从时间分析
        remaining = timeline_analysis.get("remaining_days")
        if remaining is None:
            pass  # 无日期信息，跳过时间风险判断
        elif remaining <= 3:
            risks.append("🔴 **时间极紧**：距开标仅剩3天以内，建议启用极速模式")
        elif timeline_analysis.get("remaining_days", 999) <= 7:
            risks.append("🟡 **时间较紧**：距开标不足7天，建议启用加速模式")

        # 从废标检查
        if high_risk := [d for d in disqualification_results if d["风险等级"] == "高"]:
            risks.append(f"🔴 **废标风险**：发现{len(high_risk)}项高危条款，必须逐条核验")

        # 从承诺链
        if missing_commitments := [c for c in commitment_results if c["状态"] == "❌ 缺失"]:
            risks.append(f"🔴 **承诺缺失**：{len(missing_commitments)}项承诺在企业资料库中无支撑材料")

        if not risks:
            risks.append("✅ **整体风险可控**，建议按标准流程推进")

        for r in risks:
            lines.append(f"- {r}")
        lines.append("")

    # ====== 建议 ======
    if cfg["report"]["include_recommendation"]:
        lines.append("## 八、行动建议")
        lines.append("")

        mode = timeline_analysis.get("mode", "标准模式")
        if "极速" in mode:
            lines.append("### 🚀 极速模式（3天内）")
            lines.append("")
            lines.append("1. 先看废标项清单，确认所有必备条件满足")
            lines.append("2. 直接调内容块库拼初稿，不精修排版")
            lines.append("3. 跑质量检查脚本（只看高危项）")
            lines.append("4. 肉眼过一遍 → 盖章 → 提交")
            lines.append("")
        elif "加速" in mode:
            lines.append("### 🐇 加速模式（3-7天）")
            lines.append("")
            lines.append("1. 废标项清单逐条核验")
            lines.append("2. 内容编制+质量检查并行")
            lines.append("3. 报价测算优先做")
            lines.append("4. 预留至少1天给审核和修改")
            lines.append("")
        else:
            lines.append("### 🐢 标准模式（7天以上）")
            lines.append("")
            lines.append("1. 按7阶段倒推法走完整5阶段流程")
            lines.append("2. 技术方案先交稿（预留修改时间）")
            lines.append("3. 报价测算做多稿对比")
            lines.append("4. 质量检查跑全量8条")
            lines.append("5. 归档经验，沉淀内容块")
            lines.append("")

        lines.append("---")
        lines.append("")
        lines.append(
            f"> *报告由擎标（QingBiao）投标可行性分析引擎自动生成 | "
            f"{datetime.now().strftime('%Y-%m-%d %H:%M')}*"
        )
        lines.append("")

    return "\n".join(lines)


# ======================================================================
# 命令行入口
# ======================================================================

def main():
    parser = argparse.ArgumentParser(
        prog="bid analyze",
        description="📋 投标可行性分析 — 一键拆解招标文件+评分分析+承诺审计+时间规划",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例:\n"
            "  bid analyze 招标文件.md                     # 基础拆解\n"
            "  bid analyze 招标文件.md --profile 企业资料/  # 含企业资料库深度分析\n"
            "  bid analyze 招标文件.md -o 报告.md           # 输出报告\n"
            "  bid analyze 招标文件.md --list-only          # 仅拆解，不深度分析\n"
            "  bid analyze 招标文件.md --gen-config         # 生成默认配置\n"
        ),
    )
    parser.add_argument("input", help="招标文件路径（.md）")
    parser.add_argument("--profile", "-p", default="", help="企业资料库目录路径")
    parser.add_argument("--output", "-o", default="", help="输出报告路径（.md）")
    parser.add_argument("--config", "-c", default="", help="配置文件路径")
    parser.add_argument("--list-only", action="store_true", help="仅拆解招标文件，不做深度分析")
    parser.add_argument("--gen-config", action="store_true", help="生成默认配置文件")
    parser.add_argument("--json", action="store_true", help="同时输出JSON格式")
    parser.add_argument("--open-date", default="", help="开标日期（YYYY-MM-DD），自动从招标文件提取失败时使用")

    args = parser.parse_args()

    # 生成默认配置
    if args.gen_config:
        config_path = "bid_analyze_config.yaml"
        import yaml
        with open(config_path, "w", encoding="utf-8") as f:
            yaml.dump(DEFAULT_CONFIG, f, allow_unicode=True, default_flow_style=False)
        print(f"✅ 默认配置文件已生成: {config_path}")
        return

    # 读取配置
    config = None
    if args.config:
        import yaml
        with open(args.config, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)

    # 读取文件
    print(f"📖 读取招标文件: {args.input}")
    file_data = read_files(args.input, args.profile)
    bid_text = file_data["bid_text"]
    profile_files = file_data["profile_files"]
    print(f"   - 文件大小: {file_data['meta']['file_size']} 字节")

    # 拆解招标文件
    print("🔍 拆解招标文件...")
    bid_data = parse_bid_document(bid_text)
    print(f"   - 评分项: {len(bid_data['scores'])} 项")
    print(f"   - 废标条款: {len(bid_data['invalid_clauses'])} 条")
    print(f"   - 资格要求: {len(bid_data['qualifications'])} 条")
    print(f"   - 文件清单: {len(bid_data['documents'])} 项")

    if args.list_only:
        # 仅输出拆解结果
        print("\n" + "=" * 60)
        print("📋 招标文件拆解结果")
        print("=" * 60)

        print("\n📌 项目概况:")
        for key, val in bid_data.get("project", {}).items():
            print(f"  {key}: {val}")

        print(f"\n📊 评分标准（{len(bid_data['scores'])}项）:")
        for s in bid_data["scores"]:
            print(f"  [{s['序号']}] {s['评分项']} — {s['分值']}")

        print(f"\n🚫 废标条款（{len(bid_data['invalid_clauses'])}条）:")
        for c in bid_data["invalid_clauses"]:
            print(f"  - {c[:80]}...")

        print(f"\n📋 资格要求（{len(bid_data['qualifications'])}条）:")
        for q in bid_data["qualifications"]:
            print(f"  - {q[:80]}...")

        print(f"\n📄 文件清单（{len(bid_data['documents'])}项）:")
        for d in bid_data["documents"]:
            print(f"  - {d[:80]}...")

        print(f"\n⏰ 时间节点:")
        for key, val in bid_data.get("timeline", {}).items():
            print(f"  {key}: {val}")
        return

    # 解析企业资料库
    profile_data = parse_company_profile(profile_files)
    if profile_files:
        print(f"📁 企业资料库: {len(profile_files)} 个文件")
        print(f"   - 人员: {len(profile_data['team'])} 人")
        print(f"   - 业绩: {len(profile_data['performance'])} 项")
        print(f"   - 资质: {len(profile_data['qualifications'])} 项")

    # 条款映射
    print("🔗 条款映射...")
    clause_mapping = map_clauses_to_chapters(bid_data["scores"])

    # 评分分析
    print("📊 评分分析...")
    score_analysis, score_distribution = analyze_scores(
        bid_data["scores"], profile_data, clause_mapping
    )

    # 承诺链审计
    print("🔐 承诺链审计...")
    commitment_results = audit_commitments(bid_data, profile_data)
    print(f"   - 承诺项: {len(commitment_results)} 条")
    direct = sum(1 for c in commitment_results if c["状态"] == "✅ 已匹配")
    partial = sum(1 for c in commitment_results if c["状态"] == "⚠️ 部分匹配")
    missing = sum(1 for c in commitment_results if c["状态"] == "❌ 缺失")
    print(f"   - 已匹配: {direct} | 部分匹配: {partial} | 缺失: {missing}")

    # 废标/资质检查
    print("🚫 废标/资质检查...")
    disqualification_results = check_disqualifications(bid_data, profile_data)
    high_risk = sum(1 for d in disqualification_results if d["风险等级"] == "高")
    mid_risk = sum(1 for d in disqualification_results if d["风险等级"] == "中")
    print(f"   - 高危: {high_risk} | 中危: {mid_risk}")

    # 时间节点分析
    print("⏰ 时间规划...")
    timeline_analysis = analyze_timeline(bid_data, args.open_date)
    print(f"   - {timeline_analysis.get('status', '')} | 模式: {timeline_analysis.get('mode', '')}")

    # 生成报告
    print("📝 生成报告...")
    report = generate_report(
        bid_data=bid_data,
        profile_data=profile_data,
        score_analysis=score_analysis,
        score_distribution=score_distribution,
        clause_mapping=clause_mapping,
        commitment_results=commitment_results,
        disqualification_results=disqualification_results,
        timeline_analysis=timeline_analysis,
        bid_file=args.input,
        config=config,
    )

    # 输出
    if args.output:
        output_path = args.output
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(report)
        print(f"✅ 报告已保存: {output_path}")

        if args.json:
            json_path = output_path.replace(".md", ".json")
            if json_path == output_path:
                json_path += ".json"
            json_data = {
                "meta": file_data["meta"],
                "bid": {
                    "project": bid_data["project"],
                    "scores_count": len(bid_data["scores"]),
                    "invalid_clauses_count": len(bid_data["invalid_clauses"]),
                    "qualifications_count": len(bid_data["qualifications"]),
                },
                "score_distribution": score_distribution,
                "commitments": {
                    "total": len(commitment_results),
                    "matched": direct,
                    "partial": partial,
                    "missing": missing,
                },
                "disqualifications": {
                    "high_risk": high_risk,
                    "mid_risk": mid_risk,
                },
                "timeline": {
                    "status": timeline_analysis.get("status", ""),
                    "mode": timeline_analysis.get("mode", ""),
                    "remaining_days": timeline_analysis.get("remaining_days", None),
                },
            }
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
            print(f"✅ JSON报告已保存: {json_path}")
    else:
        # 终端输出摘要
        print("\n" + "=" * 60)
        print("📋 投标可行性分析摘要")
        print("=" * 60)

        print(f"\n⏰ {timeline_analysis.get('status', '')} | 建议: {timeline_analysis.get('mode', '')}")

        print(f"\n📊 评分分布: 总分 {score_distribution.get('总分', '?')}分")
        print(f"   技术分: {score_distribution.get('技术分', '?')}分 | 价格分: {score_distribution.get('价格分', '?')}分")

        if high_risk > 0:
            print(f"\n🔴 高危废标项: {high_risk} 项 — 必须处理")

        if missing > 0:
            print(f"❌ 承诺缺失: {missing} 项")

        print(f"\n💡 建议: 使用 `bid analyze {args.input} -o 报告.md` 导出完整报告")

    print("\n✅ 分析完成")


if __name__ == "__main__":
    main()