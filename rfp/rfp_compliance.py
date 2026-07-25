#!/usr/bin/env python3
"""
招标文件合规检查器 v2.0
v2.0 变更（第二轮审核 P0 补完）：
  - 新增 check_required_sections()：检查P0必备子节（前附表/保证金/澄清/资格审查/质疑投诉）
  - run_all_checks 新增 required_sections 检查项
检查招标文件是否合规：必备章节/必备子节/排他性条款/评分分值/废标条款/时间节点/资质要求。

用法：
  python rfp_compliance.py --rfp 招标文件.md -o 合规报告.json
  python rfp_compliance.py --rfp 招标文件.md --format text
"""

import argparse
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from rfp_structure import get_chapters, LEGAL_BASIS

# 排他性条款检测正则（v2.2 扩充适配真实招标文件表述）
EXCLUSIONARY_PATTERNS = [
    (r"必须为.{0,10}品牌", "指定品牌，涉嫌排他"),
    (r"指定.{0,10}品牌", "指定品牌，涉嫌排他"),
    (r"须为.{0,10}品牌", "指定品牌，涉嫌排他"),
    (r"应选用.{0,10}品牌", "限定品牌，涉嫌排他"),
    (r"应采用.{0,10}品牌", "限定品牌，涉嫌排他"),
    (r"指定.{0,10}厂家", "指定厂家，涉嫌排他"),
    (r"指定.{0,10}生产商", "指定生产商，涉嫌排他"),
    (r"须使用.{0,10}产品", "指定产品，涉嫌排他"),
    (r"必须采用.{0,10}产品", "指定产品，涉嫌排他"),
    (r"品牌要求.{0,10}[：:]", "品牌要求清单，涉嫌排他"),
    (r"参考品牌.{0,10}[：:]", "参考品牌清单，可能排他"),
    (r"推荐品牌.{0,10}[：:]", "推荐品牌清单，可能排他"),
    (r"注册[资资]本.{0,5}不低于.{0,10}万", "注册资本门槛，可能排他"),
    (r"注册[资资]金.{0,5}不低于.{0,10}万", "注册资金门槛，可能排他"),
    (r"注册[资资]本.{0,5}不少于.{0,10}万", "注册资本门槛，可能排他"),
    (r"注册[资资]金.{0,5}不少于.{0,10}万", "注册资金门槛，可能排他"),
    (r"须为.{0,10}上市公司", "上市公司要求，排他"),
    (r"须为.{0,10}大型企业", "企业规模要求，排他"),
    (r"具有.{0,10}级资质", "资质要求可能超出项目需要"),
    (r"近三年.{0,10}业绩.{0,5}不低于.{0,10}万", "业绩门槛，可能排他"),
    (r"近三年.{0,10}业绩.{0,5}不少于.{0,10}万", "业绩门槛，可能排他"),
    (r"近三年.{0,10}合同.{0,5}金额.{0,5}不低于", "合同金额门槛，可能排他"),
    (r"近三年.{0,10}项目.{0,5}不少于.{0,5}\d", "项目数量门槛，可能排他"),
    (r"具有.{0,20}年以上.{0,10}经验", "年限要求，可能排他"),
    (r"具有.{0,20}年以上.{0,10}经营", "经营年限要求，可能排他"),
    (r"获得.{0,15}认证", "特定认证要求，可能排他"),
    (r"须具有.{0,15}认证", "特定认证要求，可能排他"),
    (r"通过.{0,15}体系认证", "体系认证要求，可能排他"),
    (r"仅在.{0,10}有", "地域限制，涉嫌排他"),
    (r"注册地.{0,5}在.{0,10}[，,。]", "地域限制，涉嫌排他"),
    (r"本地企业", "本地企业要求，涉嫌排他"),
    (r"须在.{0,10}设立.{0,5}机构", "机构设立要求，可能排他"),
]

# 废标条款关键词（v2.2 扩充适配发改委标准文件+各地表述差异）
REJECTION_KEYWORDS = [
    "废标", "否决投标", "无效投标", "按废标处理", "不予受理",
    "不予接受", "视为放弃", "拒绝投标", "取消投标资格",
    "不予评审", "不予评标", "投标无效", "应予否决",
    "视为无效", "按无效投标处理", "不予计入",
]

# 时间节点关键词
TIME_KEYWORDS = ["投标截止", "开标时间", "投标有效期", "获取招标文件", "递交投标文件"]


def check_completeness(text, project_type="services"):
    """检查必备章节完整性"""
    chapters = get_chapters(project_type)
    results = []
    for ch in chapters:
        title_variants = [ch["title"]] + ch.get("alt_titles", [])
        found = any(v in text for v in title_variants)
        results.append({
            "chapter": f"第{ch['id']}章",
            "title": ch["title"],
            "found": found,
            "mandatory": True,
            "severity": "pass" if found else "fail",
            "message": f"{'✅ 找到' if found else '❌ 缺失'} 第{ch['id']}章 {ch['title']}",
        })
    return results


# P0 必备子节（按项目类型，goods/services含质疑投诉，engineering不含）
# 每个子节提供多个别名，适配真实招标文件的不同表述
REQUIRED_SECTIONS = {
    "goods": [
        ("投标人须知前附表", ["前附表", "须知前附表"]),
        ("招标文件的澄清与修改", ["澄清与修改", "澄清和修改", "招标文件的澄清", "澄清或修改"]),
        ("投标保证金", ["保证金", "投标担保"]),
        ("资格审查", ["资格要求", "资格证明", "供应商资格审查"]),
        ("质疑与投诉", ["质疑投诉", "质疑和投诉", "质疑、投诉", "提出质疑"]),
    ],
    "services": [
        ("投标人须知前附表", ["前附表", "须知前附表"]),
        ("招标文件的澄清与修改", ["澄清与修改", "澄清和修改", "招标文件的澄清", "澄清或修改"]),
        ("投标保证金", ["保证金", "投标担保"]),
        ("资格审查", ["资格要求", "资格证明", "供应商资格审查"]),
        ("质疑与投诉", ["质疑投诉", "质疑和投诉", "质疑、投诉", "提出质疑"]),
    ],
    "engineering": [
        ("投标人须知前附表", ["前附表", "须知前附表"]),
        ("招标文件的澄清与修改", ["澄清与修改", "澄清和修改", "招标文件的澄清", "澄清或修改"]),
        ("投标保证金", ["保证金", "投标担保"]),
        ("资格审查", ["资格要求", "资格证明", "供应商资格审查"]),
    ],
}


def check_required_sections(text, project_type="services"):
    """P0: 检查必备子节完整性（前附表/保证金/澄清/资格审查/质疑投诉）
    使用主标题+别名模糊匹配，适配真实招标文件的不同表述。
    """
    required = REQUIRED_SECTIONS.get(project_type, REQUIRED_SECTIONS["services"])
    results = []
    for section_tuple in required:
        primary, aliases = section_tuple
        all_variants = [primary] + aliases
        found = any(v in text for v in all_variants)
        results.append({
            "type": "必备子节",
            "section": primary,
            "found": found,
            "severity": "pass" if found else "fail",
            "message": f"{'✅ 找到' if found else '❌ 缺失'} 必备子节：{primary}",
            "suggestion": None if found else f"请补充「{primary}」相关内容",
        })
    return results


def check_exclusionary(text):
    """检测排他性条款"""
    results = []
    for pattern, reason in EXCLUSIONARY_PATTERNS:
        matches = list(re.finditer(pattern, text))
        for m in matches:
            # 取上下文
            start = max(0, m.start() - 20)
            end = min(len(text), m.end() + 20)
            context = text[start:end].replace('\n', ' ')
            results.append({
                "type": "排他性条款",
                "severity": "warn",
                "pattern": pattern,
                "reason": reason,
                "match": m.group(),
                "context": f"...{context}...",
                "suggestion": "请核查该条款是否与项目实际需要相关，无关则删除以避免排他性投诉",
            })
    return results


def check_scoring(text, project_type="services"):
    """检查评分分值合规

    修复v2.1：只从评分表/评分标准区域提取分值，不再扫全文。
    通过定位"评标办法"章节来限定扫描范围。
    """
    results = []

    # 1. 尝试定位评分区域：第5章评标办法 到 下一章/文件末尾
    scoring_section = None
    for pattern in [
        r'第[5五][章节][^\n]*?(?:评标|评审|评分)',
        r'(?:评标|评审)办法',
        r'评分(?:标准|细则|因素|办法)',
    ]:
        m = re.search(pattern, text)
        if m:
            start = m.start()
            next_chapter = re.search(r'\n第[6六][章节]', text[start:])
            end = start + next_chapter.start() if next_chapter else len(text)
            scoring_section = text[start:end]
            break

    if not scoring_section:
        scoring_section = text

    # 2. 在评分区域内提取分值
    score_pattern = r'(\d+)\s*分'
    exclude_contexts = ['分钟', '平方米', '平方千米', '分之', '分米', '百分', '分贝',
                        '分公司', '分析', '分配', '分别', '分解', '分泌', '分裂',
                        '分担', '分享', '分割', '分辨', '区分', '成分', '满分',
                        '得分', '加分', '扣分', '评分', '评标', '总分', '合计',
                        '共计', '最高分', '最低分', '标准分']
    scores = []
    for m in re.finditer(score_pattern, scoring_section):
        num = int(m.group(1))
        ctx_start = max(0, m.start() - 15)
        ctx_end = min(len(scoring_section), m.end() + 10)
        context = scoring_section[ctx_start:ctx_end]
        if '合计' in context or '总分' in context:
            continue
        if any(kw in context for kw in exclude_contexts):
            if any(kw in context for kw in ['评分因素', '评审因素', '分值', '评审项', '评分项']):
                pass
            else:
                continue
        if num > 100:
            continue
        if num == 0:
            continue
        scores.append(num)

    total = sum(scores)

    if scores:
        if total != 100:
            results.append({
                "type": "评分分值",
                "severity": "fail" if abs(total - 100) > 5 else "warn",
                "message": f"评分总分={total}，应为100分",
                "detail": f"找到{len(scores)}个评分项：{scores}",
                "suggestion": "调整分值分配使总分=100",
            })
        else:
            results.append({
                "type": "评分分值",
                "severity": "pass",
                "message": f"评分总分=100，合规",
            })

        price_section = re.search(r'价格分?[：:]\s*(\d+)', text)
        if price_section:
            price_score = int(price_section.group(1))
            warning = check_price_range(price_score, project_type)
            if warning:
                results.append({
                    "type": "评分分值",
                    "severity": "warn",
                    "message": warning,
                    "detail": f"当前价格分={price_score}，项目类型={project_type}",
                })
    else:
        results.append({
            "type": "评分分值",
            "severity": "warn",
            "message": "未找到评分分值信息",
            "suggestion": "招标文件应包含明确的评分标准",
        })
    return results


# 各项目类型价格分合理区间
PRICE_RANGES = {
    "goods": (30, 40),
    "services": (10, 20),
    "engineering": (15, 25),
}


def check_price_range(price_score, project_type="services"):
    """按项目类型检查价格分是否在合理区间"""
    lo, hi = PRICE_RANGES.get(project_type, (5, 50))
    if price_score > 50:
        return f"价格分={price_score}，过高（超过50分），可能违反价格分上限规定"
    if price_score < 5:
        return f"价格分={price_score}，过低（不足5分），评分可能不合理"
    if not (lo <= price_score <= hi):
        return f"价格分={price_score}，偏离{project_type}类常见区间（{lo}-{hi}分），请确认是否符合项目实际"
    return None


def check_rejection_clauses(text):
    """检查废标条款"""
    results = []
    found = any(kw in text for kw in REJECTION_KEYWORDS)
    if not found:
        results.append({
            "type": "废标条款",
            "severity": "fail",
            "message": "未找到废标/否决投标条款",
            "suggestion": "招标文件必须明确列出废标情形（87号令第60条）",
        })
    else:
        # 统计废标条款数量
        count = sum(text.count(kw) for kw in REJECTION_KEYWORDS)
        results.append({
            "type": "废标条款",
            "severity": "pass",
            "message": f"找到废标相关条款{count}处",
        })

    # 检查是否引用了法律依据
    if "第60条" in text or "六十条" in text:
        results.append({
            "type": "废标条款",
            "severity": "pass",
            "message": "废标条款引用了87号令第60条",
        })
    return results


def check_time_nodes(text):
    """检查时间节点完整性"""
    results = []
    for kw in TIME_KEYWORDS:
        found = kw in text
        if not found:
            results.append({
                "type": "时间节点",
                "severity": "warn",
                "message": f"未找到「{kw}」相关时间信息",
                "suggestion": f"请补充{kw}的具体时间",
            })
        else:
            results.append({
                "type": "time_node",
                "severity": "pass",
                "message": f"已包含「{kw}」时间信息",
            })
    return results


def check_qualification(text):
    """检查资质要求合理性"""
    results = []
    # 检查资质要求是否过于具体
    qual_patterns = [
        (r"具备.{0,5}ISO\s*\d+", "ISO认证要求——确认是否与项目相关"),
        (r"具备.{0,5}CMMI.{0,5}级", "CMMI认证要求——通常仅适用于软件开发项目"),
        (r"具备.{0,5}建筑企业资质", "建筑业资质——确认是否适用于本项目类型"),
        (r"具有.{0,10}项以上.{0,10}业绩", "业绩数量要求——确认门槛是否合理"),
    ]
    for pattern, reason in qual_patterns:
        if re.search(pattern, text):
            results.append({
                "type": "资质要求",
                "severity": "warn",
                "message": reason,
                "suggestion": "请确认该资质要求与项目实际需要匹配",
            })

    # 检查是否列出了基本资格要求
    basic_qual = any(kw in text for kw in ["营业执照", "独立法人", "供应商资格"])
    if not basic_qual:
        results.append({
            "type": "资质要求",
            "severity": "warn",
            "message": "未找到基本资格要求",
            "suggestion": "应包含营业执照/独立法人等基本资格要求",
        })
    return results


def check_legal_basis(text):
    """检查法律法规引用"""
    results = []
    for law_name, law_info in LEGAL_BASIS.items():
        found = law_name in text or law_info["full_name"] in text
        if not found:
            results.append({
                "type": "法律依据",
                "severity": "warn",
                "message": f"未引用{law_name}",
                "suggestion": f"建议引用{law_info['full_name']}",
            })
        else:
            results.append({
                "type": "法律依据",
                "severity": "pass",
                "message": f"已引用{law_name}",
            })
    return results


# 红线2: 评审模糊检测——评分标准中不得出现模糊、不可量化的表述
# 法律依据：87号令第55条——评审因素应当细化和量化，与采购需求对应
VAGUE_SCORING_PATTERNS = [
    (r"酌情给分", "「酌情给分」属于模糊表述，应明确具体分值或计算公式"),
    (r"酌情.{0,4}扣分", "「酌情扣分」属于模糊表述，应明确扣分标准"),
    (r"视情况.{0,4}给分", "「视情况给分」属于模糊表述，应量化评分条件"),
    (r"由评委.{0,6}确定", "「由评委确定」缺少客观标准，应提供量化依据"),
    (r"由评标委员会.{0,6}确定.{0,4}分", "评委自由裁量分值过大，应限定范围或给计算公式"),
    (r"评委.{0,4}综合.{0,4}评定", "「评委综合评定」无量化标准，应细化评分因素"),
    (r"优.{0,4}分.{0,4}良.{0,4}分.{0,4}中.{0,4}分.{0,4}差", "等级评分法应明确各等级的具体判定标准"),
    (r"较好.{0,6}分", "「较好」无客观标准，应描述具体达到什么指标得多少分"),
    (r"一般.{0,4}分.{0,4}较好.{0,4}分", "等级评分缺少判定标准，应量化每个等级的边界"),
]


def check_vague_scoring(text):
    """红线2: 检查评分标准中是否存在模糊、不可量化的表述

    法律依据：87号令第55条——评审因素应当细化和量化，与采购需求对应。
    评分标准中出现"酌情""视情况""由评委确定"等模糊表述，投标人无法准备、
    评委自由裁量空间过大，容易被质疑/投诉。
    """
    results = []

    # 定位评分区域
    scoring_section = None
    for pattern in [
        r'第[5五][章节][^\n]*?(?:评标|评审|评分)',
        r'(?:评标|评审)办法',
        r'评分(?:标准|细则|因素|办法)',
    ]:
        m = re.search(pattern, text)
        if m:
            start = m.start()
            next_chapter = re.search(r'\n第[6六][章节]', text[start:])
            end = start + next_chapter.start() if next_chapter else len(text)
            scoring_section = text[start:end]
            break

    search_text = scoring_section if scoring_section else text

    for pattern, reason in VAGUE_SCORING_PATTERNS:
        matches = list(re.finditer(pattern, search_text))
        for m in matches:
            start = max(0, m.start() - 30)
            end = min(len(search_text), m.end() + 30)
            context = search_text[start:end].replace('\n', ' ')
            results.append({
                "type": "评审模糊",
                "severity": "warn",
                "reason": reason,
                "match": m.group(),
                "context": f"...{context}...",
                "suggestion": "将模糊表述替换为可量化的评分标准（具体分值/计算公式/明确判定条件）",
            })

    if not results:
        results.append({
            "type": "评审模糊",
            "severity": "pass",
            "message": "评分标准中未发现模糊表述",
        })

    return results


# 红线4: 等标期计算——招标公告发布至投标截止不得少于20日
# 法律依据：87号令第25条——公开招标公告发布之日起至投标截止之日止不得少于20日
# 竞争性谈判/询价另行规定，本工具以公开招标为主
DATE_PATTERN = r'(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})[日号]?'


def check_bid_period(text):
    """红线4: 检查等标期（招标公告发布至投标截止的时间间隔）

    法律依据：87号令第25条——公开招标的等标期不得少于20日。
    本函数尝试从文本中提取公告发布日期和投标截止日期，计算间隔天数。
    """
    from datetime import datetime, timedelta

    results = []

    # 尝试提取投标截止时间
    deadline_patterns = [
        r'投标截止[^\n]*?(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})',
        r'递交投标文件.{0,10}截止[^\n]*?(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})',
        r'开标时间[^\n]*?(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})',
    ]

    deadline_date = None
    for pattern in deadline_patterns:
        m = re.search(pattern, text)
        if m:
            try:
                year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
                deadline_date = datetime(year, month, day)
                break
            except ValueError:
                continue

    # 尝试提取公告发布时间
    publish_patterns = [
        r'招标公告发布[^\n]*?(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})',
        r'公告发布.{0,6}日期[^\n]*?(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})',
        r'发布.{0,6}招标公告[^\n]*?(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})',
        r'招标公告.{0,10}(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})',
    ]

    publish_date = None
    for pattern in publish_patterns:
        m = re.search(pattern, text)
        if m:
            try:
                year, month, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
                publish_date = datetime(year, month, day)
                break
            except ValueError:
                continue

    # 如果两个日期都找到了，计算等标期
    if deadline_date and publish_date:
        period = (deadline_date - publish_date).days
        if period < 20:
            results.append({
                "type": "等标期",
                "severity": "fail",
                "message": f"等标期仅{period}天，不足20日（87号令第25条）",
                "detail": f"公告发布日：{publish_date.strftime('%Y-%m-%d')}，投标截止日：{deadline_date.strftime('%Y-%m-%d')}",
                "suggestion": "延长投标截止时间至公告发布后不少于20日",
            })
        else:
            results.append({
                "type": "等标期",
                "severity": "pass",
                "message": f"等标期{period}天，满足≥20日要求",
                "detail": f"公告发布日：{publish_date.strftime('%Y-%m-%d')}，投标截止日：{deadline_date.strftime('%Y-%m-%d')}",
            })

    elif deadline_date and not publish_date:
        # 只有截止日期，检查是否有"自公告发布之日起XX日"的表述
        period_pattern = re.search(r'不少于(\d+)日|不得少于(\d+)日', text)
        if period_pattern:
            stated_period = int(period_pattern.group(1) or period_pattern.group(2))
            if stated_period < 20:
                results.append({
                    "type": "等标期",
                    "severity": "fail",
                    "message": f"文件载明等标期{stated_period}日，不足20日（87号令第25条）",
                    "suggestion": "修改为不少于20日",
                })
            else:
                results.append({
                    "type": "等标期",
                    "severity": "pass",
                    "message": f"文件载明等标期{stated_period}日，满足≥20日要求",
                })
        else:
            results.append({
                "type": "等标期",
                "severity": "warn",
                "message": "找到投标截止日期但未找到公告发布日期，无法计算等标期",
                "detail": f"投标截止日：{deadline_date.strftime('%Y-%m-%d')}",
                "suggestion": "请确保招标公告发布至投标截止不少于20日（87号令第25条）",
            })

    elif not deadline_date and not publish_date:
        # 两个日期都没找到，检查前附表中的占位符
        if "投标截止" in text or "开标时间" in text:
            results.append({
                "type": "等标期",
                "severity": "warn",
                "message": "文件包含投标截止/开标时间信息但未提取到具体日期，请人工确认等标期≥20日",
                "suggestion": "法律依据：87号令第25条，公开招标等标期不得少于20日",
            })
        else:
            results.append({
                "type": "等标期",
                "severity": "warn",
                "message": "未找到投标截止时间和公告发布时间",
                "suggestion": "招标文件应明确投标截止时间，且公告发布至截止不少于20日（87号令第25条）",
            })

    return results


def run_all_checks(text, project_type="services"):
    """运行所有检查"""
    report = {
        "check_time": __import__("datetime").datetime.now().isoformat(),
        "project_type": project_type,
        "text_length": len(text),
        "checks": {
            "completeness": check_completeness(text, project_type),
            "required_sections": check_required_sections(text, project_type),
            "exclusionary": check_exclusionary(text),
            "scoring": check_scoring(text, project_type),
            "vague_scoring": check_vague_scoring(text),
            "rejection": check_rejection_clauses(text),
            "bid_period": check_bid_period(text),
            "time_nodes": check_time_nodes(text),
            "qualification": check_qualification(text),
            "legal_basis": check_legal_basis(text),
        },
    }

    # 统计
    all_items = []
    for check_list in report["checks"].values():
        all_items.extend(check_list if isinstance(check_list, list) else [check_list])

    fail_count = sum(1 for item in all_items if item.get("severity") == "fail")
    warn_count = sum(1 for item in all_items if item.get("severity") == "warn")
    pass_count = sum(1 for item in all_items if item.get("severity") == "pass")

    report["summary"] = {
        "total": len(all_items),
        "pass": pass_count,
        "warn": warn_count,
        "fail": fail_count,
        "verdict": "不合规" if fail_count > 0 else ("需修改" if warn_count > 0 else "合规"),
    }
    return report


def format_text_report(report):
    """格式化为可读文本报告"""
    lines = []
    s = report["summary"]
    lines.append("=" * 60)
    lines.append(f"  招标文件合规检查报告")
    lines.append(f"  检查时间：{report['check_time']}")
    lines.append(f"  项目类型：{report['project_type']}")
    lines.append(f"  结论：{s['verdict']}（✅{s['pass']} ⚠️{s['warn']} ❌{s['fail']}）")
    lines.append("=" * 60)

    for check_name, items in report["checks"].items():
        if not items:
            continue
        lines.append(f"\n【{check_name}】")
        if isinstance(items, list):
            for item in items:
                sev = item.get("severity", "")
                icon = {"pass": "✅", "warn": "⚠️", "fail": "❌"}.get(sev, "•")
                # 排他性条目用 reason + match，其他用 message
                msg = item.get("message", "")
                if not msg and item.get("reason"):
                    msg = f"{item['reason']}（匹配：「{item.get('match', '')}」）"
                lines.append(f"  {icon} {msg}")
                if item.get("suggestion"):
                    lines.append(f"     -> {item['suggestion']}")
                if item.get("context"):
                    lines.append(f"     上下文：{item['context']}")
        else:
            lines.append(f"  {items}")

    return "\n".join(lines)


def read_file(path):
    """读取md或docx文件"""
    if path.endswith('.docx'):
        try:
            from docx import Document
            doc = Document(path)
            return '\n'.join(p.text for p in doc.paragraphs)
        except ImportError:
            print("错误：读取docx需要python-docx库")
            sys.exit(1)
    else:
        with open(path, 'r', encoding='utf-8') as f:
            return f.read()


def main():
    parser = argparse.ArgumentParser(description='招标文件合规检查器')
    parser.add_argument('--rfp', required=True, help='招标文件路径（.md或.docx）')
    parser.add_argument('--type', choices=['goods', 'services', 'engineering'],
                        default='services', help='项目类型')
    parser.add_argument('-o', '--output', help='输出报告路径（JSON）')
    parser.add_argument('--format', choices=['json', 'text'], default='text',
                        help='输出格式')
    args = parser.parse_args()

    text = read_file(args.rfp)
    report = run_all_checks(text, args.type)

    if args.format == 'json' or args.output:
        if args.output:
            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(report, f, ensure_ascii=False, indent=2)
            print(f"✅ 报告已保存：{args.output}")
        else:
            print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(format_text_report(report))


if __name__ == '__main__':
    main()
