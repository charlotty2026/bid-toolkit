#!/usr/bin/env python3
"""
招标文件标准结构定义 v2.0
v2.0 变更（第二轮审核 P0 补完）：
  - P0-1: 新增「资格审查」独立章节（第3章，原3-6章顺延）
  - P0-2: 投标人须知新增「投标人须知前附表」节
  - P0-3: 投标人须知新增「投标保证金」节
  - P0-4: 投标人须知新增「质疑与投诉」节（仅 goods/services）
  - P0-5: 投标人须知新增「招标文件的澄清与修改」节
  - P0-6: 新增「政府采购政策落实」章（仅 goods/services）
  - get_chapters() 支持条件章节 + 条件小节，ID 自动重编号
从5份公开招标文件样本中归纳的标准框架。
支持三种项目类型：货物类(goods)、服务类(services)、工程类(engineering)。
"""

# 项目类型
PROJECT_TYPES = {
    "goods": "货物类",
    "services": "服务类",
    "engineering": "工程类",
}

# 标准章节（所有类型通用，共7章）
# 注意：id 字段在 get_chapters() 中会自动重编号，这里仅做占位
STANDARD_CHAPTERS = [
    {
        "id": 1,
        "title": "投标邀请",
        "alt_titles": ["招标公告", "投标邀请书"],
        "sections": [
            "项目基本情况",
            "招标范围",
            "投标人资格要求",
            "招标文件获取",
            "投标文件递交",
            "开标时间与地点",
            "联系方式",
        ],
        "key_fields": [
            "项目名称", "项目编号", "采购人", "采购代理机构",
            "预算金额", "资金来源", "投标截止时间", "开标地点",
            "获取文件方式", "获取文件时间", "联系人", "联系电话",
        ],
    },
    {
        "id": 2,
        "title": "投标人须知",
        "alt_titles": ["投标须知"],
        "sections": [
            "投标人须知前附表",          # P0-2
            "总则",
            "招标文件",
            "招标文件的澄清与修改",      # P0-5
            "投标文件编制",
            "投标文件递交",
            "投标保证金",                # P0-3
            "开标与评标",
            "废标条款（否决投标条件）",
            "中标通知与签约",
        ],
        # P0-4: 质疑与投诉仅政府采购类强制
        "conditional_sections": {
            "质疑与投诉": ["goods", "services"],
        },
        "key_fields": [
            "投标文件构成", "格式要求(字体/字号/页边距)",
            "密封要求", "投标保证金", "投标有效期",
            "份数要求", "签章要求",
        ],
    },
    # P0-1: 资格审查独立章节
    {
        "id": 3,
        "title": "资格审查",
        "alt_titles": ["资格要求及审查", "投标人资格审查"],
        "sections": [
            "资格要求",
            "资格审查方式（资格预审/资格后审）",
            "资格证明文件清单",
        ],
        "key_fields": [
            "资格条件", "审查方式", "证明文件清单",
        ],
    },
    {
        "id": 4,
        "title": "采购需求",
        "alt_titles": ["招标需求", "发包人要求", "技术条件"],
        "sections": [
            "项目概况",
            "技术规格/服务要求",
            "数量清单",
            "质量标准",
            "交货/服务期限",
            "验收标准",
            "售后服务",
        ],
        "key_fields": [
            "技术参数", "服务内容", "数量", "交货期",
            "质量标准", "验收标准", "售后要求",
        ],
    },
    {
        "id": 5,
        "title": "评标办法",
        "alt_titles": ["评标方法和评标标准", "开标、评标及定标办法"],
        "sections": [
            "评标方法",
            "评审因素及分值分配",
            "废标条件",
            "加分项",
        ],
        "key_fields": [
            "评标方法(最低评标价法/综合评分法)",
            "评分因素", "分值分配", "废标条件", "加分项",
        ],
    },
    {
        "id": 6,
        "title": "合同条款",
        "alt_titles": ["合同文本", "合同条款及格式"],
        "sections": [
            "合同标的",
            "付款方式",
            "交货/服务期限",
            "违约责任",
            "质量保证",
            "知识产权",
            "争议解决",
        ],
        "key_fields": [
            "付款方式", "交货期", "违约责任",
            "质保期", "知识产权归属", "争议解决方式",
        ],
    },
    {
        "id": 7,
        "title": "投标文件格式",
        "alt_titles": [],
        "sections": [
            "投标函",
            "授权委托书",
            "投标报价表",
            "技术方案格式",
            "资格证明文件",
            "承诺函",
            "其他附件",
        ],
        "key_fields": [
            "投标函格式", "授权委托书格式", "报价表格式",
            "技术方案格式", "资格证明文件清单",
        ],
    },
]

# P0-6: 政府采购政策落实（仅 goods/services）
GOV_PROCUREMENT_CHAPTER = {
    "id": "gov-policy",  # 占位，get_chapters() 会重编号
    "title": "政府采购政策落实",
    "alt_titles": ["政府采购政策"],
    "sections": [
        "促进中小企业发展",
        "监狱企业扶持政策",
        "残疾人福利性单位政策",
        "节能产品政府采购",
        "环境标志产品政府采购",
    ],
    "key_fields": [
        "小微企业价格扣除比例(6%-10%)", "监狱企业认定材料",
        "节能产品清单", "环境标志产品清单",
    ],
    "condition": ["goods", "services"],
}

# 工程类额外章节（ID 在 get_chapters() 中自动重编号）
ENGINEERING_EXTRA_CHAPTERS = [
    {
        "id": "eng-1",
        "title": "技术条件（工程建设标准）",
        "alt_titles": [],
        "sections": ["国家标准", "行业标准", "地方标准", "企业标准"],
        "key_fields": ["适用标准清单"],
    },
    {
        "id": "eng-2",
        "title": "图纸及勘察资料",
        "alt_titles": [],
        "sections": ["施工图纸", "地质勘察报告", "现场条件说明"],
        "key_fields": ["图纸清单", "勘察资料清单"],
    },
    {
        "id": "eng-3",
        "title": "工程量清单",
        "alt_titles": [],
        "sections": ["分部分项工程量清单", "措施项目清单", "其他项目清单", "规费税金清单"],
        "key_fields": ["工程量清单格式", "计量单位", "工程量计算规则"],
    },
    {
        "id": "eng-4",
        "title": "最高投标限价",
        "alt_titles": ["招标控制价"],
        "sections": ["限价说明", "限价明细"],
        "key_fields": ["最高投标限价金额", "编制依据"],
    },
]

# 法律法规依据
LEGAL_BASIS = {
    "政府采购法": {
        "full_name": "中华人民共和国政府采购法",
        "effective_date": "2003-01-01",
        "key_articles": [
            "第22条(供应商资格)",
            "第26条(采购方式)",
            "第46条(合同签订)",
            "第52-58条(质疑与投诉)",
        ],
    },
    "招标投标法": {
        "full_name": "中华人民共和国招标投标法",
        "effective_date": "2000-01-01",
        "key_articles": [
            "第24条(投标截止)",
            "第33条(投标要求)",
            "第45条(中标通知)",
        ],
    },
    "财政部87号令": {
        "full_name": "政府采购货物和服务招标投标管理办法",
        "effective_date": "2017-10-01",
        "key_articles": [
            "第14条(招标公告内容)",
            "第18条(招标文件提供期限≥5工作日)",
            "第20条(资格审查程序)",
            "第24条(履约验收)",
            "第27条(澄清与修改，投标截止15日前)",
            "第33条(投标保证金不超过预算2%)",
            "第37条(评标方法)",
            "第44条(评标委员会组成)",
            "第60条(废标情形)",
        ],
    },
}


def get_chapters(project_type="services"):
    """获取指定项目类型的完整章节列表（含条件章节/条件小节，ID自动重编号）"""
    chapters = []

    for ch in STANDARD_CHAPTERS:
        ch_copy = dict(ch)
        sections = list(ch.get("sections", []))
        # 合并条件小节
        for section, types in ch.get("conditional_sections", {}).items():
            if project_type in types:
                sections.append(section)
        ch_copy["sections"] = sections
        # 清理 conditional_sections（不需要暴露给外部）
        ch_copy.pop("conditional_sections", None)
        chapters.append(ch_copy)

    # P0-6: 政府采购政策落实（仅 goods/services）
    if project_type in GOV_PROCUREMENT_CHAPTER.get("condition", []):
        chapters.append(dict(GOV_PROCUREMENT_CHAPTER))

    # 工程类额外章节
    if project_type == "engineering":
        chapters.extend(dict(ch) for ch in ENGINEERING_EXTRA_CHAPTERS)

    # ID 自动重编号
    for i, ch in enumerate(chapters, 1):
        ch["id"] = i

    return chapters


def get_all_key_fields(project_type="services"):
    """获取所有需要填写的字段"""
    chapters = get_chapters(project_type)
    fields = {}
    for ch in chapters:
        fields[ch["title"]] = ch["key_fields"]
    return fields


def get_compliance_checklist(project_type="services"):
    """生成合规检查清单"""
    chapters = get_chapters(project_type)
    checklist = []
    for ch in chapters:
        for section in ch["sections"]:
            checklist.append({
                "chapter": f"第{ch['id']}章",
                "title": ch["title"],
                "section": section,
                "status": "pending",
                "mandatory": True,
            })
    return checklist


if __name__ == "__main__":
    import json
    print("=== 招标文件标准结构 v2.0 ===\n")
    for ptype, ptype_name in PROJECT_TYPES.items():
        chapters = get_chapters(ptype)
        print(f"【{ptype_name}】共{len(chapters)}章：")
        for ch in chapters:
            print(f"  第{ch['id']}章 {ch['title']}")
            for s in ch["sections"]:
                print(f"    - {s}")
        print()

    print("\n=== 法律法规依据 ===")
    for law_name, law_info in LEGAL_BASIS.items():
        print(f"  {law_name}（{law_info['effective_date']}起施行）")
        for art in law_info["key_articles"]:
            print(f"    - {art}")
