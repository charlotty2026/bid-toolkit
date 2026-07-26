#!/usr/bin/env python3
"""
招标文件生成器 v2.0
v2.0 变更（第二轮审核 P0 补完）：
  - 前附表自动生成表格（不再用占位符）
  - 政府采购政策落实章节自动生成政策内容
  - 保证金/澄清/质疑/资格审查等新节加法律依据提示
  - 章节ID自动适配 rfp_structure v2.0 的新编号
从项目信息自动生成标准招标文件（Markdown + Word）。

用法：
  python rfp_generator.py --type services --project "XX后勤服务项目" --budget 500000
  python rfp_generator.py --type services --project "XX项目" --budget 500000 --docx
  python rfp_generator.py --config project.json -o 招标文件.md
"""

import argparse
import json
import os
import sys
from datetime import datetime, timedelta

# 导入标准结构
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from rfp_structure import (
    PROJECT_TYPES, STANDARD_CHAPTERS, ENGINEERING_EXTRA_CHAPTERS,
    LEGAL_BASIS, get_chapters,
)
from rfp_templates import (
    generate_procurement_content,
    generate_contract_content,
    generate_bid_format_content,
    generate_invitation_content,
    generate_notice_content,
    generate_qualification_content,
    generate_evaluation_content,
    generate_engineering_content,
)

# ===== 评分标准模板（分值汇总，用于JSON同步）=====
SCORING_TEMPLATES = {
    "goods": {"价格分": 30, "技术分": 50, "商务分": 20},
    "services": {"价格分": 10, "技术分": 60, "商务分": 30},
    "engineering": {"价格分": 20, "技术分": 60, "商务分": 20},
}

# ===== 评分标准明细（四列拆分法：评分项/分值/得分条件/投标人需提供的材料）=====
DETAILED_SCORING = {
    "services": [
        ("投标报价", 10, "满足招标文件要求的最低报价为基准价，按公式计算得分", "价格文件（投标报价表）"),
        ("服务方案", 20, "方案完整、科学、可行，覆盖全部服务需求，缺项扣分", "服务方案说明书"),
        ("人员配置", 15, "项目团队人员资质和数量满足要求，关键岗位持证上岗", "人员简历表、资格证书复印件"),
        ("服务承诺", 10, "服务质量承诺明确，响应时间合理，有量化指标", "服务承诺书"),
        ("应急预案", 10, "突发事件应对方案完整可操作，含分级响应机制", "应急预案文件"),
        ("管理制度", 5, "内部管理制度健全（考勤/培训/考核/交接班）", "管理制度文件"),
        ("企业业绩", 10, "近三年同类项目业绩，每个有效业绩得X分，最高10分", "业绩证明材料（合同复印件+验收报告）"),
        ("企业资质", 10, "持有相关资质证书（如ISO9001/人力资源许可证等）", "资质证书复印件"),
        ("财务状况", 5, "财务状况良好，近三年无亏损", "近三年财务报表"),
        ("信用记录", 5, "无不良信用记录（信用中国查询无失信/违法记录）", "信用中国查询截图"),
    ],
    "goods": [
        ("投标报价", 30, "满足招标文件要求且最低价为基准价，按公式计算", "价格文件（投标报价表）"),
        ("技术参数响应", 25, "技术参数完全响应招标要求，带★项全部满足，一般项偏离逐项扣分", "技术规格响应/偏离表"),
        ("产品质量", 15, "提供产品认证及第三方检测报告，认证齐全得满分", "产品检测报告、认证证书复印件"),
        ("售后服务方案", 10, "保修期≥X年，故障响应≤X小时，有本地服务网点", "售后服务方案、网点证明材料"),
        ("企业业绩", 8, "近三年同类货物供货业绩，每个有效业绩得X分", "业绩证明材料（合同复印件）"),
        ("企业资质", 7, "营业执照经营范围相符，相关资质齐全", "营业执照、资质证书复印件"),
        ("交货期承诺", 5, "交货期优于招标文件要求的得满分", "交货期承诺函"),
    ],
    "engineering": [
        ("投标报价", 20, "满足招标文件要求且最低价为基准价，按公式计算", "价格文件（投标报价表）"),
        ("施工组织方案", 20, "施工组织设计完整，技术方案科学可行，进度计划合理", "施工组织设计文件"),
        ("项目经理", 15, "持有相应建造师证书，近三年主持过同类工程", "项目经理简历、资格证书、业绩证明"),
        ("技术人员配置", 10, "技术团队专业齐全（施工/质量/安全），持证上岗", "人员简历表、资格证书"),
        ("质量保证措施", 10, "质量管理体系健全，措施具体可行，有验收标准", "质量保证方案"),
        ("安全文明施工", 5, "安全生产措施完善，文明施工方案到位", "安全施工方案"),
        ("企业业绩", 8, "近三年同类工程业绩，每个有效业绩得X分", "业绩证明材料（合同复印件+验收报告）"),
        ("企业资质", 7, "施工资质等级满足招标文件要求", "资质证书复印件"),
        ("财务状况", 5, "财务状况良好，具备履约能力", "近三年财务报表"),
    ],
}

# ===== 废标条款分类清单（一票否决清单）=====
# 三类：资格性废标 / 符合性废标 / 格式性废标
REJECTION_CLAUSES = {
    "一、资格性废标（资格条件不符）": [
        "投标人不符合招标文件规定的资格条件的",
        "投标人未按招标文件要求提供资格证明文件的",
        "联合体投标未提交联合体协议或联合体成员不符合资格要求的",
        "投标人处于禁止投标期内（被列入失信被执行人/重大税收违法/政府采购严重违法失信名单）",
    ],
    "二、符合性废标（实质性要求不符）": [
        "投标报价超过最高投标限价的",
        "投标报价低于成本且无法说明理由的",
        "投标有效期不足的",
        "投标文件存在重大偏差的（关键技术参数带★项不满足）",
        "交货期/服务期/工期不满足招标文件要求的",
        "存在串通投标行为的",
        "提供虚假材料谋取中标的",
    ],
    "三、格式性废标（形式要求不符）": [
        "投标文件未按招标文件要求密封的",
        "投标文件未按要求签字和盖章的",
        "未按要求提交投标保证金的（如要求）",
        "投标文件未按招标文件规定的格式编制的",
        "投标文件份数不符合招标文件要求的",
    ],
}

# P0 新节法律依据提示
SECTION_HINTS = {
    "招标文件的澄清与修改": "法律依据：87号令第27条，采购人可在投标截止15个工作日前发出澄清/修改",
    "投标保证金": "法律依据：87号令第33条，不超过预算金额的2%，且不得超过法定上限",
    "质疑与投诉": "法律依据：政府采购法第52-58条，质疑答复时限7个工作日，投诉处理时限30个工作日",
    "资格审查方式（资格预审/资格后审）": "法律依据：87号令第20条，资格审查由采购人或采购代理机构依法开展",
}

def generate_front_sheet(project_info, project_type):
    """P0-2: 生成投标人须知前附表（自动填充已知信息）"""
    rows = [
        ("1", "项目名称", str(project_info.get("project_name", "____"))),
        ("2", "项目编号", str(project_info.get("project_id", "____"))),
        ("3", "采购人", str(project_info.get("purchaser", "____"))),
        ("4", "预算金额", f"{project_info.get('budget', '____')}元"),
        ("5", "项目类型", PROJECT_TYPES.get(project_type, "服务类")),
        ("6", "投标截止时间", "____年____月____日 ____:____"),
        ("7", "开标地点", "____（详见招标公告）"),
        ("8", "投标有效期", "自投标截止日起90日历日"),
        ("9", "投标保证金", "不超过预算金额的2%（87号令第33条）"),
        ("10", "投标文件份数", "正本1份，副本4份"),
        ("11", "评标方法", "综合评分法"),
        ("12", "资金来源", str(project_info.get("fund_source", "财政资金"))),
    ]
    lines = ["| 序号 | 条款名称 | 内容 |", "|------|---------|------|"]
    for row in rows:
        lines.append(f"| {row[0]} | {row[1]} | {row[2]} |")
    lines.append("")
    lines.append("> 前附表与正文条款不一致时，以前附表为准。")
    return "\n".join(lines)

def generate_gov_policy_content():
    """P0-6: 生成政府采购政策落实章节内容"""
    policies = [
        ("促进中小企业发展",
         "对小型、微型企业产品的价格给予 **6%-10%的扣除**，用扣除后的价格参与评审。",
         "《政府采购促进中小企业发展管理办法》"),
        ("监狱企业扶持政策",
         "监狱企业视同小型、微型企业，享受相同的价格扣除政策。",
         "《关于政府采购支持监狱企业发展有关问题的通知》"),
        ("残疾人福利性单位政策",
         "残疾人福利性单位视同小型、微型企业，享受相同的价格扣除政策。",
         "《关于促进残疾人就业政府采购政策的通知》"),
        ("节能产品政府采购",
         "属于强制采购节能产品范围的，应当采购列入 **节能产品政府采购清单** 的产品。",
         "《关于调整节能产品政府采购清单的通知》"),
        ("环境标志产品政府采购",
         "属于环境标志产品政府采购清单中的产品，在同等条件下 **优先采购**。",
         "《关于调整环境标志产品政府采购清单的通知》"),
    ]
    lines = ["本项目落实以下政府采购政策：\n"]
    for title, content, basis in policies:
        lines.append(f"### {title}\n")
        lines.append(f"{content}\n")
        lines.append(f"> 法律依据：{basis}\n")
    return "\n".join(lines)

def generate_markdown(project_info, project_type="services"):
    """生成Markdown格式招标文件"""
    chapters = get_chapters(project_type)
    type_name = PROJECT_TYPES.get(project_type, "服务类")
    scoring = SCORING_TEMPLATES.get(project_type, SCORING_TEMPLATES["services"])

    lines = []
    # 文件头
    lines.append(f"# {project_info.get('project_name', '____')}招标文件\n")
    lines.append(f"> 项目编号：{project_info.get('project_id', '____')}")
    lines.append(f"> 采购人：{project_info.get('purchaser', '____')}")
    lines.append(f"> 预算金额：{project_info.get('budget', '____')}元")
    lines.append(f"> 生成时间：{datetime.now().strftime('%Y-%m-%d')}\n")
    lines.append("---\n")

    # 法律法规依据
    lines.append("## 编制依据\n")
    for law_name, law_info in LEGAL_BASIS.items():
        lines.append(f"- **{law_info['full_name']}**（{law_info['effective_date']}起施行）")
        for art in law_info["key_articles"]:
            lines.append(f"  - {art}")
    lines.append("")

    # 各章内容
    for ch in chapters:
        lines.append(f"\n---\n")
        lines.append(f"# 第{ch['id']}章 {ch['title']}\n")
        lines.append(f"> {ch.get('说明', '')}\n")

        for section in ch["sections"]:
            lines.append(f"## {section}\n")

            # P0-2: 投标人须知前附表 -> 自动生成表格
            if section == "投标人须知前附表":
                lines.append(generate_front_sheet(project_info, project_type))
                lines.append("")
                continue

            # P0-6: 政府采购政策落实 -> 自动生成政策内容
            if ch["title"] == "政府采购政策落实":
                lines.append(generate_gov_policy_content())
                lines.append("")
                continue

            # P0-核心: 采购需求 -> 三类差异化模板
            if ch["title"] == "采购需求":
                lines.append(generate_procurement_content(project_info, project_type, section))
                lines.append("")
                continue

            # P0-核心: 合同条款 -> 三类差异化模板
            if ch["title"] == "合同条款":
                lines.append(generate_contract_content(project_info, project_type, section))
                lines.append("")
                continue

            # P0-核心: 投标文件格式 -> 三类差异化模板
            if ch["title"] == "投标文件格式":
                lines.append(generate_bid_format_content(project_info, project_type, section))
                lines.append("")
                continue

            # P0-核心: 投标邀请 -> 通用模板
            if ch["title"] == "投标邀请":
                lines.append(generate_invitation_content(project_info, project_type, section))
                lines.append("")
                continue

            # P0-核心: 投标人须知 -> 通用模板
            if ch["title"] == "投标人须知":
                lines.append(generate_notice_content(project_info, project_type, section))
                lines.append("")
                continue

            # P0-核心: 资格审查 -> 通用模板
            if ch["title"] == "资格审查":
                lines.append(generate_qualification_content(project_info, project_type, section))
                lines.append("")
                continue

            # P0-核心: 评标办法 -> 通用模板
            if ch["title"] == "评标办法":
                lines.append(generate_evaluation_content(project_info, project_type, section))
                lines.append("")
                continue

            # P0-核心: 工程类专属章节 -> 通用模板
            if project_type == "engineering" and ch["title"] in (
                "技术条件（工程建设标准）", "工程图纸及勘察资料", "工程量清单", "最高投标限价"
            ):
                lines.append(generate_engineering_content(project_info, project_type, section))
                lines.append("")
                continue

            # 兜底：无模板章节（不应到达此处）
            lines.append(f"本节内容根据项目实际情况确定，投标人应按招标文件要求作出实质性响应。\n")

        # 关键字段提示
        if ch.get("key_fields"):
            lines.append("### 关键信息字段\n")
            for field in ch["key_fields"]:
                lines.append(f"- [ ] {field}")
            lines.append("")

    # 评分标准（四列拆分法）
    lines.append("\n---\n")
    lines.append("# 附件一：评分标准（综合评分法）\n")
    lines.append(f"> 项目类型：{type_name}，总分100分\n")
    detailed = DETAILED_SCORING.get(project_type, DETAILED_SCORING["services"])
    lines.append("| 序号 | 评分项 | 分值 | 得分条件 | 投标人需提供的材料 |")
    lines.append("|------|--------|------|----------|-------------------|")
    total = 0
    for idx, (item, score, condition, material) in enumerate(detailed, 1):
        lines.append(f"| {idx} | {item} | {score} | {condition} | {material} |")
        total += score
    lines.append(f"| — | **合计** | **{total}** | | |")
    lines.append("")
    lines.append("> ⚠️ 投标人应对照本表逐项准备材料，缺项可能导致该项零分。")

    # 废标条款（一票否决清单，分类列出）
    lines.append("\n---\n")
    lines.append("# 附件二：废标条款（⚠️ 一票否决清单 — 投标人必读）\n")
    lines.append("出现以下任一情形的投标文件将被否决（废标），请投标人逐项自查：\n")
    clause_num = 1
    for category, items in REJECTION_CLAUSES.items():
        lines.append(f"### {category}\n")
        for item in items:
            lines.append(f"{clause_num}. ⚠️ {item}")
            clause_num += 1
        lines.append("")
    lines.append(f"> 法律依据：{LEGAL_BASIS['财政部87号令']['full_name']}第60条")
    lines.append(f"> 投标人应在递交投标文件前逐项自查上述全部条款，任一不符即面临废标风险。")

    return "\n".join(lines)


def _convert_numbering(markdown_text, numbering="arabic"):
    """转换Markdown中的编号列表项格式。

    arabic:  1. 2. 3.       （默认，Word用List Number样式自动编号）
    multi:   1.1 1.2 1.3     （多级编号，文本内嵌）
    chinese: 一、 二、 三、   （中文编号，文本内嵌）
    none:    去掉编号         （纯文本段落）
    """
    if numbering == "arabic":
        return markdown_text

    lines = markdown_text.split('\n')
    result = []
    list_counter = 0
    chapter_num = 0
    cn_nums = [
        '一', '二', '三', '四', '五', '六', '七', '八', '九', '十',
        '十一', '十二', '十三', '十四', '十五', '十六', '十七', '十八',
        '十九', '二十', '二十一', '二十二', '二十三', '二十四', '二十五',
        '二十六', '二十七', '二十八', '二十九', '三十',
    ]

    for line in lines:
        stripped = line.strip()

        # 追踪章节号（用于 multi 模式）
        if stripped.startswith('# 第') and '章' in stripped:
            chapter_num += 1
            list_counter = 0
            result.append(line)
            continue

        # 检测编号列表项（格式：数字. 文本）
        is_numbered = (
            bool(stripped)
            and len(stripped) > 2
            and stripped[0:2].rstrip('.').isdigit()
            and '. ' in stripped[:5]
        )

        if is_numbered:
            idx = stripped.index('. ')
            content = stripped[idx + 2:]
            list_counter += 1

            if numbering == "multi":
                result.append(f"{chapter_num}.{list_counter} {content}")
            elif numbering == "chinese":
                num = cn_nums[list_counter - 1] if list_counter <= len(cn_nums) else str(list_counter)
                result.append(f"{num}、{content}")
            elif numbering == "none":
                result.append(content)
        else:
            # 非编号非空行 → 重置列表计数器（空行不重置）
            if stripped:
                list_counter = 0
            result.append(line)

    return '\n'.join(result)


def _create_numbering_instance(doc):
    """为新的编号列表创建独立的编号实例，确保编号从1重新开始。

    python-docx的List Number样式所有段落共用一个编号序列，
    不会在新列表处重置——第一章编到26，第二章接着从27开始。
    此函数创建新的numId + startOverride解决该问题。
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    numbering = doc.part.numbering_part.element

    existing_nums = numbering.findall(qn('w:num'))
    max_num_id = max([int(n.get(qn('w:numId'))) for n in existing_nums], default=0)

    abstract_nums = numbering.findall(qn('w:abstractNum'))
    if not abstract_nums:
        return None
    abs_id = abstract_nums[0].get(qn('w:abstractNumId'))

    new_num_id = max_num_id + 1

    num = OxmlElement('w:num')
    num.set(qn('w:numId'), str(new_num_id))

    abs_ref = OxmlElement('w:abstractNumId')
    abs_ref.set(qn('w:val'), abs_id)
    num.append(abs_ref)

    lvl_override = OxmlElement('w:lvlOverride')
    lvl_override.set(qn('w:ilvl'), '0')
    start_override = OxmlElement('w:startOverride')
    start_override.set(qn('w:val'), '1')
    lvl_override.append(start_override)
    num.append(lvl_override)

    numbering.append(num)
    return new_num_id


def _apply_numbering(paragraph, num_id):
    """给段落应用指定的编号实例（覆盖样式默认的numId）"""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    pPr = paragraph._element.get_or_add_pPr()
    numPr = OxmlElement('w:numPr')
    ilvl = OxmlElement('w:ilvl')
    ilvl.set(qn('w:val'), '0')
    numId_el = OxmlElement('w:numId')
    numId_el.set(qn('w:val'), str(num_id))
    numPr.append(ilvl)
    numPr.append(numId_el)
    pPr.append(numPr)


def generate_docx(markdown_text, output_path):
    """将Markdown转为Word文档，标题用原生Heading样式，插入TOC域"""
    try:
        from docx import Document
        from docx.shared import Pt, Cm, RGBColor
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
    except ImportError:
        print("错误：需要python-docx库，请运行 pip install python-docx")
        return False

    doc = Document()

    # 页面设置
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    # 初始化样式
    style_normal = doc.styles['Normal']
    style_normal.font.name = '宋体'
    style_normal.font.size = Pt(12)
    style_normal.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    for level in range(1, 5):
        style = doc.styles[f'Heading {level}']
        style.font.name = '黑体'
        style.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
        style.font.color.rgb = RGBColor(0, 0, 0)

    # 先插入TOC域（占位，打开Word后F9更新）
    _insert_toc_field(doc)

    # 解析Markdown并写入
    lines = markdown_text.split('\n')
    in_table = False
    table_rows = []
    in_numbered_list = False
    current_num_id = None

    for line in lines:
        stripped = line.strip()

        # 判断是否是编号列表项（用于检测新列表起点）
        is_numbered = (
            bool(stripped)
            and len(stripped) > 2
            and stripped[0:2].rstrip('.').isdigit()
            and '. ' in stripped[:5]
        )
        if not is_numbered:
            in_numbered_list = False

        # 跳过空行
        if not stripped:
            if in_table and table_rows:
                _write_table(doc, table_rows)
                table_rows = []
                in_table = False
            continue

        # 表格行
        if stripped.startswith('|'):
            in_table = True
            if not stripped.startswith('|--') and not stripped.startswith('| -'):
                cells = [c.strip() for c in stripped.split('|')[1:-1]]
                table_rows.append(cells)
            continue
        elif in_table and table_rows:
            _write_table(doc, table_rows)
            table_rows = []
            in_table = False

        # 标题
        if stripped.startswith('# ') and not stripped.startswith('## '):
            doc.add_heading(stripped[2:], level=1)
        elif stripped.startswith('## '):
            doc.add_heading(stripped[3:], level=2)
        elif stripped.startswith('### '):
            doc.add_heading(stripped[4:], level=3)
        # 分隔线
        elif stripped == '---':
            doc.add_page_break()
        # 引用
        elif stripped.startswith('> '):
            p = doc.add_paragraph()
            run = p.add_run(stripped[2:])
            run.italic = True
            run.font.size = Pt(10)
        # 列表
        elif stripped.startswith('- '):
            doc.add_paragraph(stripped[2:], style='List Bullet')
        elif stripped.startswith('- [ ] '):
            p = doc.add_paragraph(style='List Bullet')
            p.add_run(f"☐ {stripped[6:]}")
        elif is_numbered:
            idx = stripped.index('. ')
            # 新列表起点：创建独立编号实例
            if not in_numbered_list:
                current_num_id = _create_numbering_instance(doc)
                in_numbered_list = True
            p = doc.add_paragraph(stripped[idx+2:], style='List Number')
            if current_num_id is not None:
                _apply_numbering(p, current_num_id)
        else:
            doc.add_paragraph(stripped)

    # 收尾：如果还有表格没写
    if in_table and table_rows:
        _write_table(doc, table_rows)

    doc.save(output_path)
    return True


def _insert_toc_field(doc):
    """在文档开头插入Word目录域代码。打开Word后按Ctrl+A -> F9更新域即可生成目录。"""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    
    body = doc.element.body
    first = body[0] if len(body) > 0 else None

    # 目录标题段落（用Heading 1样式）
    toc_title = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    pStyle = OxmlElement('w:pStyle')
    pStyle.set(qn('w:val'), 'Heading1')
    pPr.append(pStyle)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center')
    pPr.append(jc)
    toc_title.append(pPr)
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = '目  录'
    r.append(t)
    toc_title.append(r)

    # TOC域段落
    toc_para = OxmlElement('w:p')
    r1 = OxmlElement('w:r')
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    r1.append(fld_begin)
    toc_para.append(r1)

    r2 = OxmlElement('w:r')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    r2.append(instr)
    toc_para.append(r2)

    r3 = OxmlElement('w:r')
    fld_sep = OxmlElement('w:fldChar')
    fld_sep.set(qn('w:fldCharType'), 'separate')
    r3.append(fld_sep)
    toc_para.append(r3)

    r4 = OxmlElement('w:r')
    t4 = OxmlElement('w:t')
    t4.text = '（打开Word后按 Ctrl+A 全选 -> F9 更新域，目录自动生成）'
    r4.append(t4)
    toc_para.append(r4)

    r5 = OxmlElement('w:r')
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    r5.append(fld_end)
    toc_para.append(r5)

    # 分页符
    page_break = OxmlElement('w:p')
    r_pb = OxmlElement('w:r')
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    r_pb.append(br)
    page_break.append(r_pb)

    if first is not None:
        body.insert(0, page_break)
        body.insert(0, toc_para)
        body.insert(0, toc_title)
    else:
        body.append(toc_title)
        body.append(toc_para)
        body.append(page_break)


def _write_table(doc, rows):
    """将收集的表格行写入Word表格"""
    if not rows:
        return
    from docx.shared import Pt
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = 'Table Grid'
    for i, row_data in enumerate(rows):
        for j, cell_text in enumerate(row_data):
            if j < len(table.rows[i].cells):
                cell = table.rows[i].cells[j]
                cell.text = cell_text
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(10)


def main():
    parser = argparse.ArgumentParser(description='招标文件生成器')
    parser.add_argument('--type', choices=['goods', 'services', 'engineering'],
                        default='services', help='项目类型（默认services）')
    parser.add_argument('--project', help='项目名称')
    parser.add_argument('--project-id', help='项目编号')
    parser.add_argument('--purchaser', help='采购人')
    parser.add_argument('--budget', type=int, help='预算金额（元，纯数字）')
    parser.add_argument('--config', help='从JSON文件读取项目信息')
    parser.add_argument('-o', '--output', default='招标文件.md', help='输出文件名')
    parser.add_argument('--docx', action='store_true', help='同时生成Word文档')
    parser.add_argument('--numbering', choices=['arabic', 'multi', 'chinese', 'none'],
                        default='arabic',
                        help='编号风格：arabic(1.2.3 默认)/multi(1.1 1.2)/chinese(一、二、)/none(不编号)')
    args = parser.parse_args()

    # 收集项目信息
    if args.config:
        with open(args.config, 'r', encoding='utf-8') as f:
            project_info = json.load(f)
    else:
        project_info = {
            'project_name': args.project or '',
            'project_id': args.project_id or '',
            'purchaser': args.purchaser or '',
            'budget': args.budget or '',
        }

    # 生成Markdown
    md_text = generate_markdown(project_info, args.type)

    # 编号风格转换（非arabic模式把编号内嵌到文本中）
    if args.numbering != "arabic":
        md_text = _convert_numbering(md_text, args.numbering)

    # 写入文件
    with open(args.output, 'w', encoding='utf-8') as f:
        f.write(md_text)
    print(f"✅ Markdown已生成：{args.output}")

    # 可选：生成Word
    if args.docx:
        docx_path = args.output.rsplit('.', 1)[0] + '.docx'
        if generate_docx(md_text, docx_path):
            print(f"✅ Word已生成：{docx_path}")
            print(f"   ⚠️ 打开Word后按 Ctrl+A 全选 -> F9 更新目录域")


if __name__ == '__main__':
    main()
