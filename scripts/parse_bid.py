#!/usr/bin/env python3
"""
招标文件拆解器 v1.0
从招标文件（PDF/Markdown）中提取结构化信息：
- 格式要求（字体/字号/行距/页边距）
- 废标红线（否决条款）
- 文件清单（必须提交的材料）
- 大纲框架（投标文件构成+格式+评分项）
- 报价规则
- 时间节点
- 资质要求

用法：
    python parse_bid.py 招标文件.pdf -o requirements.json
    python parse_bid.py 招标文件.md -o requirements.json
    python parse_bid.py 招标文件.pdf --template government -o requirements.json
"""

import os
import sys
import re
import json
import argparse
from pathlib import Path
from datetime import datetime


def _table_to_markdown(table):
    """将pdfplumber表格（列表的列表）转为Markdown格式，保留行列结构。

    评分表/废标表/资质表等表格在纯文本提取时行列会打散，
    转为Markdown后每行是一条完整记录，下游正则可直接匹配。
    """
    if not table or not table[0]:
        return ""
    cleaned = []
    for row in table:
        cleaned_row = [str(cell).strip() if cell else "" for cell in row]
        cleaned.append(cleaned_row)
    # 过滤空行
    cleaned = [row for row in cleaned if any(cell for cell in row)]
    if not cleaned:
        return ""
    lines = []
    # 表头
    lines.append("| " + " | ".join(cleaned[0]) + " |")
    lines.append("| " + " | ".join("---" for _ in cleaned[0]) + " |")
    # 数据行
    for row in cleaned[1:]:
        # 补齐列数（合并单元格可能导致行长度不一致）
        while len(row) < len(cleaned[0]):
            row.append("")
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def extract_tables_from_pdf(pdf_path):
    """从PDF提取结构化表格数据。

    返回格式：[{"page": int, "headers": list, "rows": list[list], "markdown": str}]

    用途：下游函数可直接按行列访问表格数据，不依赖正则从纯文本中解析。
    适合评分表/废标表/资质表等关键表格的精准提取。
    """
    tables_result = []
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                for table in tables:
                    if not table or not table[0]:
                        continue
                    # 过滤过小的表格（<2行或<2列=布局噪音）
                    if len(table) < 2 or len(table[0]) < 2:
                        continue
                    cleaned = []
                    for row in table:
                        cleaned_row = [str(cell).strip() if cell else "" for cell in row]
                        cleaned.append(cleaned_row)
                    cleaned = [row for row in cleaned if any(cell for cell in row)]
                    if len(cleaned) < 2:
                        continue
                    headers = cleaned[0]
                    rows = cleaned[1:] if len(cleaned) > 1 else []
                    md = _table_to_markdown(cleaned)
                    tables_result.append({
                        "page": page_idx + 1,
                        "headers": headers,
                        "rows": rows,
                        "markdown": md,
                    })
    except ImportError:
        pass  # pdfplumber未安装，静默跳过（extract_text_from_pdf会走fallback）
    except Exception as e:
        print(f"表格提取警告：{e}", file=sys.stderr)
    return tables_result


def extract_text_from_pdf(pdf_path, include_tables=True):
    """从PDF提取文本（pdfplumber主路径，PyMuPDF fallback）。

    pdfplumber优势：表格结构化提取，行列对齐。
    PyMuPDF fallback：pdfplumber失败时兜底（无表格结构化）。

    include_tables=True时，表格以Markdown格式内嵌在文本中，
    评分表/废标表等关键表格的行列结构不会被打散。
    """
    # 主路径：pdfplumber
    try:
        import pdfplumber
        text_parts = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""

                if include_tables:
                    tables = page.extract_tables()
                    if tables:
                        md_tables = []
                        for table in tables:
                            md = _table_to_markdown(table)
                            if md:
                                md_tables.append(md)
                        if md_tables:
                            page_text += "\n\n" + "\n\n".join(md_tables)

                text_parts.append(page_text)

        return "\n".join(text_parts)
    except ImportError:
        pass  # 走fallback

    # Fallback：PyMuPDF（无表格结构化，但至少能提取纯文本）
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text() + "\n"
        doc.close()
        return text
    except ImportError:
        print("错误：需要安装 pdfplumber 或 PyMuPDF：pip install pdfplumber", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"PDF读取失败：{e}", file=sys.stderr)
        sys.exit(1)


def read_input_file(file_path):
    """读取输入文件（PDF或Markdown）"""
    path = Path(file_path)
    if not path.exists():
        print(f"文件不存在：{file_path}", file=sys.stderr)
        sys.exit(1)

    if path.suffix.lower() == '.pdf':
        return extract_text_from_pdf(str(path))
    elif path.suffix.lower() in ('.md', '.txt'):
        with open(path, 'r', encoding='utf-8') as f:
            return f.read()
    else:
        print(f"不支持的文件格式：{path.suffix}", file=sys.stderr)
        print("支持：.pdf / .md / .txt", file=sys.stderr)
        sys.exit(1)


def extract_format_requirements(text):
    """提取格式要求：字体/字号/行距/页边距"""
    requirements = {}

    # 页边距
    margin_patterns = [
        r'页边距[：:]\s*上下[\s]*([0-9.]+)\s*(?:cm|厘米|MM|mm|毫米)',
        r'页边距[：:]\s*左右[\s]*([0-9.]+)\s*(?:cm|厘米|MM|mm|毫米)',
        r'上边距[\s]*([0-9.]+)\s*(?:cm|厘米)',
        r'下边距[\s]*([0-9.]+)\s*(?:cm|厘米)',
        r'左边距[\s]*([0-9.]+)\s*(?:cm|厘米)',
        r'右边距[\s]*([0-9.]+)\s*(?:cm|厘米)',
    ]
    margins = {}
    for pattern in margin_patterns:
        match = re.search(pattern, text)
        if match:
            key = "上" if "上" in pattern else "下" if "下" in pattern else "左" if "左" in pattern else "右" if "右" in pattern else "上下" if "上下" in pattern else "左右"
            margins[key] = match.group(1) + "cm"
    if margins:
        requirements['页边距'] = margins

    # 字体字号
    font_section = re.findall(
        r'(?:正文|标题|一级标题|二级标题|三级标题|表格|页眉|页脚)[：:]\s*'
        r'([宋体|黑体|仿宋|楷体|Times New Roman|Arial|微软雅黑]+)'
        r'[\s]*([一二三四五六七八九十\d]+号|[0-9.]+pt|小[一二三四五六七八九十]+)',
        text
    )
    if font_section:
        requirements['字体字号'] = [
            {"对象": item[0].strip() if item[0] else "正文",
             "字体": re.sub(r'[：:\s]', '', item[0]) if item[0] else "宋体",
             "字号": item[1].strip() if item[1] else "小四"}
            for item in font_section
        ]

    # 行距
    line_spacing = re.search(r'行距[：:]\s*([\d.]+)倍|行距[：:]\s*(固定值|单倍|1\.5倍|2倍)', text)
    if line_spacing:
        requirements['行距'] = line_spacing.group(1) + "倍" if line_spacing.group(1) else line_spacing.group(2)

    return requirements


def extract_disqualification_rules(text):
    """提取废标红线（否决条款/投标无效情形）
    v2.0: 精准匹配废标表述，只抓真正描述废标条件的句子，不再把含'否决'的段落全抓
    """
    rules = []

    # 废标触发模式：必须是"...将被否决/予以否决/投标无效/废标"这种判定性表述
    # 不抓"否决处理"(非必然废标)、不抓"可对该投标文件作否决处理"(评标委裁量权)
    negate_patterns = [
        r'其投标(?:均)?将(?:予以)?否决',
        r'其投标(?:均)?将被否决',
        r'投标(?:均)?将(?:予以)?否决',
        r'投标无效',
        r'予以废标',
        r'将被废标',
        r'即被否决',
        r'投标文件(?:将)?(?:予以)?否决',
        # v2.2: "作废标处理"句式（政府采购文件常见）
        r'作废标处理',
        r'将其投标作废标',
        r'其投标作废标',
    ]
    negate_re = re.compile('|'.join(negate_patterns))

    lines = text.split('\n')
    for i, line in enumerate(lines):
        line_stripped = line.strip()
        if not line_stripped:
            continue
        if negate_re.search(line_stripped):
            # 向上回溯找完整句子（PDF换行可能把一句话拆成多行）
            # 向上合并到句号/分号/编号开头
            merged = line_stripped
            j = i - 1
            while j >= 0 and j > i - 5:  # 最多回溯5行
                prev = lines[j].strip()
                if not prev:
                    break
                # 如果上一行以编号开头或句号结尾，说明是上一句的结束
                if re.match(r'^\d+\.\d+', prev) or prev.endswith(('。', '；', '；')):
                    merged = prev + ' ' + merged
                    break
                merged = prev + ' ' + merged
                j -= 1

            # 向下合并续行
            k = i + 1
            while k < len(lines) and k < i + 3:
                nxt = lines[k].strip()
                if not nxt:
                    break
                if re.match(r'^\d+\.\d+', nxt) or nxt.endswith(('。', '；')):
                    merged = merged + ' ' + nxt
                    break
                if negate_re.search(nxt):
                    break  # 下一行是另一条废标
                merged = merged + ' ' + nxt
                k += 1

            # 清理
            merged = re.sub(r'\s+', ' ', merged).strip()
            # 去掉页眉页脚噪音
            merged = re.sub(r'第\d+页|招标编号.*?项目名称.*?采购|^\s*-\s*\d+\s*-\s*', '', merged).strip()
            if merged and len(merged) > 10:
                rules.append(merged)

    # 去重
    seen = set()
    unique_rules = []
    for r in rules:
        key = r[:50]  # 用前50字符去重
        if key not in seen:
            seen.add(key)
            unique_rules.append(r)

    return unique_rules


def extract_document_checklist(text):
    """提取必须提交的文件清单 v2.0"""
    checklist = []

    # v2.1: 收窄触发词，去掉"应提交"/"需提交"等过宽匹配
    section_pattern = r'(?:投标文件应至少包括|投标文件应包括|投标文件由以下|投标文件包含|投标文件应包含|投标文件[组构]成|投标文件应至少包括下列部分)[^\n]*'
    matches = list(re.finditer(section_pattern, text))

    for match in matches:
        start = match.end()
        # v2.0: 扩大上下文窗口到2000字符
        context = text[start:start + 2000]
        lines = context.split('\n')

        for line in lines:
            line = line.strip()
            if not line:
                continue
            # v2.1: 跳过页码噪音如"18 -" / "19 -" / "- 18 -"
            if re.match(r'^-?\s*\d+\s*[-]?\s*$', line) or re.match(r'^-\s*\d+\s*-', line):
                continue
            # v2.0: 遇到下一个大章节标题就停（避免抓到别的章节内容）
            if re.match(r'^[一二三四五六七八九十]+[、.．]\s', line) and len(line) < 30:
                break
            # v2.1: 匹配（1）/ (1) / 1. / 1、/ ① 等，不要求)后跟分隔符
            if re.match(r'^[（(]\d+[)）]', line) or re.match(r'^\d+\s*[、.．]\s', line) or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]', line):
                item = re.sub(r'^[（(]\d+[)）]\s*', '', line).strip()
                item = re.sub(r'^\d+\s*[、.．]\s*', '', item).strip()
                item = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]\s*', '', item).strip()
                if item and len(item) > 2:
                    # v2.0: 去掉页眉页脚噪音
                    item = re.sub(r'第\d+页.*$', '', item).strip()
                    if item and len(item) > 2:
                        checklist.append(item)
            elif re.match(r'^[-•·●]\s+', line):
                item = re.sub(r'^[-•·●]\s+', '', line).strip()
                if item and len(item) > 2:
                    checklist.append(item)
        # v2.0: 不break，继续找下一个匹配（一份招标文件可能有多个文件清单段落）

    # 去重
    seen = set()
    unique = []
    for item in checklist:
        if item not in seen:
            seen.add(item)
            unique.append(item)

    return unique


def extract_outline(text):
    """提取大纲框架：从'投标文件构成'和'投标文件格式'章节提取标题"""
    outline: dict = {
        "来源_构成": [],
        "来源_格式": [],
        "来源_评分项": []
    }
    
    lines = text.split('\n')
    
    # 1. 提取"投标文件构成"章节
    in_composition = False
    for line in lines:
        line_stripped = line.strip()
        if '投标文件构成' in line_stripped or '投标文件组成' in line_stripped:
            in_composition = True
            continue
        if in_composition:
            # 遇到下一个大章节就停
            if re.match(r'^[一二三四五六七八九十]+[、.．]\s', line_stripped) and len(line_stripped) < 30:
                in_composition = False
                continue
            # 匹配标题行（有编号的）
            title_match = re.match(r'^[\d.]+[、.．\s]+(.+)', line_stripped)
            if title_match:
                title = title_match.group(1).strip()
                if len(title) > 2:
                    outline["来源_构成"].append(title)
    
    # 2. 提取"投标文件格式"章节
    in_format = False
    for line in lines:
        line_stripped = line.strip()
        if '投标文件格式' in line_stripped and '格式要求' not in line_stripped:
            in_format = True
            continue
        if in_format:
            if re.match(r'^[一二三四五六七八九十]+[、.．]\s', line_stripped) and len(line_stripped) < 30:
                in_format = False
                continue
            title_match = re.match(r'^[\d.]+[、.．\s]+(.+)', line_stripped)
            if title_match:
                title = title_match.group(1).strip()
                if len(title) > 2:
                    outline["来源_格式"].append(title)
    
    # 3. 提取评分项 - v2.0: 识别评分表结构，按"评分内容+分值"配对提取
    # 评分表通常格式: "评分内容 | 分值 | 评分标准" 后跟多行内容
    # 识别模式: 某行有"XX分"且相邻行有评分内容名称
    scoring_items = []
    lines_list = text.split('\n')
    
    # 已知的评分大类名称（从实际招标文件归纳）
    scoring_section_names = [
        '报价部分', '管理方案', '质控方案', '培训方案', '项目实施团队',
        '应急预案', '服务承诺', '合理化建议', '综合能力', '权益保障',
        '类似项目业绩', '技术方案', '商务部分', '资信部分', '价格部分',
        '业绩部分', '团队部分', '服务方案', '实施方案', '售后方案',
    ]
    
    # 扫描找"XX分"行，向上找评分内容名
    for i, line in enumerate(lines_list):
        line_s = line.strip()
        if not line_s:
            continue
        
        # 匹配 "15 分" / "10分" / "5 分" 这种独立的分值行
        score_match = re.match(r'^(\d{1,2})\s*分$', line_s)
        if not score_match:
            continue
        
        score_val = int(score_match.group(1))
        if score_val == 0 or score_val > 100:
            continue
        
        # 向上找评分内容名称（最多回溯5行）
        item_name = None
        desc_lines = []
        for j in range(i - 1, max(i - 6, -1), -1):
            prev = lines_list[j].strip()
            if not prev:
                continue
            # 跳过分值行和表头
            if re.match(r'^\d{1,2}\s*分$', prev):
                continue
            if prev in ('评分内容', '分值', '评分标准', '评分内容\n分值\n评分标准'):
                continue
            # 跳过页眉页脚
            if '招标编号' in prev or '第' in prev and '页' in prev:
                continue
            if re.match(r'^-\s*\d+\s*-$', prev):
                continue
            
            # 如果这行是已知的评分大类名称
            is_known = any(name in prev for name in scoring_section_names)
            # 或者是简短标题（2-15字，不含句号）
            is_title = len(prev) <= 15 and '。' not in prev and '，' not in prev
            
            if is_known or is_title:
                item_name = prev
                break
            else:
                # 可能是评分标准的描述行，收集
                desc_lines.insert(0, prev)
        
        # 向下找评分标准描述
        desc_after = []
        for k in range(i + 1, min(i + 8, len(lines_list))):
            nxt = lines_list[k].strip()
            if not nxt:
                continue
            # 遇到下一个分值行或表头就停
            if re.match(r'^\d{1,2}\s*分$', nxt):
                break
            if nxt in ('评分内容', '分值', '评分标准'):
                break
            if '招标编号' in nxt:
                continue
            if re.match(r'^-\s*\d+\s*-$', nxt):
                continue
            # 遇到已知评分大类名称也停
            if any(name == nxt for name in scoring_section_names):
                break
            desc_after.append(nxt)
        
        all_desc = desc_lines + desc_after
        desc_text = ' '.join(all_desc)[:200] if all_desc else ''
        
        if item_name:
            scoring_items.append({
                "项目": item_name,
                "分值": score_val,
                "评分标准": desc_text
            })
    
    # 去重（同名的保留第一个）
    seen_names = set()
    for item in scoring_items:
        if item["项目"] not in seen_names:
            outline["来源_评分项"].append(item)
            seen_names.add(item["项目"])
    
    # 计算总分
    total_score = sum(item["分值"] for item in outline["来源_评分项"])
    if total_score > 0:
        outline["评分总分"] = str(total_score)

    return outline


def extract_budget(text):
    """提取预算金额"""
    budgets = []
    patterns = [
        r'预算[金额]*[：:]\s*(?:人民币)?\s*([\d,，]+(?:\.\d+)?)\s*(?:万?元)',
        r'最高[限投]*价[：:]\s*(?:人民币)?\s*([\d,，]+(?:\.\d+)?)\s*(?:万?元)',
        r'控制价[：:]\s*(?:人民币)?\s*([\d,，]+(?:\.\d+)?)\s*(?:万?元)',
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            budgets.append(match.group(1).replace(',', '').replace('，', '') + "元")
    return budgets


def extract_timeline(text):
    """提取关键时间节点 v2.0: 放宽日期正则，兼容'2026 年07 月24 日'带空格格式"""
    timelines = []
    # v2.0: \s* 兼容"2026 年07 月24 日北京时间10:00"这种带空格的格式
    patterns = [
        (r'截[止标][时间日期]*[：:]?\s*(\d{4}\s*[-年]\s*\d{1,2}\s*[-月]\s*\d{1,2}\s*[日号]?\s*(?:北京时间)?\s*\d{1,2}\s*[：:]\s*\d{2})', '截标时间'),
        (r'开标[时间日期]*[：:]?\s*(\d{4}\s*[-年]\s*\d{1,2}\s*[-月]\s*\d{1,2}\s*[日号]?\s*(?:北京时间)?\s*\d{1,2}\s*[：:]\s*\d{2})', '开标时间'),
        (r'截[止标][时间日期]*[：:]?\s*(\d{4}\s*[-年]\s*\d{1,2}\s*[-月]\s*\d{1,2}\s*[日号]?)', '截标日期'),
        (r'开标[时间日期]*[：:]?\s*(\d{4}\s*[-年]\s*\d{1,2}\s*[-月]\s*\d{1,2}\s*[日号]?)', '开标日期'),
        (r'答疑[时间日期截止]*[：:]?\s*(\d{4}\s*[-年]\s*\d{1,2}\s*[-月]\s*\d{1,2}\s*[日号]?)', '答疑截止'),
        (r'澄清[时间日期截止]*[：:]?\s*(\d{4}\s*[-年]\s*\d{1,2}\s*[-月]\s*\d{1,2}\s*[日号]?)', '澄清截止'),
    ]
    for pattern, label in patterns:
        match = re.search(pattern, text)
        if match:
            # 清理提取结果中的多余空格
            cleaned = re.sub(r'\s+', ' ', match.group(1)).strip()
            timelines.append({"事件": label, "时间": cleaned})
    return timelines


def extract_qualifications(text):
    """提取资质要求 v2.0: 限定资格条件段落范围，过滤噪音"""
    quals = []
    lines = text.split('\n')

    # v2.0: 精准定位资格条件章节起始
    in_section = False
    section_started = False

    # v2.0: 噪音关键词过滤
    noise_kw = ['获取招标文件', '招标人：', '招标代理', '踏勘', '演示', '样品', 
                '投标货币', '最高投标限价', '联合体', '递交到', '邮编', '联系人',
                '有兴趣的', '合格潜在', '微信公众号', '所有投标文件', '关注微信',
                '报价权重', '评标基准价', '服务限价', '下浮率', '评标价']
    
    for i, line in enumerate(lines):
        line_stripped = line.strip()
        if not line_stripped:
            continue

        # v2.0: 用更精准的章节标题触发
        if not in_section:
            if re.search(r'(?:投标[人商]?资格条件|资格条件|资质要求|投标人资格要求)', line_stripped) and len(line_stripped) < 30:
                in_section = True
                section_started = False
                continue
        else:
            # 遇到下一个大章节标题就停
            if re.match(r'^[一二三四五六七八九十]+[、.．]\s', line_stripped) and len(line_stripped) < 30 and section_started:
                break
            # 跳过页眉页脚
            if '招标编号' in line_stripped or re.match(r'^-\s*\d+\s*-$', line_stripped):
                continue
            if '第' in line_stripped and '页' in line_stripped:
                continue
            # v2.0: 跳过噪音行
            if any(kw in line_stripped for kw in noise_kw):
                continue
            # v2.1: 匹配（1）/ (1) / 1. / 1、/ ① 等
            if re.match(r'^[（(]\d+[)）]', line_stripped) or re.match(r'^\d+\s*[、.．]\s', line_stripped) or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]', line_stripped):
                item = re.sub(r'^[（(]\d+[)）]\s*', '', line_stripped).strip()
                item = re.sub(r'^\d+\s*[、.．]\s*', '', item).strip()
                item = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]\s*', '', item).strip()
                # 向下合并续行（直到遇到下一个编号或空行段）
                for j in range(i + 1, min(i + 5, len(lines))):
                    nxt = lines[j].strip()
                    if not nxt:
                        break
                    if re.match(r'^[（(]\d+[)）]', nxt) or re.match(r'^\d+\s*[、.．]\s', nxt):
                        break
                    if re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]', nxt):
                        break
                    if re.match(r'^[一二三四五六七八九十]+[、.．]\s', nxt):
                        break
                    if '招标编号' in nxt or re.match(r'^-\s*\d+\s*-$', nxt):
                        break
                    item = item + ' ' + nxt
                # 清理
                item = re.sub(r'\s+', ' ', item).strip()
                if item and len(item) > 5:
                    quals.append(item)
                    section_started = True
                if len(quals) >= 15:
                    break

    return quals


def extract_deposit(text):
    """提取保证金信息 v2.0: 投标保证金+履约保证金金额/缴纳方式（支持中文数字）"""
    deposit = {}
    
    # 中文数字映射
    cn_num = {'零':0,'壹':1,'贰':2,'叁':3,'肆':4,'伍':5,'陆':6,'柒':7,'捌':8,'玖':9,'拾':10,
              '佰':100,'仟':1000,'万':10000,'亿':100000000,'圆':0,'元':0,'整':0}
    def cn_to_int(cn):
        """中文数字转int（支持拾佰仟万亿·v2.1修复拾做乘法）"""
        total = 0       # 最终结果
        section = 0     # 当前段（万/亿以内累积）
        current = 0     # 当前系数（拾/佰/仟之前）
        for ch in cn:
            if ch not in cn_num:
                continue
            val = cn_num[ch]
            if val >= 10000:  # 万/亿：段结束
                section += current
                total += section * val
                section = 0
                current = 0
            elif val >= 10:   # 拾/佰/仟：乘法（拾=一拾）
                if current == 0:
                    current = 1
                section += current * val
                current = 0
            else:             # 壹-玖
                current = val
        return total + section + current
    
    # 投标保证金 - 先试阿拉伯数字，再试中文数字
    bid_match = re.search(r'投标保证金[数额金额：:\s]*第?\d*\s*包?[：:]?\s*人民币?\s*(\d+(?:[，,]\d+)?(?:\.\d+)?)\s*万?元', text)
    if bid_match:
        amt = bid_match.group(1).replace('，','').replace(',','')
        deposit['投标保证金'] = amt + '元'
    else:
        cn_match = re.search(r'投标保证金[数额金额：:\s]*第?\d*\s*包?[：:]?\s*人民币?\s*([壹贰叁肆伍陆柒捌玖拾佰仟万亿圆元整]+)', text)
        if cn_match:
            val = cn_to_int(cn_match.group(1))
            if val > 0:
                deposit['投标保证金'] = f'{val}元'
    
    # 履约保证金
    perf_match = re.search(r'履约保证金[：:；;收取金额\s]*人民币?\s*(\d+(?:[，,]\d+)?(?:\.\d+)?)\s*万?元', text)
    if perf_match:
        amt = perf_match.group(1).replace('，','').replace(',','')
        deposit['履约保证金'] = amt + '元'
    else:
        cn_match = re.search(r'履约保证金[：:；;收取金额\s]*人民币?\s*([壹贰叁肆伍陆柒捌玖拾佰仟万亿圆元整]+)', text)
        if cn_match:
            val = cn_to_int(cn_match.group(1))
            if val > 0:
                deposit['履约保证金'] = f'{val}元'
    
    # 缴纳方式
    method_match = re.search(r'保证金[可以可]*以下列方式提交[：:]\s*(支票|汇票|本票|银行转账|金融机构|担保)', text)
    if method_match:
        deposit['缴纳方式'] = method_match.group(1)
    
    # 退还
    refund_match = re.search(r'保证金[退还退还在].*?未中标.*?(\d+)\s*个?\s*工作日', text)
    if refund_match:
        deposit['退还工作日'] = refund_match.group(1) + '个工作日'
    
    return deposit


# ============================================================
#  v3.5 新增：结构化提取函数
# ============================================================

def extract_scoring_items_from_tables(tables: list[dict]) -> list[dict]:
    """从 pdfplumber 提取的结构化表格中识别评分表并提取评分项。

    v2.5 修复：
    - 续表合并优先于col_map检查（修复pdfplumber headers错位）
    - 候选表头搜索排除须知条款表（序号>10的候选直接跳过）
    - 评分描述词验证只在item字段中检查
    - 排除文件清单和合计行

    返回：[{category: '', item: str, score: float, criteria: str}, ...]
    """
    ITEM_COL_EXACT = {'评分内容', '评分要素', '评分因素', '评审内容', '评分项目', '评审项目'}
    SCORE_COL_EXACT = {'分值', '分数', '得分', '权重', '权值', '权', '满分'}
    CRITERIA_COL_EXACT = {'评分标准', '评审标准', '评分细则', '评审细则', '评分依据', '评审依据',
                          '标准', '细则', '依据', '要求', '说明', '描述'}
    ITEM_COL_KEYS = {'评分', '内容', '要素', '因素', '项目', '评审', '指标'}
    SCORE_COL_KEYS = {'分值', '分数', '得分', '权重', '权值', '权', '满分'}
    CRITERIA_COL_KEYS = {'标准', '细则', '依据', '要求', '说明', '描述'}
    FILELIST_EXCLUDE = {'附件', '备注', '目录', '格式自拟', '格式见', '投标人须知', '投标文件格式'}
    SCORING_DESC_WORDS = {'分', '得分', '满分', '扣', '加分', '减分', '评分', '权重'}
    SCORING_SECTION_NAMES = [
        '技术', '商务', '价格', '资信', '业绩', '方案', '服务',
        '响应', '报价', '资质', '证书', '人员', '团队', '管理'
    ]

    def _classify_header(header_text: str) -> str:
        h = re.sub(r"[\s\n]+", "", header_text).strip()
        h_lower = h.lower()
        if h in ITEM_COL_EXACT:
            return "item"
        if h in SCORE_COL_EXACT:
            return "score"
        if h in CRITERIA_COL_EXACT:
            return "criteria"
        if re.match(r"^\d+$", h):
            try:
                num = int(h)
                if 0 <= num <= 100:
                    return "score"
            except:
                pass
            return "index"
        if any(k in h_lower for k in ITEM_COL_KEYS):
            return "item"
        if any(k in h_lower for k in SCORE_COL_KEYS):
            return "score"
        if any(k in h_lower for k in CRITERIA_COL_KEYS):
            return "criteria"
        if "序号" in h_lower or h == "序":
            return "index"
        return "unknown"

    def _build_col_map(hdrs: list) -> dict:
        col_map = {}
        for idx, h in enumerate(hdrs):
            ctype = _classify_header(str(h))
            if ctype in ("item", "score", "criteria") and ctype not in col_map:
                col_map[ctype] = idx
        return col_map

    def _is_scoring_table(t: dict) -> bool:
        headers = t.get("headers", [])
        rows = t.get("rows", [])
        all_text = " ".join(str(c) for c in headers if c)
        for r in rows:
            all_text += " ".join(str(c) for c in r if c)
        return any(kw in all_text for kw in ["评分", "分值", "评审", "打分", "评标", "权值", "权重", "满分"])

    def _should_exclude(t: dict) -> bool:
        rows = t.get("rows", [])
        exclude_count = 0
        seen = set()
        for r in rows:
            row_text = " ".join(str(c) for c in r if c)
            for ex_word in FILELIST_EXCLUDE:
                if ex_word in row_text and row_text not in seen:
                    exclude_count += 1
                    seen.add(row_text)
                    break
        return exclude_count >= 2

    def _headers_look_like_data_row(hdrs: list) -> bool:
        """判断headers是否实际上是数据行（pdfplumber错位）"""
        if len(hdrs) < 3:
            return False
        first = str(hdrs[0]).strip().replace("\n", " ")
        last = str(hdrs[-1]).strip().replace("\n", " ")
        # 第一列必须是纯数字
        if not re.match(r"^\d+$", first):
            return False
        # 最后一列必须是纯数字（分值）
        if not re.match(r"^\d+$", last):
            return False
        try:
            first_num = int(first)
            last_num = int(last)
            # 分值必须在0-100范围内
            if not (0 <= last_num <= 100):
                return False
            # 第一列序号必须>1（续表的序号从2开始）
            if first_num > 1:
                return True
        except:
            return False
        return False

    # 预处理：合并相邻评分续表
    merged_tables = []
    for t in tables:
        if not _is_scoring_table(t):
            continue
        if _should_exclude(t):
            continue

        # v2.5: 续表合并优先于col_map检查
        # 先检查当前表headers是否像数据行（pdfplumber将数据行误识别为headers）
        if merged_tables and _headers_look_like_data_row(t.get("headers", [])):
            prev = merged_tables[-1]
            prev_map = _build_col_map(prev.get("headers", []))
            if "item" in prev_map and "score" in prev_map:
                # 将当前headers作为数据行加入前一个表，再合并所有rows
                prev["rows"] = prev.get("rows", []) + [t.get("headers", [])] + t.get("rows", [])
                continue

        col_map = _build_col_map(t.get("headers", []))

        if "item" in col_map and "score" in col_map:
            merged_tables.append(t)
            continue

        merged_tables.append(t)

    items: list[dict] = []

    for table in merged_tables:
        headers = table.get("headers", [])
        rows = table.get("rows", [])
        if not headers and not rows:
            continue

        col_map = _build_col_map(headers)

        # v2.5: 候选表头搜索 - 排除须知条款表
        # 如果当前表头没有item+score，在前3行中找候选表头
        if "item" not in col_map or "score" not in col_map:
            for candidate_idx in range(min(3, len(rows))):
                candidate = rows[candidate_idx]
                candidate_map = _build_col_map(candidate)
                if "item" not in candidate_map or "score" not in candidate_map:
                    continue

                c_item_idx = candidate_map["item"]
                c_score_idx = candidate_map["score"]

                # v2.5: 排除须知条款表 - 候选行中score列如果是纯数字且>10，大概率是须知条款序号
                candidate_score_text = str(candidate[c_score_idx]).strip() if c_score_idx < len(candidate) else ""
                score_m = re.match(r"^(\d+)$", candidate_score_text)
                if score_m:
                    candidate_score_num = int(score_m.group(1))
                    # 须知条款表的"序号"列常为13-99，评分表的序号一般1-10
                    if candidate_score_num > 10:
                        continue

                # 验证候选表头后第一行数据是否有效
                valid_candidate = True
                if candidate_idx + 1 < len(rows):
                    first_data_row = rows[candidate_idx + 1]
                    if isinstance(first_data_row, (list, tuple)):
                        if c_item_idx < len(first_data_row) and first_data_row[c_item_idx]:
                            first_item = str(first_data_row[c_item_idx]).strip().replace("\n", " ")
                            # 数据行的item不应该是纯数字
                            if re.match(r"^\d+$", first_item):
                                valid_candidate = False
                        if valid_candidate and c_score_idx < len(first_data_row) and first_data_row[c_score_idx]:
                            first_score_text = str(first_data_row[c_score_idx]).strip()
                            if not re.search(r"\d+(?:\.\d+)?", first_score_text):
                                valid_candidate = False

                if valid_candidate:
                    headers = candidate
                    col_map = candidate_map
                    rows = rows[candidate_idx + 1:]
                    break

        if "item" not in col_map or "score" not in col_map:
            continue

        item_idx = col_map["item"]
        score_idx = col_map["score"]
        criteria_idx = col_map.get("criteria", -1)

        last_item = ""
        last_score = None
        table_items: list[dict] = []
        has_valid_score = False

        for row in rows:
            if not isinstance(row, (list, tuple)) or len(row) <= max(item_idx, score_idx):
                continue

            raw_item = str(row[item_idx]).strip().replace("\n", " ") if row[item_idx] else ""
            raw_score = str(row[score_idx]).strip() if row[score_idx] else ""

            if raw_item:
                item = raw_item
                last_item = item
            else:
                item = last_item

            score = None
            if raw_score:
                m = re.search(r"(\d+(?:\.\d+)?)", raw_score)
                if m:
                    try:
                        score = float(m.group(1))
                        if 0 <= score <= 100:
                            has_valid_score = True
                        else:
                            score = None
                    except ValueError:
                        pass
            if score is None and last_score is not None:
                score = last_score
            else:
                last_score = score

            if score is None:
                continue

            criteria = ""
            if criteria_idx >= 0 and criteria_idx < len(row) and row[criteria_idx]:
                criteria = str(row[criteria_idx]).strip().replace("\n", " ")

            if not item:
                continue

            # 跳过表头行
            skip_headers = {"评分内容", "评分要素", "评分因素", "评审内容", "评分项目",
                           "分值", "得分", "分数", "评分", "权重", "权值", "权",
                           "评分标准", "评审标准", "评分细则", "评分依据", "评审依据",
                           "序号", "序", "序号序号", "自评分值", "满分"}
            if item in skip_headers or re.match(r"^\d+$", item):
                continue

            # v2.5: 跳过合计行
            if "合计" in item or "总计" in item or "小计" in item:
                continue

            # v2.5c: 逐item验证 - 去除空格后再匹配评分词和区块名
            item_nospace = item.replace(" ", "")
            has_item_desc = any(w in item_nospace for w in SCORING_DESC_WORDS)
            has_item_section = any(s in item_nospace for s in SCORING_SECTION_NAMES)
            if not has_item_desc and not has_item_section:
                continue

            table_items.append({
                "category": "",
                "item": item,
                "score": score,
                "criteria": criteria[:300],
            })

        if not has_valid_score:
            continue

        if not table_items:
            continue

        # v2.5c: 表级别验证 - 去除空格后匹配
        table_has_desc = any(
            any(w in ti["item"].replace(" ", "") for w in SCORING_DESC_WORDS)
            or any(s in ti["item"].replace(" ", "") for s in SCORING_SECTION_NAMES)
            for ti in table_items
        )
        if not table_has_desc:
            continue

        items.extend(table_items)

    seen: set[str] = set()
    unique: list[dict] = []
    for it in items:
        key = it["item"]
        if key not in seen:
            seen.add(key)
            unique.append(it)

    return unique


def extract_scoring_items(text: str, tables: list[dict] | None = None) -> list[dict]:
    """从招标文件文本中提取评分项（技术分/商务分/价格分）。

    优先从结构化表格中提取，如果没有结果再fallback到文本正则提取。
    识别Markdown表格中的评分项/分值/评审标准，也支持纯文本格式。
    返回：[{category, item, score, criteria}, ...]
    """
    items: list[dict] = []

    # 优先从结构化表格中提取评分项
    if tables:
        table_items = extract_scoring_items_from_tables(tables)
        if table_items:
            return table_items

    # 评分类别关键词
    category_keywords: dict[str, list[str]] = {
        '技术分': ['技术', '技术方案', '技术部分', '技术评分', '技术标'],
        '商务分': ['商务', '商务部分', '商务评分', '商务方案', '商务标'],
        '价格分': ['价格', '报价', '价格部分', '报价部分', '价格评分'],
        '资信分': ['资信', '资信部分', '资信评分', '业绩', '业绩部分'],
    }

    def _guess_category(context: str) -> str:
        for cat, keywords in category_keywords.items():
            if any(kw in context for kw in keywords):
                return cat
        return ''

    # ── 模式1: Markdown表格行 | 评分内容 | 分值 | 评审标准 | ──
    md_row_pattern = re.compile(
        r'\|\s*([^|]+?)\s*\|\s*(\d{1,2})\s*分?\s*\|\s*([^|]*)\s*\|'
    )
    # 表头关键词（跳过）
    header_keywords = {'评分内容', '评分项', '评审项目', '项目', '分值', '评分标准', '评审标准'}
    # v2.1: 文件清单排除词 + 评分描述词
    filelist_exclude = {'附件', '备注', '目录', '格式自拟', '格式见', '投标人须知', '投标文件格式'}
    scoring_desc_words = {'分', '得分', '满分', '扣', '加分', '减分', '评分', '权重'}

    for line in text.split('\n'):
        line = line.strip()
        if not line or not line.startswith('|'):
            continue
        # 跳过表头分隔行 |---|---|---|
        if re.match(r'^\|[\s\-:|]+\|$', line):
            continue

        m = md_row_pattern.match(line)
        if not m:
            continue

        item_name = m.group(1).strip()
        score_str = m.group(2).strip()
        criteria = m.group(3).strip()

        # 跳过表头
        if item_name in header_keywords or score_str == '分值':
            continue

        # v2.1: 排除文件清单项
        combined = item_name + ' ' + criteria
        if any(ex in combined for ex in filelist_exclude):
            continue
        # v2.1: 要求评分描述词
        if not any(w in combined for w in scoring_desc_words):
            continue

        try:
            score = int(score_str)
        except ValueError:
            continue
        if score == 0 or score > 100:
            continue

        category = _guess_category(line + ' ' + item_name)
        items.append({
            'category': category,
            'item': item_name,
            'score': score,
            'criteria': criteria[:200],
        })

    # ── 模式2: 纯文本 "XX 分" 独立行 + 上下文回溯 ──
    if not items:
        lines = text.split('\n')
        current_category = ''
        scoring_section_names = [
            '报价部分', '管理方案', '质控方案', '培训方案', '项目实施团队',
            '应急预案', '服务承诺', '合理化建议', '综合能力', '权益保障',
            '类似项目业绩', '技术方案', '商务部分', '资信部分', '价格部分',
            '业绩部分', '团队部分', '服务方案', '实施方案', '售后方案',
            '技术分', '商务分', '价格分', '资信分',
        ]

        for i, line in enumerate(lines):
            line_s = line.strip()
            if not line_s:
                continue

            # 检测类别行
            for cat, keywords in category_keywords.items():
                if any(kw == line_s or kw + '分' in line_s for kw in keywords):
                    current_category = cat
                    break

            # 匹配 "15 分" / "10分" 独立行
            score_match = re.match(r'^(\d{1,2})\s*分$', line_s)
            if not score_match:
                continue
            score_val = int(score_match.group(1))
            if score_val == 0 or score_val > 100:
                continue

            # 向上找评分项名称
            item_name = ''
            for j in range(i - 1, max(i - 6, -1), -1):
                prev = lines[j].strip()
                if not prev:
                    continue
                if re.match(r'^\d{1,2}\s*分$', prev):
                    continue
                if prev in header_keywords:
                    continue
                if '招标编号' in prev or ('第' in prev and '页' in prev):
                    continue
                if re.match(r'^-\s*\d+\s*-$', prev):
                    continue
                is_known = any(name in prev for name in scoring_section_names)
                is_title = len(prev) <= 20 and '。' not in prev and '，' not in prev
                if is_known or is_title:
                    item_name = prev
                    break

            # 向下找评审标准描述
            criteria_lines: list[str] = []
            for k in range(i + 1, min(i + 8, len(lines))):
                nxt = lines[k].strip()
                if not nxt:
                    continue
                if re.match(r'^\d{1,2}\s*分$', nxt):
                    break
                if nxt in header_keywords:
                    break
                if '招标编号' in nxt or re.match(r'^-\s*\d+\s*-$', nxt):
                    continue
                if any(name == nxt for name in scoring_section_names):
                    break
                criteria_lines.append(nxt)

            if item_name:
                # v2.1: 排除文件清单项 + 评分描述词验证
                combined2 = item_name + ' ' + ' '.join(criteria_lines)
                if any(ex in combined2 for ex in filelist_exclude):
                    continue
                if not any(w in combined2 for w in scoring_desc_words):
                    scoring_section_names2 = [
                        '技术', '商务', '价格', '资信', '业绩', '方案', '服务',
                        '响应', '报价', '资质', '证书', '人员', '团队'
                    ]
                    if not any(s in combined2 for s in scoring_section_names2):
                        continue
                cat = current_category or _guess_category(item_name)
                items.append({
                    'category': cat,
                    'item': item_name,
                    'score': score_val,
                    'criteria': ' '.join(criteria_lines)[:200],
                })

    # 去重（同名的保留第一个）
    seen: set[str] = set()
    unique: list[dict] = []
    for item in items:
        key = item['item']
        if key not in seen:
            seen.add(key)
            unique.append(item)

    return unique


def extract_disqualification_clauses(text: str) -> list[dict]:
    """提取废标条款（废标/无效投标/否决投标/投标无效章节）。

    返回：[{clause, source_section}, ...]
    """
    clauses: list[dict] = []

    # 废标相关章节标题模式
    section_patterns = [
        r'(废标条款)',
        r'(无效投标)',
        r'(投标无效)',
        r'(否决投标)',
        r'(废标情形)',
        r'(无效投标情形)',
        r'(否决投标文件)',
        r'(废标条件)',
    ]

    # 废标触发关键词（用于全文扫描补充）
    trigger_keywords = [
        '废标', '无效投标', '否决投标', '投标无效',
        '作废标', '予以否决', '将被否决', '投标文件将否决',
    ]

    lines = text.split('\n')

    # ── 步骤1: 定位废标章节并提取编号条款 ──
    section_starts: list[tuple[int, str]] = []
    for i, line in enumerate(lines):
        line_s = line.strip()
        for sp in section_patterns:
            m = re.search(sp, line_s)
            if m and len(line_s) < 30:
                section_starts.append((i, m.group(1)))
                break

    if section_starts:
        for sec_idx, (start, sec_name) in enumerate(section_starts):
            # 确定章节结束位置
            end = len(lines)
            if sec_idx + 1 < len(section_starts):
                end = section_starts[sec_idx + 1][0]
            else:
                for j in range(start + 1, min(start + 200, len(lines))):
                    if re.match(r'^[一二三四五六七八九十]+[、.．]\s', lines[j].strip()) \
                            and len(lines[j].strip()) < 30:
                        end = j
                        break

            # 从章节内提取编号条款
            for j in range(start + 1, end):
                line_s = lines[j].strip()
                if not line_s:
                    continue
                if '招标编号' in line_s or re.match(r'^-\s*\d+\s*-$', line_s):
                    continue
                if '第' in line_s and '页' in line_s:
                    continue

                is_numbered = (
                    re.match(r'^[（(]?\d+[)）.、．]\s', line_s)
                    or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]', line_s)
                )
                if not is_numbered:
                    continue

                clause = re.sub(r'^[（(]?\d+[)）.、．]\s*', '', line_s).strip()
                clause = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]\s*', '', clause).strip()

                # 合并续行
                for k in range(j + 1, min(j + 5, end)):
                    nxt = lines[k].strip()
                    if not nxt:
                        break
                    if re.match(r'^[（(]?\d+[)）.、．]\s', nxt) \
                            or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]', nxt):
                        break
                    clause = clause + ' ' + nxt

                clause = re.sub(r'\s+', ' ', clause).strip()
                clause = re.sub(r'第\d+页|招标编号.*?$', '', clause).strip()
                if clause and len(clause) > 5:
                    clauses.append({
                        'clause': clause,
                        'source_section': sec_name,
                    })

    # ── 步骤2: 全文扫描含废标关键词的句子（补充/兜底） ──
    if not clauses:
        for i, line in enumerate(lines):
            line_s = line.strip()
            if not line_s or len(line_s) < 10:
                continue
            if any(kw in line_s for kw in trigger_keywords):
                # 向上回溯找完整句子
                merged = line_s
                for j in range(i - 1, max(i - 5, -1), -1):
                    prev = lines[j].strip()
                    if not prev:
                        break
                    if prev.endswith(('。', '；', ';')) or re.match(r'^\d+[.、]', prev):
                        merged = prev + ' ' + merged
                        break
                    merged = prev + ' ' + merged

                merged = re.sub(r'\s+', ' ', merged).strip()
                merged = re.sub(r'第\d+页|招标编号.*?$', '', merged).strip()

                if merged and len(merged) > 10:
                    source = ''
                    for sp in section_patterns:
                        m = re.search(sp, merged)
                        if m:
                            source = m.group(1)
                            break
                    if not source:
                        source = '全文扫描'
                    clauses.append({
                        'clause': merged,
                        'source_section': source,
                    })

    # 去重
    seen: set[str] = set()
    unique: list[dict] = []
    for c in clauses:
        key = c['clause'][:50]
        if key not in seen:
            seen.add(key)
            unique.append(c)

    return unique


def extract_qualification_requirements(text: str) -> list[dict]:
    """提取资质要求（投标人资格要求章节）。

    识别营业执照/资质证书/许可证/认证/业绩要求，返回结构化数据。
    返回：[{requirement, type, certificate}, ...]
    """
    requirements: list[dict] = []

    # 证书/资质关键词映射
    cert_keywords: dict[str, list[str]] = {
        '营业执照': ['营业执照', '工商营业执照', '统一社会信用代码'],
        '资质证书': ['资质证书', '资质等级', '资质证明', '建筑业企业资质', '工程资质'],
        '许可证': ['许可证', '经营许可证', '安全生产许可证', '食品经营许可证', '卫生许可证'],
        '认证': ['认证', 'ISO', '质量管理体系认证', '环境管理体系认证',
                 '职业健康安全管理体系认证', '3C认证', 'CCC认证'],
        '业绩要求': ['业绩', '类似项目', '同类项目', '项目业绩', '合同业绩'],
    }

    # 必备/可选判定关键词
    required_keywords = ['必须', '应当', '须具备', '应具备', '须具有', '应具有',
                         '须提供', '应提供', '必备']
    optional_keywords = ['可选', '加分', '优先', '宜具备', '宜具有']

    lines = text.split('\n')

    # 定位资格要求章节
    in_section = False

    for i, line in enumerate(lines):
        line_s = line.strip()
        if not line_s:
            continue

        if not in_section:
            if re.search(r'(?:投标[人商]?资格[条件要求]?|资格条件|资质要求)',
                         line_s) and len(line_s) < 30:
                in_section = True
                continue
        else:
            # 遇到下一个大章节标题就停
            if re.match(r'^[一二三四五六七八九十]+[、.．]\s', line_s) \
                    and len(line_s) < 30:
                break
            if '招标编号' in line_s or re.match(r'^-\s*\d+\s*-$', line_s):
                continue
            if '第' in line_s and '页' in line_s:
                continue

            is_numbered = (
                re.match(r'^[（(]\d+[)）]', line_s)
                or re.match(r'^\d+\s*[、.．]\s', line_s)
                or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]', line_s)
            )
            if not is_numbered:
                continue

            clause = re.sub(r'^[（(]\d+[)）]\s*', '', line_s).strip()
            clause = re.sub(r'^\d+\s*[、.．]\s*', '', clause).strip()
            clause = re.sub(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]\s*', '', clause).strip()

            # 合并续行
            for k in range(i + 1, min(i + 5, len(lines))):
                nxt = lines[k].strip()
                if not nxt:
                    break
                if re.match(r'^[（(]\d+[)）]', nxt) \
                        or re.match(r'^\d+\s*[、.．]\s', nxt) \
                        or re.match(r'^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]', nxt):
                    break
                if re.match(r'^[一二三四五六七八九十]+[、.．]\s', nxt):
                    break
                if '招标编号' in nxt or re.match(r'^-\s*\d+\s*-$', nxt):
                    break
                clause = clause + ' ' + nxt

            clause = re.sub(r'\s+', ' ', clause).strip()
            if not clause or len(clause) <= 5:
                continue

            # 判定必备/可选
            req_type = '必备'
            if any(kw in clause for kw in optional_keywords):
                req_type = '可选'
            elif any(kw in clause for kw in required_keywords):
                req_type = '必备'

            # 识别所需证书
            certs: list[str] = []
            for cert_type, keywords in cert_keywords.items():
                if any(kw in clause for kw in keywords):
                    certs.append(cert_type)

            requirements.append({
                'requirement': clause,
                'type': req_type,
                'certificate': '、'.join(certs) if certs else '',
            })

    # 兜底：从全文扫描含证书关键词的句子
    if not requirements:
        for line in lines:
            line_s = line.strip()
            if not line_s or len(line_s) < 10:
                continue
            matched_certs: list[str] = []
            for cert_type, keywords in cert_keywords.items():
                if any(kw in line_s for kw in keywords):
                    matched_certs.append(cert_type)
            if matched_certs and ('须' in line_s or '应' in line_s
                                  or '必须' in line_s or '具备' in line_s):
                req_type = '必备'
                if any(kw in line_s for kw in optional_keywords):
                    req_type = '可选'
                requirements.append({
                    'requirement': line_s,
                    'type': req_type,
                    'certificate': '、'.join(matched_certs),
                })

    # 去重
    seen: set[str] = set()
    unique: list[dict] = []
    for r in requirements:
        key = r['requirement'][:50]
        if key not in seen:
            seen.add(key)
            unique.append(r)

    return unique


def parse_bid_document(file_path):
    """主函数：拆解招标文件"""
    print(f"正在读取：{file_path}", file=sys.stderr)
    text = read_input_file(file_path)
    print(f"文本长度：{len(text)} 字符", file=sys.stderr)
    
    # PDF文件：提取结构化表格数据，合并到文本流中
    # 这样下游所有提取函数都能自然消费表格内容（评分表/废标表/资质表等）
    tables = []
    if Path(file_path).suffix.lower() == '.pdf':
        print("正在提取结构化表格...", file=sys.stderr)
        tables = extract_tables_from_pdf(file_path)
        if tables:
            table_md_parts = [t["markdown"] for t in tables if t.get("markdown")]
            if table_md_parts:
                text = text + "\n\n" + "\n\n".join(table_md_parts)
                print(f"表格提取：{len(tables)} 个表格，已合并到文本流", file=sys.stderr)
            else:
                print(f"表格提取：{len(tables)} 个表格，但无有效内容", file=sys.stderr)
        else:
            print("表格提取：未检测到表格", file=sys.stderr)
    
    result = {
        "文件名": Path(file_path).name,
        "解析时间": datetime.now().isoformat(),
        "格式要求": extract_format_requirements(text),
        "废标红线": extract_disqualification_rules(text),
        "文件清单": extract_document_checklist(text),
        "大纲框架": extract_outline(text),
        "预算": extract_budget(text),
        "时间节点": extract_timeline(text),
        "资质要求": extract_qualifications(text),
        "保证金": extract_deposit(text),
        "scoring_items": extract_scoring_items(text, tables),
        "disqualification_clauses": extract_disqualification_clauses(text),
        "qualification_requirements": extract_qualification_requirements(text),
        "_表格数据": tables,
    }
    
    # 统计
    stats = {
        "废标项": len(result["废标红线"]),
        "文件清单": len(result["文件清单"]),
        "大纲标题_构成": len(result["大纲框架"]["来源_构成"]),
        "大纲标题_格式": len(result["大纲框架"]["来源_格式"]),
        "评分项": len(result["大纲框架"]["来源_评分项"]),
        "资质要求": len(result["资质要求"]),
        "评分项_结构化": len(result["scoring_items"]),
        "废标条款_结构化": len(result["disqualification_clauses"]),
        "资质要求_结构化": len(result["qualification_requirements"]),
        "表格数": len(tables),
    }
    result["_统计"] = stats
    
    return result


def main():
    parser = argparse.ArgumentParser(description='招标文件拆解器 v1.0')
    parser.add_argument('input', help='招标文件路径（PDF/MD/TXT）')
    parser.add_argument('-o', '--output', default=None, help='输出JSON路径（默认打印到终端）')
    parser.add_argument('--pretty', action='store_true', help='格式化JSON输出')
    args = parser.parse_args()
    
    result = parse_bid_document(args.input)
    
    indent = 2 if args.pretty else None
    output_json = json.dumps(result, ensure_ascii=False, indent=indent)
    
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(output_json)
        print(f"✅ 已输出到：{args.output}", file=sys.stderr)
        
        # 打印摘要
        stats = result["_统计"]
        print(f"\n📊 拆解摘要：", file=sys.stderr)
        print(f"  废标红线：{stats['废标项']} 条", file=sys.stderr)
        print(f"  文件清单：{stats['文件清单']} 项", file=sys.stderr)
        print(f"  大纲标题（构成）：{stats['大纲标题_构成']} 个", file=sys.stderr)
        print(f"  大纲标题（格式）：{stats['大纲标题_格式']} 个", file=sys.stderr)
        print(f"  评分项：{stats['评分项']} 个", file=sys.stderr)
        print(f"  资质要求：{stats['资质要求']} 条", file=sys.stderr)
        print(f"  表格数：{stats['表格数']} 个", file=sys.stderr)
    else:
        print(output_json)


if __name__ == '__main__':
    main()
