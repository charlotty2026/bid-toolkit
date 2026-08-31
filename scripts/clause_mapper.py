#!/usr/bin/env python3
"""条款-方案映射参数 —— 招标文件评分条款自动映射到方案章节"""

import re
import json
import sys
from pathlib import Path


# ========== 1. 评分条款解析 ==========

CLAUSE_PATTERNS = [
    # 评分表格式：序号 | 评分项 | 分值 | 评分标准 | 章节
    r'(?:序号|评分[项点]|评审[项点])[^。]*?(?:分值|分数|标准)[^。]*?(?:章节|对应)[^。]*?(?:[\n\r]|$)',
    # 分值+标准的行
    r'(\d+)\s*[.、．]\s*(.+?)\s*[（(]\s*(\d+)\s*[分]',
    # 一般格式：X分 — 内容
    r'(\d+)\s*[分][：:]\s*(.+?)(?:\n|$)',
    # 表格行
    r'\|[^|]+\|\s*(\d+)\s*[分]\s*\|[^|]+\|',
]

# 常见评分项分类
SCORING_CATEGORIES = {
    '价格': ['报价', '价格', '单价', '总价', '优惠'],
    '技术方案': ['技术', '方案', '实施', '措施', '方法', '工艺', '流程', '设计'],
    '人员配置': ['人员', '团队', '项目经理', '负责人', '配置', '投入'],
    '设备投入': ['设备', '机械', '工具', '车辆', '仪器', '装备'],
    '业绩经验': ['业绩', '经验', '案例', '项目经验', '类似项目', '合同'],
    '资质资信': ['资质', '资信', '认证', '等级', '证书', '信用'],
    '服务承诺': ['服务', '售后', '响应', '承诺', '保障', '质保', '维护'],
    '进度计划': ['进度', '工期', '计划', '时间', '周期', '排期'],
    '管理能力': ['管理', '组织', '协调', '制度', '体系', '流程'],
    '安全环保': ['安全', '环保', '绿色', '健康', '环境', '职业'],
}


def parse_tender_file(filepath):
    """读取招标文件，返回文本内容"""
    path = Path(filepath)
    if not path.exists():
        print(f"❌ 文件不存在: {filepath}")
        sys.exit(1)
    with open(path, 'r', encoding='utf-8') as f:
        return f.read()


def parse_scoring_table(text):
    """从招标文件文本中提取评分项"""
    clauses = []

    # 先找评分表区域（从"评分"、"评审"到"合计"、"总分"的区域）
    table_start = None
    table_end = None
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if re.search(r'(?:评分[标表]|评审[标表]|评分[细则]|评分办法|评分标准)', line):
            table_start = i
        if table_start is not None and re.search(r'(?:合计|总分|小计|总[分计]|得分|总计)', line):
            table_end = i
            break

    if table_start is None:
        table_start = 0
    if table_end is None:
        table_end = len(lines)

    # 按行解析
    for i in range(table_start, table_end):
        line = lines[i].strip()
        if not line:
            continue

        # 跳过表头
        if re.search(r'(?:序号|评分[项点]|分值|评审[项点]|评分标准)', line):
            continue

        # 跳过合计行
        if re.search(r'(?:合计|总分|小计|得分)', line):
            continue

        # 匹配表格行：| 1 | 项目整体服务方案 | 15 | ... |
        if line.startswith('|') and line.count('|') >= 3:
            cells = [c.strip() for c in line.split('|') if c.strip()]
            if len(cells) >= 3:
                # 判断是否是评分行：第二个单元格不含"评分"且第三个单元格全是数字
                if not re.search(r'(?:序号|评分[项点]|分值|评审[项点]|评分标准)', cells[0]) and len(cells) >= 3:
                    # 尝试找分值单元格：跳过序号列，找非序号列的数字
                    score_cell = None
                    content_cell = None
                    for idx, cell in enumerate(cells):
                        if re.match(r'^\d+$', cell) and int(cell) <= 100 and idx >= 1:
                            # 确认这是分值不是序号：序号通常在首列且值 <= 条款数
                            score_cell = int(cell)
                            content_cell = cells[idx - 1] if idx > 0 else cells[0]
                            break
                    if score_cell and content_cell:
                        clause = {
                            'num': len(clauses) + 1,
                            'content': content_cell.strip(),
                            'score': score_cell,
                            'category': classify_clause(content_cell.strip()),
                            'source_line': i + 1,
                        }
                        clauses.append(clause)
                        continue

        # 匹配评分项格式：序号. 内容（分值）
        m = re.match(r'(\d+)\s*[.、．]\s*(.+?)\s*[（(]\s*(\d+)\s*[分分]', line)
        if m:
            num, content, score = m.groups()
            clause = {
                'num': int(num),
                'content': content.strip(),
                'score': int(score),
                'category': classify_clause(content.strip()),
                'source_line': i + 1,
            }
            clauses.append(clause)
            continue

        # 匹配格式：分值: 内容
        m = re.match(r'(\d+)\s*[分分][：:]\s*(.+)', line)
        if m:
            score, content = m.groups()
            clause = {
                'num': len(clauses) + 1,
                'content': content.strip(),
                'score': int(score),
                'category': classify_clause(content.strip()),
                'source_line': i + 1,
            }
            clauses.append(clause)

    return clauses


def classify_clause(content):
    """对评分项内容进行分类"""
    for category, keywords in SCORING_CATEGORIES.items():
        for kw in keywords:
            if kw in content:
                return category
    return '其他'


# ========== 2. 方案章节检测 ==========

def detect_chapters(text):
    """从方案文本中检测章节结构"""
    chapters = []
    for i, line in enumerate(text.split('\n')):
        line = line.strip()
        # 匹配章节标题：一、 二、 1. 1.1 等
        m = re.match(r'^(一|二|三|四|五|六|七|八|九|十|第[一二三四五六七八九十]+章|[1-9]\d*)\s*[.、．\s]\s*(.+)$', line)
        if m:
            num, title = m.groups()
            chapters.append({
                'num': num,
                'title': title.strip(),
                'line': i + 1,
            })
        # 匹配 ## 开头
        m = re.match(r'^#{2,3}\s+(.+)', line)
        if m:
            title = m.group(1).strip()
            chapters.append({
                'num': len(chapters) + 1,
                'title': title,
                'line': i + 1,
            })
    return chapters


# ========== 3. 映射引擎 ==========

def map_clauses_to_chapters(clauses, chapters):
    """将评分条款映射到方案章节"""
    mappings = []

    for clause in clauses:
        best_match = None
        best_score = 0

        for ch in chapters:
            # 计算相关性：评分项内容 vs 章节标题
            weight = _calc_relevance(clause['content'], ch['title'])
            if weight > best_score:
                best_score = weight
                best_match = ch

        mappings.append({
            'clause': clause,
            'matched_chapter': best_match,
            'match_score': best_score,
            'status': '已匹配' if best_score >= 0.5 else '未匹配',
        })

    return mappings


def _calc_relevance(clause_text, chapter_title):
    """计算评分项与章节标题的相关性分数"""
    score = 0
    clause_words = set(re.findall(r'[\u4e00-\u9fff]{2,}', clause_text))
    title_words = set(re.findall(r'[\u4e00-\u9fff]{2,}', chapter_title))

    # 直接匹配
    common = clause_words & title_words
    score += len(common) * 2

    # 关键词匹配（使用分类体系）
    for category, keywords in SCORING_CATEGORIES.items():
        if any(kw in clause_text for kw in keywords) and any(kw in chapter_title for kw in keywords):
            score += 1

    # 归一化
    max_possible = max(len(clause_words), 1) * 2 + 1
    return score / max_possible


# ========== 4. 报告输出 ==========

def format_mapping_report(mappings, chapters, output_path=None):
    """生成映射报告"""
    lines = []
    lines.append('=' * 80)
    lines.append('  条款-方案映射审计报告')
    lines.append('=' * 80)
    lines.append('')

    # 统计
    total = len(mappings)
    matched = sum(1 for m in mappings if m['status'] == '已匹配')
    unmatched = total - matched
    total_score = sum(m['clause']['score'] for m in mappings)
    unmatched_score = sum(m['clause']['score'] for m in mappings if m['status'] == '未匹配')

    lines.append(f'  评分条款总数: {total} 条')
    lines.append(f'  已匹配章节: {matched} 条 (覆盖率: {matched / total * 100:.1f}%)')
    lines.append(f'  未匹配章节: {unmatched} 条')
    lines.append(f'  总分值: {total_score} 分')
    lines.append(f'  未匹配分值: {unmatched_score} 分 (潜在丢分风险)')
    lines.append('')

    lines.append('─' * 80)
    lines.append('  章节结构')
    lines.append('─' * 80)
    for ch in chapters:
        lines.append(f'    {ch["num"]}. {ch["title"]}')
    lines.append('')

    lines.append('─' * 80)
    lines.append('  映射详情')
    lines.append('─' * 80)
    lines.append(f'  {"序号":<4} {"评分项":<30} {"分值":<5} {"分类":<10} {"匹配章节":<20} {"状态":<6}')
    lines.append('  ' + '─' * 75)

    for m in mappings:
        c = m['clause']
        ch_title = m['matched_chapter']['title'] if m['matched_chapter'] else '—'
        status_icon = '✅' if m['status'] == '已匹配' else '🔴'
        lines.append(f'  {c["num"]:<4} {c["content"][:28]:<30} {c["score"]:<5} {c["category"][:8]:<10} {ch_title[:18]:<20} {status_icon} {m["status"]}')

    lines.append('')
    lines.append('─' * 80)
    lines.append('  🔴 未匹配条款（丢分风险项）')
    lines.append('─' * 80)
    for m in mappings:
        if m['status'] == '未匹配':
            c = m['clause']
            lines.append(f'    • 第{c["source_line"]}行 | {c["content"]} ({c["score"]}分)')
            lines.append(f'      建议：在方案中增加「{c["category"]}」相关章节')
    lines.append('')
    lines.append(f'  结论: 潜在丢分 {unmatched_score} 分，建议补充 {unmatched} 个章节内容')
    lines.append('=' * 80)

    report = '\n'.join(lines)

    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f'📁 报告已导出: {output_path}')

    return report


# ========== 5. JSON 输出（供其他工具链调用） ==========

def export_json(mappings, chapters, output_path):
    """导出为 JSON 格式"""
    data = {
        'total_clauses': len(mappings),
        'matched': sum(1 for m in mappings if m['status'] == '已匹配'),
        'unmatched': sum(1 for m in mappings if m['status'] == '未匹配'),
        'total_score': sum(m['clause']['score'] for m in mappings),
        'unmatched_score': sum(m['clause']['score'] for m in mappings if m['status'] == '未匹配'),
        'chapters': [{'num': ch['num'], 'title': ch['title']} for ch in chapters],
        'mappings': [
            {
                'clause_num': m['clause']['num'],
                'clause_content': m['clause']['content'],
                'clause_score': m['clause']['score'],
                'clause_category': m['clause']['category'],
                'matched_chapter': m['matched_chapter']['title'] if m['matched_chapter'] else None,
                'match_score': m['match_score'],
                'status': m['status'],
            }
            for m in mappings
        ],
    }
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f'📁 JSON 已导出: {output_path}')


# ========== 主入口 ==========

def main():
    import argparse
    parser = argparse.ArgumentParser(description='条款-方案映射参数工具')
    parser.add_argument('tender_file', help='招标文件路径（.md/.txt）')
    parser.add_argument('--bid-file', '-b', help='方案文件路径（可选，不指定则只解析评分表）')
    parser.add_argument('--report', '-r', help='导出报告路径')
    parser.add_argument('--json', '-j', help='导出 JSON 路径')
    parser.add_argument('--list', action='store_true', help='仅列出评分条款，不执行映射')
    args = parser.parse_args()

    text = parse_tender_file(args.tender_file)
    clauses = parse_scoring_table(text)

    if not clauses:
        print('⚠️  未识别到评分条款。请检查招标文件格式。')
        print('   支持的格式:')
        print('     1. 评分项（分值）')
        print('     2. 分值: 内容')
        print('     3. 表格行 | 评分项 | 分值 |')
        sys.exit(1)

    print(f'📋 识别到 {len(clauses)} 个评分条款:')
    for c in clauses:
        print(f'   [{c["category"]}] {c["content"][:40]}... ({c["score"]}分)')
    print()

    if args.list:
        return

    # 检测方案章节
    chapters = []
    if args.bid_file:
        bid_text = parse_tender_file(args.bid_file)
        chapters = detect_chapters(bid_text)
        print(f'📖 检测到 {len(chapters)} 个章节')
    else:
        print('⚠️  未指定方案文件（--bid-file），仅解析评分条款')
        print()

    # 执行映射
    if chapters:
        mappings = map_clauses_to_chapters(clauses, chapters)
        report = format_mapping_report(mappings, chapters, args.report)
        print(report)

        if args.json:
            export_json(mappings, chapters, args.json)
    else:
        print('  提示: 使用 --bid-file 指定方案文件以执行映射')


if __name__ == '__main__':
    main()