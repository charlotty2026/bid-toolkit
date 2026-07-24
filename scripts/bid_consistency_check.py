#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标书前后不一致检测 v1.0
========================
检测投标文件中数字、日期、金额、名称的前后矛盾。

用法：
  python bid_consistency_check.py check 投标文件.docx
  python bid_consistency_check.py check 投标文件.docx --json
"""

import sys, re, json, argparse
from pathlib import Path
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')

try:
    from docx import Document as _Docx
    _DOCX = True
except ImportError:
    _DOCX = False


def read_file(path):
    p = Path(path)
    if not p.exists():
        return None
    if p.suffix in ('.md', '.txt'):
        return p.read_text('utf-8')
    if p.suffix == '.docx' and _DOCX:
        return '\n'.join(para.text for para in _Docx(str(p)).paragraphs)
    return None


# ---- 提取器 ----

def extract_numbers(text):
    """提取 数字+单位 组合，返回 [(line, num, unit, raw)]
    
    P0修复：先提取日期范围，再剔除日期上下文中的数字，
    避免 '2024年1月' 中的 1 被当作 '1月' 数字+单位。
    """
    results = []
    pattern = re.compile(r'(\d+(?:\.\d+)?)\s*(人|个|名|位|台|辆|套|万元|万元/年|万元/月|万|元|平方米|㎡|台/年|次/年|天|日|月|年|份|页|本|册|箱|件|套/年)')
    
    # 先提取所有日期的字符位置范围，用于排除
    date_spans = []  # [(start, end)] in flattened text per line
    date_patterns = [
        re.compile(r'\d{4}年\d{1,2}月(?:\d{1,2}日)?'),
        re.compile(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}'),
    ]
    
    for i, line in enumerate(text.split('\n'), 1):
        # 标记日期覆盖的字符区间
        date_ranges = []
        for dp in date_patterns:
            for m in dp.finditer(line):
                date_ranges.append((m.start(), m.end()))
        
        for m in pattern.finditer(line):
            # 检查匹配是否落在日期区间内
            in_date = any(ds <= m.start() < de for ds, de in date_ranges)
            if in_date:
                continue
            results.append((i, float(m.group(1)), m.group(2), m.group(0)))
    return results


def extract_dates(text):
    """提取日期，返回 [(line, date_str)]"""
    results = []
    patterns = [
        re.compile(r'(\d{4})年(\d{1,2})月(\d{1,2})日'),
        re.compile(r'(\d{4})-(\d{2})-(\d{2})'),
        re.compile(r'(\d{4})/(\d{2})/(\d{2})'),
    ]
    for i, line in enumerate(text.split('\n'), 1):
        for pat in patterns:
            for m in pat.finditer(line):
                results.append((i, f"{m.group(1)}-{m.group(2)}-{m.group(3)}"))
    return results


def extract_amounts(text):
    """提取金额（小写+大写），返回 [(line, num, raw)]"""
    results = []
    # 小写：XX万元/XX元
    for i, line in enumerate(text.split('\n'), 1):
        for m in re.finditer(r'(\d+(?:,\d{3})*(?:\.\d+)?)\s*(万元|元)', line):
            num = float(m.group(1).replace(',', ''))
            if m.group(2) == '万元':
                num *= 10000
            results.append((i, num, m.group(0)))
    # 大写金额
    big_map = {'零':0,'壹':1,'贰':2,'叁':3,'肆':4,'伍':5,'陆':6,'柒':7,'捌':8,'玖':9}
    unit_map = {'拾':10,'佰':100,'仟':1000,'万':10000,'亿':100000000}
    for i, line in enumerate(text.split('\n'), 1):
        for m in re.finditer(r'[壹贰叁肆伍陆柒捌玖拾佰仟万亿零]+元(?:[壹贰叁肆伍陆柒捌玖拾]角)?(?:[壹贰叁肆伍陆柒捌玖]分)?', line):
            raw = m.group(0)
            try:
                total = 0
                current = 0
                for ch in raw:
                    if ch in big_map:
                        current = big_map[ch]
                    elif ch in unit_map:
                        if current == 0:
                            total += unit_map[ch]
                        else:
                            total += current * unit_map[ch]
                        current = 0
                    elif ch == '元':
                        total += current
                        current = 0
                results.append((i, float(total), raw))
            except Exception:
                pass
    return results


def extract_project_names(text):
    """提取疑似项目名称，返回 [(line, name)]"""
    results = []
    for i, line in enumerate(text.split('\n'), 1):
        for m in re.finditer(r'["""\'](.{4,30}?)["""\']', line):
            name = m.group(1)
            if any(kw in name for kw in ['项目','服务','采购','管理','工程','外包']):
                results.append((i, name))
    return results


# ---- 矛盾检测 ----

def check_number_conflicts(numbers):
    """同一单位出现不同数字"""
    issues = []
    by_unit = defaultdict(list)
    for line, num, unit, raw in numbers:
        by_unit[unit].append((line, num, raw))
    for unit, entries in by_unit.items():
        nums = set(e[1] for e in entries)
        if len(nums) > 1:
            # 排除明显不同维度的（如人数1人和30人可能是不同岗位）
            if unit in ('人', '名', '位') and max(nums) / min(nums) > 5:
                continue
            issues.append({
                'type': '数字矛盾',
                'unit': unit,
                'entries': [{'line': e[0], 'value': e[1], 'raw': e[2]} for e in entries],
                'desc': f'同一单位"{unit}"出现不同数值: {sorted(nums)}'
            })
    return issues


def check_date_conflicts(dates):
    """同一类日期出现不同值
    
    P0修复：不把项目起止日期（如 2024-1-1 至 2024-12-31）误报为矛盾。
    策略：只报告同一行/同一段落中无分隔符的孤立日期差异为疑似矛盾，
    起止日期对（含 至/-/~/到/— 等分隔符）不算矛盾。
    """
    issues = []
    by_value = defaultdict(list)
    for line, d in dates:
        by_value[d].append(line)
    
    if len(by_value) <= 1:
        return issues
    
    # 检查是否有起止日期对：如果不同日期成对出现在同一行，且有分隔符，不报
    # 收集每行出现的日期
    dates_by_line = defaultdict(list)
    for line, d in dates:
        dates_by_line[line].append(d)
    
    # 如果某些不同日期只出现在"起止日期行"中（同行有2+个日期），不报为矛盾
    standalone_diff_dates = set()
    for d, lines in by_value.items():
        for ln in lines:
            if len(dates_by_line[ln]) < 2:
                # 该日期独占一行，可能是关键日期（开标日/截止日等）
                standalone_diff_dates.add(d)
    
    # 只有当有2+个独立行出现不同日期时才报告
    if len(standalone_diff_dates) > 1:
        issues.append({
            'type': '日期矛盾',
            'entries': [{'date': d, 'lines': by_value[d]} for d in sorted(standalone_diff_dates)],
            'desc': f'文件中出现 {len(standalone_diff_dates)} 个独立日期需核实: {sorted(standalone_diff_dates)}'
        })
    return issues


def check_amount_conflicts(amounts):
    """金额矛盾"""
    issues = []
    # 按相近数值分组
    if len(amounts) < 2:
        return issues
    nums = [a[1] for a in amounts]
    unique = sorted(set(nums))
    if len(unique) > 1:
        # 只报告差异大的
        big_diffs = []
        for i in range(len(unique)-1):
            if unique[i] > 0 and unique[i+1] / unique[i] > 2:
                big_diffs.append((unique[i], unique[i+1]))
        if big_diffs:
            issues.append({
                'type': '金额矛盾',
                'entries': [{'line': a[0], 'amount': a[1], 'raw': a[2]} for a in amounts],
                'desc': f'发现 {len(big_diffs)} 处金额大幅差异'
            })
    return issues


def check_name_conflicts(names):
    """名称不一致"""
    issues = []
    if len(names) < 2:
        return issues
    unique_names = set(n[1] for n in names)
    if len(unique_names) > 1:
        issues.append({
            'type': '名称不一致',
            'entries': [{'line': n[0], 'name': n[1]} for n in names],
            'desc': f'出现 {len(unique_names)} 种不同名称写法'
        })
    return issues


def run_check(filepath):
    text = read_file(filepath)
    if not text:
        print(f'❌ 无法读取文件: {filepath}', file=sys.stderr)
        return None

    numbers = extract_numbers(text)
    dates = extract_dates(text)
    amounts = extract_amounts(text)
    names = extract_project_names(text)

    all_issues = []
    all_issues += check_number_conflicts(numbers)
    all_issues += check_date_conflicts(dates)
    all_issues += check_amount_conflicts(amounts)
    all_issues += check_name_conflicts(names)

    return all_issues


def main():
    parser = argparse.ArgumentParser(description='标书前后不一致检测')
    sub = parser.add_subparsers(dest='command')
    ck = sub.add_parser('check', help='检查文件')
    ck.add_argument('file', help='文件路径 (.docx/.txt/.md)')
    ck.add_argument('--json', action='store_true', help='JSON输出')
    args = parser.parse_args()

    if args.command == 'check':
        issues = run_check(args.file)
        if issues is None:
            sys.exit(1)
        if args.json:
            print(json.dumps(issues, ensure_ascii=False, indent=2))
        else:
            if not issues:
                print('✅ 未发现前后不一致问题')
            else:
                print(f'⚠️ 发现 {len(issues)} 处不一致问题:\n')
                for i, iss in enumerate(issues, 1):
                    print(f'{i}. [{iss["type"]}] {iss["desc"]}')
                    for e in iss.get('entries', []):
                        print(f'   L{e.get("line","?")}: {e.get("raw", e.get("name", e.get("amount", e.get("date",""))))}')
                    print()
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
