#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标书查重工具 v1.0
==============================
三大检测能力：
  1. 段落级重复检测 — 同一段落在不同章节出现（相似度>80%）
  2. 跨文件数据一致性 — 日期/金额/人数/公司名在各分册间不一致
  3. 复制粘贴残留 — 旧公司名/旧项目名/他人信息残留

使用方式：
  python3 bid_duplicate_check.py 标书目录/
  python3 bid_duplicate_check.py 投标函.docx 技术方案.docx 商务标.docx
  python3 bid_duplicate_check.py 标书目录/ --format json --output 查重报告.json
"""

import os
import re
import sys
import json
import hashlib
import argparse
from collections import defaultdict
from dataclasses import dataclass, asdict
from typing import List, Dict, Tuple, Optional, Set
from datetime import datetime
from pathlib import Path
from difflib import SequenceMatcher

sys.stdout.reconfigure(encoding='utf-8')

try:
    from docx import Document
except ImportError:
    print("需要 python-docx 库，请执行: pip install python-docx")
    sys.exit(1)


# ============================================================
#  数据结构
# ============================================================
@dataclass
class DuplicateItem:
    """查重结果条目"""
    level: str          # error / warn / info
    category: str       # 内部重复/数据不一致/残留信息
    description: str    # 问题描述
    location_a: str     # 位置A
    location_b: str     # 位置B（重复对照）
    snippet_a: str      # 片段A
    snippet_b: str      # 片段B
    suggestion: str     # 修改建议

    def to_dict(self):
        return asdict(self)

    def to_text(self):
        icon = {"error": "🔴", "warn": "🟡", "info": "🔵"}.get(self.level, "⚪")
        return (f"{icon} [{self.level}] {self.category}\n"
                f"   {self.description}\n"
                f"   A: {self.location_a} → {self.snippet_a[:60]}\n"
                f"   B: {self.location_b} → {self.snippet_b[:60]}\n"
                f"   建议: {self.suggestion}")


# ============================================================
#  文件解析
# ============================================================
@dataclass
class DocParagraph:
    """段落信息"""
    text: str
    index: int
    file: str
    page_hint: int = 0  # 粗略页码

def parse_docx(file_path: str) -> List[DocParagraph]:
    """解析docx，返回段落列表"""
    paras = []
    try:
        doc = Document(file_path)
        fname = os.path.basename(file_path)
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            if text and len(text) > 5:  # 忽略太短的段落
                paras.append(DocParagraph(text=text, index=i, file=fname))
    except Exception as e:
        print(f"[ERROR] 解析失败 {file_path}: {e}")
    return paras


# ============================================================
#  检测1：段落级重复检测
# ============================================================
def check_paragraph_duplicates(all_paras: List[DocParagraph],
                                threshold: float = 0.8) -> List[DuplicateItem]:
    """检测段落级重复"""
    items = []

    # 按文件分组
    by_file = defaultdict(list)
    for p in all_paras:
        by_file[p.file].append(p)

    # 文件间对比
    files = list(by_file.keys())
    compared = set()

    for i in range(len(files)):
        for j in range(i + 1, len(files)):
            f_a, f_b = files[i], files[j]
            for pa in by_file[f_a]:
                for pb in by_file[f_b]:
                    # 跳过太短的
                    if len(pa.text) < 20 or len(pb.text) < 20:
                        continue
                    # 快速过滤：长度差异太大的不比较
                    if abs(len(pa.text) - len(pb.text)) / max(len(pa.text), len(pb.text)) > 0.3:
                        continue

                    sim = SequenceMatcher(None, pa.text, pb.text).ratio()
                    if sim >= threshold:
                        key = (pa.text[:30], pb.text[:30])
                        if key not in compared:
                            compared.add(key)
                            level = "warn" if sim < 0.95 else "error"
                            items.append(DuplicateItem(
                                level=level,
                                category="内部重复",
                                description=f"段落相似度 {sim:.0%}",
                                location_a=f"{pa.file} 段落#{pa.index+1}",
                                location_b=f"{pb.file} 段落#{pb.index+1}",
                                snippet_a=pa.text[:80],
                                snippet_b=pb.text[:80],
                                suggestion="保留最优版本，删除重复段落"
                            ))

    # 文件内重复（同一文件内不同章节出现相似段落）
    for fname, paras in by_file.items():
        for i in range(len(paras)):
            for j in range(i + 1, len(paras)):
                pa, pb = paras[i], paras[j]
                if len(pa.text) < 30 or len(pb.text) < 30:
                    continue
                if abs(pa.index - pb.index) < 5:  # 太近的跳过
                    continue
                if abs(len(pa.text) - len(pb.text)) / max(len(pa.text), len(pb.text)) > 0.3:
                    continue

                sim = SequenceMatcher(None, pa.text, pb.text).ratio()
                if sim >= threshold:
                    key = (fname, pa.text[:30], pb.text[:30])
                    if key not in compared:
                        compared.add(key)
                        level = "warn" if sim < 0.95 else "error"
                        items.append(DuplicateItem(
                            level=level,
                            category="内部重复",
                            description=f"同文件内段落重复 {sim:.0%}",
                            location_a=f"{fname} 段落#{pa.index+1}",
                            location_b=f"{fname} 段落#{pb.index+1}",
                            snippet_a=pa.text[:80],
                            snippet_b=pb.text[:80],
                            suggestion="合并或删除重复段落"
                        ))

    items.sort(key=lambda x: 0 if x.level == "error" else 1)
    return items


# ============================================================
#  检测2：跨文件数据一致性
# ============================================================
@dataclass
class DataPoint:
    """数据点"""
    value: str
    context: str
    file: str
    para_index: int

def check_data_consistency(all_paras: List[DocParagraph],
                           known_company: str = None,
                           known_project: str = None) -> List[DuplicateItem]:
    """检测跨文件数据不一致"""
    items = []

    # ── 日期一致性 ──
    date_pattern = re.compile(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日')
    dates_by_context = defaultdict(list)  # context_key -> [DataPoint]

    for p in all_paras:
        for m in date_pattern.finditer(p.text):
            date_str = m.group(0)
            # 提取上下文关键词（日期前后10字）
            start = max(0, m.start() - 10)
            end = min(len(p.text), m.end() + 10)
            ctx = p.text[start:end]

            # 按关键词分组（投标函日期、承诺函日期等）
            for kw in ["投标函", "承诺函", "声明", "委托书", "合同", "有效期",
                       "投标截止", "开标", "提交"]:
                if kw in ctx:
                    dates_by_context[kw].append(DataPoint(
                        value=date_str, context=ctx, file=p.file, para_index=p.index
                    ))

    # 检查同类日期是否一致
    for ctx_key, points in dates_by_context.items():
        values = set(p.value for p in points)
        if len(values) > 1:
            files_involved = set(p.file for p in points)
            items.append(DuplicateItem(
                level="error",
                category="数据不一致",
                description=f"「{ctx_key}」相关日期不一致",
                location_a=f"{points[0].file} 段落#{points[0].para_index+1}",
                location_b=f"{points[1].file} 段落#{points[1].para_index+1}",
                snippet_a=f"{ctx_key}: {points[0].value}",
                snippet_b=f"{ctx_key}: {points[1].value}",
                suggestion="统一所有文件中的日期"
            ))

    # ── 金额一致性 ──
    money_pattern = re.compile(r'(\d[\d,]*\.?\d*)\s*(?:万元|元|¥|￥)')
    amounts_by_context = defaultdict(list)

    for p in all_paras:
        for m in money_pattern.finditer(p.text):
            amount = m.group(0)
            start = max(0, m.start() - 15)
            end = min(len(p.text), m.end() + 15)
            ctx = p.text[start:end]

            for kw in ["总价", "报价", "金额", "预算", "控制价", "最高限价",
                       "单价", "合计", "费用"]:
                if kw in ctx:
                    amounts_by_context[kw].append(DataPoint(
                        value=amount, context=ctx, file=p.file, para_index=p.index
                    ))

    for ctx_key, points in amounts_by_context.items():
        values = set(p.value for p in points)
        if len(values) > 1:
            items.append(DuplicateItem(
                level="error",
                category="数据不一致",
                description=f"「{ctx_key}」金额不一致",
                location_a=f"{points[0].file} 段落#{points[0].para_index+1}",
                location_b=f"{points[1].file} 段落#{points[1].para_index+1}",
                snippet_a=f"{ctx_key}: {points[0].value}",
                snippet_b=f"{ctx_key}: {points[1].value}",
                suggestion="统一所有文件中的金额数据"
            ))

    # ── 人数一致性 ──
    count_pattern = re.compile(r'(\d+)\s*(?:人|名|位)(?:次|/月|/年|左右)?')
    counts_by_context = defaultdict(list)

    for p in all_paras:
        for m in count_pattern.finditer(p.text):
            count = m.group(0)
            start = max(0, m.start() - 15)
            end = min(len(p.text), m.end() + 15)
            ctx = p.text[start:end]

            for kw in ["人员", "员工", "驻场", "派遣", "配置", "到岗", "编制",
                       "团队", "服务"]:
                if kw in ctx:
                    counts_by_context[kw].append(DataPoint(
                        value=count, context=ctx, file=p.file, para_index=p.index
                    ))

    for ctx_key, points in counts_by_context.items():
        values = set(p.value for p in points)
        if len(values) > 1:
            items.append(DuplicateItem(
                level="warn",
                category="数据不一致",
                description=f"「{ctx_key}」人数不一致",
                location_a=f"{points[0].file} 段落#{points[0].para_index+1}",
                location_b=f"{points[1].file} 段落#{points[1].para_index+1}",
                snippet_a=f"{ctx_key}: {points[0].value}",
                snippet_b=f"{ctx_key}: {points[1].value}",
                suggestion="确认人数是否应一致，如不同需说明原因"
            ))

    # ── 公司名/项目名一致性 ──
    if known_company:
        for p in all_paras:
            # 检查是否出现不一致的公司名变体
            if known_company in p.text:
                continue  # 正确的名称
            # 检查是否有类似但不完全匹配的公司名
            company_words = re.findall(r'[\u4e00-\u9fa5]{2,}(?:公司|集团|有限)', p.text)
            for cw in company_words:
                if cw != known_company and len(cw) >= 4:
                    # 可能是残留的旧公司名
                    sim = SequenceMatcher(None, known_company, cw).ratio()
                    if 0.4 < sim < 0.95:
                        items.append(DuplicateItem(
                            level="warn",
                            category="残留信息",
                            description=f"疑似旧公司名残留",
                            location_a=f"{p.file} 段落#{p.index+1}",
                            location_b=f"期望: {known_company}",
                            snippet_a=p.text[:80],
                            snippet_b=f"发现: {cw}",
                            suggestion=f"将「{cw}」改为「{known_company}」"
                        ))

    if known_project:
        for p in all_paras:
            if known_project in p.text:
                continue
            project_words = re.findall(r'[\u4e00-\u9fa5]{2,}(?:项目|工程|服务|外包)', p.text)
            for pw in project_words:
                if pw != known_project and len(pw) >= 4:
                    sim = SequenceMatcher(None, known_project, pw).ratio()
                    if 0.4 < sim < 0.95:
                        items.append(DuplicateItem(
                            level="warn",
                            category="残留信息",
                            description=f"疑似旧项目名残留",
                            location_a=f"{p.file} 段落#{p.index+1}",
                            location_b=f"期望: {known_project}",
                            snippet_a=p.text[:80],
                            snippet_b=f"发现: {pw}",
                            suggestion=f"将「{pw}」改为「{known_project}」"
                        ))

    return items


# ============================================================
#  检测3：复制粘贴残留
# ============================================================
def check_copy_paste_residue(all_paras: List[DocParagraph]) -> List[DuplicateItem]:
    """检测复制粘贴残留"""
    items = []

    # ── 电话号码残留 ──
    phone_pattern = re.compile(r'(?:联系|电话|手机|Tel|tel)\s*[:：]?\s*(1[3-9]\d{9})')
    phones = defaultdict(list)
    for p in all_paras:
        for m in phone_pattern.finditer(p.text):
            phones[m.group(1)].append(p)

    # 如果出现多个不同电话，可能是残留
    if len(phones) > 3:
        for phone, paras in phones.items():
            items.append(DuplicateItem(
                level="info",
                category="残留信息",
                description=f"电话号码出现多处",
                location_a=f"{paras[0].file} 段落#{paras[0].index+1}",
                location_b=f"共{len(paras)}处",
                snippet_a=f"电话: {phone}",
                snippet_b=f"出现在: {', '.join(set(p.file for p in paras))}",
                suggestion="确认所有电话是否正确"
            ))

    # ── 邮箱残留 ──
    email_pattern = re.compile(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}')
    emails = defaultdict(list)
    for p in all_paras:
        for m in email_pattern.finditer(p.text):
            emails[m.group(0)].append(p)

    if len(emails) > 2:
        for email, paras in emails.items():
            items.append(DuplicateItem(
                level="info",
                category="残留信息",
                description=f"邮箱地址出现多处",
                location_a=f"{paras[0].file} 段落#{paras[0].index+1}",
                location_b=f"共{len(paras)}处",
                snippet_a=f"邮箱: {email}",
                snippet_b=f"出现在: {', '.join(set(p.file for p in paras))}",
                suggestion="确认所有邮箱是否正确"
            ))

    # ── 他人信息残留（身份证号）──
    id_pattern = re.compile(r'\b(\d{17}[\dXx])\b')
    ids = defaultdict(list)
    for p in all_paras:
        for m in id_pattern.finditer(p.text):
            ids[m.group(1)].append(p)

    for id_num, paras in ids.items():
        items.append(DuplicateItem(
            level="warn",
            category="残留信息",
            description=f"身份证号出现",
            location_a=f"{paras[0].file} 段落#{paras[0].index+1}",
            location_b=f"共{len(paras)}处",
            snippet_a=f"身份证: {id_num[:6]}****{id_num[-4:]}",
            snippet_b=f"出现在: {', '.join(set(p.file for p in paras))}",
            suggestion="确认身份证号是否正确且属于当前项目人员"
        ))

    return items


# ============================================================
#  主函数
# ============================================================
def check_files(file_paths: List[str],
                known_company: str = None,
                known_project: str = None,
                threshold: float = 0.8) -> List[DuplicateItem]:
    """对多个文件执行全量查重"""
    all_paras = []
    for fp in file_paths:
        paras = parse_docx(fp)
        all_paras.extend(paras)
        print(f"[INFO] {os.path.basename(fp)}: {len(paras)} 段落")

    print(f"[INFO] 总计 {len(all_paras)} 段落，开始查重...")

    items = []

    # 1. 段落级重复
    dup_items = check_paragraph_duplicates(all_paras, threshold)
    items.extend(dup_items)
    print(f"[INFO] 段落重复: {len(dup_items)} 条")

    # 2. 数据一致性
    data_items = check_data_consistency(all_paras, known_company, known_project)
    items.extend(data_items)
    print(f"[INFO] 数据不一致: {len(data_items)} 条")

    # 3. 复制粘贴残留
    residue_items = check_copy_paste_residue(all_paras)
    items.extend(residue_items)
    print(f"[INFO] 残留信息: {len(residue_items)} 条")

    # 排序
    items.sort(key=lambda x: {"error": 0, "warn": 1, "info": 2}.get(x.level, 3))

    return items


def print_report(items: List[DuplicateItem]):
    """打印查重报告"""
    print(f"\n{'='*60}")
    print(f"  标书查重报告 v1.0")
    print(f"{'='*60}")

    errors = [i for i in items if i.level == "error"]
    warns = [i for i in items if i.level == "warn"]
    infos = [i for i in items if i.level == "info"]

    if errors:
        print(f"\n🔴 严重问题 ({len(errors)})")
        print("-"*50)
        for i, item in enumerate(errors, 1):
            print(f"  {i}. {item.category}: {item.description}")
            print(f"     A: {item.location_a} → {item.snippet_a[:60]}")
            print(f"     B: {item.location_b} → {item.snippet_b[:60]}")

    if warns:
        print(f"\n🟡 警告 ({len(warns)})")
        print("-"*50)
        for i, item in enumerate(warns[:10], 1):
            print(f"  {i}. {item.category}: {item.description}")
            print(f"     {item.location_a} → {item.snippet_a[:60]}")
        if len(warns) > 10:
            print(f"  ... 还有 {len(warns)-10} 条")

    if infos:
        print(f"\n🔵 提示 ({len(infos)})")
        print("-"*50)
        for i, item in enumerate(infos[:5], 1):
            print(f"  {i}. {item.category}: {item.description}")
        if len(infos) > 5:
            print(f"  ... 还有 {len(infos)-5} 条")

    print(f"\n{'='*60}")
    print(f"  总计: {len(items)} 条 (🔴{len(errors)} 🟡{len(warns)} 🔵{len(infos)})")
    print(f"{'='*60}")


def save_items(items: List[DuplicateItem], output_path: str, fmt: str = "json"):
    """保存报告"""
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        if fmt == "json":
            json.dump({
                "generated_at": datetime.now().isoformat(),
                "total": len(items),
                "error_count": sum(1 for i in items if i.level == "error"),
                "warn_count": sum(1 for i in items if i.level == "warn"),
                "info_count": sum(1 for i in items if i.level == "info"),
                "items": [i.to_dict() for i in items]
            }, f, ensure_ascii=False, indent=2)
        else:
            for item in items:
                f.write(item.to_text() + "\n\n")
    print(f"[OK] 已保存: {output_path}")


# ============================================================
#  CLI
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="标书查重工具 v1.0")
    parser.add_argument("files", nargs="+", help="docx文件或目录")
    parser.add_argument("--company", "-c", help="当前公司名称（用于残留检测）")
    parser.add_argument("--project", "-p", help="当前项目名称（用于残留检测）")
    parser.add_argument("--threshold", "-t", type=float, default=0.8, help="相似度阈值(0-1)")
    parser.add_argument("--format", "-f", choices=["text", "json"], default="text")
    parser.add_argument("--output", "-o", help="保存报告到文件")
    args = parser.parse_args()

    # 收集所有docx文件
    file_paths = []
    for fp in args.files:
        p = Path(fp)
        if p.is_file() and p.suffix.lower() in ('.docx', '.doc'):
            file_paths.append(str(p))
        elif p.is_dir():
            for f in p.glob("*.docx"):
                file_paths.append(str(f))

    if not file_paths:
        print("[ERROR] 未找到docx文件")
        sys.exit(1)

    print(f"[INFO] 待查重文件: {len(file_paths)} 个")

    items = check_files(file_paths, args.company, args.project, args.threshold)
    print_report(items)

    if args.output:
        save_items(items, args.output, fmt="json" if args.format == "json" else "text")


if __name__ == "__main__":
    main()
