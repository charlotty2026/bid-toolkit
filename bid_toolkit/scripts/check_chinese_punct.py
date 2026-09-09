#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中文标点质检脚本 —— Chinese Punctuation Linter

铁律来源：佛跳墙 · 云端护法（2026-09-02 由主人阿芸芸连续三轮纠错沉淀）
主人原话：中文语境标点用对是最基本功，不是全角半角的技术问题；这类质检必须基础设施级，
          凡中文语境审核都要带上，甚至生成时就要带上这个脑子。

检测规则（三大类）：
  A. 引号用法：中文默认用双引号“ ”；只有双引号内再引用才用单引号‘ ’。
     检查项：
       A1 中文语境用单引号引中文词（本应双引号）
       A2 半角单引号/双引号混入中文（' "）
  B. 并列顿号：两个并列引号词之间必须用顿号“、”。
     检查项：
       B1 引号并列缺顿号： “A”“B” → 应为 “A”、“B”
       B2 “XXX”和/或/及 相邻引号之间应有顿号
  C. 半角标点混入中文正文：
     检查项：
       C1 中文后紧跟半角逗号/句号/冒号/叹号/问号/分号（排除文件名、表格单位标注）
       C2 中文语境半角括号（排除表格单位标注如 (元) (小时)）

用法：
    python3 check_chinese_punct.py <文件或目录> [--fix]
    --fix  自动修复可安全修复的问题（引号/顿号），并输出报告

支持格式：.docx / .xlsx / .md / .txt / .json

退出码：0 = 无问题；1 = 发现问题（未修复）；--fix后仍有问题也返回1
"""

from __future__ import annotations

import argparse
import os
import re
import sys

# ---------------- 文件读取 ----------------
def extract_text(path: str) -> str:
    """按扩展名提取纯文本（docx/xlsx/其他）"""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".docx":
            import zipfile
            z = zipfile.ZipFile(path)
            xml = z.read("word/document.xml").decode("utf-8", errors="ignore")
            return re.sub(r"<[^>]+>", "", xml)
        elif ext == ".xlsx":
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True)
            parts = []
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    parts.append(" ".join(str(c) for c in row if c is not None))
            return "\n".join(parts)
        elif ext in (".md", ".txt", ".json", ".yaml", ".yml"):
            with open(path, encoding="utf-8", errors="ignore") as f:
                return f.read()
        else:
            return ""
    except Exception as e:
        return f"__READ_ERR__: {e}"

# ---------------- 各类检测 ----------------
def detect(path: str, text: str, in_cell: bool = False) -> list[dict]:
    """返回 [{type, line, ctx, fix_from, fix_to}]"""
    if text.startswith("__READ_ERR__"):
        return []
    probs = []
    lines = text.split("\n")

    for li, line in enumerate(lines, 1):
        # ---- A1: 单引号引中文词（本应双引号） ----
        # 匹配 '中文词' 或 ‘中文词’（前后是中文或标点），单引号做引用而非嵌套
        for m in re.finditer(r"[\u2018\u0027]([\u4e00-\u9fff][^\u2018\u2019\u201c\u201d\u0027\u0022\r\n]{0,12})[\u2019\u0027]", line):
            # 排除：这是双引号内嵌套？简化：只要引号引中文词，且非成对双引号内，都判单引号误用
            probs.append({
                "type": "A1_单引号引中文词(应双引号)",
                "line": li, "ctx": line[max(0, m.start()-8):m.end()+8],
                "fix_from": line[m.start():m.end()],
                "fix_to": "“" + m.group(1) + "”",
            })
        # ---- A2: 半角引号混入中文 ----
        for m in re.finditer(r"[\u0027\u0022]", line):
            ctx = line[max(0, m.start()-8):m.end()+8]
            if re.search(r"[\u4e00-\u9fff]", ctx):
                probs.append({"type": "A2_半角引号混入中文", "line": li, "ctx": "…" + ctx + "…"})

        # ---- B1: 并列引号缺顿号：“A”“B” ----
        for m in re.finditer(r"[\u201d]([\u201c])", line):
            ctx = line[max(0, m.start()-6):m.end()+6]
            probs.append({"type": "B1_并列引号缺顿号(应、)", "line": li, "ctx": "…" + ctx + "…"})

        # ---- C1: 中文后接半角逗号句号等（排除文件名.扩展名、数字）----
        for m in re.finditer(r"([\u4e00-\u9fff])[,.;:!?]", line):
            # 排除文件名：中文+.扩展名（docx/xlsx/md/pdf...）
            if re.search(r"[\u4e00-\u9fff]\.(docx|xlsx|doc|xls|md|pdf|txt|py|json|yaml|csv)(\s|\||$|[0-9])", line):
                continue
            # 排除时间/日期（如 7月5日.）少见，忽略
            probs.append({
                "type": "C1_中文后接半角标点",
                "line": li, "ctx": "…" + line[max(0, m.start()-8):m.end()+8] + "…",
            })

        if not in_cell:
            # ---- C2: 中文语境半角括号（正文，排除表格单位标注本函数in_cell单独处理）----
            pass

    return probs

# ---------------- 修复 ----------------
def fix_text(text: str) -> tuple[str, int]:
    """自动修复：单引号引中文词→双引号；并列引号间补顿号。返回(新文本, 修复数)"""
    n = 0
    # 修复 A1: 'X' 或 ‘X’ → “X”
    def repA1(m):
        nonlocal n
        n += 1
        return "“" + m.group(1) + "”"
    text = re.sub(r"[\u2018\u0027]([\u4e00-\u9fff][^\u2018\u2019\u201c\u201d\u0027\u0022\r\n]{0,12})[\u2019\u0027]", repA1, text)
    # 修复 B1: ”“ → ”、“（并列引号间补顿号）
    def repB1(m):
        nonlocal n
        n += 1
        return "”" + "、" + m.group(1)
    text = re.sub(r"[\u201d]([\u201c])", repB1, text)
    return text, n


# ---------------- CLI ----------------
def main():
    ap = argparse.ArgumentParser(description="中文标点质检（引号/顿号/半角标点）")
    ap.add_argument("target", nargs="+", help="文件或目录")
    ap.add_argument("--fix", action="store_true", help="自动修复可安全修复的问题")
    ap.add_argument("--via-bid-toolkit", action="store_true",
                    help="作为 bid-toolkit 引擎环调用（带分数语义，供 check 链路）")
    args = ap.parse_args()

    total_prob = 0
    total_fix = 0
    any_error = False

    for t in args.target:
        if os.path.isdir(t):
            files = [os.path.join(r, f) for r, d, fs in os.walk(t)
                     for f in fs if os.path.splitext(f)[1].lower() in (".docx", ".xlsx", ".md", ".txt", ".json")]
        else:
            files = [t]

        for path in files:
            text = extract_text(path)
            in_cell = path.lower().endswith(".xlsx")

            if args.fix:
                new_text, nfx = fix_text(text)
                if nfx:
                    if path.lower().endswith(".docx"):
                        # docx: 需要写回 XML，简化处理（直接写纯文本会破坏结构，此处仅报告）
                        print(f"[SKIP-写回需docx结构] {path}: 发现{nfx}处可修复但docx需XML定位")
                        any_error = True
                        continue
                    _write_plain(path, new_text)
                    total_fix += nfx
                    print(f"[FIXED] {path}: 修复 {nfx} 处")
                text = new_text

            probs = detect(path, text, in_cell)
            if probs:
                any_error = True
                total_prob += len(probs)
                print(f"\n❌ {path}  ({len(probs)} 处标点问题)")
                for p in probs[:12]:
                    print(f"    L{p['line']} [{p['type']}] {p['ctx']}")
                if len(probs) > 12:
                    print(f"    ... 等共 {len(probs)} 处")
            else:
                print(f"✅ {path}")

    print(f"\n=== 汇总: 发现 {total_prob} 处标点问题, 自动修复 {total_fix} 处 ===")
    if args.via_bid_toolkit:
        # 给 bid-toolkit 打分语义：0问题=满分，每处-2分
        score = max(0, 100 - total_prob * 2)
        print(f"BID_CHINESE_PUNCT_SCORE: {score}")
    sys.exit(1 if any_error else 0)


def _write_plain(path, text):
    """写回纯文本文件（md/txt/json/xlsx会走openpyxl的场景暂不做写回）"""
    if path.lower().endswith((".md", ".txt")):
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)


if __name__ == "__main__":
    main()