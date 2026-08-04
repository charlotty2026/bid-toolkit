#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
脱敏扫描报告生成器
扫描 bid-toolkit-opensource 目录中的隐私/敏感信息
日期: 2026-07-27
"""

import subprocess
import os
import json
from datetime import datetime
from collections import defaultdict

# 配置
BASE_DIR = os.environ.get("BID_TOOLKIT_BASE_DIR", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SCAN_DATE = datetime.now().strftime("%Y-%m-%d")

# 15个敏感词（internal_team代号/内部称呼）
SENSITIVE_WORDS = [
    "internal_code",
    "internal_team",
    "internal",
    "agent_alpha",
    "agent_beta",
    "agent_gamma",
    "agent_delta",
    "agent_epsilon",
    "agent_zeta",
    "pet_name",
    "agent_eta",
    "agent_theta",
    "owner_name",
    "internal_motto",
]

# 额外隐私词（扫描中发现的）
EXTRA_PRIVACY_WORDS = [
    "owner",
    "owner",
    "侵掠如火",
    "快枪手",
    "审核工具A",
    "审核工具B",
    "自动化工具",
    "大明",
    "沪BH7T98",
    "兰州百合干",
    "维护者",
    "验证人",
    "测试人",
]

# 不算敏感但需记录的
RECORD_WORDS = ["charlotty2026"]


def grep_scan(word, base_dir):
    """使用grep扫描单个词，返回命中文件列表和行内容"""
    try:
        result = subprocess.run(
            ["grep", "-riIn", word, ".", "--exclude-dir=.git"],
            cwd=base_dir,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if result.returncode != 0 or not result.stdout.strip():
            return []
        hits = []
        for line in result.stdout.strip().split("\n"):
            if ":" in line:
                parts = line.split(":", 2)
                if len(parts) >= 3:
                    filepath, lineno, content = parts[0], parts[1], parts[2]
                    hits.append(
                        {
                            "file": filepath,
                            "line": int(lineno) if lineno.isdigit() else 0,
                            "content": content.strip(),
                        }
                    )
        return hits
    except Exception as e:
        return [{"file": "ERROR", "line": 0, "content": str(e)}]


def scan_word(word, base_dir, category="sensitive"):
    """扫描单个词并返回结构化结果"""
    hits = grep_scan(word, base_dir)
    unique_files = list(
        dict.fromkeys([h["file"] for h in hits if h["file"] != "ERROR"])
    )
    return {
        "word": word,
        "category": category,
        "hit_count": len(hits),
        "file_count": len(unique_files),
        "files": unique_files,
        "details": hits,
    }


def generate_report():
    """生成完整扫描报告"""
    report = {
        "meta": {
            "scan_date": SCAN_DATE,
            "target_dir": BASE_DIR,
            "scanner": "desensitization_scan.py",
            "total_files_scanned": "全仓库（排除.git）",
        },
        "sections": [],
    }

    # === 第一节：15个核心敏感词 ===
    section1 = {
        "title": "第一节：15个核心敏感词扫描",
        "description": "扫描internal_team代号、内部称呼、Agent名字等绝对不可外泄的词汇",
        "items": [],
    }
    for word in SENSITIVE_WORDS:
        result = scan_word(word, BASE_DIR, "sensitive")
        section1["items"].append(result)
    report["sections"].append(section1)

    # === 第二节：需记录词汇 ===
    section2 = {
        "title": "第二节：需记录词汇（非敏感但需知会）",
        "description": "GitHub用户名等公开标识，虽非隐私但需记录",
        "items": [],
    }
    for word in RECORD_WORDS:
        result = scan_word(word, BASE_DIR, "record")
        section2["items"].append(result)
    report["sections"].append(section2)

    # === 第三节：额外隐私词（实际扫描中发现） ===
    section3 = {
        "title": "第三节：额外隐私词扫描（实际命中）",
        "description": "扫描过程中发现的其他隐私/内部代号暴露",
        "items": [],
    }
    for word in EXTRA_PRIVACY_WORDS:
        result = scan_word(word, BASE_DIR, "privacy")
        if result["hit_count"] > 0:
            section3["items"].append(result)
    report["sections"].append(section3)

    return report


def print_report(report):
    """打印格式化报告"""
    print("=" * 70)
    print("脱敏扫描报告")
    print("=" * 70)
    print(f"扫描日期: {report['meta']['scan_date']}")
    print(f"目标目录: {report['meta']['target_dir']}")
    print(f"扫描范围: {report['meta']['total_files_scanned']}")
    print()

    # 第一节
    s1 = report["sections"][0]
    print(f"\n{'=' * 70}")
    print(f"{s1['title']}")
    print(f"   {s1['description']}")
    print("-" * 70)
    all_clean = True
    for item in s1["items"]:
        if item["hit_count"] == 0:
            print(f"  OK [{item['word']}] -- 0命中")
        else:
            all_clean = False
            print(
                f"  ALERT [{item['word']}] -- {item['hit_count']} 命中，{item['file_count']} 个文件"
            )
            for f in item["files"]:
                print(f"      -> {f}")
            for d in item["details"]:
                print(f"         行{d['line']}: {d['content'][:60]}")
    if all_clean:
        print("\n  全部15个敏感词零命中！本区域安全。")

    # 第二节
    s2 = report["sections"][1]
    print(f"\n{'=' * 70}")
    print(f"{s2['title']}")
    print(f"   {s2['description']}")
    print("-" * 70)
    for item in s2["items"]:
        if item["hit_count"] == 0:
            print(f"  OK [{item['word']}] -- 0命中")
        else:
            print(
                f"  ALERT [{item['word']}] -- {item['hit_count']} 命中，{item['file_count']} 个文件"
            )
            for f in item["files"]:
                print(f"      -> {f}")

    # 第三节
    s3 = report["sections"][2]
    print(f"\n{'=' * 70}")
    print(f"{s3['title']}")
    print(f"   {s3['description']}")
    print("-" * 70)
    if not s3["items"]:
        print("  OK 无额外隐私词命中")
    else:
        for item in s3["items"]:
            print(
                f"  ALERT [{item['word']}] -- {item['hit_count']} 命中，{item['file_count']} 个文件"
            )
            for f in item["files"]:
                print(f"      -> {f}")
            for d in item["details"]:
                print(f"         行{d['line']}: {d['content'][:80]}")

    # 汇总
    print(f"\n{'=' * 70}")
    print("扫描汇总")
    print("-" * 70)
    total_sens = sum(1 for i in s1["items"] if i["hit_count"] > 0)
    total_priv = len(s3["items"])
    print(f"  15个核心敏感词命中: {total_sens} 个")
    print(f"  额外隐私词命中: {total_priv} 个")
    print(f"  需记录词汇命中: {sum(1 for i in s2['items'] if i['hit_count'] > 0)} 个")
    if total_sens == 0 and total_priv == 0:
        print("\n  综合结论: 未发现隐私泄露！可以安全开源发布。")
    else:
        print("\n  综合结论: 发现隐私泄露！必须修复后才能开源发布。")
        print("  重点文件: rules/bid_rules.md（含大量内部代号、Agent角色、车牌号）")
        print("  重点文件: docs/RFP验证报告.md（含'owner'称呼）")
    print("=" * 70)

    # 输出JSON供存档
    json_path = os.path.join(BASE_DIR, "desensitization_report.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)
    print(f"\n完整JSON报告已保存: {json_path}")


if __name__ == "__main__":
    report = generate_report()
    print_report(report)
