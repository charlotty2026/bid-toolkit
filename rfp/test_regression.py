#!/usr/bin/env python3
"""验证 run_all_checks 检查项数量和回归"""

import sys, os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from rfp_compliance import run_all_checks

# 用一个实际招标文件测试
text = open("../samples/某外包项目标书.md", "r", encoding="utf-8").read()
report = run_all_checks(text, "services")

print("检查项数量:", len(report["checks"]))
for name, items in report["checks"].items():
    count = len(items) if isinstance(items, list) else 1
    print(f"  {name}: {count} 项")

print(f"\n总计: {report['summary']['total']} 项")
print(f"结论: {report['summary']['verdict']}")
print(
    f"通过: {report['summary']['pass']}, 警告: {report['summary']['warn']}, 失败: {report['summary']['fail']}"
)

# 验证 multi_package 是否存在
assert "multi_package" in report["checks"], "multi_package 检查项缺失！"
print("\n✅ run_all_checks 包含 multi_package 检查项")
