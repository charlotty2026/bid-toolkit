#!/usr/bin/env python3
"""验证兼投不兼中检查三场景"""

import sys, os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from rfp_compliance import check_multi_package_rule

# 场景1：单包项目
single_pkg = """# 某单位办公用品采购招标文件
本项目采购复印纸、墨盒等办公用品。
投标截止时间：2026年8月15日
"""

# 场景2：多包项目有规则
multi_with_rule = """# 某单位信息化设备采购招标文件
本项目分2个采购包：
第一包：服务器采购
第二包：存储设备采购
兼投不兼中规则：投标人可兼投多个包，但最多只能中1个包。中标包按包号顺序确定。
"""

# 场景3：多包项目缺规则
multi_no_rule = """# 某单位信息化设备采购招标文件
本项目分2个采购包：
第一包：服务器采购
第二包：存储设备采购
各包独立评审，分别签订合同。
"""


def test(label, text, expected_pass, expected_warn=0):
    result = check_multi_package_rule(text)
    passes = sum(1 for r in result if r["severity"] == "pass")
    warns = sum(1 for r in result if r["severity"] == "warn")
    fails = sum(1 for r in result if r["severity"] == "fail")

    ok = passes == expected_pass and warns == expected_warn
    status = "✅ PASS" if ok else "❌ FAIL"
    print(f"{status} [{label}]")
    for r in result:
        print(f"   {r['severity'].upper()}: {r['message']}")
    if not ok:
        print(
            f"   (期望 pass={expected_pass} warn={expected_warn}, 实际 pass={passes} warn={warns} fail={fails})"
        )
    print()
    return ok


all_ok = True
all_ok &= test("单包项目", single_pkg, 1, 0)
all_ok &= test("多包有规则", multi_with_rule, 1, 0)
all_ok &= test("多包缺规则", multi_no_rule, 0, 1)

print("=" * 50)
print("结果：全部通过" if all_ok else "结果：存在失败")
sys.exit(0 if all_ok else 1)
