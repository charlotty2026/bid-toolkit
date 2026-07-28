"""
覆盖检查器：检查标书标题是否覆盖了评分项的所有要点。

用法：
  python coverage_check.py 标书.docx --type 工程
  python coverage_check.py 标书.docx --score scoring.json
"""

import argparse
import json
import os
import re
import sys


def load_templates():
    """加载默认评分项模板（兼容开发模式和安装模式）"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # 尝试路径1：安装模式（site-packages/bid_toolkit/scripts/ → bid_toolkit/rules/）
    candidates = [
        os.path.join(script_dir, "..", "rules", "scoring_templates.json"),
        os.path.join(script_dir, "..", "..", "rules", "scoring_templates.json"),
        os.path.join(script_dir, "..", "..", "bid_toolkit", "rules", "scoring_templates.json"),
        os.path.join(script_dir, "..", "..", "..", "rules", "scoring_templates.json"),
    ]
    for path in candidates:
        path = os.path.normpath(path)
        if os.path.isfile(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    print(f"❌ 找不到评分模板文件 scoring_templates.json")
    print(f"   已尝试路径:")
    for p in candidates:
        print(f"   - {os.path.normpath(p)}")
    sys.exit(1)


def load_custom_score(path):
    """加载自定义评分标准文件"""
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def extract_headings(docx_path):
    """从docx提取所有标题文本"""
    try:
        from docx import Document
    except ImportError:
        print("❌ 请先安装 python-docx: pip install python-docx")
        sys.exit(1)

    doc = Document(docx_path)
    headings = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        # 判断是否为标题
        if para.style.name.startswith("Heading") or para.style.name.startswith("heading"):
            headings.append({
                "text": text,
                "style": para.style.name,
                "level": int(re.search(r"\d+", para.style.name).group()) if re.search(r"\d+", para.style.name) else 1,
            })
            continue
        # 也识别以"一、二、三"开头的段落
        cn_nums = "一二三四五六七八九十"
        if re.match(rf"^[{cn_nums}][、.．]\s*\S", text):
            headings.append({
                "text": text,
                "style": "auto-detected",
                "level": 1,
            })
            continue
        # 识别带编号的标题（1.1, 2.2.1等）
        if re.match(r"^\d+[\.\d]*\s+\S", text) and len(text) < 60:
            headings.append({
                "text": text,
                "style": "auto-detected",
                "level": 1,
            })
            continue

    return headings


def auto_detect_type(headings):
    """自动检测标书类型（工程/服务/货物）"""
    all_text = " ".join(h["text"] for h in headings)

    eng_keywords = ["施工", "工程", "安装", "建设", "土建", "机电", "公路", "装修", "改造"]
    svc_keywords = ["服务", "管理", "维护", "保洁", "保安", "物业", "派遣", "咨询", "培训", "劳务"]
    gds_keywords = ["采购", "设备", "货物", "产品", "物资", "供应", "器材", "耗材", "系统"]

    eng_score = sum(1 for k in eng_keywords if k in all_text)
    svc_score = sum(1 for k in svc_keywords if k in all_text)
    gds_score = sum(1 for k in gds_keywords if k in all_text)

    scores = {"工程": eng_score, "服务": svc_score, "货物": gds_score}
    best = max(scores, key=scores.get)
    return best, scores


def check_coverage(headings, template):
    """检查标题覆盖了哪些评分项"""
    all_text = " ".join(h["text"] for h in headings)
    all_text_lower = all_text.lower()
    heading_texts = [h["text"] for h in headings]

    results = []
    covered_count = 0
    total_items = 0

    for cat_name, cat_data in template.get("categories", {}).items():
        items = cat_data.get("items", [])
        total_items += len(items)

        for item in items:
            # 检查这个评分项是否在标题中出现
            found = item.lower() in all_text_lower

            # 更松散的匹配：评分项关键词是否在标题里
            if not found:
                # 拆成关键词（取两个汉字以上的词）
                keywords = re.findall(r"[\u4e00-\u9fff]{2,}", item)
                matched_keywords = [k for k in keywords if any(k in h for h in heading_texts)]
                partial = len(matched_keywords) / max(len(keywords), 1)
                found = partial >= 0.4  # 放宽到40%关键词匹配

            if found:
                covered_count += 1
                status = "✅"
            else:
                status = "❌"

            results.append({
                "category": cat_name,
                "item": item,
                "found": found,
                "status": status,
            })

    return results, covered_count, total_items


def full_heading_text(headings):
    """获取标题文本整体"""
    return " ".join(h["text"] for h in headings)


def format_report(docx_path, type_name, headings, results, covered, total, custom_score=False):
    """输出格式化报告"""
    print(f"\n{'='*60}")
    print(f"📋 评分项覆盖检查报告")
    print(f"{'='*60}")
    print(f"文件: {docx_path}")
    print(f"类型: {type_name} {'(自动识别)' if not custom_score else '(自定义评分标准)'}")
    print(f"标题数: {len(headings)} 个")
    print(f"评分项: {total} 项 | 已覆盖: {covered} 项 | 覆盖率: {covered}/{total} ({covered*100//max(total,1)}%)")
    print(f"{'='*60}")

    if covered < total:
        print(f"\n⚠️  有 {total-covered} 个评分项未被标题覆盖，建议补充：\n")

    last_cat = ""
    for r in results:
        if r["category"] != last_cat:
            print(f"\n  【{r['category']}】")
            last_cat = r["category"]
        print(f"    {r['status']} {r['item']}")

    print(f"\n{'='*60}")

    if covered == total:
        print(f"🎉 全部评分项均被标题覆盖！")
    elif covered >= total * 0.8:
        print(f"✅ 覆盖率超过80%，大框架完整。建议补充剩余项目。")
    else:
        print(f"⚠️  覆盖率不足80%，建议检查标书结构是否齐全。")

    print(f"{'='*60}")

    # 输出标题预览
    print(f"\n📑 当前标题结构：")
    for h in headings[:20]:
        prefix = "  " * h["level"]
        print(f"{prefix}· {h['text'][:50]}")
    if len(headings) > 20:
        print(f"  ...还有 {len(headings)-20} 个标题")


def main():
    parser = argparse.ArgumentParser(description="评分项覆盖检查器")
    parser.add_argument("docx", help="Word 标书文件路径")
    parser.add_argument("--type", "-t", choices=["工程", "服务", "货物"], help="标书类型（不指定则自动识别）")
    parser.add_argument("--score", "-s", help="自定义评分标准JSON文件路径")
    parser.add_argument("--verbose", "-v", action="store_true", help="输出详细匹配信息")

    args = parser.parse_args()

    if not os.path.isfile(args.docx):
        print(f"❌ 文件不存在: {args.docx}")
        sys.exit(1)

    # 1. 提取标题
    print(f"🔍 正在分析: {args.docx}")
    headings = extract_headings(args.docx)
    if not headings:
        print("⚠️  未检测到标题。请确保文档使用了标题样式（Heading 1/2/3等）。")
        sys.exit(0)
    print(f"   检测到 {len(headings)} 个标题")

    # 2. 加载评分标准
    custom_score = False
    if args.score:
        with open(args.score, "r", encoding="utf-8") as f:
            templates = json.load(f)
        # 如果是一份单独的评分标准，直接使用
        if "categories" in templates:
            template = templates
            type_name = "自定义"
            custom_score = True
        else:
            print("❌ 自定义评分标准格式不正确，需要包含 'categories' 字段")
            sys.exit(1)
    else:
        templates = load_templates()
        if args.type:
            type_name = args.type
        else:
            type_name, scores = auto_detect_type(headings)
            print(f"   自动识别类型: {type_name} (工程{scores['工程']}/服务{scores['服务']}/货物{scores['货物']})")
        
        if type_name not in templates:
            available = "、".join(templates.keys())
            print(f"❌ 未找到 '{type_name}' 类型的评分模板。可用类型: {available}")
            sys.exit(1)
        template = templates[type_name]

    # 3. 检查覆盖
    results, covered, total = check_coverage(headings, template)

    # 4. 输出报告
    format_report(args.docx, type_name, headings, results, covered, total, custom_score)


if __name__ == "__main__":
    main()
