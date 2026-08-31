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
    """检查标题覆盖了哪些评分项（三层匹配：精确→模糊→同义）"""
    all_text = " ".join(h["text"] for h in headings)
    all_text_lower = all_text.lower()
    heading_texts = [h["text"] for h in headings]

    # 同义/近义词映射（评分项 → 标题中可能出现的变体）
    synonym_map = {
        "人员配置方案": ["人员配置", "项目团队配置", "团队配置", "人员安排", "人力配置"],
        "项目团队经验": ["项目团队", "团队经验", "项目经验", "团队资历"],
        "项目负责人": ["项目经理", "项目负责人", "项目主管"],
        "类似项目业绩": ["类似业绩", "项目业绩", "同类项目", "成功案例", "过往业绩"],
        "售后服务方案": ["售后服务", "服务保障", "售后承诺"],
        "企业资质": ["资质证书", "企业资质", "公司资质", "营业执照"],
        "施工组织设计": ["施组", "施工组织"],
        "培训方案": ["培训计划", "培训体系", "人员培训", "员工培训"],
        "质量管理方案": ["质量保证", "质量管理", "质量体系", "质量控制", "质检"],
        "应急预案": ["应急方案", "应急措施", "突发事件", "应急处置"],
        "财务状况": ["财务", "营收", "利润", "资产"],
        "企业信誉": ["信誉", "信用", "获奖", "荣誉", "AAA"],
        "增值服务": ["增值", "额外服务", "特色服务"],
        "报价合理性": ["报价", "报价说明", "价格"],
        "费用明细": ["费用组成", "费用构成", "报价明细", "费用清单"],
        "安全措施": ["安全", "安全生产"],
        "进度计划": ["进度", "工期", "时间安排"],
    }

    results = []
    covered_count = 0
    partial_count = 0
    total_items = 0

    for cat_name, cat_data in template.get("categories", {}).items():
        items = cat_data.get("items", [])
        total_items += len(items)

        for item in items:
            # 匹配层次1：精确匹配
            found = item.lower() in all_text_lower
            match_type = "exact"

            # 匹配层次2：关键词模糊匹配
            if not found:
                keywords = re.findall(r"[\u4e00-\u9fff]{2,}", item)
                matched_keywords = [k for k in keywords if any(k in h for h in heading_texts)]
                partial = len(matched_keywords) / max(len(keywords), 1)
                if partial >= 0.4:
                    found = True
                    match_type = "fuzzy"

            # 匹配层次3：同义/近义词匹配
            if not found and item in synonym_map:
                for synonym in synonym_map[item]:
                    if synonym.lower() in all_text_lower:
                        found = True
                        match_type = "synonym"
                        break

            if found:
                covered_count += 1
                if match_type == "exact":
                    status = "✅"
                elif match_type == "fuzzy":
                    status = "🔶"
                    partial_count += 1
                else:
                    status = "🔷"
                    partial_count += 1
            else:
                status = "❌"

            results.append({
                "category": cat_name,
                "item": item,
                "found": found,
                "status": status,
                "match_type": match_type,
            })

    return results, covered_count, partial_count, total_items


def full_heading_text(headings):
    """获取标题文本整体"""
    return " ".join(h["text"] for h in headings)


def format_report(docx_path, type_name, headings, results, covered, total, custom_score=False, output_path=None):
    """输出格式化报告（支持 --output 导出为 markdown）"""
    score = covered * 5  # 每覆盖一项加5分
    max_score = total * 5

    lines = []
    lines.append(f"{'='*60}")
    lines.append(f"📋 评分项覆盖检查报告")
    lines.append(f"{'='*60}")
    lines.append(f"文件: {docx_path}")
    lines.append(f"类型: {type_name} {'(自动识别)' if not custom_score else '(自定义评分标准)'}")
    lines.append(f"标题数: {len(headings)} 个")
    lines.append(f"评分项: {total} 项 | 已覆盖: {covered} 项 | 覆盖率: {covered}/{total} ({covered*100//max(total,1)}%)")
    lines.append(f"参考评分: {score}/{max_score} 分 (每覆盖一项+5分)")
    lines.append(f"{'='*60}")

    if covered < total:
        lines.append(f"\n⚠️  有 {total-covered} 个评分项未被标题覆盖，建议补充：\n")

    last_cat = ""
    for r in results:
        if r["category"] != last_cat:
            lines.append(f"\n  【{r['category']}】")
            last_cat = r["category"]
        lines.append(f"    {r['status']} {r['item']}")

    lines.append(f"\n{'='*60}")

    if covered == total:
        lines.append(f"🎉 全部评分项均被标题覆盖！得分: {score}/{max_score}")
    elif covered >= total * 0.8:
        lines.append(f"✅ 覆盖率超过80%，大框架完整。建议补充剩余项目。得分: {score}/{max_score}")
    else:
        lines.append(f"⚠️  覆盖率不足80%，建议检查标书结构是否齐全。得分: {score}/{max_score}")

    lines.append(f"{'='*60}")

    # 标题预览
    lines.append(f"\n📑 当前标题结构：")
    for h in headings[:20]:
        prefix = "  " * min(h.get("level", 1), 5)
        lines.append(f"{prefix}· {h['text'][:50]}")
    if len(headings) > 20:
        lines.append(f"  ...还有 {len(headings)-20} 个标题")

    output = "\n".join(lines)
    print(output)

    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(output)
        print(f"\n📁 报告已导出: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="评分项覆盖检查器")
    parser.add_argument("docx", help="Word 标书文件路径")
    parser.add_argument("--type", "-t", choices=["工程", "服务", "货物"], help="标书类型（不指定则自动识别）")
    parser.add_argument("--score", "-s", help="自定义评分标准JSON文件路径")
    parser.add_argument("--output", "-o", help="导出报告到文件（.txt / .md）")
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
    results, covered, partial_matches, total = check_coverage(headings, template)

    # 4. 输出报告
    format_report(args.docx, type_name, headings, results, covered, total, custom_score, args.output)


if __name__ == "__main__":
    main()
