"""
审标报告生成器 v1.1 — 输出 5+1 张分类清单 + 统计
"""
from .scanner import ScanResult


def _category_label(cid):
    labels = {
        'primary': '🔴 废标/否决',
        'bid_types': '🔴 行业特有风险',
        'secondary': '🟡 评分/格式',
        'contract': '🔵 合同约束',
        'customization': '🔗 商务门槛',
        'certifications': '📄 证明材料',
        'time_nodes': '⏰ 时间节点',
        'emphasis_marks': '⚠️ 强调标识',
    }
    return labels.get(cid, cid)


def format_checklist_md(result):
    lines = []
    lines.append(f"# 📋 招标文件审标报告")
    lines.append(f"")
    lines.append(f"**文件**: {result.file_path}")
    lines.append(f"**字数**: {result.total_chars:,} | **行数**: {result.total_lines}")
    lines.append(f"**风险点**: {len(result.hits)} 处")
    lines.append(f"**生成时间**: {__import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append(f"")
    lines.append(f"---\n")

    # ====== 清单1：致命风险 ======
    fatals = [h for h in result.hits if h.llm_label == 'fatal']
    if fatals:
        lines.append(f"## 🔴 清单1 · 致命风险（{len(fatals)}项）")
        lines.append(f"")
        lines.append(f"| # | 风险信号 | 分类 | 行号 | 简要上下文 |")
        lines.append(f"|---|---------|------|------|-----------|")
        for i, h in enumerate(fatals[:30], 1):
            cat_short = _category_label(h.category)
            ctx = h.context[:50].replace('\n', ' ') + ('...' if len(h.context) > 50 else '')
            lines.append(f"| {i} | `{h.keyword}` | {cat_short} | L{h.line_num} | {ctx} |")
        if len(fatals) > 30:
            lines.append(f"| ... | *共 {len(fatals)} 项，仅展示前30项* | | | |")
        lines.append(f"")
    else:
        lines.append(f"## 🔴 清单1 · 致命风险")
        lines.append(f"✅ 未发现致命风险\n")

    # ====== 清单2：警告项 ======
    warns = [h for h in result.hits if h.llm_label == 'warn']
    lines.append(f"## 🟡 清单2 · 警告项（{len(warns)}项）")
    lines.append(f"")
    lines.append(f"| # | 风险信号 | 分类 | 行号 | 原因 |")
    lines.append(f"|---|---------|------|------|------|")
    for i, h in enumerate(warns[:20], 1):
        cat_short = _category_label(h.category)
        reason = h.llm_reason[:40] if h.llm_reason else ''
        lines.append(f"| {i} | `{h.keyword}` | {cat_short} | L{h.line_num} | {reason} |")
    if len(warns) > 20:
        lines.append(f"| ... | *共 {len(warns)} 项，仅展示前20项* | | | |")
    lines.append(f"")

    # ====== 清单3：证明材料 ======
    certs = [h for h in result.hits if h.category == 'certifications']
    if certs:
        lines.append(f"## 📄 清单3 · 证明材料要求（{len(certs)}项）")
        lines.append(f"")
        lines.append(f"| # | 要求 | 行号 |")
        lines.append(f"|---|------|------|")
        seen_certs = []
        for h in certs:
            key = h.keyword
            if key not in seen_certs:
                seen_certs.append(key)
                lines.append(f"| {len(seen_certs)} | `{key}` | L{h.line_num} |")
        lines.append(f"")

    # ====== 清单4：强调标识 ======
    em = [h for h in result.hits if h.category == 'emphasis_marks']
    if em:
        marks_str = '、'.join(sorted(result.emphasis_marks_found.keys()))
        lines.append(f"## ⚠️ 清单4 · 强调标识参数（{len(em)}处）")
        lines.append(f"检测到强调标识：`{marks_str}`")
        lines.append(f"")
        lines.append(f"| # | 位置 | 上下文 |")
        lines.append(f"|---|------|--------|")
        for i, h in enumerate(em[:15], 1):
            ctx = h.context[:50].replace('\n', ' ') + ('...' if len(h.context) > 50 else '')
            lines.append(f"| {i} | L{h.line_num} | {ctx} |")
        if len(em) > 15:
            lines.append(f"| ... | *共 {len(em)} 处* | |")
        lines.append(f"")

    # ====== 清单5：时间节点 ======
    times = [h for h in result.hits if h.category == 'time_nodes']
    if times:
        lines.append(f"## ⏰ 清单5 · 时间节点（{len(times)}项）")
        lines.append(f"")
        lines.append(f"| # | 信号 | 行号 |")
        lines.append(f"|---|------|------|")
        seen_times = []
        for h in times:
            key = h.keyword
            if key not in seen_times:
                seen_times.append(key)
                lines.append(f"| {len(seen_times)} | `{key}` | L{h.line_num} |")
        lines.append(f"")

    # ====== 附加：商务门槛 ======
    cust = [h for h in result.hits if h.category == 'customization']
    if cust:
        lines.append(f"## 🔗 商务定制门槛（{len(cust)}项）")
        lines.append(f"")
        lines.append(f"| # | 门槛 | 行号 |")
        lines.append(f"|---|------|------|")
        seen_cust = []
        for h in cust:
            key = h.keyword
            if key not in seen_cust:
                seen_cust.append(key)
                lines.append(f"| {len(seen_cust)} | `{key}` | L{h.line_num} |")
        lines.append(f"")

    # ====== 统计摘要 ======
    lines.append(f"---")
    lines.append(f"## 📊 统计摘要")
    lines.append(f"")
    fatal_count = len(fatals)
    warn_count = len(warns)
    cert_count = len(certs)
    em_count = len(em)
    time_count = len(times)
    cust_count = len(cust)
    contract_count = len([h for h in result.hits if h.category == 'contract'])

    lines.append(f"| 分类 | 数量 | 说明 |")
    lines.append(f"|------|------|------|")
    lines.append(f"| 🔴 致命风险 | {fatal_count} | 直接废标/否决项，必须逐条确认 |")
    lines.append(f"| 🟡 警告项 | {warn_count} | 条件性风险，需人工判断 |")
    lines.append(f"| 📄 证明材料 | {cert_count} | 招标文件要求的证明文件 |")
    lines.append(f"| ⚠️ 强调标识 | {em_count} | ▲/★等实质性响应参数 |")
    lines.append(f"| ⏰ 时间节点 | {time_count} | 截止/答疑/异议等时间 |")
    lines.append(f"| 🔗 商务门槛 | {cust_count} | 需提前与厂商建立合作 |")
    lines.append(f"| 🔵 合同约束 | {contract_count} | 中标后履约要求，不导致废标 |")
    lines.append(f"")
    lines.append(f"> 审标报告由 `bid review` 自动生成 — 确认所有 🔴 项已处理后再投标")
    lines.append(f"")

    # ====== 原文速查表 ======
    lines.append(f"---")
    lines.append(f"## 📑 风险项原文位置速查表")
    lines.append(f"")
    lines.append(f"按行号排序，方便在招标文件中精确定位。")
    lines.append(f"")
    lines.append(f"| 行号 | 分类 | 完整命中原文 |")
    lines.append(f"|------|------|-------------|")
    for h in sorted(result.hits, key=lambda x: x.line_num)[:50]:
        cat_short = _category_label(h.category)
        raw = h.context.replace('\n', ' ')[:80]
        lines.append(f"| L{h.line_num} | {cat_short} | `{h.keyword}` {raw} |")
    if len(result.hits) > 50:
        lines.append(f"| ... | *共 {len(result.hits)} 项，仅展示前50行* | |")
    lines.append(f"")

    return '\n'.join(lines)


def format_report(result):
    """终端输出（精简版）"""
    fatals = [h for h in result.hits if h.llm_label == 'fatal']
    warns = [h for h in result.hits if h.llm_label == 'warn']
    certs = [h for h in result.hits if h.category == 'certifications']
    em = [h for h in result.hits if h.category == 'emphasis_marks']
    times = [h for h in result.hits if h.category == 'time_nodes']
    cust = [h for h in result.hits if h.category == 'customization']
    contract = [h for h in result.hits if h.category == 'contract']

    print(f"\n{'='*60}")
    print(f"📋 招标文件审标报告")
    print(f"{'='*60}")
    print(f"文件: {result.file_path}")
    print(f"命中: {len(result.hits)} 处风险点 (已排除噪音+合并同类)")
    print()

    print(f"  🔴 致命风险:    {len(fatals)} 项 — 直接废标/否决，逐条确认")
    print(f"  🟡 警告:        {len(warns)} 项 — 条件性风险")
    print(f"  📄 证明材料:    {len(certs)} 项")
    print(f"  ⚠️ ▲/★标识:    {len(em)} 处")
    print(f"  ⏰ 时间节点:    {len(times)} 项")
    print(f"  🔗 商务门槛:    {len(cust)} 项")
    print(f"  🔵 合同约束:    {len(contract)} 项")

    if fatals:
        print(f"\n--- 🔴 致命风险 Top-15 ---")
        for h in fatals[:15]:
            ctx = h.context[:45].replace('\n', ' ') + ('...' if len(h.context) > 45 else '')
            cat = _category_label(h.category)
            print(f"  L{h.line_num} [{cat}] `{h.keyword[:25]}` {ctx}")
        if len(fatals) > 15:
            rest = len(fatals) - 15
            # 按分类统计剩余的
            from collections import Counter
            cats = Counter(h.category for h in fatals[15:])
            cat_detail = ' | '.join(f'{_category_label(k)}: {v}' for k, v in cats.most_common())
            print(f"  ... 剩余{rest}项 — {cat_detail}")
            print(f"  完整清单: bid review --output report.md")

    print(f"\n💡 `bid review 招标文件.pdf -o report.md` 导出完整清单")
    print(f"{'='*60}\n")
