#!/usr/bin/env python3
"""
承诺链三源追踪 —— 扫描标书中的承诺，逐条对照企业资料库验证证据支撑。

承诺模式识别：
  1. 人员承诺：「指派X年以上经验的项目经理」「配备注册XX师」
  2. 设备承诺：「提供X台XX设备」「配备XX台」
  3. 业绩承诺：「服务过X家XX客户」「完成X个同类项目」
  4. 资质承诺：「具备XX资质」「取得XX认证」
  5. 时间承诺：「X小时内响应」「X日内到场」

三源证据：
  源1: company_profile/team.md       — 人员资质
  源2: company_profile/performance.md — 业绩案例
  源3: company_profile/qualifications.md — 资质证书

用法：
  python commitment_scanner.py 标书.md --profile company_profile/
  python commitment_scanner.py 标书.md --profile company_profile/ --report 承诺链审计.md
"""

import re
import sys
from pathlib import Path


# ── 承诺模式识别 ──

# 人员承诺：X年经验 / 注册XX师 / 持有XX证书
PATTERN_PERSONNEL = [
    (r'(\d+)[-~]?(\d+)?\s*年\s*以[上内]\s*(?:相关\s*)?(?:工作|从业|行业|项目)经验',
     'personnel', '年经验'),
    (r'(?:指派|配备|安排|拟派|投入|配置)\s*.{1,20}(?:项目经理|项目负责人|技术负责人|工程师|专员|主管|经理|人员|团队)',
     'personnel', '人员安排'),
    (r'(?:注册|持[有证]|具备|具有)\s*.{1,10}(?:工程师|建造师|造价师|监理师|安全员|会计师|经济师|评估师|咨询师|设计师|建筑师)',
     'personnel', '注册证书'),
    (r'(?:持有|取得|通过|具备)\s*.{1,20}(?:证书|资格证|上岗证|资质证|执业资格)',
     'personnel', '持证要求'),
]

# 设备承诺：X台设备 / 配备XX
PATTERN_EQUIPMENT = [
    (r'(\d+)\s*[台套辆](?:\s*以[上内])?\s*.{1,15}(?:设备|车辆|仪器|机械|工具|电脑|系统|平台)',
     'equipment', '设备数量'),
    (r'(?:配备|提供|配置|投入|采购)\s*.{1,20}(?:设备|车辆|仪器|机械|工具|电脑)',
     'equipment', '设备配备'),
]

# 业绩承诺：X家客户 / 服务过XX
PATTERN_PERFORMANCE = [
    (r'(?:服务过|服务过|承接过|完成过|实施过|交付过|具备)\s*.{1,30}(?:项目|客户|案例|经验|业绩)',
     'performance', '业绩经验'),
    (r'(\d+)\s*[个家项次]\s*(?:以[上内])?\s*(?:同类|类似|相关|同类型)\s*(?:项目|客户|案例|业绩)',
     'performance', '项目数量'),
    (r'(\d+)\s*年\s*(?:以[上内])?\s*(?:行业|领域|同类)\s*(?:经验|积累|沉淀)',
     'performance', '行业年限'),
]

# 资质承诺：具备XX资质
PATTERN_QUALIFICATION = [
    (r'(?:具有|具备|取得|持有|通过)\s*.{1,20}(?:资质|认证|许可|批准|备案|登记)',
     'qualification', '资质持有'),
    (r'(?:通过|获得|取得)\s*.{1,20}(?:ISO|认证|体系|标准|评审)',
     'qualification', '体系认证'),
]

# 时间承诺：X小时内响应
PATTERN_TIMELINE = [
    (r'(\d+)\s*(?:小时|分钟|天|日|个工作日)\s*(?:内|以[内上])\s*(?:响应|到场|到达|处理|完成|解决|反馈|修复)',
     'timeline', '时限响应'),
    (r'(?:7×24|24小时|全天候|全年无休|随时)',
     'timeline', '全天候服务'),
]

# 综合承诺模式
ALL_PATTERNS = (
    PATTERN_PERSONNEL +
    PATTERN_EQUIPMENT +
    PATTERN_PERFORMANCE +
    PATTERN_QUALIFICATION +
    PATTERN_TIMELINE
)


# ── 证据匹配引擎 ──

def load_profile(profile_dir: str) -> dict:
    """加载企业资料库，返回结构化数据"""
    prof = {'company_name': '', 'team': [], 'qualifications': [], 'performance': [], 'equipment': []}
    pdir = Path(profile_dir)
    if not pdir.exists():
        return prof

    # 团队 → 人员资质列表
    team_file = pdir / 'team.md'
    if team_file.exists():
        text = team_file.read_text(encoding='utf-8')
        # 提取人员信息（每行一个，格式：- 张三/项目经理/5年经验/注册建造师）
        for line in text.split('\n'):
            line = line.strip()
            if line.startswith('- ') or line.startswith('* '):
                prof['team'].append(line.lstrip('- *').strip())

    # 资质
    qual_file = pdir / 'qualifications.md'
    if qual_file.exists():
        for line in qual_file.read_text(encoding='utf-8').split('\n'):
            line = line.strip()
            if line.startswith('- ') or line.startswith('* '):
                prof['qualifications'].append(line.lstrip('- *').strip())

    # 业绩
    perf_file = pdir / 'performance.md'
    if perf_file.exists():
        for line in perf_file.read_text(encoding='utf-8').split('\n'):
            line = line.strip()
            if line.startswith('- ') or line.startswith('* '):
                prof['performance'].append(line.lstrip('- *').strip())

    # 公司名
    info_file = pdir / 'company_info.md'
    if info_file.exists():
        for line in info_file.read_text(encoding='utf-8').split('\n'):
            m = re.search(r'## 公司名称\s*\n\s*(.+?)\s*$', line)
            if m:
                prof['company_name'] = m.group(1).strip()

    return prof


def match_evidence(commitment_text: str, category: str, profile: dict) -> tuple:
    """
    匹配承诺到证据。
    返回：(matched: bool, evidence: str, score: int)
        score: 0=无证据, 1=弱匹配(领域匹配但不精确), 2=强匹配(具体证据)
    """
    matched = False
    evidence = ""
    score = 0

    if category == 'personnel':
        # 人员承诺 → 匹配 team 列表
        for member in profile['team']:
            # 检查承诺中的关键词是否在人员描述中出现
            keywords = re.findall(r'[^\s，,。；;]{2,10}', commitment_text)
            hit_count = 0
            for kw in keywords:
                if kw in member:
                    hit_count += 1
            if hit_count >= 2:
                matched = True
                evidence = member
                score = 2
                break
        if not matched:
            # 弱匹配：检查是否有"年经验"关键词
            if '年' in commitment_text and profile['team']:
                matched = True
                evidence = f"团队有{len(profile['team'])}人，具体年限需核实"
                score = 1

    elif category == 'qualification':
        # 资质承诺 → 匹配 qualifications 列表
        for qual in profile['qualifications']:
            keywords = re.findall(r'[^\s，,。；;]{2,10}', commitment_text)
            for kw in keywords:
                if kw in qual:
                    matched = True
                    evidence = qual
                    score = 2
                    break
            if matched:
                break
        if not matched and profile['qualifications']:
            matched = True
            evidence = f"有{len(profile['qualifications'])}项资质，具体需核实"
            score = 1

    elif category == 'performance':
        # 业绩承诺 → 匹配 performance 列表
        for proj in profile['performance']:
            keywords = re.findall(r'[^\s，,。；;]{2,10}', commitment_text)
            hit_count = 0
            for kw in keywords:
                if kw in proj:
                    hit_count += 1
            if hit_count >= 2:
                matched = True
                evidence = proj
                score = 2
                break
        if not matched and profile['performance']:
            matched = True
            evidence = f"有{len(profile['performance'])}项业绩，具体需核实"
            score = 1

    elif category == 'equipment':
        # 设备承诺 → 弱匹配（需要 equipment 列表）
        if profile.get('equipment'):
            matched = True
            evidence = f"设备列表有{len(profile['equipment'])}项，具体需核实"
            score = 1
        else:
            matched = False
            evidence = "⚠️ 企业资料库无设备清单，无法验证设备承诺"
            score = 0

    elif category == 'timeline':
        # 时间承诺 → 无法从企业资料库验证，标记为"需人工确认"
        matched = True
        evidence = "⏰ 时间承诺需人工确认：是否具备相应响应能力"
        score = 1

    return matched, evidence, score


def scan_commitments(text: str, profile: dict) -> list:
    """
    扫描文本中的所有承诺，返回审计结果列表。
    每个结果：{line, text, category, subcategory, evidence, score, matched}
    """
    results = []
    lines = text.split('\n')
    seen = set()  # 去重

    for line_no, line in enumerate(lines, 1):
        stripped = line.strip()
        if not stripped or len(stripped) < 8:
            continue

        for pattern, category, subcat in ALL_PATTERNS:
            for m in re.finditer(pattern, stripped):
                committed_text = m.group(0)
                if committed_text in seen:
                    continue
                seen.add(committed_text)

                matched, evidence, score = match_evidence(committed_text, category, profile)

                results.append({
                    'line': line_no,
                    'text': stripped[:100],  # 截断
                    'commitment': committed_text,
                    'category': category,
                    'subcategory': subcat,
                    'matched': matched,
                    'score': score,
                    'evidence': evidence,
                })
                break  # 每行只匹配第一个模式

    return results


def format_report(results: list, profile: dict) -> str:
    """生成人类可读的审计报告"""
    total = len(results)
    matched = sum(1 for r in results if r['matched'] and r['score'] >= 2)
    weak = sum(1 for r in results if r['matched'] and r['score'] == 1)
    missing = sum(1 for r in results if not r['matched'])

    lines = []
    lines.append("=" * 60)
    lines.append("  承诺链三源追踪审计报告")
    lines.append("=" * 60)
    lines.append(f"  企业: {profile.get('company_name', '未设置')}")
    lines.append(f"  承诺总数: {total} 条")
    lines.append(f"  ✅ 强证据支撑: {matched} 条")
    lines.append(f"  ⚠️  弱证据支撑: {weak} 条")
    lines.append(f"  ❌ 无证据支撑: {missing} 条")
    lines.append("=" * 60)

    if missing > 0:
        lines.append("\n🔴 无证据支撑的承诺（需要重点关注）:")
        for r in results:
            if not r['matched']:
                lines.append(f"  L{r['line']} | {r['subcategory']}: {r['commitment']}")
                lines.append(f"    → 第{r['line']}行: {r['text'][:80]}")

    if weak > 0:
        lines.append("\n🟡 弱证据支撑的承诺（建议补充具体证据）:")
        for r in results:
            if r['matched'] and r['score'] == 1:
                lines.append(f"  L{r['line']} | {r['subcategory']}: {r['commitment']}")
                lines.append(f"    → {r['evidence']}")

    if matched > 0:
        lines.append("\n✅ 强证据支撑的承诺:")
        for r in results:
            if r['score'] >= 2:
                lines.append(f"  L{r['line']} | {r['subcategory']}: {r['commitment']}")
                lines.append(f"    → ✅ {r['evidence']}")

    lines.append("\n" + "=" * 60)
    conc = ""
    if missing == 0 and weak == 0:
        conc = "所有承诺均有证据支撑 ✅"
    elif missing == 0 and weak > 0:
        conc = "承诺有企业资料支撑，但部分证据不够具体 🟡"
    else:
        conc = f"有 {missing} 条承诺无证据支撑，建议补充企业资料或修改承诺 🔴"
    lines.append("  审计结论: " + conc)

    return '\n'.join(lines)


def main():
    import argparse
    parser = argparse.ArgumentParser(description='承诺链三源追踪——标书承诺逐条验证')
    parser.add_argument('file', help='标书文件路径（.md/.txt）')
    parser.add_argument('--profile', '-p', default='company_profile', help='企业资料库目录')
    parser.add_argument('--report', '-o', help='输出审计报告到文件')
    args = parser.parse_args()

    file_path = Path(args.file)
    if not file_path.exists():
        print(f"❌ 文件不存在: {file_path}")
        sys.exit(1)

    text = file_path.read_text(encoding='utf-8')
    profile = load_profile(args.profile)
    results = scan_commitments(text, profile)
    report = format_report(results, profile)

    print(report)

    if args.report:
        Path(args.report).write_text(report, encoding='utf-8')
        print(f"\n📄 报告已保存: {args.report}")


if __name__ == '__main__':
    main()