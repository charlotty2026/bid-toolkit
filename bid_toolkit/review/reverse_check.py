"""
反向覆盖检查 — Layer 3
=======================
检查招标文件中的风险项是否在投标书中得到回应。

流程：
1. 读审标报告（ScanResult）
2. 读投标书（Markdown或Word）
3. 对每个 fatal 项，检查投标书中是否有对应内容
4. 输出覆盖状态：✅已回应 / ⚠️部分回应 / ❌未回应
"""
import re
from pathlib import Path
from .scanner import ScanResult


def extract_sections_from_bid(bid_path):
    """从投标书中提取章节标题"""
    path = Path(bid_path)
    suffix = path.suffix.lower()

    headings = []
    if suffix == '.docx':
        try:
            from docx import Document
            doc = Document(str(path))
            for para in doc.paragraphs:
                text = para.text.strip()
                if text and para.style.name.startswith(('Heading', 'heading')):
                    headings.append(text)
                # 也识别 # 开头的
                elif text.startswith('#'):
                    headings.append(text.lstrip('#').strip())
        except ImportError:
            pass
    else:
        # 当做文本文件读
        try:
            with open(path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line.startswith('# ') or line.startswith('## '):
                        headings.append(line.lstrip('#').strip())
        except Exception:
            pass

    return headings


def _keyword_to_search_terms(keyword):
    """把判词转成搜索词"""
    # 去除否定词前缀，提取主干
    kw = keyword
    for prefix in ['未', '不', '不得', '视为', '按']:
        if kw.startswith(prefix):
            kw = kw[len(prefix):]
            break
    # 如果太短（<2字）则用原词
    if len(kw) < 2:
        return [keyword]
    return [kw, keyword]


def reverse_coverage_check(result, bid_path):
    """对审标结果中的 fatal 项，检查投标书覆盖情况"""
    fatals = [h for h in result.hits if h.category in ('primary', 'bid_types')]
    if not fatals:
        return {'status': 'no_risks', 'total': 0, 'items': []}

    headings = extract_sections_from_bid(bid_path)
    all_text = ' '.join(headings)
    all_text_lower = all_text.lower()

    items = []
    covered = 0
    for h in fatals:
        search_terms = _keyword_to_search_terms(h.keyword)
        found = False
        for term in search_terms:
            if term.lower() in all_text_lower:
                found = True
                break
        # 再从整个文件搜一次（如果章节没命中，读文件全文）
        if not found:
            try:
                full_text = Path(bid_path).read_text('utf-8', errors='ignore')
                for term in search_terms:
                    if term.lower() in full_text.lower():
                        found = True
                        break
            except Exception:
                pass

        status = 'covered' if found else 'missing'
        if found:
            covered += 1
        items.append({
            'keyword': h.keyword,
            'category': h.category,
            'line_num': h.line_num,
            'status': status,
            'suggestion': '' if found else f'建议补充"{h.keyword}"相关章节',
        })

    return {
        'status': 'ok',
        'total': len(fatals),
        'covered': covered,
        'missing': len(fatals) - covered,
        'coverage_rate': f'{covered}/{len(fatals)} ({covered*100//max(len(fatals),1)}%)',
        'items': items,
    }


def print_coverage_report(report):
    """打印覆盖检查报告"""
    print(f"\n{'='*60}")
    print(f"🔍 反向覆盖检查（风险项→投标书回应）")
    print(f"{'='*60}")
    if report['status'] == 'no_risks':
        print(f"✅ 无风险项需要覆盖检查")
        return

    print(f"总风险项: {report['total']}")
    print(f"已回应:   {report['covered']} ({report['coverage_rate']})")
    print(f"未回应:   {report['missing']}")
    print()

    missing = [i for i in report['items'] if i['status'] == 'missing']
    if missing:
        print(f"⚠️  以下风险项在投标书中未找到对应回应：")
        for i in missing[:10]:
            print(f"  ❌ L{i['line_num']} `{i['keyword']}` → {i['suggestion']}")
        if len(missing) > 10:
            print(f"  ... 还有 {len(missing)-10} 项")
    else:
        print(f"✅ 所有风险项均有回应")

    print(f"{'='*60}\n")
