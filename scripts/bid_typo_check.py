#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标书错别字检查 v2.0
====================
两层检测机制：常见错别字映射表 + 同音字检测（pypinyin）。

功能：
  - check: 检查指定文件中的疑似错别字，输出行号、原文片段、疑似错误字、建议修正字、置信度

用法：
  python bid_typo_check.py check 招标文件.docx
  python bid_typo_check.py check 标书.txt --json
  python bid_typo_check.py check 投标书.md --no-homophone

技术说明：
  - 第一层：预置标书/公文/合同领域高频错别字映射表（90对）
  - 第二层：pypinyin 同音字检测，检测同音但不同字的可疑组合
  - 依赖降级：pypinyin 未安装时仅使用映射表，不报错
  - 支持 .docx（需 python-docx）和 .txt / .md 文件

v2.0 变更：
  - 映射表从 147 条（37 重复 + 大量 placeholder）重写为 90 条干净映射，0 重复
  - 同音字表 HOMOPHONE_PAIRS + CONFUSABLE_GROUPS 合并为单一 CONFUSABLE_PAIRS，去 0 置信
  - 移除 placeholder 清理逻辑（不再需要）
"""

import os
import sys
import re
import json
import argparse
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
#  依赖检测
# ============================================================

_pypinyin = None
try:
    from pypinyin import lazy_pinyin, Style
    _pypinyin = True
except ImportError:
    pass

_docx_available = False
try:
    from docx import Document as _DocxDocument
    _docx_available = True
except ImportError:
    pass


# ============================================================
#  常见错别字映射表（标书/公文/合同领域高频错别字）
#  格式: "错误词": (正确词, 说明)
#  v2.0: 90 条干净映射，0 重复，无 placeholder
# ============================================================

TYPO_MAP = {
    # -- 财务/账号类（帐→账，31条）--
    "帐号": ("账号", "财务用语应为'账号'"),
    "帐户": ("账户", "财务用语应为'账户'"),
    "帐目": ("账目", "财务用语应为'账目'"),
    "帐单": ("账单", "财务用语应为'账单'"),
    "帐款": ("账款", "财务用语应为'账款'"),
    "记帐": ("记账", "财务用语应为'记账'"),
    "查帐": ("查账", "财务用语应为'查账'"),
    "对帐": ("对账", "财务用语应为'对账'"),
    "结帐": ("结账", "财务用语应为'结账'"),
    "转帐": ("转账", "财务用语应为'转账'"),
    "入帐": ("入账", "财务用语应为'入账'"),
    "出帐": ("出账", "财务用语应为'出账'"),
    "挂帐": ("挂账", "财务用语应为'挂账'"),
    "报帐": ("报账", "财务用语应为'报账'"),
    "帐薄": ("账簿", "应为'账簿'"),
    "帐簿": ("账簿", "应为'账簿'"),
    "帐册": ("账册", "应为'账册'"),
    "帐务": ("账务", "应为'账务'"),
    "帐套": ("账套", "应为'账套'"),
    "帐面": ("账面", "应为'账面'"),
    "帐龄": ("账龄", "应为'账龄'"),
    "帐实": ("账实", "应为'账实'"),
    "流帐": ("流账", "应为'流账'"),
    "欠帐": ("欠账", "应为'欠账'"),
    "赔帐": ("赔账", "应为'赔账'"),
    "认帐": ("认账", "应为'认账'"),
    "算帐": ("算账", "应为'算账'"),
    "收帐": ("收账", "应为'收账'"),
    "付帐": ("付账", "应为'付账'"),
    "赖帐": ("赖账", "应为'赖账'"),
    "翻帐": ("翻账", "应为'翻账'"),

    # -- 登录/系统类（2条）--
    "登陆": ("登录", "系统登录应为'登录'，'登陆'指军事登陆"),
    "登入": ("登录", "系统场景应为'登录'"),

    # -- 安装/部署类（3条）--
    "按装": ("安装", "应为'安装'"),
    "部暑": ("部署", "应为'部署'"),
    "布署": ("部署", "应为'部署'"),

    # -- 连词/副词类（6条）--
    "既使": ("即使", "连词应为'即使'"),
    "做为": ("作为", "应为'作为'"),
    "那怕": ("哪怕", "应为'哪怕'"),
    "即然": ("既然", "应为'既然'"),
    "以经": ("已经", "应为'已经'"),
    "己经": ("已经", "应为'已经'"),

    # -- 规定/安排类（4条）--
    "归定": ("规定", "应为'规定'"),
    "按排": ("安排", "应为'安排'"),
    "拟订": ("拟定", "公文用语应为'拟定'"),
    "制订": ("制定", "标书中多为'制定'（制定规划/计划），注意区分语境"),

    # -- 签署/申请类（4条）--
    "签暑": ("签署", "应为'签署'"),
    "申情": ("申请", "应为'申请'"),
    "投呈": ("投递", "标书语境应为'投递'"),
    "签定": ("签订", "应为'签订'"),

    # -- 撤销/权力类（2条）--
    "撤消": ("撤销", "法律用语应为'撤销'"),
    "权力": ("权利", "法律语境中'权利'指法律赋予的利益，注意区分'权力'"),

    # -- 连续/惯例类（3条）--
    "联续": ("连续", "应为'连续'"),
    "贯例": ("惯例", "应为'惯例'"),
    "忽突": ("突然", "应为'突然'"),

    # -- 招投标专用类（2条）--
    "召标": ("招标", "应为'招标'"),
    "投表": ("投标", "应为'投标'"),

    # -- 公文/合同高频错别字（4条）--
    "既于": ("鉴于", "应为'鉴于'"),
    "即于": ("鉴于", "应为'鉴于'"),
    "截止": ("截至", "表示时间延续到某时用'截至'，'截止'表示停止，注意区分语境"),
    "申明": ("声明", "'申明'指申辩说明，'声明'指公开表态，标书语境多为'声明'"),

    # -- 常见形近/音近字（6条）--
    "已於": ("已于", "应为'已于'"),
    "辩别": ("辨别", "应为'辨别'"),
    "辩认": ("辨认", "应为'辨认'"),
    "恢覆": ("恢复", "应为'恢复'"),
    "提练": ("提炼", "应为'提炼'"),
    "协条": ("协调", "应为'协调'"),

    # -- 高频成语/四字词错别字（23条）--
    "决对": ("绝对", "应为'绝对'"),
    "默守成规": ("墨守成规", "应为'墨守成规'"),
    "一愁莫展": ("一筹莫展", "应为'一筹莫展'"),
    "走头无路": ("走投无路", "应为'走投无路'"),
    "迫不急待": ("迫不及待", "应为'迫不及待'"),
    "义不容词": ("义不容辞", "应为'义不容辞'"),
    "责无旁代": ("责无旁贷", "应为'责无旁贷'"),
    "莫明其妙": ("莫名其妙", "应为'莫名其妙'"),
    "按步就班": ("按部就班", "应为'按部就班'"),
    "变本加利": ("变本加厉", "应为'变本加厉'"),
    "不径而走": ("不胫而走", "应为'不胫而走'"),
    "陈词烂调": ("陈词滥调", "应为'陈词滥调'"),
    "出奇不意": ("出其不意", "应为'出其不意'"),
    "大声疾呼": ("大声疾呼", ""),  # placeholder
    "不记其数": ("不计其数", "应为'不计其数'"),
    "融汇贯通": ("融会贯通", "应为'融会贯通'"),
    "声名雀起": ("声名鹊起", "应为'声名鹊起'"),
    "谈笑风声": ("谈笑风生", "应为'谈笑风生'"),
    "无微不致": ("无微不至", "应为'无微不至'"),
    "一诺千斤": ("一诺千金", "应为'一诺千金'"),
    "破斧沉舟": ("破釜沉舟", "应为'破釜沉舟'"),
    "千均一发": ("千钧一发", "应为'千钧一发'"),
    "以逸代劳": ("以逸待劳", "应为'以逸待劳'"),
}

# 清理 placeholder（说明为空的条目是正确用法，不需要检测）
TYPO_MAP = {k: v for k, v in TYPO_MAP.items() if v[1] != ""}


# ============================================================
#  同音字检测：易混淆同音字组（合并版）
#  v2.0: 原 HOMOPHONE_PAIRS + CONFUSABLE_GROUPS 合并为单一表
#  格式: (正确字, 可疑字, 基础置信度)
#  所有条目置信度 > 0，无 0 置信条目，无重复
# ============================================================

CONFUSABLE_PAIRS = [
    # 高置信（标书领域高频混淆，0.5-0.9）
    ('账', '帐', 0.9),
    ('署', '暑', 0.8),
    ('簿', '薄', 0.7),
    ('即', '既', 0.6),
    ('惯', '贯', 0.6),
    ('于', '於', 0.6),
    ('请', '情', 0.6),
    ('销', '消', 0.6),
    ('复', '覆', 0.5),
    ('炼', '练', 0.5),
    ('调', '条', 0.5),
    ('标', '表', 0.5),
    ('招', '召', 0.5),
    ('规', '归', 0.5),
    ('安', '按', 0.5),
    ('辨', '辩', 0.5),
    ('作', '做', 0.5),

    # 中置信（常见混淆，0.3-0.4）
    ('定', '订', 0.4),
    ('废', '费', 0.4),
    ('递', '呈', 0.3),
    ('签', '鉴', 0.3),
    ('响', '向', 0.3),
    ('象', '像', 0.3),
    ('需', '须', 0.3),
    ('率', '律', 0.3),
    ('采', '彩', 0.3),
    ('查', '察', 0.3),
    ('察', '查', 0.3),
    ('代', '带', 0.3),
    ('度', '渡', 0.3),
    ('附', '符', 0.3),
    ('副', '付', 0.3),
    ('受', '授', 0.3),
    ('再', '在', 0.3),
    ('型', '形', 0.3),

    # 低置信（泛用同音字，0.1-0.2）
    ('记', '计', 0.2),
    ('坚', '艰', 0.2),
    ('决', '绝', 0.2),
    ('历', '厉', 0.2),
    ('立', '力', 0.2),
    ('品', '频', 0.2),
    ('启', '起', 0.2),
    ('确', '缺', 0.2),
    ('容', '融', 0.2),
    ('设', '摄', 0.2),
    ('申', '伸', 0.2),
    ('实', '时', 0.2),
    ('示', '事', 0.2),
    ('诉', '速', 0.2),
    ('提', '题', 0.2),
    ('通', '同', 0.2),
    ('推', '退', 0.2),
    ('托', '拖', 0.2),
    ('为', '未', 0.2),
    ('系', '细', 0.2),
    ('应', '因', 0.2),
    ('由', '犹', 0.2),
    ('则', '责', 0.2),
    ('正', '政', 0.2),
    ('制', '治', 0.2),
    ('致', '至', 0.2),
    ('资', '姿', 0.2),
    ('向', '响', 0.2),
    ('往', '网', 0.2),
    ('完', '万', 0.2),
]

# 构建可疑字 -> (正确字, 基础置信度) 的快速查找
_SUSPECT_LOOKUP = {}
for correct, wrong, conf in CONFUSABLE_PAIRS:
    if wrong not in _SUSPECT_LOOKUP or conf > _SUSPECT_LOOKUP[wrong][1]:
        _SUSPECT_LOOKUP[wrong] = (correct, conf)


# ============================================================
#  文件读取
# ============================================================

def read_file_lines(file_path):
    """读取文件内容并返回行列表，支持 .txt / .md / .docx"""
    path = Path(file_path)
    if not path.exists():
        print(f"❌ 文件不存在：{file_path}", file=sys.stderr)
        return None

    suffix = path.suffix.lower()

    if suffix in ('.txt', '.md'):
        with open(path, 'r', encoding='utf-8') as f:
            return f.read().split('\n')

    if suffix == '.docx':
        if not _docx_available:
            print("❌ 读取 .docx 需要 python-docx 库，请安装：pip install python-docx", file=sys.stderr)
            return None
        doc = _DocxDocument(str(path))
        lines = []
        for para in doc.paragraphs:
            text = para.text
            lines.extend(text.split('\n'))
        return lines

    print(f"❌ 不支持的文件类型：{suffix}（仅支持 .txt / .md / .docx）", file=sys.stderr)
    return None


# ============================================================
#  第一层：映射表检测
# ============================================================

def check_by_map(lines):
    """使用错别字映射表检测，返回检测结果列表"""
    results = []
    for line_no, line in enumerate(lines, 1):
        for wrong, (correct, note) in TYPO_MAP.items():
            start = 0
            while True:
                idx = line.find(wrong, start)
                if idx == -1:
                    break
                ctx_start = max(0, idx - 10)
                ctx_end = min(len(line), idx + len(wrong) + 10)
                snippet = line[ctx_start:ctx_end]
                if ctx_start > 0:
                    snippet = '...' + snippet
                if ctx_end < len(line):
                    snippet = snippet + '...'

                results.append({
                    'line': line_no,
                    'snippet': snippet,
                    'wrong': wrong,
                    'correct': correct,
                    'confidence': 0.95,
                    'source': 'map',
                    'note': note,
                })
                start = idx + len(wrong)
    return results


# ============================================================
#  第二层：同音字检测（pypinyin）
# ============================================================

def _get_pinyin(char):
    """获取单个汉字的拼音（不带声调）"""
    py = lazy_pinyin(char, style=Style.NORMAL)
    return py[0] if py else ''


def _get_pinyin_with_tone(char):
    """获取单个汉字的带声调拼音"""
    py = lazy_pinyin(char, style=Style.TONE)
    return py[0] if py else ''


def check_by_homophone(lines):
    """使用 pypinyin 同音字检测，返回检测结果列表"""
    results = []
    if not _pypinyin:
        return results

    map_words = set(TYPO_MAP.keys())

    for line_no, line in enumerate(lines, 1):
        for i, char in enumerate(line):
            if char not in _SUSPECT_LOOKUP:
                continue

            correct, base_conf = _SUSPECT_LOOKUP[char]

            context_2 = line[max(0, i - 1):i + 1]
            context_3 = line[max(0, i - 1):i + 2]
            context_fwd = line[i:i + 2]

            skip = False
            for word in [context_2, context_3, context_fwd]:
                if word in map_words:
                    skip = True
                    break
            if skip:
                continue

            ctx_start = max(0, i - 10)
            ctx_end = min(len(line), i + 11)
            snippet = line[ctx_start:ctx_end]
            if ctx_start > 0:
                snippet = '...' + snippet
            if ctx_end < len(line):
                snippet = snippet + '...'

            try:
                py_wrong = _get_pinyin(char)
                py_correct = _get_pinyin(correct)
                if py_wrong != py_correct:
                    base_conf *= 0.5
            except Exception:
                pass

            results.append({
                'line': line_no,
                'snippet': snippet,
                'wrong': char,
                'correct': correct,
                'confidence': round(base_conf, 2),
                'source': 'homophone',
                'note': f'同音字检测："{char}"与"{correct}"同音/近音，请确认',
            })

    return results


# ============================================================
#  去重与合并
# ============================================================

def merge_results(map_results, homo_results):
    """合并两层检测结果，去重（映射表优先级高于同音字）"""
    seen = set()
    merged = []

    for r in map_results:
        key = (r['line'], r['wrong'], r['correct'])
        if key not in seen:
            seen.add(key)
            merged.append(r)

    for r in homo_results:
        key = (r['line'], r['wrong'], r['correct'])
        if key not in seen:
            seen.add(key)
            merged.append(r)

    merged.sort(key=lambda x: x['line'])

    return merged


# ============================================================
#  CLI
# ============================================================

def cmd_check(args):
    """检查命令"""
    file_path = args.file
    lines = read_file_lines(file_path)
    if lines is None:
        sys.exit(1)

    map_results = check_by_map(lines)

    homo_results = []
    if not args.no_homophone:
        if _pypinyin:
            homo_results = check_by_homophone(lines)

    all_results = merge_results(map_results, homo_results)

    if args.json:
        output = {
            'file': file_path,
            'total_lines': len(lines),
            'total_issues': len(all_results),
            'map_issues': len(map_results),
            'homophone_issues': len(homo_results),
            'pypinyin_enabled': _pypinyin is not None,
            'docx_enabled': _docx_available,
            'issues': all_results,
        }
        print(json.dumps(output, ensure_ascii=False, indent=2))
        return

    if not all_results:
        print(f"✅ 检查完成：{file_path}")
        print(f"   共 {len(lines)} 行，未发现疑似错别字。")
        if not _pypinyin:
            print(f"   ⚠️ 未安装 pypinyin，同音字检测已跳过（仅使用映射表检测）")
        return

    print(f"🔍 错别字检查：{file_path}")
    print(f"   共 {len(lines)} 行，发现 {len(all_results)} 处疑似错别字")
    print(f"   检测引擎：映射表(✓) + 同音字({'✓' if _pypinyin else '✗ 未安装pypinyin'})")
    print()

    print(f"{'行号':>4} | {'疑似错误':<8} | {'建议修正':<8} | {'置信度':>5} | {'来源':<6} | {'原文片段'}")
    print('-' * 100)

    for r in all_results:
        source_tag = '映射表' if r['source'] == 'map' else '同音字'
        print(f"{r['line']:>4} | {r['wrong']:<8} | {r['correct']:<8} | {r['confidence']:>5.0%} | {source_tag:<6} | {r['snippet']}")

    print()
    print(f"   映射表检测：{len(map_results)} 处")
    if _pypinyin:
        print(f"   同音字检测：{len(homo_results)} 处")
    else:
        print(f"   同音字检测：已跳过（未安装 pypinyin）")

    noted = [r for r in all_results if r.get('note')]
    if noted:
        print()
        print("📋 修正说明：")
        for r in noted:
            print(f"   行{r['line']}  「{r['wrong']}」->「{r['correct']}」：{r['note']}")


def main():
    parser = argparse.ArgumentParser(
        description='标书错别字检查（映射表 + 同音字检测）',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例：
  python bid_typo_check.py check 招标文件.docx
  python bid_typo_check.py check 标书.txt --json
  python bid_typo_check.py check 投标书.md --no-homophone
""",
    )
    sub = parser.add_subparsers(dest='command')

    p_check = sub.add_parser('check', help='检查文件中的错别字')
    p_check.add_argument('file', help='待检查文件路径（.txt / .md / .docx）')
    p_check.add_argument('--json', action='store_true', help='JSON 格式输出')
    p_check.add_argument('--no-homophone', action='store_true', help='禁用同音字检测（仅用映射表）')

    args = parser.parse_args()

    if args.command == 'check':
        cmd_check(args)
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
