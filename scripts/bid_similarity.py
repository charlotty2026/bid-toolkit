#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标书查重工具 v1.0
==================
基于SimHash算法的标书文本相似度检测，支持段落级对比。

功能：
  - compare: 比较两个文件的相似度，输出重复段落
  - check:   拿一个文件和历史库中所有文件比对，找出最相似的
  - add:     把已完成标书加入历史库

用法：
  python bid_similarity.py compare 文件A.md 文件B.md [--threshold 0.8] [--json]
  python bid_similarity.py check 新标书.md --library ./bid_library/ [--threshold 0.8] [--json]
  python bid_similarity.py add 标书.md --library ./bid_library/ --name "项目名"

技术说明：
  - SimHash: 64位指纹，汉明距离≤3视为相似
  - 分词：优先使用jieba，未安装则退化为3-gram
  - 支持 .md / .txt / .docx 输入（docx需python-docx）
"""

import os
import sys
import re
import json
import hashlib
import argparse
from pathlib import Path
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
#  依赖检测
# ============================================================

# jieba分词（可选）
_jieba = None
try:
    import jieba
    _jieba = jieba
except ImportError:
    pass

# python-docx（可选，仅docx需要）
_docx_available = False
try:
    from docx import Document as _DocxDocument
    _docx_available = True
except ImportError:
    pass


# ============================================================
#  文件读取
# ============================================================

def read_file(file_path):
    """读取文件内容，支持 .md / .txt / .docx"""
    path = Path(file_path)
    if not path.exists():
        print(f"❌ 文件不存在：{file_path}", file=sys.stderr)
        sys.exit(1)

    suffix = path.suffix.lower()

    if suffix in ('.md', '.txt'):
        with open(path, 'r', encoding='utf-8') as f:
            return f.read()

    if suffix == '.docx':
        if not _docx_available:
            print("❌ 读取docx需要 python-docx: pip install python-docx", file=sys.stderr)
            sys.exit(1)
        doc = _DocxDocument(str(path))
        paragraphs = []
        for para in doc.paragraphs:
            paragraphs.append(para.text)
        return '\n\n'.join(paragraphs)

    # 尝试当作文本读取
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception:
        print(f"❌ 不支持的文件格式：{suffix}（支持 .md / .txt / .docx）", file=sys.stderr)
        sys.exit(1)


# ============================================================
#  段落拆分
# ============================================================

def split_paragraphs(text):
    """把文本按段落拆分，返回非空段落列表（去掉markdown标题标记）"""
    # 按双换行或单换行分割
    raw_paras = re.split(r'\n\s*\n', text)

    paragraphs = []
    for para in raw_paras:
        # 去掉markdown标题前缀（#号）
        cleaned = re.sub(r'^#{1,6}\s*', '', para.strip())
        # 去掉markdown列表标记
        cleaned = re.sub(r'^[\s]*[-*+]\s+', '', cleaned, flags=re.MULTILINE)
        # 去掉markdown加粗/斜体标记
        cleaned = re.sub(r'\*{1,3}([^*]+)\*{1,3}', r'\1', cleaned)
        # 合并多行为一行
        cleaned = re.sub(r'\n+', ' ', cleaned).strip()
        # 过滤太短的段落（少于10个字符的跳过）
        if len(cleaned) >= 10:
            paragraphs.append(cleaned)

    return paragraphs


# ============================================================
#  分词
# ============================================================

def tokenize(text):
    """分词：优先jieba，退化方案用3-gram"""
    if _jieba is not None:
        # jieba分词，过滤纯标点和空白
        words = _jieba.cut(text)
        tokens = [w.strip() for w in words if w.strip() and not re.match(r'^[\s\W]+$', w)]
        if tokens:
            return tokens
        # 如果分词结果为空，退化到3-gram

    # 3-gram退化方案
    # 去掉空白字符后按3个字符一组
    clean_text = re.sub(r'\s+', '', text)
    if len(clean_text) < 3:
        return [clean_text] if clean_text else []
    return [clean_text[i:i+3] for i in range(len(clean_text) - 2)]


# ============================================================
#  SimHash 算法
# ============================================================

def _hash64(token):
    """对单个token计算64位哈希（MD5取前8字节）"""
    h = hashlib.md5(token.encode('utf-8')).digest()
    # 取前8字节转为整数
    return int.from_bytes(h[:8], 'big')


def simhash(tokens):
    """
    计算文本的SimHash指纹（64位）
    :param tokens: 分词后的token列表
    :return: 64位整数指纹
    """
    if not tokens:
        return 0

    # 统计token频率作为权重
    freq = defaultdict(int)
    for t in tokens:
        freq[t] += 1

    # 64位的权重向量，初始化为0
    v = [0] * 64

    for token, weight in freq.items():
        h = _hash64(token)
        for i in range(64):
            bit = (h >> i) & 1
            if bit:
                v[i] += weight
            else:
                v[i] -= weight

    # 生成指纹
    fingerprint = 0
    for i in range(64):
        if v[i] > 0:
            fingerprint |= (1 << i)

    return fingerprint


def hamming_distance(hash1, hash2):
    """计算两个64位哈希的汉明距离"""
    x = hash1 ^ hash2
    # Brian Kernighan算法计数
    count = 0
    while x:
        x &= x - 1
        count += 1
    return count


def simhash_similarity(hash1, hash2):
    """
    根据汉明距离计算相似度（0~1）
    汉明距离0=完全相同，64=完全不同
    """
    dist = hamming_distance(hash1, hash2)
    return 1.0 - dist / 64.0


# ============================================================
#  段落级对比
# ============================================================

def compare_paragraphs(paras_a, paras_b, threshold=0.8):
    """
    段落级对比：计算所有段落对的相似度，找出超过阈值的
    :return: 匹配列表 [{idx_a, idx_b, similarity, text_a, text_b}]
    """
    # 预计算所有段落的simhash
    hashes_a = [simhash(tokenize(p)) for p in paras_a]
    hashes_b = [simhash(tokenize(p)) for p in paras_b]

    matches = []
    for i, (ha, pa) in enumerate(zip(hashes_a, paras_a)):
        for j, (hb, pb) in enumerate(zip(hashes_b, paras_b)):
            sim = simhash_similarity(ha, hb)
            if sim >= threshold:
                matches.append({
                    'idx_a': i + 1,  # 段落编号从1开始
                    'idx_b': j + 1,
                    'similarity': sim,
                    'text_a': pa,
                    'text_b': pb,
                })

    # 按相似度降序排列
    matches.sort(key=lambda x: x['similarity'], reverse=True)
    return matches


def overall_similarity(paras_a, paras_b):
    """
    计算两个文档的整体相似度
    方法：对所有段落对取最高相似度的均值，再综合文档级SimHash
    """
    if not paras_a or not paras_b:
        return 0.0

    # 文档级SimHash相似度
    full_text_a = ' '.join(paras_a)
    full_text_b = ' '.join(paras_b)
    hash_a = simhash(tokenize(full_text_a))
    hash_b = simhash(tokenize(full_text_b))
    doc_sim = simhash_similarity(hash_a, hash_b)

    # 段落级最佳匹配平均
    hashes_a = [simhash(tokenize(p)) for p in paras_a]
    hashes_b = [simhash(tokenize(p)) for p in paras_b]

    # 对A中每个段落，找B中最相似的
    best_sims_a = []
    for ha in hashes_a:
        best = max(simhash_similarity(ha, hb) for hb in hashes_b)
        best_sims_a.append(best)

    # 对B中每个段落，找A中最相似的
    best_sims_b = []
    for hb in hashes_b:
        best = max(simhash_similarity(ha, hb) for ha in hashes_a)
        best_sims_b.append(best)

    # 综合相似度 = 文档级(40%) + 段落级A→B(30%) + 段落级B→A(30%)
    para_sim_a = sum(best_sims_a) / len(best_sims_a)
    para_sim_b = sum(best_sims_b) / len(best_sims_b)
    overall = doc_sim * 0.4 + para_sim_a * 0.3 + para_sim_b * 0.3

    return min(overall, 1.0)


# ============================================================
#  目录结构提取
# ============================================================

def extract_headings(text):
    """从文本中提取标题层级结构

    支持 markdown (#) 和 docx 转文本后的 [Heading] 标记。
    返回: [(level, title), ...]
    """
    headings = []
    for line in text.split('\n'):
        line = line.strip()
        if not line:
            continue
        # markdown 标题: # / ## / ###
        m = re.match(r'^(#{1,6})\s+(.+)', line)
        if m:
            level = len(m.group(1))
            title = m.group(2).strip()
            headings.append((level, title))
            continue
        # docx 转文本可能保留 [Heading 1]xxx 格式（非标准，兼容性处理）
        m = re.match(r'^\[Heading\s+(\d+)\]\s*(.+)', line)
        if m:
            level = int(m.group(1))
            title = m.group(2).strip()
            headings.append((level, title))
            continue
        # 中文数字标题：一、 / （一） / 1. （粗匹配，level=2）
        m = re.match(r'^[一二三四五六七八九十]+[、.]\s*(.+)', line)
        if m and len(line) < 50:
            headings.append((2, line))
            continue
        m = re.match(r'^[（(][一二三四五六七八九十]+[)）]\s*(.+)', line)
        if m and len(line) < 50:
            headings.append((3, line))
    return headings


def compare_structure(headings_a, headings_b):
    """目录结构相似度比对

    三个维度:
    1. 标题数量比（结构规模是否接近）
    2. 标题文本SimHash（标题用词是否雷同）
    3. 层级序列比（章节骨架是否一致）

    返回: {'similarity': 0~1, 'details': str}
    """
    if not headings_a or not headings_b:
        return {'similarity': 0.0, 'details': '标题不足，无法比对目录结构'}

    # 维度1: 数量比
    count_ratio = min(len(headings_a), len(headings_b)) / max(len(headings_a), len(headings_b))

    # 维度2: 标题文本SimHash
    titles_a = ' '.join(h[1] for h in headings_a)
    titles_b = ' '.join(h[1] for h in headings_b)
    hash_a = simhash(tokenize(titles_a))
    hash_b = simhash(tokenize(titles_b))
    title_sim = simhash_similarity(hash_a, hash_b)

    # 维度3: 层级序列比（只看层级不看文本）
    levels_a = [h[0] for h in headings_a]
    levels_b = [h[0] for h in headings_b]
    # 用最长公共子序列比
    lcs_len = _lcs_length(levels_a, levels_b)
    level_sim = (2 * lcs_len) / (len(levels_a) + len(levels_b)) if (levels_a and levels_b) else 0

    # 综合权重: 文本SimHash 50% + 层级 30% + 数量 20%
    overall = title_sim * 0.5 + level_sim * 0.3 + count_ratio * 0.2
    overall = min(overall, 1.0)

    details = (f'标题数量: {len(headings_a)} vs {len(headings_b)} '
               f'(比{count_ratio:.2f}), '
               f'标题文本相似{title_sim:.1%}, '
               f'层级结构相似{level_sim:.1%}')

    return {'similarity': overall, 'details': details}


def _lcs_length(seq_a, seq_b):
    """最长公共子序列长度（动态规划）"""
    m, n = len(seq_a), len(seq_b)
    dp = [[0] * (n + 1) for _ in range(m + 1)]
    for i in range(1, m + 1):
        for j in range(1, n + 1):
            if seq_a[i - 1] == seq_b[j - 1]:
                dp[i][j] = dp[i - 1][j - 1] + 1
            else:
                dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])
    return dp[m][n]


# ============================================================
#  元数据提取与比对
# ============================================================

METADATA_PATTERNS = {
    '项目名称': [
        r'项目名称\s*[:：]\s*(.+?)[\n。，；]',
        r'工程名称\s*[:：]\s*(.+?)[\n。，；]',
    ],
    '项目编号': [
        r'项目编号\s*[:：]\s*(.+?)[\n。，；]',
        r'招标编号\s*[:：]\s*(.+?)[\n。，；]',
        r'标段编号\s*[:：]\s*(.+?)[\n。，；]',
    ],
    '招标单位': [
        r'招标(?:人|单位|方)\s*[:：]\s*(.+?)[\n。，；]',
        r'采购(?:人|单位|方)\s*[:：]\s*(.+?)[\n。，；]',
    ],
    '投标单位': [
        r'投标(?:人|单位|方)\s*[:：]\s*(.+?)[\n。，；]',
        r'供应商\s*[:：]\s*(.+?)[\n。，；]',
    ],
    '总报价': [
        r'(?:总报价|投标报价|总[价金额计])\s*[:：为是]?\s*(\d[\d,]*(?:\.\d+)?)\s*(万元|元|万)',
    ],
    '服务期限': [
        r'(?:服务期[限制]|项目期[限制]|合同期[限制])\s*[:：为]?\s*(\d+)\s*(?:个?月|年|天)',
    ],
    '投标有效期': [
        r'投标有效期\s*[:：为]?\s*(\d+)\s*天',
    ],
}


def extract_metadata(text):
    """从标书文本中提取关键元数据

    返回: {field: [value1, value2, ...]}  （同一字段可能出现多次）
    """
    metadata = {}
    for field, patterns in METADATA_PATTERNS.items():
        for pat in patterns:
            for m in re.finditer(pat, text):
                # 最后一个group是值（金额/期限取group1+group2，文本取group1）
                if m.lastindex and m.lastindex >= 2:
                    value = f'{m.group(1)}{m.group(2)}'
                else:
                    value = m.group(1).strip()
                if field not in metadata:
                    metadata[field] = []
                if value not in metadata[field]:
                    metadata[field].append(value)
    return metadata


def compare_metadata(meta_a, meta_b):
    """元数据比对

    返回: {
        'matches': [(field, value)],   # 两份文档都有且相同的字段
        'conflicts': [(field, val_a, val_b)],  # 两份文档都有但值不同
        'a_only': [(field, value)],
        'b_only': [(field, value)],
        'cross_project_risk': str,  # 串项目风险评估
    }
    """
    matches = []
    conflicts = []
    a_only = []
    b_only = []

    all_fields = set(meta_a.keys()) | set(meta_b.keys())
    for field in all_fields:
        vals_a = meta_a.get(field, [])
        vals_b = meta_b.get(field, [])

        if vals_a and vals_b:
            # 取第一个值做比较（标书中同一字段通常只有一个值）
            va, vb = vals_a[0], vals_b[0]
            if va == vb:
                matches.append((field, va))
            else:
                conflicts.append((field, va, vb))
        elif vals_a:
            a_only.append((field, vals_a[0]))
        elif vals_b:
            b_only.append((field, vals_b[0]))

    # 串项目风险评估：项目名称/编号/招标单位不同 = 可能串项目
    risk_fields = ['项目名称', '项目编号', '招标单位']
    risk_signals = []
    for field in risk_fields:
        for f, va, vb in conflicts:
            if field == f:
                risk_signals.append(f'{field}: 「{va}」≠「{vb}」')

    cross_project_risk = None
    if risk_signals:
        cross_project_risk = '⚠️ 疑似串项目：' + '；'.join(risk_signals)

    return {
        'matches': matches,
        'conflicts': conflicts,
        'a_only': a_only,
        'b_only': b_only,
        'cross_project_risk': cross_project_risk,
    }


# ============================================================
#  多维度深度比对
# ============================================================

def deep_compare(text_a, text_b):
    """多维度深度比对

    三维度:
    1. 文本相似度（SimHash段落级）
    2. 目录结构相似度
    3. 元数据比对（含串项目检测）

    返回: {
        'text_similarity': float,
        'structure_similarity': float,
        'structure_details': str,
        'metadata': {...},
        'matches': [...],  # 重复段落
    }
    """
    paras_a = split_paragraphs(text_a)
    paras_b = split_paragraphs(text_b)

    # 1. 文本相似度
    text_sim = overall_similarity(paras_a, paras_b)
    matches = compare_paragraphs(paras_a, paras_b)

    # 2. 目录结构
    headings_a = extract_headings(text_a)
    headings_b = extract_headings(text_b)
    struct_result = compare_structure(headings_a, headings_b)

    # 3. 元数据
    meta_a = extract_metadata(text_a)
    meta_b = extract_metadata(text_b)
    meta_result = compare_metadata(meta_a, meta_b)

    return {
        'text_similarity': text_sim,
        'match_count': len(matches),
        'matches': matches,
        'structure_similarity': struct_result['similarity'],
        'structure_details': struct_result['details'],
        'headings_a': len(headings_a),
        'headings_b': len(headings_b),
        'metadata_a': meta_a,
        'metadata_b': meta_b,
        'metadata_comparison': meta_result,
    }


# ============================================================
#  输出格式化
# ============================================================

def truncate_text(text, max_len=50):
    """截断文本，超出部分用省略号"""
    text = text.replace('\n', ' ').strip()
    if len(text) > max_len:
        return text[:max_len] + '...'
    return text


def print_compare_report(file_a_name, file_b_name, overall_sim, matches, threshold):
    """打印compare模式报告"""
    print()
    print('📊 标书查重报告')
    print('=' * 50)
    print(f'文件A: {file_a_name}')
    print(f'文件B: {file_b_name}')
    print(f'阈值: {threshold * 100:.0f}%')
    print('-' * 50)
    print(f'整体相似度: {overall_sim * 100:.1f}%')
    print(f'重复段落: {len(matches)}处')

    if matches:
        print()
        for i, m in enumerate(matches, 1):
            print(f'  {i}. 文件A第{m["idx_a"]}段 ↔ 文件B第{m["idx_b"]}段 '
                  f'(相似度{m["similarity"] * 100:.0f}%)')
            print(f'     「{truncate_text(m["text_a"], 60)}」')

    print()
    if overall_sim >= threshold:
        print('⚠️  相似度超过阈值，请注意查重风险！')
    else:
        print('✅ 相似度在安全范围内。')
    print()


def print_check_report(file_name, results, threshold):
    """打印check模式报告"""
    print()
    print('📊 标书库查重报告')
    print('=' * 50)
    print(f'目标文件: {file_name}')
    print(f'库文件数: {len(results)}')
    print(f'阈值: {threshold * 100:.0f}%')
    print('-' * 50)

    if not results:
        print('库中无文件可比对。')
        print()
        return

    # 按相似度降序
    results.sort(key=lambda x: x['overall_similarity'], reverse=True)

    for i, r in enumerate(results, 1):
        sim = r['overall_similarity']
        mark = '⚠️' if sim >= threshold else '  '
        print(f'{mark} {i}. {r["library_file"]} — 相似度 {sim * 100:.1f}%')
        if r['matches']:
            print(f'     重复段落: {len(r["matches"])}处')
            # 只显示前3处
            for m in r['matches'][:3]:
                print(f'     · 新标书第{m["idx_a"]}段 ↔ {r["library_file"]}第{m["idx_b"]}段 '
                      f'(相似度{m["similarity"] * 100:.0f}%)')
                print(f'       「{truncate_text(m["text_a"], 50)}」')
            if len(r['matches']) > 3:
                print(f'     ... 还有{len(r["matches"]) - 3}处')
        print()

    # 最高相似度警告
    if results and results[0]['overall_similarity'] >= threshold:
        print(f'🚨 最高相似度 {results[0]["overall_similarity"] * 100:.1f}%，'
              f'与「{results[0]["library_file"]}」高度重复！')
    else:
        print('✅ 未发现高相似度匹配。')
    print()


# ============================================================
#  命令处理
# ============================================================

def cmd_compare(args):
    """compare子命令：比较两个文件"""
    text_a = read_file(args.file_a)
    text_b = read_file(args.file_b)

    paras_a = split_paragraphs(text_a)
    paras_b = split_paragraphs(text_b)

    overall_sim = overall_similarity(paras_a, paras_b)
    matches = compare_paragraphs(paras_a, paras_b, args.threshold)

    # 终端输出（--json模式下不打印文本报告）
    if not args.json:
        print_compare_report(
            Path(args.file_a).name,
            Path(args.file_b).name,
            overall_sim,
            matches,
            args.threshold,
        )

    # JSON输出
    if args.json:
        result = {
            'mode': 'compare',
            'file_a': args.file_a,
            'file_b': args.file_b,
            'threshold': args.threshold,
            'overall_similarity': round(overall_sim, 4),
            'match_count': len(matches),
            'matches': [
                {
                    'idx_a': m['idx_a'],
                    'idx_b': m['idx_b'],
                    'similarity': round(m['similarity'], 4),
                    'text_a': m['text_a'],
                    'text_b': m['text_b'],
                }
                for m in matches
            ],
        }
        print(json.dumps(result, ensure_ascii=False, indent=2))


def cmd_check(args):
    """check子命令：和库目录下所有文件比对"""
    text_new = read_file(args.file)
    paras_new = split_paragraphs(text_new)

    lib_dir = Path(args.library)
    if not lib_dir.exists():
        print(f"❌ 库目录不存在：{args.library}", file=sys.stderr)
        sys.exit(1)

    # 收集库中所有支持格式的文件
    lib_files = []
    for ext in ('*.md', '*.txt', '*.docx'):
        lib_files.extend(lib_dir.glob(ext))

    if not lib_files:
        print(f"⚠️  库目录 {args.library} 中没有可比对的文件", file=sys.stderr)
        if args.json:
            print(json.dumps({
                'mode': 'check',
                'file': args.file,
                'library': args.library,
                'results': [],
            }, ensure_ascii=False, indent=2))
        return

    results = []
    for lib_file in lib_files:
        text_lib = read_file(str(lib_file))
        paras_lib = split_paragraphs(text_lib)

        overall_sim = overall_similarity(paras_new, paras_lib)
        matches = compare_paragraphs(paras_new, paras_lib, args.threshold)

        results.append({
            'library_file': lib_file.name,
            'library_path': str(lib_file),
            'overall_similarity': overall_sim,
            'match_count': len(matches),
            'matches': matches,
        })

    # 终端输出（--json模式下不打印文本报告）
    if not args.json:
        print_check_report(Path(args.file).name, results, args.threshold)

    # JSON输出
    if args.json:
        json_result = {
            'mode': 'check',
            'file': args.file,
            'library': args.library,
            'threshold': args.threshold,
            'results': [
                {
                    'library_file': r['library_file'],
                    'library_path': r['library_path'],
                    'overall_similarity': round(r['overall_similarity'], 4),
                    'match_count': r['match_count'],
                    'matches': [
                        {
                            'idx_a': m['idx_a'],
                            'idx_b': m['idx_b'],
                            'similarity': round(m['similarity'], 4),
                            'text_a': m['text_a'],
                            'text_b': m['text_b'],
                        }
                        for m in r['matches']
                    ],
                }
                for r in results
            ],
        }
        print(json.dumps(json_result, ensure_ascii=False, indent=2))


def cmd_add(args):
    """add子命令：把标书加入历史库"""
    lib_dir = Path(args.library)
    lib_dir.mkdir(parents=True, exist_ok=True)

    src_path = Path(args.file)
    if not src_path.exists():
        print(f"❌ 文件不存在：{args.file}", file=sys.stderr)
        sys.exit(1)

    # 确定目标文件名
    name = args.name if args.name else src_path.stem
    # 文件名安全处理
    safe_name = re.sub(r'[^\w\u4e00-\u9fff\-]', '_', name)
    dest_name = f"{safe_name}{src_path.suffix}"
    dest_path = lib_dir / dest_name

    # 如果同名文件已存在，加序号
    counter = 1
    while dest_path.exists():
        dest_name = f"{safe_name}_{counter}{src_path.suffix}"
        dest_path = lib_dir / dest_name
        counter += 1

    # 复制文件
    import shutil
    shutil.copy2(str(src_path), str(dest_path))

    print(f"✅ 已添加到历史库：{dest_path}")
    print(f"   名称：{name}")
    print(f"   格式：{src_path.suffix}")


def cmd_deep(args):
    """deep子命令：多维度深度比对（文本+目录结构+元数据）"""
    text_a = read_file(args.file_a)
    text_b = read_file(args.file_b)

    result = deep_compare(text_a, text_b)

    if args.json:
        # JSON模式
        json_result = {
            'mode': 'deep',
            'file_a': args.file_a,
            'file_b': args.file_b,
            'text_similarity': round(result['text_similarity'], 4),
            'match_count': result['match_count'],
            'structure_similarity': round(result['structure_similarity'], 4),
            'structure_details': result['structure_details'],
            'headings_a': result['headings_a'],
            'headings_b': result['headings_b'],
            'metadata_a': result['metadata_a'],
            'metadata_b': result['metadata_b'],
            'metadata_comparison': {
                'matches': [(f, v) for f, v in result['metadata_comparison']['matches']],
                'conflicts': [(f, va, vb) for f, va, vb in result['metadata_comparison']['conflicts']],
                'a_only': [(f, v) for f, v in result['metadata_comparison']['a_only']],
                'b_only': [(f, v) for f, v in result['metadata_comparison']['b_only']],
                'cross_project_risk': result['metadata_comparison']['cross_project_risk'],
            },
            'matches': [
                {
                    'idx_a': m['idx_a'],
                    'idx_b': m['idx_b'],
                    'similarity': round(m['similarity'], 4),
                    'text_a': m['text_a'],
                    'text_b': m['text_b'],
                }
                for m in result['matches'][:20]  # 最多输出20条
            ],
        }
        print(json.dumps(json_result, ensure_ascii=False, indent=2))
        return

    # 终端报告
    name_a = Path(args.file_a).name
    name_b = Path(args.file_b).name

    print(f'\n{"="*60}')
    print(f'  多维度深度比对报告')
    print(f'  {name_a}  vs  {name_b}')
    print(f'{"="*60}\n')

    # --- 维度1: 文本相似度 ---
    ts = result['text_similarity']
    bar = '█' * int(ts * 20) + '░' * (20 - int(ts * 20))
    print(f'📊 维度1: 文本相似度')
    print(f'   [{bar}] {ts:.1%}')
    print(f'   重复段落: {result["match_count"]} 段')

    if result['match_count'] > 0:
        print(f'\n   重复段落详情（前5条）:')
        for i, m in enumerate(result['matches'][:5]):
            print(f'   {i+1}. 段落A#{m["idx_a"]} ↔ 段落B#{m["idx_b"]} '
                  f'(相似{m["similarity"]:.1%})')
            print(f'      A: {truncate_text(m["text_a"], 60)}')
            print(f'      B: {truncate_text(m["text_b"], 60)}')

    # --- 维度2: 目录结构 ---
    ss = result['structure_similarity']
    bar2 = '█' * int(ss * 20) + '░' * (20 - int(ss * 20))
    print(f'\n📊 维度2: 目录结构相似度')
    print(f'   [{bar2}] {ss:.1%}')
    print(f'   {result["structure_details"]}')

    # --- 维度3: 元数据 ---
    mc = result['metadata_comparison']
    print(f'\n📊 维度3: 元数据比对')

    if mc['matches']:
        print(f'   ✅ 一致字段:')
        for f, v in mc['matches']:
            print(f'      {f}: {v}')

    if mc['conflicts']:
        print(f'   ⚠️ 冲突字段:')
        for f, va, vb in mc['conflicts']:
            print(f'      {f}: A=「{va}」 vs B=「{vb}」')

    if mc['a_only']:
        print(f'   📄 仅A有:')
        for f, v in mc['a_only']:
            print(f'      {f}: {v}')

    if mc['b_only']:
        print(f'   📄 仅B有:')
        for f, v in mc['b_only']:
            print(f'      {f}: {v}')

    if not mc['matches'] and not mc['conflicts'] and not mc['a_only'] and not mc['b_only']:
        print(f'   （未提取到元数据）')

    # --- 串项目风险 ---
    if mc['cross_project_risk']:
        print(f'\n🚨 {mc["cross_project_risk"]}')

    # --- 综合评估 ---
    combined = ts * 0.5 + ss * 0.3 + (0.2 if mc['matches'] else 0)
    combined = min(combined, 1.0)
    print(f'\n{"─"*60}')
    print(f'  综合相似度: {combined:.1%}')
    if combined >= 0.8:
        print(f'  ⛔ 高风险: 疑似换皮/复制标书，请人工核查')
    elif combined >= 0.5:
        print(f'  ⚠️  中风险: 存在较多重叠，建议核查重复段落')
    else:
        print(f'  ✅ 低风险: 相似度较低')
    print(f'{"─"*60}\n')


# ============================================================
#  CLI入口
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description='标书查重工具 - 基于SimHash的文本相似度检测',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例：
  # 比较两个文件
  python bid_similarity.py compare 标书A.md 标书B.md

  # 检查新标书和历史库的相似度
  python bid_similarity.py check 新标书.md --library ./bid_library/

  # 添加标书到历史库
  python bid_similarity.py add 旧标书.md --library ./bid_library/ --name "2024市政项目"

  # 输出JSON格式
  python bid_similarity.py compare A.md B.md --json > report.json

  # 多维度深度比对（文本+目录结构+元数据+串项目检测）
  python bid_similarity.py deep 标书A.md 标书B.md
  python bid_similarity.py deep 标书A.md 标书B.md --json > deep_report.json
        """,
    )

    subparsers = parser.add_subparsers(dest='command', help='子命令')

    # compare 子命令
    p_compare = subparsers.add_parser('compare', help='比较两个文件的相似度')
    p_compare.add_argument('file_a', help='文件A路径')
    p_compare.add_argument('file_b', help='文件B路径')
    p_compare.add_argument('--threshold', type=float, default=0.8,
                           help='相似度阈值，默认0.8（80%%）')
    p_compare.add_argument('--json', action='store_true', help='输出JSON格式')

    # check 子命令
    p_check = subparsers.add_parser('check', help='和库目录下所有文件比对')
    p_check.add_argument('file', help='待检查的标书文件')
    p_check.add_argument('--library', default='./bid_library/',
                         help='历史库目录，默认 ./bid_library/')
    p_check.add_argument('--threshold', type=float, default=0.8,
                         help='相似度阈值，默认0.8（80%%）')
    p_check.add_argument('--json', action='store_true', help='输出JSON格式')

    # add 子命令
    p_add = subparsers.add_parser('add', help='添加标书到历史库')
    p_add.add_argument('file', help='标书文件路径')
    p_add.add_argument('--library', default='./bid_library/',
                       help='历史库目录，默认 ./bid_library/')
    p_add.add_argument('--name', default=None, help='库中显示名称（默认用文件名）')

    # deep 子命令
    p_deep = subparsers.add_parser('deep', help='多维度深度比对（文本+目录结构+元数据）')
    p_deep.add_argument('file_a', help='文件A路径')
    p_deep.add_argument('file_b', help='文件B路径')
    p_deep.add_argument('--json', action='store_true', help='输出JSON格式')

    args = parser.parse_args()

    if args.command is None:
        parser.print_help()
        sys.exit(1)

    # 显示分词模式提示
    if _jieba is not None:
        pass  # 静默使用jieba
    else:
        # 首次运行时提示
        pass

    if args.command == 'compare':
        cmd_compare(args)
    elif args.command == 'check':
        cmd_check(args)
    elif args.command == 'add':
        cmd_add(args)
    elif args.command == 'deep':
        cmd_deep(args)


if __name__ == '__main__':
    main()
