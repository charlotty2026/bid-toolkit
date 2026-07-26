#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
企业知识库全文检索 v1.0
========================
基于 BM25 算法的企业知识库（标书库）全文检索，支持段落级定位。

功能：
  - search: 关键词搜索，BM25 排序，返回最相关的段落片段
  - index:  建立索引并查看统计信息
  - info:   查看知识库概览

用法：
  python bid_search.py search "物业管理方案" --library ./samples/
  python bid_search.py search "保安 消防" --library ./samples/ --top 5 --json
  python bid_search.py index --library ./samples/
  python bid_search.py info --library ./samples/

技术说明：
  - BM25: 经典文本相关性排序算法（k1=1.5, b=0.75）
  - 分词：优先使用 jieba，未安装则退化为 2-gram
  - 索引粒度：段落级（按空行/标题分块），搜索结果直接定位到相关段落
  - 支持 .md / .txt / .docx 输入（docx 需 python-docx）
"""

import os
import sys
import re
import json
import math
import argparse
from pathlib import Path
from collections import defaultdict, Counter

sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
#  依赖检测
# ============================================================

_jieba = None
try:
    import jieba
    _jieba = jieba
    # 抑制 jieba 初始化日志
    jieba.setLogLevel(20)
except ImportError:
    pass

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
        return None

    suffix = path.suffix.lower()

    if suffix in ('.md', '.txt'):
        with open(path, 'r', encoding='utf-8') as f:
            return f.read()

    if suffix == '.docx':
        if not _docx_available:
            return None
        doc = _DocxDocument(str(path))
        paragraphs = []
        for para in doc.paragraphs:
            paragraphs.append(para.text)
        return '\n'.join(paragraphs)

    return None


def split_paragraphs(text):
    """将文本按段落分块（空行分隔 + 标题行独立成段）"""
    paragraphs = []
    current = []

    for line in text.split('\n'):
        stripped = line.strip()
        # 空行 = 段落分隔
        if not stripped:
            if current:
                paragraphs.append('\n'.join(current))
                current = []
        # Markdown 标题行独立成段
        elif re.match(r'^#{1,6}\s', stripped):
            if current:
                paragraphs.append('\n'.join(current))
                current = []
            paragraphs.append(stripped)
        else:
            current.append(line)

    if current:
        paragraphs.append('\n'.join(current))

    # 过滤太短的段落（<10字符的可能是页眉页脚），但保留有内容的标题行
    result = []
    for p in paragraphs:
        stripped = p.strip()
        # 标题行：去掉 # 号后 >= 2 字符就保留
        if re.match(r'^#{1,6}\s', stripped):
            title_text = re.sub(r'^#{1,6}\s+', '', stripped)
            if len(title_text) >= 2:
                result.append(p)
        elif len(stripped) >= 10:
            result.append(p)
    return result


# ============================================================
#  分词
# ============================================================

def tokenize(text):
    """分词：jieba 优先，未安装退化为 2-gram"""
    # 去标点+空白
    text = re.sub(r'[^\w\u4e00-\u9fff]', ' ', text)
    text = text.strip()

    if _jieba:
        words = [w for w in _jieba.cut(text) if w.strip() and len(w) >= 1]
        # 过滤单字符英文/数字噪声
        return [w for w in words if not (len(w) == 1 and w.isascii())]
    else:
        # 退化：中文 2-gram + 英文单词
        tokens = []
        i = 0
        while i < len(text):
            if '\u4e00' <= text[i] <= '\u9fff':
                if i + 1 < len(text) and '\u4e00' <= text[i + 1] <= '\u9fff':
                    tokens.append(text[i:i + 2])
                i += 1
            elif text[i].isalpha():
                j = i
                while j < len(text) and text[j].isalpha():
                    j += 1
                if j - i >= 2:
                    tokens.append(text[i:j])
                i = j
            else:
                i += 1
        return tokens


# ============================================================
#  BM25 索引
# ============================================================

class BM25Index:
    """BM25 全文检索索引（段落级）"""

    def __init__(self, k1=1.5, b=0.75):
        self.k1 = k1
        self.b = b
        self.docs = []          # [{id, file, para, tokens, len}]
        self.df = defaultdict(int)  # 词 -> 出现在多少段落
        self.avgdl = 0          # 平均段落长度
        self.N = 0              # 段落总数

    def build(self, library_path):
        """从知识库目录构建索引"""
        lib = Path(library_path)
        if not lib.exists():
            print(f"❌ 知识库目录不存在：{library_path}", file=sys.stderr)
            return False

        files = sorted(
            f for f in lib.iterdir()
            if f.is_file() and f.suffix.lower() in ('.md', '.txt', '.docx')
        )

        if not files:
            print(f"⚠️ 知识库目录中没有可索引的文件：{library_path}", file=sys.stderr)
            return False

        self.docs = []
        self.df = defaultdict(int)
        total_len = 0

        for fpath in files:
            text = read_file(fpath)
            if not text:
                continue

            paragraphs = split_paragraphs(text)
            for idx, para in enumerate(paragraphs):
                tokens = tokenize(para)
                if not tokens:
                    continue

                doc_id = len(self.docs)
                self.docs.append({
                    'id': doc_id,
                    'file': fpath.name,
                    'para_idx': idx,
                    'text': para,
                    'tokens': tokens,
                    'len': len(tokens),
                })

                total_len += len(tokens)
                # 统计 df（每个词在当前段落只计一次）
                for word in set(tokens):
                    self.df[word] += 1

        self.N = len(self.docs)
        self.avgdl = total_len / self.N if self.N > 0 else 0
        return True

    def _idf(self, word):
        """计算 IDF"""
        df = self.df.get(word, 0)
        if df == 0:
            return 0
        return math.log((self.N - df + 0.5) / (df + 0.5) + 1)

    def search(self, query, top_k=10):
        """搜索，返回 top_k 个最相关的段落"""
        query_tokens = tokenize(query)
        if not query_tokens:
            return []

        # 计算每个段落的 BM25 得分
        scores = []
        for doc in self.docs:
            score = 0.0
            tf_counter = Counter(doc['tokens'])
            dl = doc['len']

            for word in query_tokens:
                tf = tf_counter.get(word, 0)
                if tf == 0:
                    continue
                idf = self._idf(word)
                # BM25 公式
                numerator = tf * (self.k1 + 1)
                denominator = tf + self.k1 * (1 - self.b + self.b * dl / self.avgdl)
                score += idf * numerator / denominator

            if score > 0:
                scores.append((score, doc))

        # 按得分排序
        scores.sort(key=lambda x: x[0], reverse=True)

        # 构造结果
        results = []
        for rank, (score, doc) in enumerate(scores[:top_k], 1):
            # 高亮匹配词
            snippet = self._highlight(doc['text'], query_tokens)
            results.append({
                'rank': rank,
                'score': round(score, 4),
                'file': doc['file'],
                'para_idx': doc['para_idx'],
                'snippet': snippet,
                'matched_tokens': list(set(
                    w for w in query_tokens if w in doc['tokens']
                )),
            })
        return results

    def _highlight(self, text, query_tokens, max_len=200):
        """截取片段并高亮匹配词"""
        # 找第一个匹配位置
        first_pos = -1
        for token in query_tokens:
            pos = text.find(token)
            if pos >= 0 and (first_pos < 0 or pos < first_pos):
                first_pos = pos

        if first_pos < 0:
            snippet = text[:max_len]
        else:
            start = max(0, first_pos - max_len // 3)
            end = min(len(text), start + max_len)
            snippet = text[start:end]
            if start > 0:
                snippet = '...' + snippet
            if end < len(text):
                snippet = snippet + '...'

        # 高亮（用【】包裹匹配词）
        for token in sorted(query_tokens, key=len, reverse=True):
            snippet = snippet.replace(token, f'【{token}】')

        return snippet.replace('\n', ' ')

    def info(self):
        """返回索引统计信息"""
        files = set(d['file'] for d in self.docs)
        return {
            'total_paragraphs': self.N,
            'total_files': len(files),
            'files': sorted(files),
            'avg_para_len': round(self.avgdl, 1),
            'vocab_size': len(self.df),
            'jieba_enabled': _jieba is not None,
        }


# ============================================================
#  CLI
# ============================================================

def cmd_search(args):
    """搜索命令"""
    index = BM25Index()
    if not index.build(args.library):
        sys.exit(1)

    if not index.N:
        print("⚠️ 知识库为空，无法搜索", file=sys.stderr)
        sys.exit(1)

    results = index.search(args.query, top_k=args.top)

    if args.json:
        print(json.dumps(results, ensure_ascii=False, indent=2))
        return

    if not results:
        print(f"🔍 搜索「{args.query}」无匹配结果")
        return

    print(f"🔍 搜索「{args.query}」找到 {len(results)} 条结果（共 {index.N} 段落）\n")
    print(f"{'排名':>4} | {'分数':>7} | {'文件':<20} | {'片段'}")
    print('-' * 100)

    for r in results:
        print(f"  #{r['rank']:<2} | {r['score']:>7.4f} | {r['file']:<20} | {r['snippet']}")
        if r['matched_tokens']:
            print(f"       匹配词：{', '.join(r['matched_tokens'])}")
        print()


def cmd_index(args):
    """建立索引命令"""
    index = BM25Index()
    if not index.build(args.library):
        sys.exit(1)

    info = index.info()
    print(f"📚 知识库索引构建完成\n")
    print(f"  段落总数：{info['total_paragraphs']}")
    print(f"  文件数量：{info['total_files']}")
    print(f"  词汇量：  {info['vocab_size']}")
    print(f"  平均段落长度：{info['avg_para_len']} 词")
    print(f"  分词引擎：{'jieba' if info['jieba_enabled'] else '2-gram（未装jieba）'}")
    print(f"\n  文件列表：")
    for f in info['files']:
        print(f"    - {f}")


def cmd_info(args):
    """查看知识库概览"""
    lib = Path(args.library)
    if not lib.exists():
        print(f"❌ 知识库目录不存在：{args.library}", file=sys.stderr)
        sys.exit(1)

    files = sorted(
        f for f in lib.iterdir()
        if f.is_file() and f.suffix.lower() in ('.md', '.txt', '.docx')
    )

    if not files:
        print(f"⚠️ 知识库目录为空：{args.library}")
        return

    print(f"📚 知识库概览：{args.library}\n")
    print(f"{'文件名':<30} | {'大小':>8} | {'类型'}")
    print('-' * 60)

    total_size = 0
    for f in files:
        size = f.stat().st_size
        total_size += size
        size_str = f"{size / 1024:.1f}KB" if size >= 1024 else f"{size}B"
        print(f"{f.name:<30} | {size_str:>8} | {f.suffix[1:].upper()}")

    print(f"\n  共 {len(files)} 个文件，总大小 {total_size / 1024:.1f}KB")


def main():
    parser = argparse.ArgumentParser(
        description='企业知识库全文检索（BM25）',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest='command')

    # search
    p_search = sub.add_parser('search', help='关键词搜索')
    p_search.add_argument('query', help='搜索关键词')
    p_search.add_argument('--library', '-l', default='./samples/', help='知识库目录（默认 ./samples/）')
    p_search.add_argument('--top', '-t', type=int, default=10, help='返回结果数量（默认10）')
    p_search.add_argument('--json', action='store_true', help='JSON 格式输出')

    # index
    p_index = sub.add_parser('index', help='建立索引并查看统计')
    p_index.add_argument('--library', '-l', default='./samples/', help='知识库目录')

    # info
    p_info = sub.add_parser('info', help='查看知识库概览')
    p_info.add_argument('--library', '-l', default='./samples/', help='知识库目录')

    args = parser.parse_args()

    if args.command == 'search':
        cmd_search(args)
    elif args.command == 'index':
        cmd_index(args)
    elif args.command == 'info':
        cmd_info(args)
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
