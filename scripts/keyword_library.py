#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标书判词库管理器 v1.0
管理禁用词/敏感词/规范词，支持积累、审批、扫描、导出。

三种词类：
  - forbidden_words: 禁用词（出现即FAIL）
  - sensitive_words: 敏感词（出现WARN，给替代建议）
  - compliant_words: 规范词（鼓励使用的合规表述）

用法：
  python keyword_library.py list [--type forbidden] [--status approved]
  python keyword_library.py add --word "绝对没问题" --type forbidden --category "夸大承诺" [--note "备注"]
  python keyword_library.py remove --word "绝对没问题" --type forbidden
  python keyword_library.py suggest --word "某新词" --type forbidden  # AI发现→pending
  python keyword_library.py approve --word "某新词" --type forbidden  # pending→approved
  python keyword_library.py scan 文件.md [--json]
  python keyword_library.py export --output forbidden_words.txt
  python keyword_library.py import --input config.yaml
"""

import os, sys, json, argparse
from pathlib import Path
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')

# 判词库默认路径
DEFAULT_LIBRARY = Path(__file__).parent.parent / 'keyword_library.json'

# ===== 库文件读写 =====
def load_library(lib_path=None):
    """加载判词库JSON"""
    path = Path(lib_path) if lib_path else DEFAULT_LIBRARY
    if path.exists():
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    # 空库骨架
    return {
        'forbidden_words': [],
        'sensitive_words': [],
        'compliant_words': [],
    }

def save_library(library, lib_path=None):
    """保存判词库JSON"""
    path = Path(lib_path) if lib_path else DEFAULT_LIBRARY
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(library, f, ensure_ascii=False, indent=2)
    return path

# ===== 词操作 =====
WORD_TYPES = ['forbidden', 'sensitive', 'compliant']

def _get_list(library, word_type):
    """根据类型获取词列表"""
    key = f'{word_type}_words'
    if key not in library:
        library[key] = []
    return library[key]

def add_word(library, word, word_type, category='', note='', added_by='user', status='approved'):
    """添加一个词到判词库"""
    word_list = _get_list(library, word_type)
    # 去重
    for entry in word_list:
        if entry['word'] == word:
            print(f'⚠️  词「{word}」已存在于{word_type}（状态：{entry.get("status", "approved")}）')
            return False
    entry = {
        'word': word,
        'category': category or '未分类',
        'added_by': added_by,
        'added_date': datetime.now().strftime('%Y-%m-%d'),
        'status': status,
        'note': note,
    }
    if word_type == 'sensitive':
        entry['suggestion'] = ''
    word_list.append(entry)
    return True

def remove_word(library, word, word_type):
    """从判词库删除一个词"""
    word_list = _get_list(library, word_type)
    for i, entry in enumerate(word_list):
        if entry['word'] == word:
            word_list.pop(i)
            return True
    return False

def find_word(library, word, word_type=None):
    """查找词，返回(type, entry)或None"""
    types = [word_type] if word_type else WORD_TYPES
    for t in types:
        for entry in _get_list(library, t):
            if entry['word'] == word:
                return t, entry
    return None

def suggest_word(library, word, word_type, category='', note=''):
    """AI发现新词→pending状态，待人工审批"""
    return add_word(library, word, word_type, category, note, added_by='AI', status='pending')

def approve_word(library, word, word_type):
    """审批pending词→approved"""
    word_list = _get_list(library, word_type)
    for entry in word_list:
        if entry['word'] == word and entry.get('status') == 'pending':
            entry['status'] = 'approved'
            entry['approved_date'] = datetime.now().strftime('%Y-%m-%d')
            return True
    return False

# ===== 文件扫描 =====
def scan_file(file_path, library, status_filter='approved'):
    """扫描文件中的禁用词和敏感词
    
    status_filter: 'approved'=只扫已审批词, 'all'=扫所有词
    """
    path = Path(file_path)
    if not path.exists():
        print(f'❌ 文件不存在: {file_path}')
        sys.exit(1)
    
    # 读取文件
    if path.suffix == '.docx':
        try:
            from docx import Document
            doc = Document(str(path))
            text = '\n'.join(p.text for p in doc.paragraphs)
        except ImportError:
            print('❌ 需要python-docx: pip install python-docx')
            sys.exit(1)
    else:
        with open(path, 'r', encoding='utf-8') as f:
            text = f.read()
    
    results = {'forbidden': [], 'sensitive': [], 'stats': {'total_chars': len(text)}}
    
    # 扫描禁用词
    for entry in _get_list(library, 'forbidden'):
        if status_filter == 'approved' and entry.get('status') != 'approved':
            continue
        if entry['word'] in text:
            count = text.count(entry['word'])
            results['forbidden'].append({
                'word': entry['word'],
                'category': entry.get('category', ''),
                'count': count,
                'note': entry.get('note', ''),
            })
    
    # 扫描敏感词
    for entry in _get_list(library, 'sensitive'):
        if status_filter == 'approved' and entry.get('status') != 'approved':
            continue
        if entry['word'] in text:
            count = text.count(entry['word'])
            results['sensitive'].append({
                'word': entry['word'],
                'category': entry.get('category', ''),
                'count': count,
                'suggestion': entry.get('suggestion', ''),
            })
    
    return results

# ===== 导入导出 =====
def export_words(library, output_path, word_type='forbidden', status_filter='approved'):
    """导出纯词表（每行一个词），供bid_engine.py的config.yaml使用"""
    word_list = _get_list(library, word_type)
    words = [e['word'] for e in word_list if status_filter == 'all' or e.get('status') == status_filter]
    with open(output_path, 'w', encoding='utf-8') as f:
        for w in words:
            f.write(w + '\n')
    return len(words)

def import_from_config(config_path, library):
    """从config.yaml导入现有forbidden_words列表"""
    try:
        import yaml
    except ImportError:
        print('❌ 需要pyyaml: pip install pyyaml')
        sys.exit(1)
    
    with open(config_path, 'r', encoding='utf-8') as f:
        cfg = yaml.safe_load(f)
    
    quality = cfg.get('quality', {})
    forbidden = quality.get('forbidden_words', [])
    
    imported = 0
    for word in forbidden:
        if add_word(library, word, 'forbidden', category='从config导入', added_by='system'):
            imported += 1
    return imported

# ===== 打印 =====
def print_list(library, word_type=None, status_filter=None):
    """打印判词库列表"""
    types = [word_type] if word_type else WORD_TYPES
    for t in types:
        word_list = _get_list(library, t)
        if status_filter:
            word_list = [e for e in word_list if e.get('status') == status_filter]
        
        type_names = {'forbidden': '🚫 禁用词', 'sensitive': '⚠️  敏感词', 'compliant': '✅ 规范词'}
        print(f'\n{type_names.get(t, t)} ({len(word_list)}条)')
        print('-' * 50)
        if not word_list:
            print('  （空）')
            continue
        for entry in word_list:
            status_icon = '🟢' if entry.get('status') == 'approved' else '🟡'
            line = f'  {status_icon} [{entry.get("category", "未分类")}] {entry["word"]}'
            if entry.get('note'):
                line += f'  # {entry["note"]}'
            if entry.get('suggestion'):
                line += f'  → 建议替换: {entry["suggestion"]}'
            print(line)

def print_scan_report(results):
    """打印扫描报告"""
    print('\n' + '=' * 60)
    print('🔍 判词库扫描报告')
    print('=' * 60)
    print(f'📊 文档字数: {results["stats"]["total_chars"]}')
    
    if results['forbidden']:
        print(f'\n🚫 禁用词: {len(results["forbidden"])}种')
        for item in results['forbidden']:
            print(f'  ❌ "{item["word"]}" x{item["count"]}次  [{item["category"]}]')
            if item['note']:
                print(f'      备注: {item["note"]}')
    else:
        print('\n✅ 未发现禁用词')
    
    if results['sensitive']:
        print(f'\n⚠️  敏感词: {len(results["sensitive"])}种')
        for item in results['sensitive']:
            line = f'  ⚠️  "{item["word"]}" x{item["count"]}次  [{item["category"]}]'
            if item['suggestion']:
                line += f'  → 建议替换: {item["suggestion"]}'
            print(line)
    else:
        print('\n✅ 未发现敏感词')
    
    total_issues = len(results['forbidden']) + len(results['sensitive'])
    status = '✅ PASS' if total_issues == 0 else '❌ FAIL'
    print(f'\n🏁 结论: {status}')
    print('=' * 60)

# ===== 主入口 =====
def main():
    parser = argparse.ArgumentParser(description='标书判词库管理器 v1.0')
    parser.add_argument('--library', default=None, help='判词库JSON路径（默认keyword_library.json）')
    sub = parser.add_subparsers(dest='command')
    
    # list
    p_list = sub.add_parser('list', help='列出判词库内容')
    p_list.add_argument('--type', choices=WORD_TYPES, help='只看某类词')
    p_list.add_argument('--status', choices=['approved', 'pending'], help='只看某状态')
    
    # add
    p_add = sub.add_parser('add', help='添加词')
    p_add.add_argument('--word', required=True)
    p_add.add_argument('--type', required=True, choices=WORD_TYPES)
    p_add.add_argument('--category', default='')
    p_add.add_argument('--note', default='')
    p_add.add_argument('--suggestion', default='', help='敏感词替代建议')
    
    # remove
    p_rm = sub.add_parser('remove', help='删除词')
    p_rm.add_argument('--word', required=True)
    p_rm.add_argument('--type', required=True, choices=WORD_TYPES)
    
    # suggest (AI发现新词)
    p_sug = sub.add_parser('suggest', help='AI发现新词→pending待审')
    p_sug.add_argument('--word', required=True)
    p_sug.add_argument('--type', required=True, choices=WORD_TYPES)
    p_sug.add_argument('--category', default='')
    p_sug.add_argument('--note', default='')
    
    # approve
    p_app = sub.add_parser('approve', help='审批pending词→approved')
    p_app.add_argument('--word', required=True)
    p_app.add_argument('--type', required=True, choices=WORD_TYPES)
    
    # scan
    p_scan = sub.add_parser('scan', help='扫描文件')
    p_scan.add_argument('file', help='待扫描文件(md/txt/docx)')
    p_scan.add_argument('--json', action='store_true')
    p_scan.add_argument('--all-status', action='store_true', help='扫描所有状态(含pending)')
    
    # export
    p_exp = sub.add_parser('export', help='导出纯词表')
    p_exp.add_argument('--output', required=True)
    p_exp.add_argument('--type', default='forbidden', choices=WORD_TYPES)
    
    # import
    p_imp = sub.add_parser('import', help='从config.yaml导入')
    p_imp.add_argument('--input', required=True)
    
    args = parser.parse_args()
    lib_path = args.library or DEFAULT_LIBRARY
    library = load_library(lib_path)
    
    if args.command == 'list':
        print_list(library, args.type, args.status)
    
    elif args.command == 'add':
        if add_word(library, args.word, args.type, args.category, args.note):
            # 敏感词加替代建议
            if args.type == 'sensitive' and args.suggestion:
                for e in _get_list(library, 'sensitive'):
                    if e['word'] == args.word:
                        e['suggestion'] = args.suggestion
            save_library(library, lib_path)
            print(f'✅ 已添加: [{args.type}] {args.word}')
        # add_word内部会打印重复提示
    
    elif args.command == 'remove':
        if remove_word(library, args.word, args.type):
            save_library(library, lib_path)
            print(f'✅ 已删除: [{args.type}] {args.word}')
        else:
            print(f'❌ 未找到: [{args.type}] {args.word}')
    
    elif args.command == 'suggest':
        if suggest_word(library, args.word, args.type, args.category, args.note):
            save_library(library, lib_path)
            print(f'🟡 已加入待审(pending): [{args.type}] {args.word}')
            print('   用 approve 命令审批通过后生效')
    
    elif args.command == 'approve':
        if approve_word(library, args.word, args.type):
            save_library(library, lib_path)
            print(f'✅ 已审批通过: [{args.type}] {args.word}')
        else:
            print(f'❌ 未找到pending状态的: [{args.type}] {args.word}')
    
    elif args.command == 'scan':
        status_filter = 'all' if args.all_status else 'approved'
        results = scan_file(args.file, library, status_filter)
        if args.json:
            print(json.dumps(results, ensure_ascii=False, indent=2))
        else:
            print_scan_report(results)
    
    elif args.command == 'export':
        count = export_words(library, args.output, args.type)
        print(f'✅ 已导出{count}个{args.type}词到 {args.output}')
    
    elif args.command == 'import':
        imported = import_from_config(args.input, library)
        save_library(library, lib_path)
        print(f'✅ 从 {args.input} 导入 {imported} 个禁用词')
    
    else:
        parser.print_help()

if __name__ == '__main__':
    main()
