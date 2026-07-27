#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标书自动化引擎 v3.2
一行命令完成：Markdown → 全角半角修复 → Word生成 → 格式自检

用法:
  python bid_engine.py 标书.md                    # 生成Word
  python bid_engine.py 标书.md -o 输出.docx       # 指定输出
  python bid_engine.py 标书.md --scan             # 仅扫描问题
  python bid_engine.py 标书.md --check            # 生成后自检
  python bid_engine.py 标书.md --暗标             # 暗标模式(去公司标识)
  python bid_engine.py 标书.md --template government  # 使用政府采购模板
  python bid_engine.py 标书.md --template enterprise  # 使用企业投标模板
  python bid_engine.py 标书.md --template engineering # 使用工程类模板
  python bid_engine.py 标书.md --config my.yaml  # 使用自定义配置
"""

import os, sys, re, argparse, json
from datetime import datetime
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

try:
    from docx import Document
    from docx.shared import Pt, Inches, Twips, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("❌ 需要 python-docx: pip install python-docx")
    sys.exit(1)

try:
    import yaml
except ImportError:
    print("⚠️  未安装pyyaml，将使用默认配置。安装命令: pip install pyyaml")
    yaml = None

# ===== 默认配置 =====
DEFAULT_CONFIG = {
    'body': {'font': '宋体', 'size': 12, 'bold': False, 'first_indent_chars': 2, 'line_spacing': 1.5, 'alignment': 'justify'},
    'headings': {
        1: {'font': '宋体', 'size': 16, 'bold': True, 'space_before': 12, 'space_after': 6},
        2: {'font': '宋体', 'size': 15, 'bold': True, 'space_before': 8, 'space_after': 4},
        3: {'font': '宋体', 'size': 14, 'bold': True, 'space_before': 6, 'space_after': 3},
        4: {'font': '宋体', 'size': 14, 'bold': True, 'space_before': 4, 'space_after': 2},
    },
    'table': {'font': '宋体', 'size': 10.5, 'bold_header': True, 'style': 'Table Grid'},
    'page': {'margin_top': 2.54, 'margin_bottom': 2.54, 'margin_left': 2.00, 'margin_right': 2.00},
    'anonymous': {'enabled': False, 'replace_words': ['我公司', '本公司'], 'replace_with': '投标人', 'company_names': [], 'company_addresses': []},
    'quality': {'check_placeholder': True, 'check_punctuation': True, 'check_forbidden_words': True, 'forbidden_words': ['保证中标', '100%成功率', '最优价格', '独家技术', '唯一选择', '最高质量']},
}

def load_config(config_path=None, template_name=None):
    """加载配置文件，优先级：config_path > template > 自动查找config.yaml > 默认"""
    import copy
    config = copy.deepcopy(DEFAULT_CONFIG)

    # 1) 加载模板
    if template_name and yaml:
        tpl_path = Path(__file__).parent.parent / 'templates' / f'{template_name}.yaml'
        if tpl_path.exists():
            with open(tpl_path, 'r', encoding='utf-8') as f:
                tpl = yaml.safe_load(f)
                if tpl:
                    _deep_merge(config, tpl)
            print(f'📋 已加载模板: {template_name}')
        else:
            print(f'⚠️  模板不存在: {tpl_path}，使用默认配置')

    # 2) 加载config.yaml（覆盖模板）
    if config_path and yaml:
        with open(config_path, 'r', encoding='utf-8') as f:
            custom = yaml.safe_load(f)
            if custom:
                _deep_merge(config, custom)
        print(f'📋 已加载配置: {config_path}')
    elif yaml:
        default_path = Path(__file__).parent.parent / 'config.yaml'
        if default_path.exists():
            with open(default_path, 'r', encoding='utf-8') as f:
                custom = yaml.safe_load(f)
                if custom:
                    _deep_merge(config, custom)

    return config

def _deep_merge(base, override):
    """递归合并字典"""
    for k, v in override.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            _deep_merge(base[k], v)
        else:
            base[k] = v

def build_specs(config):
    """从配置构建内部规格字典"""
    body = config.get('body', DEFAULT_CONFIG['body'])
    headings_cfg = config.get('headings', {})
    table_cfg = config.get('table', DEFAULT_CONFIG['table'])
    page_cfg = config.get('page', DEFAULT_CONFIG['page'])

    heading_spec = {}
    for level in [1, 2, 3, 4]:
        h = headings_cfg.get(level, headings_cfg.get(str(level), DEFAULT_CONFIG['headings'].get(level, {})))
        heading_spec[level] = {
            'size': h.get('size', 12),
            'bold': h.get('bold', True),
            'name': h.get('font', '宋体'),
            'space_before': h.get('space_before', 6),
            'space_after': h.get('space_after', 3),
        }

    body_spec = {
        'font': body.get('font', '宋体'),
        'size': body.get('size', 12),
        'bold': body.get('bold', False),
        'first_indent_chars': body.get('first_indent_chars', 2),
        'line_spacing': body.get('line_spacing', 1.5),
    }

    table_spec = {
        'font': table_cfg.get('font', '宋体'),
        'size': table_cfg.get('size', 10.5),
        'bold_header': table_cfg.get('bold_header', True),
        'style': table_cfg.get('style', 'Table Grid'),
    }

    page_spec = {
        'margin_top': page_cfg.get('margin_top', 2.54),
        'margin_bottom': page_cfg.get('margin_bottom', 2.54),
        'margin_left': page_cfg.get('margin_left', 2.00),
        'margin_right': page_cfg.get('margin_right', 2.00),
    }

    return heading_spec, body_spec, table_spec, page_spec


# ===== 下划线/占位符保留逻辑 =====
UNDERLINE_PATTERNS = [
    # 致：_________（招标人） → 保留下划线和括号说明
    (r'(致[：:]\s*)(_{3,})(\s*（[^）]+）)', r'\1{value}\3'),
    # 根据：_________（招标文件编号）
    (r'(根据[：:]\s*)(_{3,})(\s*（[^）]+）)', r'\1{value}\3'),
    # 项目名称：_________
    (r'(项目名称[：:]\s*)(_{3,})', r'\1{value}'),
    # 通用：文字：_____（说明）
    (r'([：:]\s*)(_{3,})(\s*（[^）]+）)', r'\1{value}\3'),
    # 通用：文字：_____（无括号说明）
    (r'([：:]\s*)(_{3,})', r'\1{value}'),
    # 括号说明保留：（招标人）（供应商）等
    (r'（(招标人|投标人|供应商|采购人|代理机构|甲方|乙方|丙方)）', r'（\1）'),
]

def preserve_underlines(text, replacements=None):
    """
    保留下划线占位符的格式，同时支持填入实际值。
    
    处理逻辑：
    1. 识别 "致：_________（招标人）" 这类模式
    2. 保留下划线和括号说明
    3. 如果提供了replacements字典，将下划线替换为实际值
    4. 如果没提供，保留下划线原样输出
    
    replacements 示例:
    {
        "致：": "致：上海交通大学",
        "项目名称：": "项目名称：信息化系统建设",
    }
    """
    if not replacements:
        return text
    
    lines = text.split('\n')
    result = []
    for line in lines:
        for pattern, template in UNDERLINE_PATTERNS:
            match = re.search(pattern, line)
            if match:
                # 找到匹配的下划线模式
                for key, value in replacements.items():
                    if key in line:
                        # 替换下划线部分，保留下划线后的括号说明
                        line = re.sub(
                            r'(' + re.escape(key) + r'\s*)(_{3,})(\s*（[^）]+）)',
                            lambda m: m.group(1) + value + m.group(3),
                            line
                        )
                        # 无括号说明的情况
                        if '（' not in line:
                            line = re.sub(
                                r'(' + re.escape(key) + r'\s*)(_{3,})',
                                lambda m: m.group(1) + value,
                                line
                            )
                        break
        result.append(line)
    return '\n'.join(result)


# ===== 文本格式化工具 =====
def set_run_font(run, font_name='宋体', size=12, bold=False):
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    for attr in ['w:asciiTheme', 'w:hAnsiTheme', 'w:eastAsiaTheme']:
        try: del rFonts.attrib[qn(attr)]
        except KeyError: pass

def init_doc_styles(doc, heading_spec, body_spec):
    try:
        style = doc.styles['Normal']
        style.font.name = body_spec['font']
        style.font.size = Pt(body_spec['size'])
        style.paragraph_format.line_spacing = body_spec['line_spacing']
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    except KeyError: pass
    for level, spec in heading_spec.items():
        try:
            style = doc.styles[f'Heading {level}']
            style.font.name = spec['name']
            style.font.size = Pt(spec['size'])
            style.font.bold = spec['bold']
            rPr = style.element.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.insert(0, rFonts)
            rFonts.set(qn('w:ascii'), spec['name'])
            rFonts.set(qn('w:hAnsi'), spec['name'])
            rFonts.set(qn('w:eastAsia'), spec['name'])
        except KeyError: pass

# ===== 全角半角检测 =====
HALF_TO_FULL = {
    ',': '，', '.': '。', '!': '！', '?': '？', ':': '：', ';': '；',
    '(': '（', ')': '）', '[': '【', ']': '】', '<': '《', '>': '》',
    '\"': '\"', "'": "'", '~': '～', '-': '—',
}
FULLWIDTH_TRANS = str.maketrans(
    '０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ'
    'ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ',
    '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
)

def fix_punctuation(text):
    """全角半角自动修复，跳过编号和小数中的点号"""
    fixed = text.translate(FULLWIDTH_TRANS)
    result = []
    in_cn = True
    changes = 0
    for i, ch in enumerate(fixed):
        # 点号特殊处理：前后都是数字时不替换（编号1.1、小数3.14等）
        if ch == '.' and i > 0 and i < len(fixed) - 1:
            prev_ch = fixed[i - 1]
            next_ch = fixed[i + 1]
            if prev_ch.isdigit() and next_ch.isdigit():
                result.append(ch)
                continue
        if ch in HALF_TO_FULL and in_cn:
            result.append(HALF_TO_FULL[ch])
            changes += 1
        else:
            result.append(ch)
        if '\u4e00' <= ch <= '\u9fff' or '\u3000' <= ch <= '\u303f': in_cn = True
        elif ch.isascii() and ch.isalpha(): in_cn = False
    return ''.join(result), changes

def scan_punctuation(text):
    issues = []
    for i, line in enumerate(text.split('\n')):
        for half, full in HALF_TO_FULL.items():
            for m in re.finditer(rf'[\u4e00-\u9fff]{re.escape(half)}[\u4e00-\u9fff]', line):
                issues.append({'line': i+1, 'type': '半角标点混入中文', 'char': half, 'should_be': full,
                    'context': line[max(0,m.start()-8):m.end()+8]})
        for m in re.finditer(r'[\uff10-\uff19\uff21-\uff3a\uff41-\uff5a]+', line):
            issues.append({'line': i+1, 'type': '全角数字/字母', 'char': m.group(),
                'context': line[max(0,m.start()-5):m.end()+5]})
    return issues

# ===== Markdown解析 =====
def parse_md_table(lines, start_idx):
    headers, rows, idx = [], [], start_idx
    if idx < len(lines) and lines[idx].strip().startswith('|'):
        headers = [c.strip() for c in lines[idx].strip().strip('|').split('|')]
        idx += 1
    if idx < len(lines) and re.match(r'^[\s|\-:]+$', lines[idx]): idx += 1
    while idx < len(lines) and lines[idx].strip().startswith('|'):
        rows.append([c.strip() for c in lines[idx].strip().strip('|').split('|')])
        idx += 1
    return headers, rows, idx

# ===== Word文档构建 =====
def add_heading(doc, text, level, heading_spec):
    spec = heading_spec.get(level, heading_spec.get(4, {}))
    p = doc.add_heading(text, level=level)
    for run in p.runs: set_run_font(run, spec['name'], spec['size'], spec['bold'])
    p.paragraph_format.space_before = Pt(spec['space_before'])
    p.paragraph_format.space_after = Pt(spec['space_after'])
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return p

# ===== Mermaid图表渲染 =====
def render_mermaid(code, output_png=None, cache_dir=None):
    """将Mermaid代码渲染为PNG图片。
    
    依赖：mermaid-cli (npm install -g @mermaid-js/mermaid-cli)
    如果未安装mmdc命令，返回None，调用方插入占位文本。
    """
    import subprocess, tempfile, hashlib
    if cache_dir is None:
        cache_dir = Path(tempfile.gettempdir()) / 'mermaid_cache'
    cache_dir = Path(cache_dir)
    cache_dir.mkdir(parents=True, exist_ok=True)
    
    # 内容哈希做缓存键
    content_hash = hashlib.md5(code.encode()).hexdigest()[:12]
    if output_png is None:
        output_png = cache_dir / f'mermaid_{content_hash}.png'
    else:
        output_png = Path(output_png)
    
    # 缓存命中
    if output_png.exists() and output_png.stat().st_size > 0:
        return str(output_png)
    
    # 写临时.mmd文件
    mmd_file = cache_dir / f'mermaid_{content_hash}.mmd'
    mmd_file.write_text(code, encoding='utf-8')
    
    try:
        result = subprocess.run(
            ['mmdc', '-i', str(mmd_file), '-o', str(output_png),
             '-b', 'white', '-w', 1200],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode == 0 and output_png.exists():
            return str(output_png)
        else:
            print(f'⚠️  Mermaid渲染失败: {result.stderr[:200]}')
            return None
    except FileNotFoundError:
        print('⚠️  未安装mermaid-cli (mmdc)，跳过图表渲染。安装: npm install -g @mermaid-js/mermaid-cli')
        return None
    except subprocess.TimeoutExpired:
        print('⚠️  Mermaid渲染超时(30s)，跳过')
        return None
    except Exception as e:
        print(f'⚠️  Mermaid渲染异常: {e}')
        return None

def add_mermaid_block(doc, code, body_spec):
    """在Word文档中插入Mermaid图表（或占位文本）"""
    png_path = render_mermaid(code)
    if png_path:
        # 插入图片（居中）
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        try:
            from docx.shared import Inches
            run.add_picture(png_path, width=Inches(5.5))
        except Exception as e:
            # 图片插入失败，退化为占位
            p.clear()
            run = p.add_run(f'[Mermaid图表渲染失败: {e}]')
            set_run_font(run, body_spec['font'], body_spec['size'], False)
            run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
    else:
        # mmdc未安装，插入占位文本
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f'[Mermaid图表 - 需安装mmdc渲染]\n{code[:200]}')
        set_run_font(run, body_spec['font'], body_spec['size'] - 1, False)
        run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)

# ===== 目录域插入 =====
def insert_toc(doc, title='目  录', toc_levels='1-3'):
    """在文档开头插入Word目录域代码(TOC field)。
    
    打开Word后按Ctrl+A → F9更新域即可自动生成目录。
    toc_levels: '1-3' 表示显示到三级标题。
    
    铁律：必须在生成完所有标题后调用，且插在文档最前面。
    """
    # 在文档开头插入一个空段落
    new_para = OxmlElement('w:p')
    body = doc.element.body
    body.insert(0, new_para)
    
    # 在空段落里插入目录标题
    toc_title = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    pStyle = OxmlElement('w:pStyle')
    pStyle.set(qn('w:val'), 'Heading1')
    pPr.append(pStyle)
    # 居中
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center')
    pPr.append(jc)
    toc_title.append(pPr)
    r = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = title
    r.append(t)
    toc_title.append(r)
    body.insert(0, toc_title)
    
    # 插入TOC域代码
    toc_para = OxmlElement('w:p')
    r = OxmlElement('w:r')
    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')
    r.append(fldChar_begin)
    toc_para.append(r)
    
    r2 = OxmlElement('w:r')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = f' TOC \\o "{toc_levels}" \\h \\z \\u '
    r2.append(instrText)
    toc_para.append(r2)
    
    r3 = OxmlElement('w:r')
    fldChar_sep = OxmlElement('w:fldChar')
    fldChar_sep.set(qn('w:fldCharType'), 'separate')
    r3.append(fldChar_sep)
    toc_para.append(r3)
    
    r4 = OxmlElement('w:r')
    t2 = OxmlElement('w:t')
    t2.text = '（打开Word后按 Ctrl+A 全选 → F9 更新域，目录自动生成）'
    r4.append(t2)
    toc_para.append(r4)
    
    r5 = OxmlElement('w:r')
    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')
    r5.append(fldChar_end)
    toc_para.append(r5)
    
    # 插入到标题之后（index 2）
    body.insert(2, toc_para)
    
    # 在目录后面加分页符
    page_break = OxmlElement('w:p')
    r_brk = OxmlElement('w:r')
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    r_brk.append(br)
    page_break.append(r_brk)
    body.insert(3, page_break)
    
    return True

# ===== 假标题检测 =====
def check_heading_styles(doc):
    """检查文档中是否有'假标题'——看起来像标题但没用Heading样式的段落。
    
    常见假标题模式：
    1. Normal样式 + 加粗 + 短文本(≤30字) + 无缩进
    2. Normal样式 + 短文本 + 开头是编号(如'1.'/'一、'/'第一章')
    
    Word目录只识别Heading 1-4，假标题不会出现在目录里。
    """
    issues = []
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        style = para.style.name if para.style else ''
        
        # 只检查Normal样式的段落
        if style != 'Normal':
            continue
        
        # 跳过空段落和超长段落(不可能是标题)
        if len(text) < 2 or len(text) > 50:
            continue
        
        # 跳过首行缩进的段落（明显是正文）
        indent = para.paragraph_format.first_line_indent
        if indent is not None and indent != 0:
            continue
        
        # 检测模式1：加粗 + 短文本 + 无缩进
        is_bold = any(run.bold for run in para.runs if run.bold is not None)
        
        # 检测模式2：编号开头
        has_numbering = bool(re.match(
            r'^(第[一二三四五六七八九十百]+[章节条款]|[一二三四五六七八九十]+[、．.]'
            r'|\d+[\.\、]|\d+\.\d+|[\(（]\d+[\)）])',
            text
        ))
        
        if is_bold and not has_numbering:
            # 加粗但没编号——可能不是标题，给个warning
            issues.append({
                'para_idx': i + 1,
                'text': text[:40],
                'type': '疑似假标题',
                'detail': 'Normal样式+加粗但无Heading样式，目录不会收录',
                'severity': 'warn',
                'fix': f'选中段落 → Word样式栏改为 Heading 1/2/3'
            })
        elif has_numbering:
            # 有编号但不是Heading样式——很可能真是标题
            issues.append({
                'para_idx': i + 1,
                'text': text[:40],
                'type': '疑似标题未设样式',
                'detail': f'检测到"{text[:8]}"开头，可能是标题但未用Heading样式，目录不会收录',
                'severity': 'fail',
                'fix': f'选中段落 → Word样式栏改为对应的 Heading 级别'
            })
    
    return issues

def add_body(doc, text, body_spec):
    p = doc.add_paragraph(text, style='Normal')
    for run in p.runs: set_run_font(run, body_spec['font'], body_spec['size'], body_spec['bold'])
    p.paragraph_format.first_line_indent = Twips(480)
    p.paragraph_format.line_spacing = body_spec['line_spacing']
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p

def add_table(doc, headers, rows, table_spec):
    if not headers: return None
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = table_spec['style']
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        run = p.add_run(header)
        set_run_font(run, table_spec['font'], table_spec['size'], table_spec['bold_header'])
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for row_data in rows:
        row_cells = table.add_row().cells
        for i, val in enumerate(row_data):
            row_cells[i].text = ''
            p = row_cells[i].paragraphs[0]
            run = p.add_run(val)
            set_run_font(run, table_spec['font'], table_spec['size'], False)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Twips(0)
    return table

# ===== 暗标模式 =====
def apply_anonymous(text, anonymous_config):
    """应用暗标模式，替换公司标识"""
    if not anonymous_config.get('enabled', False):
        return text
    replace_with = anonymous_config.get('replace_with', '投标人')
    # 替换通用词
    for word in anonymous_config.get('replace_words', ['我公司', '本公司']):
        text = text.replace(word, replace_with)
    # 替换公司名称
    for name in anonymous_config.get('company_names', []):
        text = text.replace(name, replace_with)
    # 替换地址
    for addr in anonymous_config.get('company_addresses', []):
        text = text.replace(addr, '【地址已隐藏】')
    print(f'🔒 暗标模式：已替换公司标识')
    return text

# ===== 质检规则 =====

# 标书常见错别字字典
# 格式: "错误词": ("正确词", "说明")
# 注意：只收录在标书语境下确定错误的词组，避免单字匹配误报
BID_TYPO_DICT = {
    # 财务规范用字（2001年起财政部规范用「账」）
    "帐号": ("账号", "财务规范用「账」"),
    "帐户": ("账户", "财务规范用「账」"),
    "帐单": ("账单", "财务规范用「账」"),
    "帐目": ("账目", "财务规范用「账」"),
    "帐务": ("账务", "财务规范用「账」"),
    "结帐": ("结账", "财务规范用「账」"),
    "对帐": ("对账", "财务规范用「账」"),
    "记帐": ("记账", "财务规范用「账」"),
    "转帐": ("转账", "财务规范用「账」"),
    # 系统登录用字
    "登陆系统": ("登录系统", "系统登录用「录」"),
    "登陆平台": ("登录平台", "系统登录用「录」"),
    "登陆网站": ("登录网站", "系统登录用「录」"),
    # 行政用字
    "做为": ("作为", "作为用「作」"),
    "按装": ("安装", "安装用「安」"),
    "部暑": ("部署", "部署用「署」"),
    "布署": ("部署", "部署用「署」"),
    # 两岸用语（标书应使用大陆用语）
    "软体": ("软件", "大陆用语用「软件」"),
    "硬体": ("硬件", "大陆用语用「硬件」"),
    "网路": ("网络", "大陆用语用「网络」"),
    "萤幕": ("屏幕", "大陆用语用「屏幕」"),
    "滑鼠": ("鼠标", "大陆用语用「鼠标」"),
    "讯息": ("信息", "大陆用语用「信息」"),
    "资料库": ("数据库", "大陆用语用「数据库」"),
    "视窗": ("窗口", "大陆用语用「窗口」"),
    "专案管理": ("项目管理", "大陆用语用「项目」"),
    # 常见形近/同音字
    "既使": ("即使", "即使用「即」"),
    "己经": ("已经", "已经用「已」"),
    "自已": ("自己", "自己用「己」"),
}

def check_typos(text):
    """检查常见错别字（标书场景高频）

    只匹配确定错误的词组，返回上下文便于人工确认。
    """
    issues = []
    for wrong, (correct, note) in BID_TYPO_DICT.items():
        start = 0
        while True:
            pos = text.find(wrong, start)
            if pos < 0:
                break
            # 取上下文
            ctx_start = max(0, pos - 10)
            ctx_end = min(len(text), pos + len(wrong) + 10)
            context = text[ctx_start:ctx_end].replace('\n', ' ')
            issues.append({
                'type': f'错别字「{wrong}」应为「{correct}」',
                'word': wrong,
                'correct': correct,
                'note': note,
                'context': context,
            })
            start = pos + len(wrong)
    return issues

# 不一致检测字段模式
# 每个字段定义一组正则，提取值后做跨段落比对
CONSISTENCY_PATTERNS = {
    '投标有效期': [
        r'投标有效期\s*[:：为]?\s*(\d+)\s*天',
        r'有效期\s*[:：为]?\s*(\d+)\s*个?天',
    ],
    '总报价': [
        r'(?:总报价|总[价金额计]|合计金额|报价总[价额])\s*[:：为是]?\s*(\d[\d,]*(?:\.\d+)?)\s*(万元|元|万)',
        r'投标(?:总)?报价\s*[:：为是]?\s*(\d[\d,]*(?:\.\d+)?)\s*(万元|元|万)',
    ],
    '人员总数': [
        r'(?:总人数|人员[总数配置量]|配备人员)\s*[:：为共]?\s*(\d+)\s*人',
        r'共\s*(\d+)\s*人',
    ],
    '服务期限': [
        r'(?:服务期[限制]|项目期[限制]|合同期[限制]|合作期[限制])\s*[:：为]?\s*(\d+)\s*(?:个?月|年|天)',
    ],
}

def check_consistency(doc):
    """检查前后不一致（日期/金额/人数/期限跨段落比对）

    对每个字段，在全文中提取所有出现的值，如果同一字段出现不同值则告警。
    """
    import re
    issues = []
    paragraphs = [(i, p.text) for i, p in enumerate(doc.paragraphs)]

    for field, patterns in CONSISTENCY_PATTERNS.items():
        found = []  # [(para_idx, value, unit, context)]
        for idx, text in paragraphs:
            for pat in patterns:
                for m in re.finditer(pat, text):
                    value = m.group(1).replace(',', '')
                    unit = m.group(2) if m.lastindex and m.lastindex >= 2 else ''
                    ctx_start = max(0, m.start() - 10)
                    ctx_end = min(len(text), m.end() + 10)
                    context = text[ctx_start:ctx_end].replace('\n', ' ')
                    found.append((idx + 1, value, unit, context))

        if len(found) < 2:
            continue

        # 归一化比较（金额统一换算为元）
        def normalize(value, unit):
            v = float(value)
            if unit in ('万元', '万'):
                v *= 10000
            return v

        # 按归一化值分组
        groups = {}
        for para_idx, value, unit, context in found:
            norm = normalize(value, unit) if field == '总报价' else float(value)
            key = f'{value}{unit}' if unit else value
            if norm not in groups:
                groups[norm] = []
            groups[norm].append((para_idx, key, context))

        # 多个不同值 = 不一致
        if len(groups) >= 2:
            desc_parts = []
            for norm, entries in groups.items():
                locations = [f'段落{e[0]}({e[1]})' for e in entries]
                desc_parts.append(' / '.join(locations))
            issues.append({
                'type': f'{field}前后不一致',
                'detail': ' vs '.join(desc_parts),
                'values': list(groups.keys()),
            })

    return issues

def check_forbidden_words(text, forbidden_words):
    """检查禁用词"""
    issues = []
    for word in forbidden_words:
        if word in text:
            issues.append({'type': '禁用词', 'word': word})
    return issues

def check_placeholders(text):
    """检查占位符"""
    patterns = [r'请填写', r'XXX', r'TODO', r'____', r'待填', r'略']
    issues = []
    for pat in patterns:
        for m in re.finditer(pat, text):
            issues.append({'type': '占位符', 'word': m.group(), 'pos': m.start()})
    return issues

def check_scoring_coverage(matrix_file, content_file):
    """评分项覆盖矩阵检查：比对矩阵与正文，确保每个评分项都有对应章节响应
    
    注意：本函数仅支持markdown格式的正文文件（通过# heading匹配）。
    如需检查docx格式，请先转换为markdown再使用。
    """
    with open(matrix_file, 'r', encoding='utf-8') as f:
        matrix_text = f.read()
    with open(content_file, 'r', encoding='utf-8') as f:
        content_text = f.read()

    # 解析矩阵表格行：| 编号 | 内容 | 分值 | 对应章节 | 状态 |
    matrix_rows = []
    for line in matrix_text.split('\n'):
        line = line.strip()
        if line.startswith('|') and not line.startswith('|---') and not line.startswith('| 评分项'):
            cells = [c.strip() for c in line.split('|')[1:-1]]
            if len(cells) >= 5:
                matrix_rows.append({
                    'id': cells[0],
                    'content': cells[1],
                    'score': cells[2],
                    'chapter': cells[3],
                    'status': cells[4],
                })

    results = {'total': len(matrix_rows), 'covered': 0, 'uncovered': [], 'status_mismatch': []}
    for row in matrix_rows:
        # 检查状态是否为已响应
        if '✅' not in row['status'] and '已响应' not in row['status']:
            results['status_mismatch'].append(row)
            continue
        # 提取章节编号做多模式匹配
        # 如"第四章2.1" -> 同时尝试匹配"第四章"、"四"、"2.1"、"4.2.1"等
        chapter_raw = row['chapter'].strip()
        chapter_refs = set()
        chapter_refs.add(chapter_raw)  # 原文全称
        chapter_refs.add(chapter_raw.replace('第', '').replace('章', '').replace('节', '').strip())  # 去掉"第章节"
        # 提取中文序号（一二三...）转为阿拉伯数字也加进去
        cn_num_map = {'一': '1', '二': '2', '三': '3', '四': '4', '五': '5',
                      '六': '6', '七': '7', '八': '8', '九': '9', '十': '10'}
        cn_match = re.search(r'第([一二三四五六七八九十]+)章', chapter_raw)
        if cn_match:
            cn_str = cn_match.group(1)
            if cn_str == '十':
                chapter_refs.add('第10章')
                chapter_refs.add('10')
            elif '十' in cn_str:
                # 如"十一"->"11", "二十"->"20"
                parts = cn_str.split('十')
                tens = cn_num_map.get(parts[0], '0') if parts[0] else '1'
                ones = cn_num_map.get(parts[1], '0') if len(parts) > 1 and parts[1] else '0'
                arabic = str(int(tens) * 10 + int(ones))
                chapter_refs.add(f'第{arabic}章')
                chapter_refs.add(arabic)
            elif cn_str in cn_num_map:
                chapter_refs.add(f'第{cn_num_map[cn_str]}章')
                chapter_refs.add(cn_num_map[cn_str])
        # 提取纯数字编号（如"2.1"）
        num_match = re.search(r'(\d+\.?\d*)', chapter_raw)
        if num_match:
            chapter_refs.add(num_match.group(1))
        found = False
        for heading_line in content_text.split('\n'):
            if heading_line.strip().startswith('#'):
                heading_text = heading_line.lstrip('#').strip()
                for ref in chapter_refs:
                    if ref in heading_text:
                        found = True
                        break
                if found:
                    break
        if found:
            results['covered'] += 1
        else:
            results['uncovered'].append(row)

    return results

def check_priority_issues(content_file, dark_mode=False):
    """P0/P1/P2合规分级检查：扫描正文中的合规风险关键词
    
    注意：P2级问题（措辞优化、排版细节、图片清晰度）需要人工检查，本函数不做自动扫描。
    dark_mode: 是否暗标模式。暗标模式下才触发身份泄露检测。
    """
    with open(content_file, 'r', encoding='utf-8') as f:
        text = f.read()

    p0_patterns = [
        # 资质过期相关
        # 资质有效期检查（覆盖多种日期格式）
        (r'有效期至\s*20[12]\d年', 'P0', '资质有效期检查'),
        (r'有效期至\s*20[12]\d[-./]\d{1,2}([-./]\d{1,2})?', 'P0', '资质有效期检查（数字日期格式）'),
        (r'有效期[：:]\s*20[12]\d年\d{1,2}月', 'P0', '资质有效期检查（冒号格式）'),
        (r'20[12]\d[-./]\d{1,2}[-./]\d{1,2}\s*过期', 'P0', '资质过期风险'),
        # 项目名称占位符（说明还没填）
        (r'【.+?项目.+?】', 'P0', '项目名称未填（占位符）'),
        (r'XXX.{0,5}项目', 'P0', '项目名称未填（XXX）'),
        # 偏离表虚假响应
        (r'(?:无偏离|均响应|完全响应).{0,20}(?:参数|技术)', 'P0', '偏离表需核实是否与产品彩页一致'),
    ]

    p1_patterns = [
        # 格式问题
        # 只匹配整行只有加粗文字的情况，避免误杀正文中的加粗强调
        (r'^\s*\*\*[^\*]+\*\*\s*$', 'P1', '整行加粗可能冒充标题（应使用#号标记）'),
        # 前后不一致信号
        (r'(?:详见|见|参见).{0,10}(?:第|附件)', 'P1', '交叉引用需人工确认一致性'),
        # 金额占位
        (r'【金额.+?】', 'P1', '金额未填（占位符）'),
        (r'【人名.+?】', 'P1', '人名未填（占位符）'),
        (r'【日期.+?】', 'P1', '日期未填（占位符）'),
    ]

    issues = []
    all_patterns = list(p0_patterns)
    # 暗标泄露身份检测仅在暗标模式下触发，避免正常引用误报
    if dark_mode:
        all_patterns.append(
            (r'(?:投标人|供应商|响应人).{0,5}(?:公司|有限|集团)', 'P0', '暗标可能泄露身份（需人工确认）')
        )
    for pattern, level, desc in all_patterns + p1_patterns:
        for m in re.finditer(pattern, text):
            line_num = text[:m.start()].count('\n') + 1
            issues.append({
                'level': level,
                'desc': desc,
                'match': m.group()[:50],
                'line': line_num,
            })

    # 统计
    p0_count = len([i for i in issues if i['level'] == 'P0'])
    p1_count = len([i for i in issues if i['level'] == 'P1'])
    return {'issues': issues, 'p0_count': p0_count, 'p1_count': p1_count,
            'deliverable': p0_count == 0 and p1_count <= 2}

# ===== 主转换函数 =====
def md_to_docx(md_text, output_path, auto_fix=True, dark_mode=False, config=None, no_toc=False):
    """Markdown → Word转换"""
    if config is None:
        config = DEFAULT_CONFIG
    
    # 构建规格
    heading_spec, body_spec, table_spec, page_spec = build_specs(config)
    
    # 全角半角修复
    if auto_fix:
        fixed_text, changes = fix_punctuation(md_text)
        if changes > 0:
            print(f'🔧 自动修复 {changes} 处全角半角混用')
            md_text = fixed_text
    
    # 暗标模式
    if dark_mode:
        anonymous_cfg = config.get('anonymous', DEFAULT_CONFIG['anonymous'])
        anonymous_cfg['enabled'] = True
        md_text = apply_anonymous(md_text, anonymous_cfg)
    
    # 创建文档
    doc = Document()
    init_doc_styles(doc, heading_spec, body_spec)
    
    # 设置页边距
    for section in doc.sections:
        section.top_margin = Cm(page_spec['margin_top'])
        section.bottom_margin = Cm(page_spec['margin_bottom'])
        section.left_margin = Cm(page_spec['margin_left'])
        section.right_margin = Cm(page_spec['margin_right'])
    
    # 解析Markdown
    lines = md_text.split('\n')
    i, in_table = 0, False
    table_headers, table_rows = [], []
    while i < len(lines):
        line = lines[i].rstrip()
        # Mermaid代码块检测
        if line.strip().startswith('```mermaid'):
            i += 1
            mermaid_code = []
            while i < len(lines) and not lines[i].strip().startswith('```'):
                mermaid_code.append(lines[i])
                i += 1
            i += 1  # 跳过结束```
            add_mermaid_block(doc, '\n'.join(mermaid_code), body_spec)
            continue
        # 普通代码块跳过（非mermaid）
        if line.strip().startswith('```') and not line.strip().startswith('```mermaid'):
            i += 1
            while i < len(lines) and not lines[i].strip().startswith('```'):
                i += 1
            i += 1
            continue
        # 表格处理
        if line.startswith('|') and line.endswith('|'):
            if not in_table:
                table_headers, table_rows, i = parse_md_table(lines, i)
                in_table = True
                continue
            else:
                headers, rows, i = parse_md_table(lines, i)
                if headers: add_table(doc, headers, rows, table_spec)
                continue
        if in_table:
            add_table(doc, table_headers, table_rows, table_spec)
            in_table = False
        # 标题
        h_match = re.match(r'^(#{1,4})\s+(.+)$', line)
        if h_match:
            add_heading(doc, h_match.group(2).strip(), len(h_match.group(1)), heading_spec)
            i += 1; continue
        # 引用块
        if line.startswith('>'):
            text = line.lstrip('> ').strip()
            if text:
                p = doc.add_paragraph()
                run = p.add_run(text)
                set_run_font(run, body_spec['font'], body_spec['size'], False)
                run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
                p.paragraph_format.left_indent = Cm(1)
            i += 1; continue
        # 列表
        list_match = re.match(r'^(\s*)[-*]\s+(.+)$', line)
        if list_match:
            text = list_match.group(2).strip()
            p = doc.add_paragraph(style='List Bullet')
            p.clear()
            run = p.add_run(text)
            set_run_font(run, body_spec['font'], body_spec['size'], False)
            p.paragraph_format.line_spacing = body_spec['line_spacing']
            i += 1; continue
        # 空行跳过
        if not line.strip():
            i += 1; continue
        # 普通段落
        clean_text = re.sub(r'\*\*(.+?)\*\*', r'\1', line)
        clean_text = re.sub(r'\*(.+?)\*', r'\1', clean_text)
        add_body(doc, clean_text, body_spec)
        i += 1
    if in_table: add_table(doc, table_headers, table_rows, table_spec)
    # 插入目录域（默认开启，--no-toc可关闭）
    if not no_toc:
        insert_toc(doc)
    doc.save(output_path)
    return output_path

# ===== 质检函数 =====
def check_docx(docx_path, config=None):
    if config is None:
        config = DEFAULT_CONFIG
    quality_cfg = config.get('quality', DEFAULT_CONFIG['quality'])
    forbidden_words = quality_cfg.get('forbidden_words', [])
    
    doc = Document(docx_path)
    results = {'pass': [], 'warn': [], 'fail': [],
        'stats': {'paragraphs': len(doc.paragraphs), 'tables': len(doc.tables),
            'total_chars': sum(len(p.text) for p in doc.paragraphs)}}
    
    for i, para in enumerate(doc.paragraphs):
        if para.style.name.startswith('Heading'):
            if para.alignment not in (WD_ALIGN_PARAGRAPH.LEFT, None):
                results['warn'].append(f'段落{i+1}: 标题未左对齐')
        if para.style.name == 'Normal':
            indent = para.paragraph_format.first_line_indent
            if indent is None or indent == 0:
                results['warn'].append(f'段落{i+1}: 正文无首行缩进')
    
    # 全角半角检查
    if quality_cfg.get('check_punctuation', True):
        for i, para in enumerate(doc.paragraphs):
            issues = scan_punctuation(para.text)
            for issue in issues:
                results['fail'].append(f'段落{i+1}: {issue["type"]}')
    
    # 占位符检查
    if quality_cfg.get('check_placeholder', True):
        for i, para in enumerate(doc.paragraphs):
            ph_issues = check_placeholders(para.text)
            for issue in ph_issues:
                results['fail'].append(f'段落{i+1}: 未填写占位符 "{issue["word"]}"')
    
    # 禁用词检查
    if quality_cfg.get('check_forbidden_words', True) and forbidden_words:
        for i, para in enumerate(doc.paragraphs):
            fw_issues = check_forbidden_words(para.text, forbidden_words)
            for issue in fw_issues:
                results['fail'].append(f'段落{i+1}: 禁用词 "{issue["word"]}"')
    
    # 错别字检查
    if quality_cfg.get('check_typos', True):
        for i, para in enumerate(doc.paragraphs):
            typo_issues = check_typos(para.text)
            for issue in typo_issues:
                results['warn'].append(
                    f'段落{i+1}: 错别字「{issue["word"]}」应为「{issue["correct"]}」'
                    f'({issue["note"]}) 上下文: ...{issue["context"]}...'
                )
    
    # 前后不一致检测
    if quality_cfg.get('check_consistency', True):
        consistency_issues = check_consistency(doc)
        for issue in consistency_issues:
            results['fail'].append(f'{issue["type"]}: {issue["detail"]}')
    
    # 假标题检测（Heading样式完整性）
    heading_issues = check_heading_styles(doc)
    for issue in heading_issues:
        target = results['fail'] if issue['severity'] == 'fail' else results['warn']
        target.append(f"段落{issue['para_idx']}: {issue['type']} — \"{issue['text']}\" ({issue['detail']})")
    
    total = len(results['pass']) + len(results['warn']) + len(results['fail'])
    results['summary'] = {'total_checks': total, 'pass': len(results['pass']),
        'warn': len(results['warn']), 'fail': len(results['fail']),
        'status': 'PASS' if len(results['fail']) == 0 else 'FAIL'}
    return results

def print_check_report(results):
    print('\n' + '='*60)
    print('📋 标书格式自检报告')
    print('='*60)
    print(f'📊 统计: {results["stats"]["paragraphs"]}段落, '
          f'{results["stats"]["tables"]}表格, {results["stats"]["total_chars"]}字')
    print(f'✅ 通过: {results["summary"]["pass"]}')
    print(f'⚠️  警告: {results["summary"]["warn"]}')
    print(f'❌ 失败: {results["summary"]["fail"]}')
    if results['warn']:
        print('\n⚠️  警告项:')
        for w in results['warn'][:20]: print(f'  {w}')
    if results['fail']:
        print('\n❌ 失败项:')
        for f in results['fail'][:20]: print(f'  {f}')
    print(f'\n🏁 结论: {"✅ PASS" if results["summary"]["status"] == "PASS" else "❌ FAIL"}')
    print('='*60)

# ===== 主入口 =====
def main():
    parser = argparse.ArgumentParser(description='标书自动化引擎 v3.2')
    parser.add_argument('input', nargs='?', default=None, help='Markdown输入文件（--check-scoring/--check-priority模式不需要）')
    parser.add_argument('-o', '--output', default=None, help='输出docx路径')
    parser.add_argument('--scan', action='store_true', help='仅扫描全角半角')
    parser.add_argument('--check', action='store_true', help='生成后自检')
    parser.add_argument('--no-fix', action='store_true', help='跳过全角修复')
    parser.add_argument('--暗标', action='store_true', help='暗标模式(去公司标识)')
    parser.add_argument('--template', type=str, default=None,
                        choices=['government', 'enterprise', 'engineering'],
                        help='使用预设模板: government=政府采购, enterprise=企业投标, engineering=工程类')
    parser.add_argument('--config', type=str, default=None,
                        help='指定config.yaml配置文件路径')
    parser.add_argument('--json', action='store_true', help='JSON输出')
    parser.add_argument('--no-toc', action='store_true', help='不插入目录域（默认自动生成目录）')
    parser.add_argument('--check-scoring', nargs=2, metavar=('MATRIX', 'CONTENT'),
                        help='评分项覆盖矩阵检查：传入矩阵文件和正文文件')
    parser.add_argument('--check-priority', type=str, default=None,
                        help='P0/P1/P2合规分级检查：传入正文md文件')
    args = parser.parse_args()
    
    # --check-scoring 模式：评分项覆盖矩阵检查
    if args.check_scoring:
        matrix_file, content_file = args.check_scoring
        if not os.path.exists(matrix_file):
            print(f'❌ 矩阵文件不存在: {matrix_file}'); sys.exit(1)
        if not os.path.exists(content_file):
            print(f'❌ 正文文件不存在: {content_file}'); sys.exit(1)
        results = check_scoring_coverage(matrix_file, content_file)
        print('\n' + '='*60)
        print('📋 评分项覆盖矩阵检查报告')
        print('='*60)
        print(f'📊 评分项总数: {results["total"]}')
        print(f'✅ 已覆盖: {results["covered"]}')
        print(f'❌ 未覆盖: {len(results["uncovered"])}')
        print(f'⚠️  状态异常: {len(results["status_mismatch"])}')
        if results['uncovered']:
            print('\n❌ 未覆盖的评分项:')
            for row in results['uncovered']:
                print(f'  {row["id"]} | {row["content"]} | 应在: {row["chapter"]}')
        if results['status_mismatch']:
            print('\n⚠️  状态未标"已响应"的评分项:')
            for row in results['status_mismatch']:
                print(f'  {row["id"]} | {row["content"]} | 状态: {row["status"]}')
        status = '✅ PASS' if not results['uncovered'] and not results['status_mismatch'] else '❌ FAIL'
        print(f'\n🏁 结论: {status}')
        print('='*60)
        sys.exit(0)
    
    # --check-priority 模式：P0/P1/P2合规分级检查
    if args.check_priority:
        if not os.path.exists(args.check_priority):
            print(f'❌ 文件不存在: {args.check_priority}'); sys.exit(1)
        results = check_priority_issues(args.check_priority, dark_mode=args.暗标)
        print('\n' + '='*60)
        print('📋 P0/P1/P2 合规分级检查报告')
        print('='*60)
        print(f'🔴 P0-致命: {results["p0_count"]} 个')
        print(f'🟡 P1-重要: {results["p1_count"]} 个')
        for issue in results['issues']:
            icon = '🔴' if issue['level'] == 'P0' else '🟡'
            print(f'  {icon} 行{issue["line"]} | {issue["desc"]} | 匹配: "{issue["match"]}"')
        deliverable = '✅ 可交付' if results['deliverable'] else '❌ 不可交付（P0>0或P1>2）'
        print(f'\n🏁 结论: {deliverable}')
        print('='*60)
        sys.exit(0)
    
    if not os.path.exists(args.input):
        print(f'❌ 文件不存在: {args.input}'); sys.exit(1)
    
    # 加载配置
    config = load_config(config_path=args.config, template_name=args.template)
    
    # --check 模式：如果是.docx文件，直接质检不读文本
    if args.check and args.input.endswith('.docx'):
        results = check_docx(args.input, config=config)
        if args.json: print(json.dumps(results, ensure_ascii=False, indent=2))
        else: print_check_report(results)
        sys.exit(0)
    
    with open(args.input, 'r', encoding='utf-8') as f: md_text = f.read()
    
    if args.scan:
        issues = scan_punctuation(md_text)
        if args.json:
            print(json.dumps({'issues': issues, 'count': len(issues)}, ensure_ascii=False))
        elif issues:
            print(f'\n⚠️  发现 {len(issues)} 处全角半角问题:\n')
            for issue in issues:
                print(f'  行{issue["line"]}: [{issue["type"]}] "{issue["char"]}"')
                print(f'        上下文: …{issue["context"]}…\n')
        else:
            print('✅ 未发现全角半角问题')
        sys.exit(0)
    
    output = args.output or os.path.splitext(args.input)[0] + '_排版.docx'
    md_to_docx(md_text, output, auto_fix=not args.no_fix, dark_mode=args.暗标, config=config, no_toc=args.no_toc)
    print(f'✅ 已生成: {output}')
    
    if args.check:
        results = check_docx(output, config=config)
        if args.json: print(json.dumps(results, ensure_ascii=False, indent=2))
        else: print_check_report(results)
    sys.exit(0)

if __name__ == '__main__': main()
