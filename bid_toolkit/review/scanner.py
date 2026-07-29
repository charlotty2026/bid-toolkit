"""
审标扫描器 v1.1 — Layer 1 + Layer 2
=====================================
Layer 1: 判词库逐行扫描 → 标记风险位置
Layer 2: 规则+LLM 上下文判断 → 区分致命/警告/信息

v1.1 改进:
- 排除模式：过滤目录行、章节标题行、页码行等噪音
- 命中合并：同段落(连续3行内)同分类关键词合并为一个风险点
- Layer 2 规则增强：精确化评分/要求/否决的判断逻辑
"""
import os, sys, re, json
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class ScanHit:
    category: str       # primary / secondary / contract / customization / certification / time_nodes / emphasis / bid_types
    keyword: str        # 命中的词
    match_type: str     # literal / regex / emphasis
    line_num: int       # 行号（1-indexed）
    context: str        # 命中位置前后50字
    page_num: int = 0   # 页码（可选）
    llm_label: str = "" # LLM判断结果: fatal / warn / info（Layer 2 填充）
    llm_reason: str = ""


@dataclass
class ScanResult:
    file_path: str
    total_chars: int
    total_lines: int
    hits: list = field(default_factory=list)
    emphasis_marks_found: dict = field(default_factory=dict)
    excluded_lines: int = 0  # 被排除模式过滤的行数

    @property
    def fatals(self):
        """致命项（显式标记为fatal的）"""
        return [h for h in self.hits if h.llm_label == 'fatal']

    @property
    def warnings(self):
        """警告项"""
        return [h for h in self.hits if h.llm_label == 'warn']

    @property
    def info_items(self):
        """信息项"""
        return [h for h in self.hits if h.llm_label == 'info']


# ========== 排除模式（过滤噪音） ==========

EXCLUDE_PATTERNS = [
    # 目录行
    r'^[　 \t]*目[ 　]*录[ 　]*$',
    r'^[　 \t]*目录[ 　]*CONTENTS',
    # 章节标题 — 第X章 / 第X部分 / 第X节
    r'^第[一二三四五六七八九十百千]+[章节部分]',
    r'^第\d+[章节部分]',
    # 页码 / 页眉页脚
    r'^\d+[－-]\d+$',
    r'^\d+/\d+$',
    r'^[-—]\s*\d+\s*[-—]$',
    r'^第\s*\d+\s*页$',
    r'^-\s*\d+\s*-$',
    # 纯数字行（页码）
    r'^\d+$',
    # 表头分隔行
    r'^[—－=]{5,}$',
    r'^[|┃][—－= ]+[|┃]',
    # 文件署名
    r'^招标[人单]?[：:].*$',
    r'^采购[人代][：:].*$',
    r'^代理机构[：:].*$',
    # 纯标点/符号行
    r'^[、，。．·…—－\s]{3,}$',
]

EXCLUDE_COMPILED = [re.compile(p) for p in EXCLUDE_PATTERNS]


def is_excluded_line(line):
    """判断是否为噪音行（目录/章节标题/页码等）"""
    for pat in EXCLUDE_COMPILED:
        if pat.match(line.strip()):
            return True
    # 长度过滤：少于3个完整中文字符的行（数字+单位的纯技术行保留）
    stripped = line.strip()
    if len(stripped) < 4:
        cn_chars = len(re.findall(r'[\u4e00-\u9fff]', stripped))
        if cn_chars < 2:
            return True
    return False


# ========== 词库加载 ==========

_KEYWORDS_CACHE = None

def load_keywords(path=None):
    global _KEYWORDS_CACHE
    if _KEYWORDS_CACHE is not None:
        return _KEYWORDS_CACHE

    if path is None:
        path = Path(__file__).parent / 'keywords_risk.json'
    with open(path, 'r', encoding='utf-8') as f:
        _KEYWORDS_CACHE = json.load(f)
    return _KEYWORDS_CACHE


# ========== 文本提取 ==========

def extract_text(file_path):
    """从 PDF/DOCX/MD/TXT 提取纯文本（带行号）"""
    path = Path(file_path)
    suffix = path.suffix.lower()

    if suffix == '.pdf':
        return _extract_pdf(path)
    elif suffix == '.docx':
        return _extract_docx(path)
    elif suffix in ('.md', '.txt', '.json', '.yaml', '.yml'):
        return _extract_text_file(path)
    else:
        return _extract_text_file(path)


def _extract_pdf(path):
    try:
        import pdfplumber
        lines = []
        with pdfplumber.open(path) as pdf:
            for page_idx, page in enumerate(pdf.pages, 1):
                text = page.extract_text() or ''
                for line in text.split('\n'):
                    stripped = line.strip()
                    if stripped and not is_excluded_line(stripped):
                        lines.append({'line': stripped, 'page': page_idx})
        # PyMuPDF 补充（pdfplumber 可能丢字）
        try:
            import fitz
            doc = fitz.open(str(path))
            fitz_lines = []
            for page_idx, page in enumerate(doc, 1):
                text = page.get_text() or ''
                for line in text.split('\n'):
                    stripped = line.strip()
                    if stripped and not is_excluded_line(stripped):
                        fitz_lines.append({'line': stripped, 'page': page_idx})
            if len(fitz_lines) > len(lines):
                lines = fitz_lines
            doc.close()
        except ImportError:
            pass
        return lines
    except ImportError:
        print("❌ 需要 pdfplumber: pip install pdfplumber")
        sys.exit(1)


def _extract_docx(path):
    try:
        from docx import Document
        doc = Document(str(path))
        lines = []
        page = 1
        for para in doc.paragraphs:
            text = para.text.strip()
            if text and not is_excluded_line(text):
                lines.append({'line': text, 'page': page})
            if para.runs and any(r.text and '\x0c' in r.text for r in para.runs):
                page += 1
        return lines
    except ImportError:
        print("❌ 需要 python-docx: pip install python-docx")
        sys.exit(1)


def _extract_text_file(path):
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    lines = []
    for line in content.split('\n'):
        stripped = line.strip()
        if stripped and not is_excluded_line(stripped):
            lines.append({'line': stripped, 'page': 0})
    return lines


# ========== Layer 1: 判词扫描 ==========

def _scan_literal(lines, keywords, category_id):
    """精确匹配扫描"""
    hits = []
    for idx, entry in enumerate(lines):
        line = entry['line']
        page = entry['page']
        for kw in keywords:
            if kw in line:
                pos = line.index(kw)
                start = max(0, pos - 50)
                end = min(len(line), pos + len(kw) + 50)
                context = ('...' if start > 0 else '') + line[start:end] + ('...' if end < len(line) else '')
                hits.append(ScanHit(
                    category=category_id,
                    keyword=kw,
                    match_type='literal',
                    line_num=idx + 1,
                    context=context,
                    page_num=page,
                ))
    return hits


def _scan_regex(lines, patterns, category_id):
    """正则匹配扫描"""
    hits = []
    for idx, entry in enumerate(lines):
        line = entry['line']
        page = entry['page']
        for pat in patterns:
            try:
                m = re.search(pat, line)
                if m:
                    matched = m.group()
                    start = max(0, m.start() - 50)
                    end = min(len(line), m.end() + 50)
                    context = ('...' if start > 0 else '') + line[start:end] + ('...' if end < len(line) else '')
                    hits.append(ScanHit(
                        category=category_id,
                        keyword=matched,
                        match_type='regex',
                        line_num=idx + 1,
                        context=context,
                        page_num=page,
                    ))
            except re.error:
                pass  # 跳过无效正则
    return hits


def _scan_emphasis(lines, marks, min_occurrences=3):
    """强调标识扫描"""
    mark_count = {}
    for entry in lines:
        line = entry['line']
        for m in marks:
            count = line.count(m)
            if count > 0:
                mark_count[m] = mark_count.get(m, 0) + count

    active_marks = {m for m, c in mark_count.items() if c >= min_occurrences}
    # ▲和★在低阈值也保留
    for m in ('▲', '★'):
        if m in mark_count and mark_count[m] >= 1:
            active_marks.add(m)

    hits = []
    for idx, entry in enumerate(lines):
        line = entry['line']
        page = entry['page']
        for m in active_marks:
            if m in line:
                pos = line.index(m)
                if pos < 10:  # 行首附近
                    start = max(0, pos - 30)
                    end = min(len(line), pos + len(m) + 80)
                    context = ('...' if start > 0 else '') + line[start:end] + ('...' if end < len(line) else '')
                    hits.append(ScanHit(
                        category='emphasis_marks',
                        keyword=f'{m} 标识参数',
                        match_type='emphasis',
                        line_num=idx + 1,
                        context=context,
                        page_num=page,
                    ))
    return hits, active_marks


# ========== 命中合并（同段落合并） ==========

def _merge_adjacent_hits(hits, max_gap=3):
    """同段落(连续max_gap行内)命中同一分类的关键词合并为一个风险点"""
    if not hits:
        return []

    # 按行号排序
    sorted_hits = sorted(hits, key=lambda h: (h.line_num, h.keyword))
    merged = []

    current_group = [sorted_hits[0]]

    for h in sorted_hits[1:]:
        prev = current_group[-1]
        # 同分类 + 行号差距 <= max_gap
        if h.category == prev.category and (h.line_num - prev.line_num) <= max_gap:
            current_group.append(h)
        else:
            # 合并当前组
            merged.append(_merge_group(current_group))
            current_group = [h]

    if current_group:
        merged.append(_merge_group(current_group))

    return merged


def _merge_group(group):
    """合并一组同分类相邻命中"""
    if len(group) == 1:
        return group[0]

    # 取第1个作为主条目
    first = group[0]
    keywords = list(dict.fromkeys(h.keyword for h in group))  # 去重有序

    # 合并关键词：取最多3个代表性词
    if len(keywords) <= 3:
        kw_display = ' / '.join(keywords)
    else:
        kw_display = ' / '.join(keywords[:3]) + f' 等{len(keywords)}个'

    # 合并行号范围
    line_start = min(h.line_num for h in group)
    line_end = max(h.line_num for h in group)
    line_display = f'{line_start}–{line_end}' if line_end > line_start else str(line_start)

    # 取最长的上下文
    best_context = max((h.context for h in group), key=len)

    # 合并LLM标签：取最严重的
    labels = [h.llm_label for h in group if h.llm_label]
    best_label = 'fatal' if 'fatal' in labels else ('warn' if 'warn' in labels else first.llm_label)
    reasons = [h.llm_reason for h in group if h.llm_reason]

    return ScanHit(
        category=first.category,
        keyword=kw_display,
        match_type=first.match_type,
        line_num=line_start,
        context=f'[段落 {line_display}] ' + (reasons[0] if reasons else '') + ' | ' + best_context[:80],
        page_num=first.page_num,
        llm_label=best_label,
        llm_reason='; '.join(reasons[:3]) if reasons else '',
    )


# ========== Layer 2: 上下文判断（规则+可选LLM） ==========

def _llm_judge_rules(result):
    """规则引擎判断 — 对所有命中做三级分类"""
    if not result.hits:
        return

    for h in result.hits:
        ctx = h.context
        ctx_lower = ctx.lower()

        # ===== 分类无关的通用规则 =====
        # 评分类内容：出现在评分标准/细则中的，降级为warn或info
        if any(w in ctx_lower for w in ['评分标准', '评分细则', '评分项', '评审标准', '分值', '满分', '得分']):
            if any(w in ctx_lower for w in ['否则', '视为', '予以', '按废标', '按无效', '否决']):
                h.llm_label = 'fatal'
                h.llm_reason = '评分标准中的否决条款'
            else:
                h.llm_label = 'warn' if any(w in ctx_lower for w in ['不得分', '0分', '零分']) else 'info'
                h.llm_reason = '出现在评分标准中' if h.llm_label != 'warn' else '评分失分项'
            continue

        # ===== 按分类判断 =====
        if h.category in ('primary', 'bid_types'):
            # 一级判决词 + 行业特有
            if any(w in ctx for w in ['必须', '应当', '须提供']) and \
               not any(w in ctx for w in ['否则', '视为', '按废标处理', '按无效处理', '按否决处理', '予以否决', '拒绝', '将']):
                h.llm_label = 'warn'
                h.llm_reason = '属必要条件但未带否决后果'
            elif h.keyword in ('不接受联合体', '不接受联合体投标'):
                h.llm_label = 'info'
                h.llm_reason = '联合体限制，独立投标无影响'
            elif '投标有效期' in h.keyword:
                h.llm_label = 'info'
                h.llm_reason = '时间约束，非否决'
            else:
                h.llm_label = 'fatal'
                h.llm_reason = '直接否决/废标风险'

        elif h.category in ('secondary',):
            # 二级判决：评分/格式/保证金等
            if any(w in h.keyword for w in ['不得分', '0分', '零分', '没收', '不予退还']):
                h.llm_label = 'fatal'
                h.llm_reason = '直接影响得分或保证金'
            elif any(w in h.keyword for w in ['必须', '应当', '密封', '法定代表', '盖章']):
                h.llm_label = 'warn'
                h.llm_reason = '格式/签字盖章类要求'
            else:
                h.llm_label = 'info'
                h.llm_reason = '一般性提示'

        elif h.category in ('contract',):
            h.llm_label = 'info'
            h.llm_reason = '合同履行期约束，不影响投标有效性'

        elif h.category in ('certifications', 'time_nodes'):
            h.llm_label = 'info'
            h.llm_reason = '证明材料/时间节点提示'

        elif h.category in ('customization',):
            h.llm_label = 'warn'
            h.llm_reason = '商务门槛，需提前确认能否满足'

        elif h.category in ('emphasis_marks',):
            # 强调标识 — 如果上下文是参考/详见/例如，降为warn
            if any(w in ctx for w in ['参考', '详见', '例如', '如']):
                h.llm_label = 'warn'
                h.llm_reason = '强调标识出现在参考/示例说明中，非实质性响应'
            else:
                h.llm_label = 'fatal'
                h.llm_reason = '实质性响应参数，不响应即否决'

        else:
            # 兜底
            h.llm_label = 'warn'
            h.llm_reason = '未分类风险项'


def _llm_judge_api(result, llm_client):
    """调用真实LLM API做判断（需传入已初始化的llm_client）"""
    to_judge = [h for h in result.hits if h.category in ('primary', 'bid_types')]
    if not to_judge:
        return

    # 先跑规则兜底
    _llm_judge_rules(result)

    # 批量构建LLM判断请求
    print(f"🧠 LLM API 判断: {len(to_judge)} 个命中...")
    batch_size = 10
    for start_idx in range(0, len(to_judge), batch_size):
        batch = to_judge[start_idx:start_idx + batch_size]
        prompt = "判断以下每条上下文中的风险是「fatal」(直接废标/否决)还是「warn」(条件性风险/需注意)还是「info」(无关紧要的引用说明)。只输出json数组。\n\n"
        for i, h in enumerate(batch):
            prompt += f'{i}: keyword="{h.keyword}" context="{h.context[:80]}"\n'

        try:
            resp = llm_client.chat(prompt, temperature=0.1)
            # 解析响应（简单找JSON数组）
            import ast
            resp = resp.strip()
            if resp.startswith('```'):
                resp = resp.split('\n', 1)[1].rsplit('```', 1)[0]
            try:
                labels = json.loads(resp)
                if isinstance(labels, list) and len(labels) == len(batch):
                    for h, label in zip(batch, labels):
                        if isinstance(label, dict):
                            h.llm_label = label.get('label', h.llm_label)
                            h.llm_reason = label.get('reason', h.llm_reason)
            except (json.JSONDecodeError, TypeError):
                pass  # LLM返回格式不对，用规则结果兜底
        except Exception:
            pass  # LLM调用失败，用规则结果兜底


# ========== 主扫描入口 ==========

def scan_tender(file_path, keywords_path=None, with_llm=False, llm_client=None):
    """执行全量扫描（Layer 1 + Layer 2）"""
    kw = load_keywords(keywords_path)

    # 提取文本（排除模式已内置）
    lines = extract_text(file_path)
    if not lines:
        print(f"⚠️  未从文件提取到任何文本: {file_path}")
        return ScanResult(file_path=file_path, total_chars=0, total_lines=0)

    total_chars = sum(len(e['line']) for e in lines)
    result = ScanResult(
        file_path=file_path,
        total_chars=total_chars,
        total_lines=len(lines),
    )

    # 逐分类扫描
    for cat in kw['categories']:
        cid = cat['id']
        match_type = cat.get('match', 'literal')

        if cid == 'emphasis_marks':
            mark_hits, active_marks = _scan_emphasis(
                lines,
                cat.get('marks', []),
                cat.get('min_occurrences', 3),
            )
            result.hits.extend(mark_hits)
            result.emphasis_marks_found = {m: True for m in active_marks}

        elif match_type == 'regex':
            pat_hits = _scan_regex(lines, cat.get('patterns', []), cid)
            result.hits.extend(pat_hits)

        else:  # literal
            lit_hits = _scan_literal(lines, cat.get('words', []), cid)
            result.hits.extend(lit_hits)

    # 去重（同行同词只保留一个）
    seen = set()
    deduped = []
    for h in result.hits:
        key = (h.line_num, h.keyword)
        if key not in seen:
            seen.add(key)
            deduped.append(h)
    result.hits = deduped

    # 同段落合并
    result.hits = _merge_adjacent_hits(result.hits, max_gap=3)

    # 按行号排序
    result.hits.sort(key=lambda h: (h.line_num,))

    # Layer 2: 规则判断（必需）
    _llm_judge_rules(result)

    # Layer 2: LLM API 判断（可选）
    if with_llm and llm_client:
        _llm_judge_api(result, llm_client)
    elif with_llm:
        print("  ⚠️  未传入 llm_client，使用规则引擎判断")

    return result
