#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
投标文件智能生成器 v3.5
======================
从招标文件到投标文件初稿的一键生成。

核心流程：
  1. 调用 parse_bid.py 解析招标文件，提取结构化信息
  2. 根据 bid_type_detection.yaml 自动识别标书类型（货物/服务/工程）
  3. 加载对应大纲模板和企业素材
  4. 商务标部分用模板占位（不调用LLM）
  5. 技术标部分逐章节调用LLM生成内容
  6. 组装为完整Markdown，可选转Word
  7. 输出 needs_manual 清单（占位符 + 待确认项）

不是替代人工，是生成80%骨架让投标人专注于20%的关键内容。

用法：
  # 全自动：解析→识别类型→生成投标文件
  python scripts/bid_generator.py generate --rfp XX项目招标文件.pdf --company company_profile/ --output 投标文件.md

  # 指定类型
  python scripts/bid_generator.py generate --rfp XX项目招标文件.pdf --type services --company company_profile/ --output 投标文件.md

  # 生成后直接转Word
  python scripts/bid_generator.py generate --rfp XX项目招标文件.pdf --company company_profile/ --output 投标文件.md --docx

  # 只解析不生成（查看招标文件拆解结果）
  python scripts/bid_generator.py parse --rfp XX项目招标文件.pdf

License: MIT
"""

import os
import sys
import re
import json
import argparse
import logging
import hashlib
import time
from pathlib import Path
from datetime import datetime
from typing import Optional

# ===== 路径设置 =====
SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
TEMPLATES_DIR = PROJECT_ROOT / 'templates'

sys.path.insert(0, str(SCRIPT_DIR))
sys.stdout.reconfigure(encoding='utf-8')

# ===== 依赖导入 =====
try:
    import yaml
except ImportError:
    yaml = None
    print('⚠️  未安装pyyaml，将无法加载yaml配置。安装: pip install pyyaml', file=sys.stderr)

# 导入项目已有模块
from llm_client import LLMClient, load_user_config
from parse_bid import parse_bid_document

# ===== 日志 =====
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S',
)
logger = logging.getLogger('bid_generator')

# ===== 常量 =====

BID_TYPE_NAMES: dict[str, str] = {
    'goods': '货物类',
    'service': '服务类',
    'engineering': '工程类',
}

# 大纲模板文件名映射
OUTLINE_FILES: dict[str, str] = {
    'goods': 'goods_outline.md',
    'service': 'service_outline.md',
    'engineering': 'engineering_outline.md',
}

# 格式模板映射（用于bid_engine的template_name参数）
FORMAT_TEMPLATES: dict[str, str] = {
    'goods': 'goods',
    'service': 'government',
    'engineering': 'engineering',
}

# 企业素材文件列表
COMPANY_PROFILE_FILES: list[str] = [
    'company_info.md',
    'qualifications.md',
    'team.md',
    'performance.md',
    'pitfalls.md',
]

# 商务标关键词（用于判断大纲章节属于商务标还是技术标）
COMMERCIAL_KEYWORDS: list[str] = [
    '投标函', '开标一览表', '报价明细', '报价表', '中小企业声明函',
    '投标保证金', '授权委托书', '营业执照', '资质证书', '保证金',
]

# 系统提示词
SYSTEM_PROMPT: str = (
    '你是专业的投标文件撰写专家，严格按照招标文件要求和大纲结构生成内容。'
    '生成规则：\n'
    '1. 严格按照大纲结构生成，不遗漏章节\n'
    '2. 遇到需要填写具体数据（如人数、金额、时间）的位置，使用 {占位符} 标记\n'
    '3. 不得编造虚假业绩或资质\n'
    '4. 语言正式、专业，使用书面语\n'
    '5. 如招标文件有明确要求（如响应时间、人员数量），必须逐字引用并响应\n'
    '6. 不要出现"保证中标""100%成功率"等绝对化表述\n'
    '7. 输出纯Markdown格式，包含标题层级和表格'
)


# =====================================================================
# 标书类型识别
# =====================================================================

def detect_bid_type(rfp_text: str, detection_yaml_path: Path) -> tuple[str, dict]:
    """根据招标文件文本自动识别标书类型（货物/服务/工程）。

    使用 bid_type_detection.yaml 中的关键词规则进行评分：
    - 强信号关键词：每个计2分
    - 中信号关键词：每个计1分
    - 达到 confidence_threshold 才定性

    参数：
        rfp_text: 招标文件全文
        detection_yaml_path: bid_type_detection.yaml 路径

    返回：
        (类型代码, 识别详情字典)
        类型代码: 'goods' / 'service' / 'engineering'
        识别详情: {'type': str, 'name': str, 'scores': dict, 'details': dict}
    """
    if yaml is None or not detection_yaml_path.exists():
        logger.warning('类型识别规则文件不存在或yaml未安装，默认使用服务类')
        return 'service', {'type': 'service', 'name': '服务类', 'scores': {}, 'details': {'fallback': True}}

    with open(detection_yaml_path, 'r', encoding='utf-8') as f:
        detection_data = yaml.safe_load(f) or {}

    rules = detection_data.get('detection_rules', {})
    scores: dict[str, int] = {}
    details: dict[str, dict] = {}

    for type_key, rule in rules.items():
        score = 0
        matched_strong: list[str] = []
        matched_medium: list[str] = []

        keywords = rule.get('keywords', {})
        strong_kw = keywords.get('strong', [])
        medium_kw = keywords.get('medium', [])

        for kw in strong_kw:
            count = rfp_text.count(kw)
            if count > 0:
                score += 2 * count
                matched_strong.append(f'{kw}({count})')

        for kw in medium_kw:
            count = rfp_text.count(kw)
            if count > 0:
                score += 1 * count
                matched_medium.append(f'{kw}({count})')

        # 排除条件检查
        exclude_kw = rule.get('exclude_if_dominant', [])
        exclude_hits = sum(rfp_text.count(kw) for kw in exclude_kw)
        if exclude_hits > 5:
            score = max(0, score - exclude_hits)

        threshold = rule.get('confidence_threshold', 3)
        scores[type_key] = score
        details[type_key] = {
            'score': score,
            'threshold': threshold,
            'matched_strong': matched_strong,
            'matched_medium': matched_medium,
            'passed': score >= threshold,
        }

    # 选最高分且达标的类型
    best_type = 'service'  # 默认
    best_score = 0
    for type_key, detail in details.items():
        if detail['passed'] and detail['score'] > best_score:
            best_score = detail['score']
            best_type = type_key

    # 如果都没有达标，选最高分的
    if best_score == 0:
        best_type = max(scores, key=lambda k: scores.get(k, 0)) if scores else 'service'

    return best_type, {
        'type': best_type,
        'name': BID_TYPE_NAMES.get(best_type, '未分类'),
        'scores': scores,
        'details': details,
    }


# =====================================================================
# 大纲模板解析
# =====================================================================

def parse_outline(outline_path: Path) -> list[dict]:
    """将大纲模板Markdown解析为章节列表。

    大纲模板结构：
      # 注释行（含"第X部分"分隔标记）
      ## 章节标题
      章节内容...
      ## 章节标题
      章节内容...

    返回：
        [{'index': int, 'title': str, 'content': str, 'part': str, 'is_commercial': bool}, ...]
        part: '商务标' / '技术标' / '其他'
    """
    if not outline_path.exists():
        logger.error(f'大纲模板不存在: {outline_path}')
        return []

    with open(outline_path, 'r', encoding='utf-8') as f:
        text = f.read()

    lines = text.split('\n')
    sections: list[dict] = []
    current_part = '商务标'  # 默认从商务标开始
    current_title = ''
    current_lines: list[str] = []
    section_index = 0

    for line in lines:
        stripped = line.strip()

        # 检测"第X部分"标记（注释行）
        if stripped.startswith('#') and not stripped.startswith('##'):
            if '第一部分' in stripped or '商务标' in stripped:
                current_part = '商务标'
            elif '第二部分' in stripped or '技术标' in stripped:
                current_part = '技术标'
            elif '第三部分' in stripped or '其他' in stripped:
                current_part = '其他'
            continue

        # 检测 ## 级别标题（章节）
        if stripped.startswith('## ') and not stripped.startswith('### '):
            # 保存前一个章节
            if current_title:
                is_commercial = _is_commercial_section(current_title, current_part)
                sections.append({
                    'index': section_index,
                    'title': current_title,
                    'content': '\n'.join(current_lines).strip(),
                    'part': current_part,
                    'is_commercial': is_commercial,
                })
                section_index += 1

            current_title = stripped[3:].strip()
            current_lines = []
        else:
            if current_title:
                current_lines.append(line)

    # 保存最后一个章节
    if current_title:
        is_commercial = _is_commercial_section(current_title, current_part)
        sections.append({
            'index': section_index,
            'title': current_title,
            'content': '\n'.join(current_lines).strip(),
            'part': current_part,
            'is_commercial': is_commercial,
        })
        section_index += 1

    return sections


def _is_commercial_section(title: str, part: str) -> bool:
    """判断章节是否属于商务标（不调用LLM，直接用模板）。

    参数：
        title: 章节标题
        part: 所属部分（商务标/技术标/其他）

    返回：
        True 表示该章节使用模板占位，不调用LLM
    """
    if part == '商务标':
        return True
    if part == '其他':
        return True
    # 技术标中的业绩表也用模板（需要用户手动填写真实业绩）
    if '业绩' in title:
        return True
    return False


# =====================================================================
# 企业素材加载
# =====================================================================

def load_company_profiles(company_dir: Path) -> dict[str, str]:
    """加载企业素材目录下的所有Markdown文件。

    参数：
        company_dir: company_profile/ 目录路径

    返回：
        {'company_info': str, 'qualifications': str, 'team': str, 'performance': str, 'pitfalls': str}
        文件不存在时值为空字符串
    """
    profiles: dict[str, str] = {}

    for filename in COMPANY_PROFILE_FILES:
        key = filename.replace('.md', '')
        filepath = company_dir / filename
        if filepath.exists():
            try:
                content = filepath.read_text(encoding='utf-8')
                # 去掉HTML注释块（模板说明）
                content = re.sub(r'<!--[\s\S]*?-->', '', content).strip()
                profiles[key] = content
            except Exception as e:
                logger.warning(f'读取企业素材失败 {filename}: {e}')
                profiles[key] = ''
        else:
            logger.warning(f'企业素材文件不存在: {filepath}')
            profiles[key] = ''

    return profiles


# =====================================================================
# RFP摘要构建
# =====================================================================

def build_rfp_summary(rfp_data: dict) -> str:
    """从解析后的招标文件数据构建简洁摘要，用于LLM上下文。

    参数：
        rfp_data: parse_bid_document() 返回的字典

    返回：
        招标文件要求摘要文本（用于LLM prompt）
    """
    parts: list[str] = []

    # 基本信息
    parts.append(f'招标文件：{rfp_data.get("文件名", "未知")}')

    # 预算
    budgets = rfp_data.get('预算', [])
    if budgets:
        parts.append(f'项目预算：{", ".join(budgets)}')

    # 时间节点
    timelines = rfp_data.get('时间节点', [])
    if timelines:
        timeline_str = '; '.join(f'{t["事件"]}: {t["时间"]}' for t in timelines)
        parts.append(f'时间节点：{timeline_str}')

    # 资质要求
    quals = rfp_data.get('资质要求', [])
    if quals:
        parts.append('资质要求：')
        for i, q in enumerate(quals, 1):
            parts.append(f'  ({i}) {q}')

    # 废标红线
    disqual = rfp_data.get('废标红线', [])
    if disqual:
        parts.append('⚠️ 废标红线（必须遵守，否则投标无效）：')
        for i, rule in enumerate(disqual, 1):
            parts.append(f'  ({i}) {rule}')

    # 文件清单
    checklist = rfp_data.get('文件清单', [])
    if checklist:
        parts.append('投标文件清单：')
        for i, item in enumerate(checklist, 1):
            parts.append(f'  {i}. {item}')

    # 评分项
    outline = rfp_data.get('大纲框架', {})
    scoring_items = outline.get('来源_评分项', [])
    if scoring_items:
        parts.append('评分项：')
        for item in scoring_items:
            parts.append(f'  - {item.get("项目", "")}：{item.get("分值", "")}分 {item.get("评分标准", "")}')
        total = outline.get('评分总分', '')
        if total:
            parts.append(f'  评分总分：{total}分')

    # 保证金
    deposit = rfp_data.get('保证金', {})
    if deposit:
        deposit_str = '; '.join(f'{k}: {v}' for k, v in deposit.items())
        parts.append(f'保证金：{deposit_str}')

    # 格式要求
    fmt_req = rfp_data.get('格式要求', {})
    if fmt_req:
        fmt_parts = []
        if '行距' in fmt_req:
            fmt_parts.append(f'行距{fmt_req["行距"]}')
        if '字体字号' in fmt_req:
            for item in fmt_req['字体字号']:
                fmt_parts.append(f'{item.get("对象", "")}用{item.get("字体", "")}{item.get("字号", "")}')
        if fmt_parts:
            parts.append(f'格式要求：{", ".join(fmt_parts)}')

    return '\n'.join(parts)


# =====================================================================
# 章节内容生成
# =====================================================================

def generate_commercial_section(section: dict, rfp_summary: str, company_profiles: dict) -> str:
    """生成商务标章节（直接使用模板内容，不调用LLM）。

    商务标部分（投标函/开标一览表/报价表等）有固定格式，
    需要按招标文件给定格式填写，不适合AI生成，直接保留模板占位。

    参数：
        section: 章节字典 {'title', 'content', 'part', ...}
        rfp_summary: 招标文件摘要
        company_profiles: 企业素材

    返回：
        章节Markdown文本
    """
    title = section['title']
    content = section['content']

    # 对于资质证书章节，尝试注入企业资质信息
    if '资质证书' in title and company_profiles.get('qualifications'):
        content = content + '\n\n' + company_profiles['qualifications']

    # 对于营业执照章节，注入企业基本信息
    if '营业执照' in title and company_profiles.get('company_info'):
        content = content + '\n\n' + company_profiles['company_info']

    # 对于业绩章节，注入企业业绩
    if '业绩' in title and company_profiles.get('performance'):
        content = content + '\n\n' + company_profiles['performance']

    return f'## {title}\n\n{content}'


def generate_technical_section(
    section: dict,
    rfp_summary: str,
    company_profiles: dict,
    llm_client: LLMClient,
    max_retries: int = 3,
) -> tuple[str, bool]:
    """使用LLM生成技术标章节内容。

    每个章节独立生成，失败重试max_retries次。
    生成失败时回退到模板内容。

    参数：
        section: 章节字典
        rfp_summary: 招标文件摘要
        company_profiles: 企业素材
        llm_client: LLM客户端
        max_retries: 最大重试次数

    返回：
        (章节Markdown文本, 是否成功用LLM生成)
    """
    title = section['title']
    outline_content = section['content']

    # 构建企业素材摘要（截取关键信息，控制token）
    company_brief = _build_company_brief(company_profiles)

    # 构建用户提示词
    user_prompt = (
        f'## 招标文件要求摘要\n{rfp_summary}\n\n'
        f'## 企业素材\n{company_brief}\n\n'
        f'## 当前章节大纲\n'
        f'章节标题：{title}\n'
        f'大纲要求：\n{outline_content}\n\n'
        f'## 生成指令\n'
        f'请根据以上信息，为「{title}」章节生成投标文件内容。\n'
        f'要求：\n'
        f'1. 严格按大纲结构生成，包含所有子标题\n'
        f'2. 需要填写具体数据的地方用 {{占位符}} 标记\n'
        f'3. 内容要具体、专业，不要写空话套话\n'
        f'4. 如果招标文件有具体要求（如人员数量、响应时间），必须引用并响应\n'
        f'5. 输出以「## {title}」开头'
    )

    # 重试生成
    for attempt in range(1, max_retries + 1):
        try:
            logger.info(f'  生成章节 [{title}] (第{attempt}/{max_retries}次)')
            response = llm_client.chat_simple(
                system_prompt=SYSTEM_PROMPT,
                user_prompt=user_prompt,
                temperature=0.5,
                max_tokens=4096,
            )

            if response and len(response.strip()) > 50:
                # 确保以正确的标题开头
                cleaned = response.strip()
                if not cleaned.startswith(f'## {title}'):
                    # 尝试找到标题行并修正
                    if f'## {title}' in cleaned:
                        idx = cleaned.index(f'## {title}')
                        cleaned = cleaned[idx:]
                    else:
                        cleaned = f'## {title}\n\n' + cleaned
                return cleaned, True

            logger.warning(f'  章节生成内容过短，重试: [{title}]')

        except Exception as e:
            logger.error(f'  章节生成异常 [{title}] 第{attempt}次: {e}')

        if attempt < max_retries:
            wait = 2 ** (attempt - 1)
            logger.info(f'  等待{wait}秒后重试...')
            time.sleep(wait)

    # 全部重试失败，回退到模板
    logger.warning(f'  ⚠️ 章节LLM生成失败，使用模板: [{title}]')
    return f'## {title}\n\n{outline_content}', False


def _build_company_brief(company_profiles: dict) -> str:
    """构建企业素材摘要（用于LLM上下文，控制长度）。

    参数：
        company_profiles: 企业素材字典

    返回：
        企业素材摘要文本（截断到合理长度）
    """
    parts: list[str] = []

    if company_profiles.get('company_info'):
        parts.append(f'### 企业基本信息\n{company_profiles["company_info"][:800]}')

    if company_profiles.get('team'):
        parts.append(f'### 团队信息\n{company_profiles["team"][:600]}')

    if company_profiles.get('performance'):
        parts.append(f'### 业绩案例\n{company_profiles["performance"][:600]}')

    if company_profiles.get('qualifications'):
        parts.append(f'### 资质证书\n{company_profiles["qualifications"][:400]}')

    if company_profiles.get('pitfalls'):
        parts.append(f'### ⚠️ 注意事项（历史踩坑，生成时必须避免）\n{company_profiles["pitfalls"][:400]}')

    return '\n\n'.join(parts) if parts else '（企业素材未填写）'


# =====================================================================
# 断点续传
# =====================================================================

def get_cache_dir(output_path: Path) -> Path:
    """获取断点续传缓存目录。

    缓存目录位于输出文件同级的 .bid_gen_cache/ 子目录下。

    参数：
        output_path: 输出文件路径

    返回：
        缓存目录Path对象
    """
    cache_dir = output_path.parent / '.bid_gen_cache'
    cache_dir.mkdir(parents=True, exist_ok=True)
    return cache_dir


def get_cache_key(output_path: Path) -> str:
    """根据输出文件路径生成缓存键。

    参数：
        output_path: 输出文件路径

    返回：
        缓存键字符串（用于区分不同项目的缓存）
    """
    path_str = str(output_path.resolve())
    return hashlib.md5(path_str.encode()).hexdigest()[:8]


def load_checkpoint(cache_dir: Path, cache_key: str) -> dict[int, str]:
    """加载断点续传数据。

    参数：
        cache_dir: 缓存目录
        cache_key: 缓存键

    返回：
        {section_index: content} 字典，已生成的章节
    """
    checkpoint_file = cache_dir / f'checkpoint_{cache_key}.json'
    if not checkpoint_file.exists():
        return {}

    try:
        with open(checkpoint_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        logger.info(f'📂 发现断点续传数据：{len(data)} 个章节已生成')
        return {int(k): v for k, v in data.items()}
    except Exception as e:
        logger.warning(f'加载断点续传数据失败: {e}')
        return {}


def save_checkpoint(cache_dir: Path, cache_key: str, checkpoint: dict[int, str]) -> None:
    """保存断点续传数据。

    参数：
        cache_dir: 缓存目录
        cache_key: 缓存键
        checkpoint: {section_index: content} 字典
    """
    checkpoint_file = cache_dir / f'checkpoint_{cache_key}.json'
    try:
        with open(checkpoint_file, 'w', encoding='utf-8') as f:
            json.dump(checkpoint, f, ensure_ascii=False, indent=2)
    except Exception as e:
        logger.warning(f'保存断点续传数据失败: {e}')


def clear_checkpoint(cache_dir: Path, cache_key: str) -> None:
    """生成完成后清除断点续传数据。

    参数：
        cache_dir: 缓存目录
        cache_key: 缓存键
    """
    checkpoint_file = cache_dir / f'checkpoint_{cache_key}.json'
    try:
        if checkpoint_file.exists():
            checkpoint_file.unlink()
            logger.info('🧹 已清除断点续传缓存')
    except Exception as e:
        logger.warning(f'清除断点续传缓存失败: {e}')


# =====================================================================
# 占位符提取
# =====================================================================

def extract_placeholders(text: str) -> list[str]:
    """从文本中提取 {占位符} 模式。

    参数：
        text: 待提取的文本

    返回：
        占位符列表（去重，保持顺序）
    """
    pattern = r'\{([^}]+)\}'
    matches = re.findall(pattern, text)
    # 去重保持顺序
    seen: set[str] = set()
    unique: list[str] = []
    for m in matches:
        if m not in seen:
            seen.add(m)
            unique.append(m)
    return unique


# =====================================================================
# 文档头部构建
# =====================================================================

def build_header(bid_type: str, placeholder_count: int, confirm_count: int) -> str:
    """构建生成文档的头部注释。

    参数：
        bid_type: 标书类型代码
        placeholder_count: 占位符数量
        confirm_count: 待确认章节数量

    返回：
        Markdown注释头部
    """
    type_name = BID_TYPE_NAMES.get(bid_type, '未分类')
    now = datetime.now().strftime('%Y-%m-%d %H:%M')

    return (
        f'<!-- \n'
        f'bid-toolkit 自动生成 | 类型: {type_name} | 生成时间: {now}\n'
        f'⚠️ 这是AI生成的初稿，必须人工审核后使用\n'
        f'待填写占位符: {placeholder_count}个 | 待确认章节: {confirm_count}个\n'
        f'-->\n'
    )


# =====================================================================
# needs_manual 清单
# =====================================================================

def build_needs_manual_report(
    placeholders: list[str],
    failed_sections: list[str],
    template_sections: list[str],
) -> str:
    """构建 needs_manual 清单（需要手动处理的内容）。

    参数：
        placeholders: 所有占位符列表
        failed_sections: LLM生成失败的章节列表
        template_sections: 使用模板的商务标章节列表

    返回：
        needs_manual 清单文本
    """
    lines: list[str] = []
    lines.append('=' * 60)
    lines.append('📋 needs_manual 清单 — 以下内容需要手动处理')
    lines.append('=' * 60)

    # 占位符
    if placeholders:
        lines.append(f'\n🔹 待填写占位符（{len(placeholders)}个）：')
        for i, ph in enumerate(placeholders, 1):
            lines.append(f'  {i}. {{{ph}}}')
    else:
        lines.append('\n✅ 未发现占位符')

    # LLM生成失败的章节
    if failed_sections:
        lines.append(f'\n⚠️ LLM生成失败（已用模板回退，需重点检查）：')
        for i, title in enumerate(failed_sections, 1):
            lines.append(f'  {i}. {title}')
    else:
        lines.append('\n✅ 所有技术标章节均由LLM成功生成')

    # 商务标模板章节
    if template_sections:
        lines.append(f'\n📌 商务标模板章节（需按招标文件格式填写）：')
        for i, title in enumerate(template_sections, 1):
            lines.append(f'  {i}. {title}')

    lines.append('\n' + '=' * 60)
    lines.append('⚠️ 请逐项检查以上内容，确保投标文件完整、准确。')
    lines.append('=' * 60)

    return '\n'.join(lines)


# =====================================================================
# Word转换
# =====================================================================

def convert_to_docx(md_path: Path, bid_type: str, config_path: Optional[str] = None) -> Optional[str]:
    """将生成的Markdown转换为Word文档。

    调用 bid_engine.py 的 md_to_docx 函数完成转换。

    参数：
        md_path: Markdown文件路径
        bid_type: 标书类型（用于选择格式模板）
        config_path: 自定义配置文件路径

    返回：
        生成的docx文件路径，失败返回None
    """
    try:
        from bid_engine import md_to_docx, load_config

        # 选择格式模板
        template_name = FORMAT_TEMPLATES.get(bid_type, 'government')

        # 加载配置
        config = load_config(config_path=config_path, template_name=template_name)

        # 读取Markdown
        md_text = md_path.read_text(encoding='utf-8')

        # 输出路径
        docx_path = md_path.with_suffix('.docx')

        # 转换
        md_to_docx(md_text, str(docx_path), auto_fix=True, config=config)
        logger.info(f'📝 Word文档已生成: {docx_path}')
        return str(docx_path)

    except ImportError:
        logger.error('❌ 无法导入bid_engine模块，请确保bid_engine.py在同一目录')
        return None
    except Exception as e:
        logger.error(f'❌ Word转换失败: {e}')
        return None


# =====================================================================
# 主生成函数
# =====================================================================

def generate_bid_document(
    rfp_path: str,
    output_path: str,
    company_dir: str,
    bid_type: Optional[str] = None,
    docx: bool = False,
    config_path: Optional[str] = None,
) -> bool:
    """投标文件生成主函数。

    参数：
        rfp_path: 招标文件路径（PDF/MD/TXT）
        output_path: 输出投标文件路径（.md）
        company_dir: 企业素材目录路径
        bid_type: 指定标书类型（None=自动识别）
        docx: 是否同时生成Word文档
        config_path: 自定义配置文件路径

    返回：
        True=生成成功，False=生成失败
    """
    output_file = Path(output_path)
    company_path = Path(company_dir)

    # 验证输入
    if not Path(rfp_path).exists():
        logger.error(f'❌ 招标文件不存在: {rfp_path}')
        return False
    if not company_path.exists():
        logger.error(f'❌ 企业素材目录不存在: {company_dir}')
        return False

    # 加载用户配置
    user_config = load_user_config(config_path) if config_path else load_user_config()
    if not config_path:
        # 自动查找配置
        auto_config = TEMPLATES_DIR / 'user_config.yaml'
        if auto_config.exists():
            user_config = load_user_config(str(auto_config))

    # ===== 步骤1: 解析招标文件 =====
    logger.info('=' * 50)
    logger.info('📌 步骤1/7: 解析招标文件')
    logger.info('=' * 50)

    try:
        rfp_data = parse_bid_document(rfp_path)
    except SystemExit:
        logger.error('❌ 招标文件解析失败')
        return False
    except Exception as e:
        logger.error(f'❌ 招标文件解析异常: {e}')
        return False

    # 获取招标文件全文（用于类型识别）
    from parse_bid import read_input_file
    try:
        rfp_text = read_input_file(rfp_path)
    except SystemExit:
        rfp_text = ''
    except Exception:
        rfp_text = ''

    logger.info(f'  解析完成: {len(rfp_data.get("废标红线", []))} 条废标红线, '
                f'{len(rfp_data.get("资质要求", []))} 条资质要求')

    # ===== 步骤2: 识别标书类型 =====
    logger.info('=' * 50)
    logger.info('📌 步骤2/7: 识别标书类型')
    logger.info('=' * 50)

    detection_yaml = TEMPLATES_DIR / 'bid_type_detection.yaml'

    if bid_type:
        if bid_type not in BID_TYPE_NAMES:
            logger.error(f'❌ 不支持的标书类型: {bid_type}（可选: {", ".join(BID_TYPE_NAMES.keys())}）')
            return False
        logger.info(f'  用户指定类型: {BID_TYPE_NAMES[bid_type]}')
    else:
        bid_type, detection_info = detect_bid_type(rfp_text, detection_yaml)
        logger.info(f'  自动识别类型: {detection_info["name"]}')
        for type_key, detail in detection_info.get('details', {}).items():
            status = '✅' if detail['passed'] else '  '
            logger.info(f'    {status} {BID_TYPE_NAMES.get(type_key, type_key)}: '
                        f'{detail["score"]}分 (阈值{detail["threshold"]})')

    # ===== 步骤3: 加载大纲模板 =====
    logger.info('=' * 50)
    logger.info('📌 步骤3/7: 加载大纲模板')
    logger.info('=' * 50)

    outline_file = OUTLINE_FILES.get(bid_type, 'service_outline.md')
    outline_path = TEMPLATES_DIR / outline_file
    sections = parse_outline(outline_path)

    if not sections:
        logger.error(f'❌ 大纲模板解析失败: {outline_path}')
        return False

    commercial_count = sum(1 for s in sections if s['is_commercial'])
    technical_count = len(sections) - commercial_count
    logger.info(f'  大纲: {len(sections)} 个章节 (商务标{commercial_count} / 技术标{technical_count})')

    # ===== 步骤4: 加载企业素材 =====
    logger.info('=' * 50)
    logger.info('📌 步骤4/7: 加载企业素材')
    logger.info('=' * 50)

    company_profiles = load_company_profiles(company_path)
    filled_count = sum(1 for v in company_profiles.values() if v)
    logger.info(f'  企业素材: {filled_count}/{len(COMPANY_PROFILE_FILES)} 个文件已填写')

    # ===== 步骤5: 构建RFP摘要 =====
    logger.info('=' * 50)
    logger.info('📌 步骤5/7: 构建招标文件摘要')
    logger.info('=' * 50)

    rfp_summary = build_rfp_summary(rfp_data)
    logger.info(f'  摘要长度: {len(rfp_summary)} 字符')

    # ===== 步骤6: 逐章节生成 =====
    logger.info('=' * 50)
    logger.info('📌 步骤6/7: 逐章节生成投标文件内容')
    logger.info('=' * 50)

    # 初始化LLM客户端
    llm_client = LLMClient(config_path=str(TEMPLATES_DIR / 'user_config.yaml') if (TEMPLATES_DIR / 'user_config.yaml').exists() else None)
    if not llm_client.is_available():
        logger.warning('  ⚠️ LLM客户端未就绪（API Key未配置），技术标章节将使用模板')

    # 断点续传
    cache_dir = get_cache_dir(output_file)
    cache_key = get_cache_key(output_file)
    checkpoint = load_checkpoint(cache_dir, cache_key)

    generated_sections: list[str] = []
    failed_sections: list[str] = []
    template_section_titles: list[str] = []
    all_placeholders: list[str] = []

    for section in sections:
        idx = section['index']
        title = section['title']

        # 检查断点续传
        if idx in checkpoint:
            logger.info(f'  [{idx+1}/{len(sections)}] {title} (断点续传，跳过)')
            content = checkpoint[idx]
            generated_sections.append(content)
            all_placeholders.extend(extract_placeholders(content))
            continue

        logger.info(f'  [{idx+1}/{len(sections)}] 生成: {title}')

        if section['is_commercial']:
            # 商务标：直接用模板
            content = generate_commercial_section(section, rfp_summary, company_profiles)
            template_section_titles.append(title)
        elif llm_client.is_available():
            # 技术标：调用LLM
            content, success = generate_technical_section(section, rfp_summary, company_profiles, llm_client)
            if not success:
                failed_sections.append(title)
        else:
            # LLM不可用，回退到模板
            content = generate_commercial_section(section, rfp_summary, company_profiles)
            failed_sections.append(title)

        generated_sections.append(content)

        # 提取占位符
        all_placeholders.extend(extract_placeholders(content))

        # 保存断点续传
        checkpoint[idx] = content
        save_checkpoint(cache_dir, cache_key, checkpoint)

    # 占位符去重
    seen: set[str] = set()
    unique_placeholders: list[str] = []
    for ph in all_placeholders:
        if ph not in seen:
            seen.add(ph)
            unique_placeholders.append(ph)

    # ===== 步骤7: 组装输出 =====
    logger.info('=' * 50)
    logger.info('📌 步骤7/7: 组装并输出投标文件')
    logger.info('=' * 50)

    # 构建文档头部
    header = build_header(bid_type, len(unique_placeholders), len(failed_sections))

    # 组装完整文档
    full_md = header + '\n\n' + '\n\n---\n\n'.join(generated_sections) + '\n'

    # 写入输出文件
    try:
        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text(full_md, encoding='utf-8')
        logger.info(f'  ✅ 投标文件已生成: {output_file}')
    except Exception as e:
        logger.error(f'❌ 写入输出文件失败: {e}')
        return False

    # 清除断点续传缓存
    clear_checkpoint(cache_dir, cache_key)

    # 打印 needs_manual 清单
    report = build_needs_manual_report(unique_placeholders, failed_sections, template_section_titles)
    print('\n' + report, file=sys.stderr)

    # 可选：转换Word
    if docx:
        logger.info('  正在转换为Word文档...')
        docx_path = convert_to_docx(output_file, bid_type, config_path)
        if docx_path:
            logger.info(f'  ✅ Word文档已生成: {docx_path}')
        else:
            logger.warning('  ⚠️ Word转换失败，请手动使用 bid_engine.py 转换')

    logger.info('\n' + '=' * 50)
    logger.info('🎉 投标文件生成完成！')
    logger.info(f'   类型: {BID_TYPE_NAMES.get(bid_type, "未知")}')
    logger.info(f'   章节: {len(sections)} 个')
    logger.info(f'   占位符: {len(unique_placeholders)} 个待填写')
    logger.info(f'   失败章节: {len(failed_sections)} 个')
    logger.info('=' * 50)

    return True


# =====================================================================
# parse 子命令
# =====================================================================

def parse_rfp_only(rfp_path: str, output: Optional[str] = None, pretty: bool = False) -> None:
    """只解析招标文件，不生成投标文件。

    参数：
        rfp_path: 招标文件路径
        output: 输出JSON路径（None=打印到终端）
        pretty: 是否格式化JSON输出
    """
    if not Path(rfp_path).exists():
        print(f'❌ 文件不存在: {rfp_path}', file=sys.stderr)
        sys.exit(1)

    try:
        result = parse_bid_document(rfp_path)
    except SystemExit:
        sys.exit(1)
    except Exception as e:
        print(f'❌ 解析失败: {e}', file=sys.stderr)
        sys.exit(1)

    indent = 2 if pretty else None
    output_json = json.dumps(result, ensure_ascii=False, indent=indent)

    if output:
        with open(output, 'w', encoding='utf-8') as f:
            f.write(output_json)
        print(f'✅ 已输出到: {output}', file=sys.stderr)
    else:
        print(output_json)

    # 打印摘要
    stats = result.get('_统计', {})
    if stats:
        print(f'\n📊 拆解摘要:', file=sys.stderr)
        print(f'  废标红线: {stats.get("废标项", 0)} 条', file=sys.stderr)
        print(f'  文件清单: {stats.get("文件清单", 0)} 项', file=sys.stderr)
        print(f'  评分项: {stats.get("评分项", 0)} 个', file=sys.stderr)
        print(f'  资质要求: {stats.get("资质要求", 0)} 条', file=sys.stderr)
        print(f'  表格数: {stats.get("表格数", 0)} 个', file=sys.stderr)

    # 类型识别
    from parse_bid import read_input_file
    try:
        rfp_text = read_input_file(rfp_path)
        detection_yaml = TEMPLATES_DIR / 'bid_type_detection.yaml'
        bid_type, detection_info = detect_bid_type(rfp_text, detection_yaml)
        print(f'\n🔍 标书类型识别: {detection_info["name"]}', file=sys.stderr)
        for type_key, detail in detection_info.get('details', {}).items():
            status = '✅' if detail['passed'] else '  '
            print(f'  {status} {BID_TYPE_NAMES.get(type_key, type_key)}: '
                  f'{detail["score"]}分 (阈值{detail["threshold"]})', file=sys.stderr)
    except Exception:
        pass


# =====================================================================
# CLI 入口
# =====================================================================

def main() -> None:
    """CLI入口函数。

    支持两个子命令：
      generate — 解析招标文件并生成投标文件
      parse    — 只解析招标文件，查看拆解结果
    """
    parser = argparse.ArgumentParser(
        description='投标文件智能生成器 v3.5 — 从招标文件到投标文件初稿的一键生成',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            '示例:\n'
            '  # 全自动生成\n'
            '  python scripts/bid_generator.py generate --rfp XX项目.pdf --company company_profile/ --output 投标文件.md\n\n'
            '  # 指定类型 + 生成Word\n'
            '  python scripts/bid_generator.py generate --rfp XX项目.pdf --type services --company company_profile/ --output 投标文件.md --docx\n\n'
            '  # 只解析招标文件\n'
            '  python scripts/bid_generator.py parse --rfp XX项目.pdf --pretty'
        ),
    )

    subparsers = parser.add_subparsers(dest='command', help='子命令')

    # generate 子命令
    gen_parser = subparsers.add_parser('generate', help='生成投标文件')
    gen_parser.add_argument('--rfp', required=True, help='招标文件路径（PDF/MD/TXT）')
    gen_parser.add_argument('--type', choices=['goods', 'service', 'engineering'],
                            default=None, help='指定标书类型（不指定则自动识别）')
    gen_parser.add_argument('--company', required=True, help='企业素材目录路径')
    gen_parser.add_argument('--output', required=True, help='输出投标文件路径（.md）')
    gen_parser.add_argument('--docx', action='store_true', help='同时生成Word文档')
    gen_parser.add_argument('--config', default=None, help='自定义配置文件路径')

    # parse 子命令
    parse_parser = subparsers.add_parser('parse', help='只解析招标文件')
    parse_parser.add_argument('--rfp', required=True, help='招标文件路径（PDF/MD/TXT）')
    parse_parser.add_argument('-o', '--output', default=None, help='输出JSON路径（默认打印到终端）')
    parse_parser.add_argument('--pretty', action='store_true', help='格式化JSON输出')

    args = parser.parse_args()

    if args.command == 'generate':
        success = generate_bid_document(
            rfp_path=args.rfp,
            output_path=args.output,
            company_dir=args.company,
            bid_type=args.type,
            docx=args.docx,
            config_path=args.config,
        )
        sys.exit(0 if success else 1)

    elif args.command == 'parse':
        parse_rfp_only(rfp_path=args.rfp, output=args.output, pretty=args.pretty)

    else:
        parser.print_help()
        sys.exit(0)


if __name__ == '__main__':
    main()
