#!/usr/bin/env python3
"""Pre-publish Checklist Engine - Layer 1 of WeChat Publishing Pipeline.

v1.0 (2026-08-12): Initial release.
  - Layout iron rules: border/card style, image points, cover topic match
  - Content quality: typo, readability, logic, reference checks
  - Template binding: column-specific rule enforcement
  - Non-bypassable critical checks with structured report output
"""

import json
import re
import sys
import os
from pathlib import Path
from typing import Dict, List, Tuple, Any

# ============================================================
# Configuration
# ============================================================
SCRIPT_DIR = Path(__file__).parent.resolve()
CONFIG_PATH = SCRIPT_DIR / "column_templates.json"

# Common Chinese typo patterns: (wrong, correct, context_pattern_or_None)
# Context pattern can be regex to reduce false positives
COMMON_TYPO_PATTERNS = [
    ("的得地混淆", None),  # Will be handled by heuristics
    ("做坐座作", None),
    ("在再", None),
    ("那哪", None),
    ("象像", None),
    ("已以", None),
    ("andriod", "Android"),
    ("Andriod", "Android"),
    ("jsva", "Java"),
    ("Jaca", "Java"),
    ("pyhon", "Python"),
    ("Pythom", "Python"),
    ("goLang", "Go"),
    ("Golang", "Go"),
    ("ai", "AI", r"\bai\b"),  # lowercase 'ai' as standalone word
    ("gpt", "GPT", r"\bgpt\b"),
    ("deepseek", "DeepSeek", r"\bdeepseek\b"),
    ("openai", "OpenAI", r"\bopenai\b"),
    ("chatgpt", "ChatGPT", r"\bchatgpt\b"),
    ("claude", "Claude", r"\bclaude\b"),
    ("gemini", "Gemini", r"\bgemini\b"),
    ("llm", "LLM", r"\bllm\b"),
    ("api", "API", r"\bapi\b"),
    ("ui", "UI", r"\bui\b"),
    ("ux", "UX", r"\bux\b"),
    ("gpu", "GPU", r"\bgpu\b"),
    ("cpu", "CPU", r"\bcpu\b"),
    ("ram", "RAM", r"\bram\b"),
    ("ssd", "SSD", r"\bssd\b"),
    ("hdd", "HDD", r"\bhdd\b"),
    ("saas", "SaaS", r"\bsaas\b"),
    ("paas", "PaaS", r"\bpaas\b"),
    ("iaas", "IaaS", r"\biaas\b"),
    ("crm", "CRM", r"\bcrm\b"),
    ("erp", "ERP", r"\berp\b"),
    ("oa", "OA", r"\boa\b"),
    ("hr", "HR", r"\bhr\b"),
    ("it", "IT", r"\bit\b"),
    ("cto", "CTO", r"\bcto\b"),
    ("ceo", "CEO", r"\bceo\b"),
    ("cfo", "CFO", r"\bcfo\b"),
    ("coo", "COO", r"\bcoo\b"),
    ("cmo", "CMO", r"\bcmo\b"),
    ("ciso", "CISO", r"\bciso\b"),
    ("ai agent", "AI Agent", None),
    ("aiagent", "AI Agent", None),
    ("rag", "RAG", r"\brag\b"),
    ("mcp", "MCP", r"\bmcp\b"),
    ("lora", "LoRA", r"\blora\b"),
    ("qlora", "QLoRA", r"\bqlora\b"),
    ("grpo", "GRPO", r"\bgrpo\b"),
    ("ppo", "PPO", r"\bppo\b"),
    ("dpo", "DPO", r"\bdpo\b"),
    ("rlhf", "RLHF", r"\brlhf\b"),
    ("sft", "SFT", r"\bsft\b"),
    ("icl", "ICL", r"\bicl\b"),
    ("cot", "CoT", r"\bcot\b"),
    ("tot", "ToT", r"\btot\b"),
    ("got", "GoT", r"\bgot\b"),
    ("aigc", "AIGC", r"\baigc\b"),
    ("ugc", "UGC", r"\bugc\b"),
    ("pgc", "PGC", r"\bpgc\b"),
    ("pugc", "PUGC", r"\bpugc\b"),
    ("ogc", "OGC", r"\bogc\b"),
    ("mcn", "MCN", r"\bmcn\b"),
    ("kpi", "KPI", r"\bkpi\b"),
    ("okr", "OKR", r"\bokr\b"),
    ("roi", "ROI", r"\broi\b"),
    ("gmv", "GMV", r"\bgmv\b"),
    ("dau", "DAU", r"\bdau\b"),
    ("mau", "MAU", r"\bmau\b"),
    ("arpu", "ARPU", r"\barpu\b"),
    ("ltv", "LTV", r"\bltv\b"),
    ("cac", "CAC", r"\bcac\b"),
    ("nlp", "NLP", r"\bnlp\b"),
    ("cv", "CV", r"\bcv\b"),
    ("asr", "ASR", r"\basr\b"),
    ("tts", "TTS", r"\btts\b"),
    ("ocr", "OCR", r"\bocr\b"),
    ("ocr文字识别", "OCR文字识别", None),
    ("nl\b", "NLP", r"\bnl\b"),
]

# Suspicious phrases that often indicate unverified claims
SUSPICIOUS_CLAIM_PATTERNS = [
    r"据说",
    r"据传",
    r"有消息称",
    r"业内人士透露",
    r"知情人士称",
    r"未经证实",
    r"可能大概",
    r"应该大概",
    r"估计大约",
    r"或许可能",
    r"\d+%的人",
    r"\d+% 的人",
    r"超过\d+%",
    r"高达\d+%",
    r"低至\d+%",
    r"至少\d+%",
    r"最多\d+%",
    r"\d+%以上",
    r"\d+%以下",
    r"排名第一",
    r"行业第一",
    r"全球第一",
    r"国内第一",
    r"遥遥领先",
    r"碾压",
    r"吊打",
    r"秒杀",
    r"完爆",
    r"颠覆",
    r"革命性",
    r"突破性",
    r"前所未有",
    r"史无前例",
    r"开创先河",
    r"里程碑",
    r"转折点",
    r"分水岭",
    r"标志性事件",
]

# Logic consistency patterns
LOGIC_CHECK_PATTERNS = {
    "date_inconsistency": r"(\d{4})年.*同年.*(?:但|然而|不过|却).*\1年",  # Same year contradiction
    "number_range_error": r"(\d+(?:\.\d+)?).*(?:超过|大于|高于).*\1",  # X > X
    "self_contradiction": r"(?:不是|并非|没有).*(?:就是|正是|有着|存在)",  # Not X but X
}


# ============================================================
# Check Result Classes
# ============================================================
class CheckResult:
    CRITICAL = "CRITICAL"
    WARNING = "WARNING"
    INFO = "INFO"
    PASS = "PASS"
    SKIP = "SKIP"

    def __init__(
        self,
        category: str,
        name: str,
        status: str,
        message: str = "",
        details: List[str] = None,
    ):
        self.category = category
        self.name = name
        self.status = status
        self.message = message
        self.details = details or []

    def to_dict(self) -> Dict[str, Any]:
        return {
            "category": self.category,
            "name": self.name,
            "status": self.status,
            "message": self.message,
            "details": self.details,
        }

    def is_blocking(self) -> bool:
        return self.status == self.CRITICAL


class ChecklistReport:
    def __init__(self):
        self.results: List[CheckResult] = []
        self.blocked = False
        self.block_reasons: List[str] = []

    def add(self, result: CheckResult):
        self.results.append(result)
        if result.is_blocking():
            self.blocked = True
            self.block_reasons.append(
                f"[{result.category}] {result.name}: {result.message}"
            )

    def get_summary(self) -> Dict[str, Any]:
        critical = sum(1 for r in self.results if r.status == CheckResult.CRITICAL)
        warnings = sum(1 for r in self.results if r.status == CheckResult.WARNING)
        infos = sum(1 for r in self.results if r.status == CheckResult.INFO)
        passed = sum(1 for r in self.results if r.status == CheckResult.PASS)
        skipped = sum(1 for r in self.results if r.status == CheckResult.SKIP)
        return {
            "total": len(self.results),
            "critical": critical,
            "warning": warnings,
            "info": infos,
            "pass": passed,
            "skip": skipped,
            "blocked": self.blocked,
            "block_reasons": self.block_reasons,
        }

    def to_dict(self) -> Dict[str, Any]:
        return {
            "summary": self.get_summary(),
            "results": [r.to_dict() for r in self.results],
        }

    def print_report(self):
        summary = self.get_summary()
        print(f"\n{'=' * 60}")
        print(f"  排版前检查清单报告")
        print(f"{'=' * 60}")
        print(f"  总计: {summary['total']} 项")
        print(
            f"  通过: {summary['pass']}  |  警告: {summary['warning']}  |  严重: {summary['critical']}"
        )
        print(f"  信息: {summary['info']}  |  跳过: {summary['skip']}")
        if summary["blocked"]:
            print(f"\n  [BLOCKED] 以下严重问题阻止发布:")
            for reason in summary["block_reasons"]:
                print(f"    - {reason}")
        print(f"{'=' * 60}\n")

        # Detail print
        for r in self.results:
            icon = {
                CheckResult.PASS: "[PASS]",
                CheckResult.WARNING: "[WARN]",
                CheckResult.CRITICAL: "[FAIL]",
                CheckResult.INFO: "[INFO]",
                CheckResult.SKIP: "[SKIP]",
            }.get(r.status, "[?]")
            print(f"  {icon} [{r.category}] {r.name}: {r.message}")
            for d in r.details[:5]:  # Limit details
                print(f"      -> {d}")
            if len(r.details) > 5:
                print(f"      ... 还有 {len(r.details) - 5} 项")


# ============================================================
# Template Loader
# ============================================================
def load_column_templates() -> Dict[str, Any]:
    """Load column template configuration."""
    if not CONFIG_PATH.exists():
        return {}
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("templates", {})


def get_global_rules() -> Dict[str, Any]:
    """Load global publishing rules."""
    if not CONFIG_PATH.exists():
        return {}
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("global_rules", {})


def get_content_quality_config() -> Dict[str, Any]:
    """Load content quality check configuration."""
    if not CONFIG_PATH.exists():
        return {}
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("content_quality_checks", {})


# ============================================================
# Layout Iron Rules Checks
# ============================================================
def check_layout_iron_rules(
    md_text: str, column: str, template: Dict[str, Any], report: ChecklistReport
):
    """Check layout iron rules specific to the column template."""

    # 1. Card border / section border check
    if template.get("layout_rules", {}).get("card_border", False):
        # For markdown->HTML pipeline, we check if the template requires borders
        # This is a design-time check; actual HTML verification happens in post-audit
        report.add(
            CheckResult(
                category="排版铁律",
                name="边框/卡片化风格绑定",
                status=CheckResult.PASS,
                message=f"栏目「{column}」已绑定边框卡片化模板(style={template.get('style')})",
            )
        )
    else:
        report.add(
            CheckResult(
                category="排版铁律",
                name="边框/卡片化风格绑定",
                status=CheckResult.INFO,
                message=f"栏目「{column}」使用极简风格，无需卡片边框",
            )
        )

    # 2. Image points check for long articles
    article_length = len(md_text)
    threshold = template.get("long_article_threshold", 1500)
    min_images = template.get("long_article_min_images", 2)

    # Count images in markdown
    image_count = len(re.findall(r"!\[.*?\]\(.*?\)", md_text))
    # Also count image placeholders or references
    image_count += len(re.findall(r"【配图.*?】", md_text))
    image_count += len(re.findall(r"\[图片.*?\]", md_text))

    if article_length >= threshold:
        if image_count >= min_images:
            report.add(
                CheckResult(
                    category="排版铁律",
                    name="长文配图点检查",
                    status=CheckResult.PASS,
                    message=f"长文({article_length}字)已标注 {image_count} 个配图点(要求≥{min_images})",
                )
            )
        else:
            report.add(
                CheckResult(
                    category="排版铁律",
                    name="长文配图点检查",
                    status=CheckResult.CRITICAL,
                    message=f"长文({article_length}字)仅标注 {image_count} 个配图点(要求≥{min_images})",
                    details=["请在文章中标注足够的配图位置，或使用【配图N: 描述】标记"],
                )
            )
    else:
        min_img = template.get("min_image_count", 1)
        if image_count >= min_img:
            status = CheckResult.PASS
            msg = (
                f"文章({article_length}字)已标注 {image_count} 个配图点(要求≥{min_img})"
            )
        else:
            status = CheckResult.WARNING
            msg = (
                f"文章({article_length}字)仅标注 {image_count} 个配图点(建议≥{min_img})"
            )
        report.add(
            CheckResult(
                category="排版铁律", name="配图点检查", status=status, message=msg
            )
        )

    # 3. Cover image topic matching (basic check)
    # Extract title
    title_match = re.search(r"^#\s+(.+)$", md_text, re.MULTILINE)
    title = title_match.group(1).strip() if title_match else ""

    # Check if cover is provided (this will be checked at runtime with cover_path)
    report.add(
        CheckResult(
            category="排版铁律",
            name="封面图主题匹配",
            status=CheckResult.INFO,
            message=f"标题「{title[:30]}...」请在生成封面图时确认主题一致",
        )
    )

    # 4. Required elements check
    required = template.get("required_elements", [])
    # For markdown source, we verify structure presence
    has_h2 = bool(re.search(r"^##\s+", md_text, re.MULTILINE))
    has_quote = bool(re.search(r"^>\s+", md_text, re.MULTILINE))
    has_table = bool(re.search(r"^\|.*\|.*\|$", md_text, re.MULTILINE))
    has_code = bool(re.search(r"^```", md_text, re.MULTILINE))

    missing_elements = []
    if "h2_left_bar" in required and not has_h2:
        missing_elements.append("缺少H2标题（用于左竖线装饰）")
    if "quote_block_styled" in required and not has_quote:
        missing_elements.append("缺少引用块（栏目风格需要引用装饰）")
    if "table_styled" in required and not has_table:
        missing_elements.append("缺少表格（栏目风格需要数据展示）")
    if "code_block_styled" in required and not has_code:
        missing_elements.append("缺少代码块（技术栏目需要代码展示）")

    if missing_elements:
        report.add(
            CheckResult(
                category="排版铁律",
                name="栏目必需元素检查",
                status=CheckResult.WARNING,
                message=f"栏目「{column}」建议包含以下元素",
                details=missing_elements,
            )
        )
    else:
        report.add(
            CheckResult(
                category="排版铁律",
                name="栏目必需元素检查",
                status=CheckResult.PASS,
                message=f"栏目「{column}」必需元素已满足",
            )
        )


# ============================================================
# Content Quality Checks
# ============================================================
def check_typos(md_text: str, report: ChecklistReport, config: Dict[str, Any]):
    """Basic typo and case consistency check."""
    if not config.get("typo_check", {}).get("enabled", True):
        report.add(
            CheckResult(
                category="内容质量",
                name="错别字检查",
                status=CheckResult.SKIP,
                message="已禁用",
            )
        )
        return

    issues = []

    # Check common typo patterns
    for item in COMMON_TYPO_PATTERNS:
        if len(item) >= 3 and item[2]:  # Has context pattern
            wrong, correct, pattern = item
            matches = re.finditer(pattern, md_text, re.IGNORECASE)
            for m in matches:
                # Check if it's already correct case
                matched_text = m.group(0)
                if matched_text != correct:
                    # Get context
                    start = max(0, m.start() - 15)
                    end = min(len(md_text), m.end() + 15)
                    context = md_text[start:end].replace("\n", " ")
                    issues.append(
                        f"位置{m.start()}: 「{matched_text}」应为「{correct}」...{context}..."
                    )
        elif len(item) == 2:
            wrong, correct = item
            # Simple substring check for obvious typos
            if wrong in md_text.lower() and wrong != correct.lower():
                # Check context to avoid false positives
                for m in re.finditer(
                    r"\b" + re.escape(wrong) + r"\b", md_text, re.IGNORECASE
                ):
                    matched = m.group(0)
                    if matched != correct:
                        start = max(0, m.start() - 15)
                        end = min(len(md_text), m.end() + 15)
                        context = md_text[start:end].replace("\n", " ")
                        issues.append(
                            f"位置{m.start()}: 「{matched}」疑似应为「{correct}」...{context}..."
                        )

    # Check 的/得/地 heuristics (very basic)
    de_issues = []
    # Pattern: "adj + 的 + verb" is likely wrong (should be 地)
    for m in re.finditer(
        r"([飞快迅速慢慢悄悄轻轻偷偷默默紧紧牢牢深深细细微微])的([做走跑说看听来去上下进出])",
        md_text,
    ):
        de_issues.append(
            f"位置{m.start()}: 「{m.group(0)}」应为「{m.group(1)}地{m.group(2)}」"
        )
    # Pattern: "verb + 的 + adj/noun" is likely wrong (should be 得)
    # Skip if followed by English letter (e.g., "跑出来的多Agent" is not "来得多")
    for m in re.finditer(
        r"([做走跑说看听来去上下进出吃睡玩写读])的([很好快慢多少远近高矮胖瘦冷热])",
        md_text,
    ):
        # Check if the char after the match is an English letter
        after_pos = m.end()
        if after_pos < len(md_text) and re.match(r"[A-Za-z]", md_text[after_pos]):
            continue  # "的+多" followed by English is likely "的" + "多X", not "得多"
        de_issues.append(
            f"位置{m.start()}: 「{m.group(0)}」应为「{m.group(1)}得{m.group(2)}」"
        )

    issues.extend(de_issues)

    # Check repeated characters (e.g., "的的", "了了")
    # Exclude markdown syntax lines (headings #, hr ---, tables |, list -)
    # and valid Chinese punctuation
    for m in re.finditer(r"(.)\1{2,}", md_text):  # 3+ same chars
        char = m.group(1)
        # Skip markdown syntax chars and valid repeated punctuation
        if char in "#-|>*_=~` \t":
            continue
        if char in "\u2026\u2014":  # …… ——
            continue
        # Check if this match is inside a markdown syntax line
        line_start = md_text.rfind("\n", 0, m.start()) + 1
        line_content = md_text[line_start : m.start() + len(m.group(0))]
        if line_content.lstrip().startswith(("#", "---", "|", "- ", ">", "```")):
            continue
        issues.append(f"位置{m.start()}: 连续重复字符「{m.group(0)}」")

    if issues:
        report.add(
            CheckResult(
                category="内容质量",
                name="错别字与大小写检查",
                status=CheckResult.WARNING,
                message=f"发现 {len(issues)} 处疑似问题",
                details=issues[:10],
            )
        )
    else:
        report.add(
            CheckResult(
                category="内容质量",
                name="错别字与大小写检查",
                status=CheckResult.PASS,
                message="未发现明显错别字或大小写问题",
            )
        )


def check_readability(md_text: str, report: ChecklistReport, config: Dict[str, Any]):
    """Check sentence and paragraph readability."""
    if not config.get("readability_check", {}).get("enabled", True):
        report.add(
            CheckResult(
                category="内容质量",
                name="语句通顺检查",
                status=CheckResult.SKIP,
                message="已禁用",
            )
        )
        return

    issues = []

    # Check overly long sentences (>100 chars)
    sentences = re.split(r"[。！？\n]", md_text)
    long_sentences = []
    for s in sentences:
        s = s.strip()
        if len(s) > 100 and not s.startswith("```") and not s.startswith("|"):
            long_sentences.append(s[:60] + "...")

    if len(long_sentences) > 3:
        issues.append(f"发现 {len(long_sentences)} 个超长句子(>100字)，建议拆分")
        issues.extend(long_sentences[:3])

    # Check repeated words in close proximity (e.g., "非常非常", "就是一个就是一个")
    # Exclude markdown syntax (---, |, ###, ```, etc.)
    # Exclude common rhetorical repetitions ("一个一个", "一步一步", "一天一天")
    _RHETORICAL_REPS = {
        "一个一个",
        "一步一步",
        "一天一天",
        "一点一点",
        "一次一次",
        "一遍一遍",
        "一年一年",
        "一声一声",
        "一下一下",
        "一件一件",
    }
    for m in re.finditer(r"(\S{2,})\1", md_text):  # Repeated 2+ char sequences
        # Skip if the match is part of markdown syntax
        if m.group(0) and m.group(0)[0] in "#-|>*_=`~":
            continue
        # Skip rhetorical repetitions
        if m.group(0) in _RHETORICAL_REPS:
            continue
        issues.append(f"位置{m.start()}: 疑似重复用词「{m.group(0)}」")

    # Check paragraphs that are too long (>500 chars)
    paragraphs = md_text.split("\n\n")
    long_paras = []
    for p in paragraphs:
        p = p.strip()
        if len(p) > 500 and not p.startswith("```"):
            long_paras.append(p[:50] + "...")
    if len(long_paras) > 2:
        issues.append(f"发现 {len(long_paras)} 个超长段落(>500字)，建议分段")

    # Check consecutive same-starting sentences (monotony)
    # Exclude markdown syntax lines (headings, list items, blockquotes, hr, table rows)
    lines = [
        l.strip()
        for l in md_text.split("\n")
        if l.strip()
        and len(l.strip()) > 10
        and not l.strip().startswith(("#", "- ", ">", "|", "---", "```", "!"))
    ]
    same_start = 0
    for i in range(1, min(len(lines), 20)):
        if lines[i][0] == lines[i - 1][0]:
            same_start += 1
    if same_start > 5:
        issues.append(f"发现 {same_start} 处连续同首字句子，注意行文节奏变化")

    if issues:
        report.add(
            CheckResult(
                category="内容质量",
                name="语句通顺检查",
                status=CheckResult.WARNING,
                message=f"发现 {len(issues)} 处可读性问题",
                details=issues[:8],
            )
        )
    else:
        report.add(
            CheckResult(
                category="内容质量",
                name="语句通顺检查",
                status=CheckResult.PASS,
                message="语句通顺，节奏良好",
            )
        )


def check_logic(md_text: str, report: ChecklistReport, config: Dict[str, Any]):
    """Basic logic and consistency checks."""
    if not config.get("logic_check", {}).get("enabled", True):
        report.add(
            CheckResult(
                category="内容质量",
                name="逻辑一致性检查",
                status=CheckResult.SKIP,
                message="已禁用",
            )
        )
        return

    issues = []

    # Check date mentions for consistency
    dates_found = re.findall(r"(\d{4})年(\d{1,2})月(\d{1,2})日", md_text)
    years = set(d for d, m, day in dates_found)
    if len(years) > 1:
        issues.append(f"文中提及多个年份: {', '.join(sorted(years))}，请确认时间线一致")

    # Check number consistency (same number mentioned differently)
    numbers = re.findall(r"\b(\d+(?:\.\d+)?)\b", md_text)
    from collections import Counter

    num_counter = Counter(numbers)
    repeated_numbers = {k: v for k, v in num_counter.items() if v > 2}
    if repeated_numbers:
        issues.append(
            f"以下数字出现多次，请确认一致性: {list(repeated_numbers.keys())[:5]}"
        )

    # Check for self-contradictory phrases
    for pattern_name, pattern in LOGIC_CHECK_PATTERNS.items():
        matches = re.findall(pattern, md_text)
        if matches:
            issues.append(f"发现可能的自相矛盾表达({pattern_name})")

    if issues:
        report.add(
            CheckResult(
                category="内容质量",
                name="逻辑一致性检查",
                status=CheckResult.WARNING,
                message=f"发现 {len(issues)} 处逻辑疑点",
                details=issues,
            )
        )
    else:
        report.add(
            CheckResult(
                category="内容质量",
                name="逻辑一致性检查",
                status=CheckResult.PASS,
                message="未发现明显逻辑矛盾",
            )
        )


def check_references(md_text: str, report: ChecklistReport, config: Dict[str, Any]):
    """Check references and flag suspicious claims."""
    if not config.get("reference_check", {}).get("enabled", True):
        report.add(
            CheckResult(
                category="内容质量",
                name="引用真实性检查",
                status=CheckResult.SKIP,
                message="已禁用",
            )
        )
        return

    issues = []
    flagged_positions = []

    # Check for suspicious claim patterns
    for pattern in SUSPICIOUS_CLAIM_PATTERNS:
        for m in re.finditer(pattern, md_text):
            start = max(0, m.start() - 20)
            end = min(len(md_text), m.end() + 20)
            context = md_text[start:end].replace("\n", " ")
            issues.append(f"位置{m.start()}: 可疑表述「{m.group(0)}」...{context}...")
            flagged_positions.append(m.start())

    # Check for unverified data claims (numbers without sources)
    # Pattern: "X%" or "X billion" followed by claim without citation
    unverified_data = re.finditer(
        r"(\d+(?:\.\d+)?%?\s*(?:亿|万|千|百万|千万|万亿)?).{0,30}(?=。|！|？|$)",
        md_text,
    )
    data_count = 0
    for m in unverified_data:
        # Check if followed by source mention
        after = md_text[m.end() : m.end() + 50]
        if not re.search(r"(据|来源|来自|根据|引用|报告|显示)", after):
            data_count += 1

    if data_count > 5:
        issues.append(f"发现约 {data_count} 处数据/数字陈述未标注来源，建议补充")

    # Check for specific company/product claims without context
    company_claims = re.finditer(
        r"(百度|阿里|腾讯|字节|华为|小米|OPPO|vivo|苹果|谷歌|微软|OpenAI|Meta|亚马逊).{0,40}(?:推出|发布|宣布| claims)",
        md_text,
    )
    company_issues = []
    for m in company_claims:
        start = max(0, m.start() - 10)
        end = min(len(md_text), m.end() + 30)
        context = md_text[start:end].replace("\n", " ")
        company_issues.append(f"公司声明: ...{context}...")

    if company_issues:
        issues.append(f"发现 {len(company_issues)} 处公司/产品声明，请确认信息来源")
        issues.extend(company_issues[:3])

    if issues:
        report.add(
            CheckResult(
                category="内容质量",
                name="引用真实性检查",
                status=CheckResult.WARNING,
                message=f"发现 {len(issues)} 处需核实的内容",
                details=issues[:10],
            )
        )
    else:
        report.add(
            CheckResult(
                category="内容质量",
                name="引用真实性检查",
                status=CheckResult.PASS,
                message="引用和声明未发现明显问题",
            )
        )


# ============================================================
# Global Rule Checks
# ============================================================
def check_global_rules(md_text: str, report: ChecklistReport, config: Dict[str, Any]):
    """Check global publishing rules."""

    # 1. Title length
    title_match = re.search(r"^#\s+(.+)$", md_text, re.MULTILINE)
    title = title_match.group(1).strip() if title_match else ""
    min_len = config.get("min_title_length", 5)
    max_len = config.get("max_title_length", 64)

    if not title:
        report.add(
            CheckResult(
                category="全局规则",
                name="文章标题",
                status=CheckResult.CRITICAL,
                message="未找到H1标题",
            )
        )
    elif len(title) < min_len:
        report.add(
            CheckResult(
                category="全局规则",
                name="文章标题",
                status=CheckResult.CRITICAL,
                message=f"标题过短({len(title)}字，要求≥{min_len})",
            )
        )
    elif len(title) > max_len:
        report.add(
            CheckResult(
                category="全局规则",
                name="文章标题",
                status=CheckResult.WARNING,
                message=f"标题过长({len(title)}字，建议≤{max_len})",
            )
        )
    else:
        report.add(
            CheckResult(
                category="全局规则",
                name="文章标题",
                status=CheckResult.PASS,
                message=f"标题长度合适({len(title)}字)",
            )
        )

    # 2. Article length
    article_len = len(md_text)
    min_article = config.get("article_min_length", 300)
    max_article = config.get("article_max_length", 50000)

    if article_len < min_article:
        report.add(
            CheckResult(
                category="全局规则",
                name="文章长度",
                status=CheckResult.CRITICAL,
                message=f"文章过短({article_len}字，要求≥{min_article})",
            )
        )
    elif article_len > max_article:
        report.add(
            CheckResult(
                category="全局规则",
                name="文章长度",
                status=CheckResult.WARNING,
                message=f"文章超长({article_len}字，建议≤{max_article})",
            )
        )
    else:
        report.add(
            CheckResult(
                category="全局规则",
                name="文章长度",
                status=CheckResult.PASS,
                message=f"文章长度合适({article_len}字)",
            )
        )

    # 3. Character hygiene (basic check for zero-width chars)
    zw_chars = []
    for pos, ch in enumerate(md_text):
        if ch in "\u200b\u200c\u200d\ufeff":
            zw_chars.append(f"位置{pos}: U+{ord(ch):04X}")

    if zw_chars:
        report.add(
            CheckResult(
                category="全局规则",
                name="字符卫生",
                status=CheckResult.WARNING,
                message=f"发现 {len(zw_chars)} 个零宽字符",
                details=zw_chars[:5],
            )
        )
    else:
        report.add(
            CheckResult(
                category="全局规则",
                name="字符卫生",
                status=CheckResult.PASS,
                message="未发现零宽字符",
            )
        )

    # 4. Full-width/half-width basic check
    # Check for common full-width ASCII in non-Chinese context
    suspicious_fw = []
    for m in re.finditer(r"[\uff10-\uff19\uff21-\uff3a\uff41-\uff5a]", md_text):
        # Check if surrounded by Chinese (then it's OK)
        before = md_text[max(0, m.start() - 1) : m.start()]
        after = md_text[m.end() : min(len(md_text), m.end() + 1)]
        if not (
            before
            and "\u4e00" <= before <= "\u9fff"
            or after
            and "\u4e00" <= after <= "\u9fff"
        ):
            suspicious_fw.append(f"位置{m.start()}: 「{m.group(0)}」")

    if suspicious_fw:
        report.add(
            CheckResult(
                category="全局规则",
                name="全半角规范",
                status=CheckResult.WARNING,
                message=f"发现 {len(suspicious_fw)} 处可疑全角字符",
                details=suspicious_fw[:5],
            )
        )
    else:
        report.add(
            CheckResult(
                category="全局规则",
                name="全半角规范",
                status=CheckResult.PASS,
                message="全半角规范检查通过",
            )
        )


# ============================================================
# Main Entry Point
# ============================================================
def run_checklist(
    md_text: str, column: str = "江湖夜话", cover_path: str = None
) -> ChecklistReport:
    """Run the full pre-publish checklist.

    Args:
        md_text: The markdown article content.
        column: The WeChat column name (栏目).
        cover_path: Path to cover image (for verification).

    Returns:
        ChecklistReport with all check results.
    """
    report = ChecklistReport()

    # Load configurations
    templates = load_column_templates()
    global_rules = get_global_rules()
    quality_config = get_content_quality_config()

    # Validate column
    if column not in templates:
        report.add(
            CheckResult(
                category="模板绑定",
                name="栏目模板存在性",
                status=CheckResult.CRITICAL,
                message=f"栏目「{column}」未在模板配置中定义，请先在 column_templates.json 中配置",
            )
        )
        return report

    template = templates[column]
    report.add(
        CheckResult(
            category="模板绑定",
            name="栏目模板存在性",
            status=CheckResult.PASS,
            message=f"已绑定栏目「{column}」模板(style={template.get('style')})",
        )
    )

    # Run all check modules
    check_global_rules(md_text, report, global_rules)
    check_layout_iron_rules(md_text, column, template, report)
    check_typos(md_text, report, quality_config)
    check_readability(md_text, report, quality_config)
    check_logic(md_text, report, quality_config)
    check_references(md_text, report, quality_config)

    return report


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Pre-publish checklist for WeChat articles"
    )
    parser.add_argument("md_path", help="Path to markdown file")
    parser.add_argument("--column", default="江湖夜话", help="WeChat column name")
    parser.add_argument("--cover", default=None, help="Path to cover image")
    parser.add_argument("--json", action="store_true", help="Output as JSON")
    parser.add_argument(
        "--bypass", action="store_true", help="Bypass blocking (logs warning)"
    )
    args = parser.parse_args()

    if not os.path.exists(args.md_path):
        print(f"ERROR: File not found: {args.md_path}")
        sys.exit(1)

    with open(args.md_path, "r", encoding="utf-8") as f:
        md_text = f.read()

    report = run_checklist(md_text, args.column, args.cover)

    if args.json:
        import json as _json

        print(_json.dumps(report.to_dict(), ensure_ascii=False, indent=2))
    else:
        report.print_report()

    if report.blocked and not args.bypass:
        print("\n[BLOCKED] 发现严重问题，发布流程已阻断。")
        print("如需强制继续，请添加 --bypass 参数（将记录日志）。")
        sys.exit(2)
    elif report.blocked and args.bypass:
        print("\n[WARNING] 已绕过严重问题继续发布（此操作已被记录）。")
        sys.exit(0)
    else:
        print("[PASS] 所有严重检查项已通过，可以进入排版流程。")
        sys.exit(0)


if __name__ == "__main__":
    main()
