# -*- coding: utf-8 -*-
"""
原文锚定比对引擎 (Source Anchoring Engine)
==========================================

核心能力：投标文件中的每一条承诺，必须能在招标文件中找到原文依据。

与铁律①过度承诺的关系：
  - ①过度承诺：用静态关键词表检查（"保险金额""违约金"等固定列表）
  - 原文锚定：动态分析投标文件中所有承诺性语句，逐条反查招标文件原文
  - 原文锚定是①的升级版，从"关键词列表"升级到"语义级锚定"

工作流：
  1. 从投标文件提取承诺性语句（含"承诺/保证/确保/负责"等动词的句子）
  2. 对每条承诺提取核心关键词
  3. 在招标文件原文中搜索匹配
  4. 无原文依据的标记⚠️（可能是AI编的承诺）

纯代码实现，不需要LLM。
"""

import re
from typing import List, Dict, Optional, Tuple
from collections import Counter


class SourceAnchor:
    """原文锚定比对引擎"""

    # 承诺性动词（触发承诺提取）
    COMMITMENT_VERBS = [
        "承诺", "保证", "确保", "负责", "承担", "提供", "配备", "投入",
        "派遣", "安排", "设立", "建立", "实行", "执行", "遵守",
    ]

    # 将来时态标记（"我方将..."）
    FUTURE_MARKERS = ["将", "将会", "拟", "计划", "准备"]

    # 承诺性句式正则
    COMMITMENT_PATTERNS = [
        # "我方承诺/保证/确保..."
        r'(?:我方|本公司|投标人|我们)(?:郑重)?(?:承诺|保证|确保)([^。！\n]{10,150})',
        # "承诺/保证/确保..."（不带主语）
        r'(?:承诺|保证|确保)([^。！\n]{10,150})',
        # "负责/承担/提供..."
        r'(?:负责|承担)([^。！\n]{10,150})',
        # "将/将会..."
        r'(?:我方|本公司|投标人)?(?:将|将会)([^。！\n]{10,150})',
    ]

    # 数字承诺模式（含具体数字的承诺更需要锚定）
    NUMBER_PATTERNS = [
        (r'\d+(?:\.\d+)?\s*(?:人|名|位)', '人员数量'),
        (r'\d+(?:\.\d+)?\s*(?:万|元)', '金额'),
        (r'\d+(?:\.\d+)?\s*(?:小时|天|日|月|年)', '时间'),
        (r'\d+(?:\.\d+)?\s*%', '百分比'),
        (r'\d+(?:\.\d+)?\s*(?:套|台|辆|个)', '设备数量'),
    ]

    # 停用词（提取关键词时过滤）
    STOP_WORDS = {
        "的", "了", "在", "是", "我", "有", "和", "就", "不", "人", "都",
        "一", "一个", "上", "也", "很", "到", "说", "要", "去", "你",
        "会", "着", "没有", "看", "好", "自己", "这", "那", "与", "及",
        "或", "等", "为", "以", "于", "对", "从", "向", "把", "被",
        "让", "使", "给", "由", "按", "根据", "按照", "依据",
        "进行", "实施", "开展", "做好", "加强", "推进",
        "本", "该", "其", "之", "所", "可", "能", "需", "应",
        "并", "且", "但", "而", "则", "即", "若", "如",
        "以下", "以上", "以内", "以外", "之前", "之后",
        "方面", "方式", "过程", "情况", "状态", "水平",
    }

    # 最小关键词长度
    MIN_KEYWORD_LEN = 2
    # 锚定匹配阈值
    ANCHOR_THRESHOLD_HIGH = 0.6   # >=60%关键词匹配 = 已锚定
    ANCHOR_THRESHOLD_LOW = 0.3    # 30-60% = 部分锚定
                                    # <30% = 无原文依据

    def __init__(self, tender_text: str = ""):
        """
        Args:
            tender_text: 招标文件原文（纯文本）
        """
        self.tender_text = tender_text or ""
        self.tender_sentences = self._split_sentences(self.tender_text) if tender_text else []

    def extract_commitments(self, content: str) -> List[Dict]:
        """
        从投标文件中提取所有承诺性语句。

        Returns:
            [{"sentence": "...", "verb": "承诺", "has_number": True,
              "number_type": "人员数量", "line": 42, "context": "..."}]
        """
        commitments = []
        lines = content.split("\n")

        for line_idx, line in enumerate(lines, 1):
            line = line.strip()
            if not line or len(line) < 10:
                continue

            # 跳过标题行
            if line.startswith("#") or re.match(r'^第[一二三四五六七八九十\d]+[章节条]', line):
                continue

            # 按句分割
            sentences = re.split(r'[。！；\n]', line)
            for sent in sentences:
                sent = sent.strip()
                if len(sent) < 10:
                    continue

                matched_verb = None
                matched_pattern = None

                # 检查承诺性句式
                for pattern in self.COMMITMENT_PATTERNS:
                    m = re.search(pattern, sent)
                    if m:
                        matched_verb = m.group(0)[:4]
                        matched_pattern = pattern
                        break

                # 检查承诺性动词（直接出现）
                if not matched_verb:
                    for verb in self.COMMITMENT_VERBS:
                        if verb in sent:
                            matched_verb = verb
                            break

                if not matched_verb:
                    continue

                # 检查是否含数字承诺
                number_type = None
                for pattern, ntype in self.NUMBER_PATTERNS:
                    if re.search(pattern, sent):
                        number_type = ntype
                        break

                # 提取上下文（前后各20字）
                sent_idx = line.find(sent)
                ctx_start = max(0, sent_idx - 20)
                ctx_end = min(len(line), sent_idx + len(sent) + 20)
                context = line[ctx_start:ctx_end]

                commitments.append({
                    "sentence": sent,
                    "verb": matched_verb,
                    "has_number": number_type is not None,
                    "number_type": number_type,
                    "line": line_idx,
                    "context": context,
                })

        return commitments

    def extract_keywords(self, sentence: str) -> List[str]:
        """
        从句子中提取核心关键词（用于招标文件搜索）。

        策略：
        1. 去除承诺性动词和停用词
        2. 提取2-4字的实词
        3. 保留专业术语和数字
        """
        keywords = []

        # 提取中文词组（2-4字连续中文）
        words = re.findall(r'[\u4e00-\u9fff]{2,4}', sentence)

        for word in words:
            # 过滤停用词
            if word in self.STOP_WORDS:
                continue
            # 过滤纯承诺动词
            if word in self.COMMITMENT_VERBS:
                continue
            # 过滤太短或太长的
            if len(word) < self.MIN_KEYWORD_LEN:
                continue
            keywords.append(word)

        # 提取数字+单位组合（如"50人""100万"）
        for pattern, _ in self.NUMBER_PATTERNS:
            for m in re.finditer(pattern, sentence):
                keywords.append(m.group())

        # 去重保持顺序
        seen = set()
        unique = []
        for kw in keywords:
            if kw not in seen:
                seen.add(kw)
                unique.append(kw)

        return unique

    def anchor_commitment(self, commitment: Dict) -> Dict:
        """
        将单条承诺锚定到招标文件原文。

        Returns:
            {
                "sentence": "...",
                "verb": "承诺",
                "keywords": ["物业服务", "人员配置"],
                "anchored": True/False,
                "anchor_level": "anchored" / "partial" / "unanchored",
                "matched_tender_text": "...",  # 匹配到的招标文件原文片段
                "match_ratio": 0.8,
            }
        """
        sentence = commitment["sentence"]
        keywords = self.extract_keywords(sentence)

        if not keywords:
            return {
                **commitment,
                "keywords": [],
                "anchored": False,
                "anchor_level": "unanchored",
                "matched_tender_text": "",
                "match_ratio": 0.0,
                "note": "无法提取关键词",
            }

        if not self.tender_text:
            return {
                **commitment,
                "keywords": keywords,
                "anchored": False,
                "anchor_level": "unanchored",
                "matched_tender_text": "",
                "match_ratio": 0.0,
                "note": "无招标文件原文",
            }

        # 在招标文件中搜索每个关键词
        matched_keywords = 0
        best_match_text = ""

        for kw in keywords:
            # 直接搜索
            if kw in self.tender_text:
                matched_keywords += 1
                if not best_match_text:
                    idx = self.tender_text.find(kw)
                    best_match_text = self.tender_text[
                        max(0, idx - 20):idx + len(kw) + 20
                    ]
            else:
                # 模糊搜索：关键词的子串
                if len(kw) >= 3:
                    substr = kw[:2]
                    if substr in self.tender_text:
                        matched_keywords += 0.5
                        if not best_match_text:
                            idx = self.tender_text.find(substr)
                            best_match_text = self.tender_text[
                                max(0, idx - 20):idx + len(substr) + 20
                            ]

        match_ratio = matched_keywords / len(keywords) if keywords else 0

        # 判定锚定级别
        if match_ratio >= self.ANCHOR_THRESHOLD_HIGH:
            anchor_level = "anchored"
            anchored = True
        elif match_ratio >= self.ANCHOR_THRESHOLD_LOW:
            anchor_level = "partial"
            anchored = False
        else:
            anchor_level = "unanchored"
            anchored = False

        return {
            **commitment,
            "keywords": keywords,
            "anchored": anchored,
            "anchor_level": anchor_level,
            "matched_tender_text": best_match_text,
            "match_ratio": round(match_ratio, 2),
        }

    def anchor_all(self, content: str) -> Dict:
        """
        完整锚定比对流程：提取承诺 -> 锚定到招标文件 -> 生成报告。

        Args:
            content: 投标文件文本内容

        Returns:
            {
                "total_commitments": int,
                "anchored": int,
                "partial": int,
                "unanchored": int,
                "anchor_rate": float,  # 锚定率
                "unanchored_with_numbers": int,  # 无依据且含数字的高危承诺
                "details": [...],
                "timestamp": "...",
            }
        """
        commitments = self.extract_commitments(content)

        if not commitments:
            return {
                "total_commitments": 0,
                "anchored": 0,
                "partial": 0,
                "unanchored": 0,
                "anchor_rate": 1.0,
                "unanchored_with_numbers": 0,
                "details": [],
                "message": "未检测到承诺性语句",
            }

        details = []
        for commit in commitments:
            anchored = self.anchor_commitment(commit)
            details.append(anchored)

        # 统计
        anchored_count = sum(1 for d in details if d["anchor_level"] == "anchored")
        partial_count = sum(1 for d in details if d["anchor_level"] == "partial")
        unanchored_count = sum(1 for d in details if d["anchor_level"] == "unanchored")

        # 无依据且含数字的=高危（可能编了具体数据）
        unanchored_numbers = sum(
            1 for d in details
            if d["anchor_level"] == "unanchored" and d.get("has_number", False)
        )

        anchor_rate = anchored_count / len(commitments) if commitments else 0

        return {
            "total_commitments": len(commitments),
            "anchored": anchored_count,
            "partial": partial_count,
            "unanchored": unanchored_count,
            "anchor_rate": round(anchor_rate, 2),
            "unanchored_with_numbers": unanchored_numbers,
            "details": details,
        }

    @staticmethod
    def format_report(report: Dict) -> str:
        """格式化锚定报告为人类可读文本"""
        lines = [
            "=" * 60,
            "原文锚定比对报告",
            "=" * 60,
            "",
        ]

        if report["total_commitments"] == 0:
            lines.append("未检测到承诺性语句")
            return "\n".join(lines)

        # 总览
        rate = report["anchor_rate"]
        if rate >= 0.8:
            rate_icon = "✅"
        elif rate >= 0.5:
            rate_icon = "🟡"
        else:
            rate_icon = "🔴"

        lines.append(f"📊 承诺总数: {report['total_commitments']}")
        lines.append(f"   ✅ 已锚定: {report['anchored']}（有招标文件原文依据）")
        lines.append(f"   🟡 部分锚定: {report['partial']}（关键词部分匹配）")
        lines.append(f"   ⚠️  无原文依据: {report['unanchored']}（招标文件未提及）")
        lines.append(f"   {rate_icon} 锚定率: {rate:.0%}")
        lines.append("")

        if report["unanchored_with_numbers"] > 0:
            lines.append(f"🔴 高危: {report['unanchored_with_numbers']} 条无依据且含具体数字")
            lines.append("   这些承诺含具体数字但招标文件未提及，极可能是AI编造！")
            lines.append("")

        # 无原文依据的承诺（高危优先）
        unanchored = [d for d in report["details"] if d["anchor_level"] == "unanchored"]
        if unanchored:
            lines.append("━" * 40)
            lines.append("⚠️  无原文依据的承诺（需人工核实）")
            lines.append("━" * 40)

            # 按危险程度排序：含数字的优先
            unanchored.sort(key=lambda x: (x.get("has_number", False), x["line"]), reverse=True)

            for d in unanchored[:20]:  # 最多显示20条
                num_flag = " 🔴含数字" if d.get("has_number") else ""
                lines.append(f"  行{d['line']}{num_flag}: {d['sentence'][:60]}...")
                lines.append(f"  关键词: {', '.join(d.get('keywords', [])[:5])}")
                lines.append("")

            if len(unanchored) > 20:
                lines.append(f"  ... 还有 {len(unanchored) - 20} 条未显示")
                lines.append("")

        # 部分锚定
        partial = [d for d in report["details"] if d["anchor_level"] == "partial"]
        if partial:
            lines.append("━" * 40)
            lines.append("🟡 部分锚定的承诺（建议核实）")
            lines.append("━" * 40)
            for d in partial[:10]:
                lines.append(f"  行{d['line']} [{d['match_ratio']:.0%}]: {d['sentence'][:60]}...")
                if d.get("matched_tender_text"):
                    lines.append(f"  招标文件匹配: ...{d['matched_tender_text']}...")
                lines.append("")

        # 已锚定（简要统计）
        anchored = [d for d in report["details"] if d["anchor_level"] == "anchored"]
        if anchored:
            lines.append("━" * 40)
            lines.append(f"✅ 已锚定的承诺: {len(anchored)} 条（有招标文件原文依据）")
            lines.append("━" * 40)

        # 结论
        lines.append("")
        lines.append("=" * 60)
        if rate >= 0.8 and report["unanchored_with_numbers"] == 0:
            lines.append("✅ 锚定率良好，承诺基本都有招标文件依据")
        elif report["unanchored_with_numbers"] > 0:
            lines.append("🔴 发现无原文依据且含数字的承诺，必须人工核实！")
            lines.append("   这些极可能是AI编造的数据，删除或补充真实来源")
        elif rate < 0.5:
            lines.append("⚠️  锚定率偏低，大量承诺无招标文件依据")
            lines.append("   建议逐条核实，删除无依据的承诺")
        else:
            lines.append("🟡 锚定率一般，建议核实无依据的承诺")

        return "\n".join(lines)

    def _split_sentences(self, text: str) -> List[str]:
        """将文本按句分割"""
        return [s.strip() for s in re.split(r'[。！？\n]', text) if s.strip()]
