# -*- coding: utf-8 -*-
"""
投标人身份锚定 (Bidder Identity Anchor)
========================================

铁律第⑨项的前置数据源。

根因：Agent"知道"投标人身份，但写标书时不会自动激活这个认知。
      模型在生成文本时优先生成"看起来流畅"的句子，而非"业务逻辑正确"的句子。
      "与医院签订劳动合同"语言上通顺，但业务逻辑完全反了。

解法：写标书前强制填写身份卡，引擎在内容生成后逐条比对，
      发现违反身份关系的表述直接打回（fatal级，不通过=废标风险）。

身份卡结构（JSON文件，用 profile --set-file 加载）：
{
    "bidder_name": "XX公司",
    "business_type": "XX服务外包",
    "role": "用人单位",           # 投标人在用工关系中的角色
    "client_role": "用工单位",     # 招标方在用工关系中的角色
    "contract_relationship": "...", # 一句话描述合同关系
    "forbidden_expressions": [...], # 绝对不能出现的表述
    "must_have_facts": {...},       # 必须正确体现的事实
}

内置预设：
  - generic: 通用模板（需手动填写）

企业专属身份卡请放在不入库的本地路径，用 --set-file 加载。
"""

import json
from pathlib import Path
from typing import Optional, Dict, List


# ===== 预设身份卡 =====
# 内置只有 generic 通用模板。
# 企业专属身份卡（含公司名称/用工关系等）请用 --set-file <json路径> 加载，
# 不要把真实企业信息写进代码（开源安全）。

PRESET_PROFILES = {
    "generic": {
        "bidder_name": "（请填写投标人名称）",
        "business_type": "（请填写业务类型）",
        "role": "（请填写：用人单位/服务提供方/承包方）",
        "client_role": "（请填写：用工单位/采购人/发包方）",
        "contract_relationship": "（请描述合同关系）",
        "forbidden_expressions": [],
        "must_have_facts": {},
        "notes": [],
    },
}


class BidderProfile:
    """投标人身份卡管理"""

    def __init__(self, profile_data: Optional[dict] = None):
        self.data = profile_data or {}

    @classmethod
    def from_preset(cls, preset_name: str) -> "BidderProfile":
        """从预设加载身份卡"""
        if preset_name not in PRESET_PROFILES:
            raise ValueError(
                f"未知预设: {preset_name}，可选: {list(PRESET_PROFILES.keys())}"
            )
        return cls(PRESET_PROFILES[preset_name])

    @classmethod
    def from_file(cls, path: Path) -> "BidderProfile":
        """从JSON文件加载身份卡"""
        with open(path, "r", encoding="utf-8") as f:
            return cls(json.load(f))

    def save(self, path: Path):
        """保存身份卡到JSON文件"""
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def is_complete(self) -> bool:
        """检查身份卡是否填写完整"""
        if not self.data:
            return False
        required = ["bidder_name", "role", "client_role", "contract_relationship"]
        for key in required:
            val = self.data.get(key, "")
            if not val or val.startswith("（请填写"):
                return False
        return True

    def check_content(self, content: str) -> List[dict]:
        """
        检查内容是否违反身份卡

        Returns:
            issues列表，每个issue包含:
            - type: 问题类型
            - expression: 触发的表述
            - context: 上下文
            - severity: "fatal"（废标级）
            - suggestion: 修正建议
        """
        issues = []

        if not self.is_complete():
            return [{
                "type": "身份卡未填写",
                "severity": "fatal",
                "suggestion": "请先执行 profile --set <预设名> 设置投标人身份卡，再写标书内容",
            }]

        # 1. 检查禁止表述
        for expr in self.data.get("forbidden_expressions", []):
            if expr in content:
                idx = content.find(expr)
                context = content[max(0, idx - 20):idx + len(expr) + 20]
                issues.append({
                    "type": "身份关系错误",
                    "expression": expr,
                    "context": context,
                    "severity": "fatal",
                    "suggestion": f"删除「{expr}」--{self.data['bidder_name']}是{self.data['role']}，{self.data.get('client_role', '招标方')}是{self.data.get('client_role', '用工单位')}",
                })

        # 2. 检查"签订劳动合同"的主语是否正确
        # 如果内容提到"签订劳动合同"但没提到投标人名称，标红提醒
        bidder_name = self.data.get("bidder_name", "")
        short_name = bidder_name.replace("上海", "").replace("有限公司", "").replace("股份", "") if bidder_name else ""

        if "签订劳动合同" in content or "签劳动合同" in content:
            # 检查是否在投标人名称附近出现
            contract_mentions = []
            for pattern in ["签订劳动合同", "签劳动合同"]:
                start = 0
                while True:
                    idx = content.find(pattern, start)
                    if idx == -1:
                        break
                    # 看前50字内有没有投标人名称或简称
                    nearby = content[max(0, idx - 50):idx]
                    has_bidder = any(name in nearby for name in [bidder_name, short_name] if name)
                    contract_mentions.append((idx, has_bidder, nearby))
                    start = idx + 1

            for idx, has_bidder, nearby in contract_mentions:
                if not has_bidder:
                    context = content[max(0, idx - 30):idx + 20]
                    issues.append({
                        "type": "劳动合同签订方不明",
                        "context": context,
                        "severity": "fatal",
                        "suggestion": f"「签订劳动合同」的主语应为{bidder_name}（{self.data['role']}），而非{self.data.get('client_role', '招标方')}",
                    })

        # 3. 检查"用人单位"/"用工单位"是否用对
        if self.data.get("role") and self.data.get("client_role"):
            # 如果内容提到"用人单位"但没关联到投标人，提醒
            if "用人单位" in content:
                idx = content.find("用人单位")
                nearby = content[max(0, idx - 30):idx + 10]
                if bidder_name and bidder_name not in nearby and short_name not in nearby:
                    issues.append({
                        "type": "用人单位指向不明",
                        "context": content[max(0, idx - 20):idx + 20],
                        "severity": "high",
                        "suggestion": f"「用人单位」应指向{bidder_name}，确认上下文没有把{self.data.get('client_role', '招标方')}写成用人单位",
                    })

        return issues

    def summary(self) -> str:
        """返回身份卡摘要文本"""
        if not self.is_complete():
            return "⚠️ 身份卡未填写，请先执行 profile --set <预设名>"

        lines = [
            "=" * 50,
            "投标人身份卡",
            "=" * 50,
            f"投标人: {self.data.get('bidder_name', '?')}",
            f"业务类型: {self.data.get('business_type', '?')}",
            f"角色: {self.data['role']}",
            f"招标方角色: {self.data['client_role']}",
            f"合同关系: {self.data.get('contract_relationship', '?')}",
            "",
        ]

        forbidden = self.data.get("forbidden_expressions", [])
        if forbidden:
            lines.append(f"禁止表述 ({len(forbidden)}条):")
            for expr in forbidden:
                lines.append(f"  ❌ {expr}")
            lines.append("")

        facts = self.data.get("must_have_facts", {})
        if facts:
            lines.append("必须正确体现的事实:")
            for key, val in facts.items():
                lines.append(f"  ✅ {key}: {val}")
            lines.append("")

        notes = self.data.get("notes", [])
        if notes:
            lines.append("注意事项:")
            for note in notes:
                lines.append(f"  ⚠️ {note}")

        lines.append("=" * 50)
        return "\n".join(lines)
