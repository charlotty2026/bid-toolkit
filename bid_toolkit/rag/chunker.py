"""切块与评分项打标。

通道 B（历史投标库）切块策略（对齐项目书第 17 章）：招标文件/历史标书按
「评分项」切块而非固定长度 —— 保留评分项完整性。每块打上最相关的评分项
标签，供后续「按评分项过滤召回」使用。
"""
from __future__ import annotations

import json
from typing import Any, Dict, List


def load_score_items(score_json_path: str) -> List[str]:
    """从评标办法评分项 JSON 提取评审因素，作为切块标签字典。"""
    try:
        data = json.load(open(score_json_path, encoding="utf-8"))
    except Exception:
        return []
    items: List[str] = []
    for d in data.get("明细", []):
        f = d.get("评审因素", "")
        if f:
            items.append(f)
    # 去重保序
    seen = set()
    uniq = []
    for i in items:
        if i not in seen:
            seen.add(i)
            uniq.append(i)
    return uniq


def chunk_text(text: str, chunk_size: int = 800, overlap: int = 120) -> List[str]:
    """简单滑动窗口切块（无外部依赖）。chunk_size<=0 时整体为一块。"""
    if chunk_size <= 0:
        return [text]
    chunks: List[str] = []
    start = 0
    step = max(1, chunk_size - overlap)
    n = len(text)
    while start < n:
        chunks.append(text[start : start + chunk_size])
        start += step
    return chunks


def chunk_document(
    text: str,
    source: str,
    project: str,
    score_items: List[str],
    chunk_size: int = 800,
    overlap: int = 120,
    page_offset: int = 0,
) -> List[Dict[str, Any]]:
    """切块并打评分项标签：每块匹配命中的评分项（包含即标），取首个为主标签。"""
    raw = chunk_text(text, chunk_size, overlap)
    out: List[Dict[str, Any]] = []
    for i, c in enumerate(raw):
        matched = [si for si in score_items if si and si in c]
        label = matched[0] if matched else ""
        out.append(
            {
                "id": f"{source}-{page_offset}-{i}",
                "text": c,
                "project": project,
                "source": source,
                "page": page_offset,
                "score_item": label,
                "score_items": matched,
            }
        )
    return out
