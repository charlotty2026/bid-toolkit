"""零依赖 BM25 检索后端（embedding_backend=none 时的检索路径）。

刻意不引入 jieba / sentence-transformers / chromadb：评审或用户克隆仓库后，
不装任何可选依赖、不配任何 API Key 也能跑通 `bid rag ingest/query`。

分词策略：中文按 bigram 切分，英文与数字按连续片段成词。bigram 对中文
短查询的召回效果接近分词器，且完全无依赖、无模型下载。

索引以 JSON 持久化，加载时重建内存结构（标书场景分块量级在千级，重建开销可忽略）。
"""
from __future__ import annotations

import json
import math
import os
import re
from collections import Counter
from typing import Any, Dict, List, Optional

_TOKEN_RE = re.compile(r"[\u4e00-\u9fff]+|[a-zA-Z0-9]+")


def tokenize(text: str) -> List[str]:
    """中文 bigram + 英文数字整词。"""
    toks: List[str] = []
    for seg in _TOKEN_RE.findall((text or "").lower()):
        if "\u4e00" <= seg[0] <= "\u9fff":
            if len(seg) == 1:
                toks.append(seg)
            else:
                toks.extend(seg[i:i + 2] for i in range(len(seg) - 1))
        else:
            toks.append(seg)
    return toks


class BM25Index:
    """BM25 倒排索引，支持 project 与 score_item 两级过滤。"""

    def __init__(self, path: str = "./.rag_bm25.json", k1: float = 1.5, b: float = 0.75):
        self.path = path
        self.k1 = k1
        self.b = b
        self.docs: List[Dict[str, Any]] = []
        self.doc_tf: List[Counter] = []
        self.doc_len: List[int] = []
        self.df: Counter = Counter()
        self.avgdl = 0.0
        self._load()

    # ---- 持久化 ----
    def _load(self) -> None:
        if not os.path.exists(self.path):
            return
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (json.JSONDecodeError, OSError):
            return
        if not isinstance(data, dict) or not isinstance(data.get("docs"), list):
            return
        self.docs = data["docs"]
        self._rebuild()

    def save(self) -> None:
        _dir = os.path.dirname(os.path.abspath(self.path))
        if _dir:
            os.makedirs(_dir, exist_ok=True)
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(
                {"version": 1, "k1": self.k1, "b": self.b, "docs": self.docs},
                f,
                ensure_ascii=False,
            )

    def _rebuild(self) -> None:
        self.doc_tf = []
        self.doc_len = []
        self.df = Counter()
        for d in self.docs:
            tf = Counter(tokenize(d.get("text", "")))
            self.doc_tf.append(tf)
            self.doc_len.append(sum(tf.values()))
            for term in tf:
                self.df[term] += 1
        total = sum(self.doc_len)
        self.avgdl = total / len(self.docs) if self.docs else 0.0

    # ---- 写入 ----
    def add(self, chunks: List[Dict[str, Any]]) -> int:
        """chunks: [{id, project, text, score_item, source, page}]，按 id 去重更新。"""
        existing = {(d.get("project"), d.get("id")): i for i, d in enumerate(self.docs)}
        added = 0
        for c in chunks:
            key = (c.get("project"), c.get("id"))
            rec = {
                "id": c.get("id"),
                "project": c.get("project"),
                "text": c.get("text", ""),
                "score_item": c.get("score_item", ""),
                "source": c.get("source", ""),
                "page": c.get("page", 0),
            }
            if key in existing:
                self.docs[existing[key]] = rec
            else:
                existing[key] = len(self.docs)
                self.docs.append(rec)
                added += 1
        self._rebuild()
        self.save()
        return added

    # ---- 检索 ----
    def query(
        self,
        text: str,
        top_k: int = 5,
        project: Optional[str] = None,
        score_item: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        q_tokens = tokenize(text)
        if not q_tokens or not self.docs:
            return []

        candidates = [
            i
            for i, d in enumerate(self.docs)
            if (project is None or d.get("project") == project)
            and (not score_item or d.get("score_item") == score_item)
        ]
        if not candidates:
            return []

        n = len(self.docs)
        scored: List[tuple] = []
        for i in candidates:
            tf = self.doc_tf[i]
            dl = self.doc_len[i] or 1
            score = 0.0
            for qt in q_tokens:
                f = tf.get(qt, 0)
                if not f:
                    continue
                df = self.df.get(qt, 0)
                idf = math.log(1 + (n - df + 0.5) / (df + 0.5))
                denom = f + self.k1 * (1 - self.b + self.b * dl / (self.avgdl or 1.0))
                score += idf * (f * (self.k1 + 1)) / denom
            if score > 0:
                scored.append((score, i))

        scored.sort(key=lambda x: -x[0])
        out = []
        for score, i in scored[:top_k]:
            d = self.docs[i]
            out.append(
                {
                    "text": d.get("text", ""),
                    "meta": {
                        "score_item": d.get("score_item", ""),
                        "source": d.get("source", ""),
                        "page": d.get("page", 0),
                    },
                    "score": float(score),
                }
            )
        return out

    def count(self, project: Optional[str] = None) -> int:
        if project is None:
            return len(self.docs)
        return sum(1 for d in self.docs if d.get("project") == project)

    def clear(self, project: Optional[str] = None) -> int:
        before = len(self.docs)
        if project is None:
            self.docs = []
        else:
            self.docs = [d for d in self.docs if d.get("project") != project]
        self._rebuild()
        self.save()
        return before - len(self.docs)
