"""检索：向量召回 + 评分项元数据过滤。

MVP 先做向量召回（云端 bge-m3 / 本地 bge-m3）。后续接入 BM25 混合检索
（复用 bid_search 的段落级全文索引）+ RRF 融合，对齐《擎标历史RAG方案设计》
第 3 节「混合检索」设计。
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional

from .embeddings import EmbeddingService
from .vector_store import VectorStore


class Retriever:
    def __init__(self, store: VectorStore, embed: EmbeddingService, project: str):
        self.store = store
        self.embed = embed
        self.project = project

    def retrieve(
        self, query: str, top_k: int = 5, score_item: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        # embedding 不可用时（默认 none，或 local 探测失败）走零依赖 BM25
        if not self.embed.available:
            return self.store.query_text(query, top_k, self.project, score_item)
        vec = self.embed.embed([query])[0]
        # 优先混合检索（localvec 同时持有向量与 BM25，RRF 融合更稳；
        # 若后端无 query_hybrid，VectorStore 会自动回退纯向量检索）
        return self.store.query_hybrid(query, vec, top_k, self.project, score_item)
