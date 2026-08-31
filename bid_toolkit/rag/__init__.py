"""bid-toolkit RAG 包：历史标书语义检索（云端多租户 / 双形态部署）。

通道 B（历史投标库）MVP：切块历史标书 + 评分项打标 + 向量索引 +
「输入评分项 → 召回历史标杆章节」。云端优先（pgvector + bge-m3 API），
缺省自动回退本地（chromadb + 本地 bge-m3）。
"""
from __future__ import annotations

from .cli import run
from .config import RAGConfig

__all__ = ["RAGConfig", "run"]
