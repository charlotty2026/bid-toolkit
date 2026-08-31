"""入库流程：读历史标书 → 切块 → embedding → 写入向量库。"""
from __future__ import annotations

import os
from typing import Any, Dict, List, Optional

from .chunker import chunk_document, load_score_items
from .config import RAGConfig
from .embeddings import EmbeddingService
from .vector_store import VectorStore


def read_document(path: str, limit_pages: int = 0) -> str:
    """读取历史标书文本。PDF 支持 limit_pages 抽样（大文件调试用）。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        import fitz

        doc = fitz.open(path)
        n = len(doc)
        if limit_pages:
            n = min(n, limit_pages)
        texts = [doc[i].get_text() for i in range(n)]
        return "\n".join(texts)
    if ext == ".docx":
        from docx import Document

        d = Document(path)
        return "\n".join(p.text for p in d.paragraphs)
    if ext in (".md", ".txt"):
        return open(path, encoding="utf-8").read()
    raise ValueError(f"不支持的格式: {ext}")


def ingest(
    path: str,
    project: str = "default",
    score_json: Optional[str] = None,
    limit_pages: int = 0,
    cfg: Optional[RAGConfig] = None,
) -> Dict[str, Any]:
    cfg = cfg or RAGConfig.load()
    embed = EmbeddingService(cfg)
    # 先确定真实可用性（会触发逐级降级并把 cfg.vector_store 同步为 bm25/none），
    # 再据此构建向量库，避免降级后仍用 localvec 收 None 向量而崩
    _ = embed.available
    store = VectorStore(cfg)
    text = read_document(path, limit_pages)
    score_items = load_score_items(score_json) if score_json else []
    chunks = chunk_document(
        text,
        source=os.path.basename(path),
        project=project,
        score_items=score_items,
        chunk_size=cfg.chunk_size,
        overlap=cfg.chunk_overlap,
    )
    # embedding 不可用时（默认 none，或 local 探测失败）只写 BM25 索引，不产出向量
    embeddings = embed.embed([c["text"] for c in chunks]) if embed.available else None
    store.add(chunks, embeddings)
    return {
        "project": project,
        "chunks": len(chunks),
        "score_items": len(score_items),
        "embedding": embed.kind,
        "vector_store": cfg.vector_store,
    }
