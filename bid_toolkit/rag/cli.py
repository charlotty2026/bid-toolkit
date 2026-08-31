"""bid rag CLI 入口（由 bid_toolkit.__init__.main 分发调用）。"""
from __future__ import annotations

from .config import RAGConfig
from .embeddings import EmbeddingService
from .ingest import ingest
from .retriever import Retriever
from .vector_store import VectorStore


def run(args) -> None:
    sub = getattr(args, "rag_sub", None)
    if sub == "ingest":
        res = ingest(
            args.path,
            project=args.project,
            score_json=args.score_json,
            limit_pages=args.limit,
        )
        print(
            f"[RAG] 入库完成：项目={res['project']} 切块={res['chunks']} "
            f"评分项标签={res['score_items']} embedding={res['embedding']} "
            f"向量库={res['vector_store']}"
        )
    elif sub == "query":
        cfg = RAGConfig.load()
        embed = EmbeddingService(cfg)
        # 先确定可用性并同步 vector_store，再据此构建 store（降级后应为 bm25）
        _ = embed.available
        store = VectorStore(cfg)
        r = Retriever(store, embed, args.project)
        results = r.retrieve(args.query, top_k=args.top_k, score_item=args.score_item)
        print(f"\n[RAG] 查询「{args.query}」召回 {len(results)} 条：\n")
        for i, it in enumerate(results, 1):
            m = it["meta"]
            print(
                f"── {i}. score={it['score']:.3f} | 评分项={m.get('score_item','')} "
                f"| 来源={m.get('source','')} | 页={m.get('page',0)}"
            )
            print(it["text"][:300])
            print()
    elif sub == "status":
        cfg = RAGConfig.load()
        store = VectorStore(cfg)
        print(
            f"[RAG] 向量库={cfg.vector_store} | 项目={args.project} | "
            f"块数={store.count(args.project)}"
        )
    else:
        print("用法: bid rag {ingest|query|status}")
