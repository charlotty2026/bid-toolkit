"""向量库适配器：pgvector（云端多租户）/ chromadb（本地回退）。

多租户隔离：所有写入携带 project 字段，查询时按 project 过滤（pgvector 走
tenant_id 行级隔离语义；chromadb 走 where 元数据过滤）。评分项标签同样作为
元数据，支持「按评分项过滤召回」。
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional

from .bm25 import BM25Index
from .config import RAGConfig

# bge-m3 维度固定 1024
_EMBED_DIM = 1024


class VectorStore:
    """统一检索入口。bm25 为零依赖后端，chromadb / pgvector 为可选增强。"""

    def __init__(self, cfg: RAGConfig):
        self.cfg = cfg
        if cfg.vector_store == "pgvector" and cfg.pg_dsn:
            self._impl = _PgVector(cfg)
        elif cfg.vector_store in ("chromadb", "pgvector"):
            try:
                self._impl = _Chroma(cfg)
            except ImportError:
                print("[RAG] chromadb 未安装，向量库回退本地 localvec")
                self._impl = _LocalVec(cfg)
        elif cfg.vector_store == "localvec":
            try:
                self._impl = _LocalVec(cfg)
            except ImportError:
                print("[RAG] numpy 未安装，向量库回退零依赖 BM25")
                self._impl = _BM25(cfg)
        else:
            self._impl = _BM25(cfg)

    @property
    def backend(self) -> str:
        return type(self._impl).__name__.lstrip("_").lower()

    def add(self, chunks: List[Dict[str, Any]],
            embeddings: Optional[List[List[float]]] = None) -> None:
        self._impl.add(chunks, embeddings)

    def query(self, vector: List[float], top_k: int, project: str,
              score_item: Optional[str] = None) -> List[Dict[str, Any]]:
        return self._impl.query(vector, top_k, project, score_item)

    def query_hybrid(self, text: str, vector: List[float], top_k: int, project: str,
                     score_item: Optional[str] = None) -> List[Dict[str, Any]]:
        """混合检索（向量 RRF 融合 BM25）。后端不支持时回退纯向量检索。"""
        fn = getattr(self._impl, "query_hybrid", None)
        if fn is None:
            return self.query(vector, top_k, project, score_item)
        return fn(text, vector, top_k, project, score_item)

    def query_text(self, text: str, top_k: int, project: str,
                   score_item: Optional[str] = None) -> List[Dict[str, Any]]:
        """纯文本检索（BM25 路径）。向量后端不支持时会给出明确指引。"""
        fn = getattr(self._impl, "query_text", None)
        if fn is None:
            raise RuntimeError(
                "当前向量库不支持纯文本检索。"
                "零依赖用法请设 BID_RAG_EMBED_BACKEND=none，或补齐 embedding 配置。"
            )
        return fn(text, top_k, project, score_item)

    def count(self, project: str) -> int:
        return self._impl.count(project)


class _BM25:
    """零依赖 BM25 后端：无需模型、无需 API Key、无需外部服务。"""

    def __init__(self, cfg: RAGConfig):
        self.index = BM25Index(cfg.bm25_path)

    def add(self, chunks, embeddings=None):
        return self.index.add(chunks)

    def query(self, vector, top_k, project, score_item=None):
        raise RuntimeError("BM25 后端不支持向量检索，请改用 query_text()")

    def query_text(self, text, top_k, project, score_item=None):
        return self.index.query(text, top_k, project, score_item)

    def count(self, project):
        return self.index.count(project)


class _LocalVec:
    """本地向量库：只需 numpy，无外部服务、无数据库。

    同时维护 BM25 倒排索引，检索时把「语义向量 Top-N」与「关键词 Top-N」
    用 RRF（Reciprocal Rank Fusion）融合 —— 招投标文本里专有名词、条款号、
    数字指标很多，纯语义容易漏掉精确匹配，纯关键词又抓不住同义表述，
    融合后两者互补。
    """

    RRF_K = 60          # RRF 平滑常数，抑制头部排名的过度支配
    RECALL_MULT = 3     # 每路先取 top_k*3 再融合，给融合留出重排空间

    def __init__(self, cfg: RAGConfig):
        import numpy as np  # noqa: F401 - 仅探测依赖，缺失时由上层回退 BM25

        self.cfg = cfg
        self.path = cfg.vec_path
        self.bm25 = BM25Index(cfg.bm25_path)
        self._vecs = None
        self._metas: List[Dict[str, Any]] = []
        self._load()

    # ---- 持久化 ----
    def _meta_path(self) -> str:
        return self.path + ".meta.json"

    def _load(self) -> None:
        import json
        import os

        import numpy as np

        if os.path.exists(self.path) and os.path.exists(self._meta_path()):
            with np.load(self.path) as z:
                self._vecs = z["vectors"]
            with open(self._meta_path(), encoding="utf-8") as f:
                self._metas = json.load(f)
        else:
            self._vecs = np.zeros((0, 0), dtype="float32")
            self._metas = []

    def _save(self) -> None:
        import json
        import os

        import numpy as np

        d = os.path.dirname(os.path.abspath(self.path))
        if d:
            os.makedirs(d, exist_ok=True)
        np.savez_compressed(self.path, vectors=self._vecs)
        with open(self._meta_path(), "w", encoding="utf-8") as f:
            json.dump(self._metas, f, ensure_ascii=False)

    # ---- 写入 ----
    def add(self, chunks, embeddings=None):
        import numpy as np

        if not embeddings:
            raise RuntimeError("localvec 需要 embedding 向量")
        new = np.asarray(embeddings, dtype="float32")
        # 归一化，使点积等价于余弦相似度
        norms = np.linalg.norm(new, axis=1, keepdims=True)
        new = new / np.clip(norms, 1e-12, None)

        if self._vecs is None or self._vecs.size == 0:
            self._vecs = new
        else:
            if self._vecs.shape[1] != new.shape[1]:
                raise RuntimeError(
                    f"向量维度不一致：索引 {self._vecs.shape[1]} vs 新数据 {new.shape[1]}。"
                    f"更换 embedding 模型后请删除 {self.path} 重新入库。"
                )
            self._vecs = np.concatenate([self._vecs, new], axis=0)

        for c in chunks:
            self._metas.append(
                {
                    "text": c["text"],
                    "project": c["project"],
                    "score_item": c.get("score_item", ""),
                    "source": c.get("source", ""),
                    "page": c.get("page", 0),
                }
            )
        self._save()
        # 同一批内容双写 BM25，为混合检索备好关键词侧
        self.bm25.add(chunks)

    # ---- 检索 ----
    def _candidates(self, project, score_item):
        idx = []
        for i, m in enumerate(self._metas):
            if m["project"] != project:
                continue
            if score_item and m.get("score_item") != score_item:
                continue
            idx.append(i)
        return idx

    def _vector_rank(self, vector, top_n, project, score_item):
        import numpy as np

        cand = self._candidates(project, score_item)
        if not cand or self._vecs is None or self._vecs.size == 0:
            return []
        q = np.asarray(vector, dtype="float32")
        q = q / max(float(np.linalg.norm(q)), 1e-12)
        sims = self._vecs[cand] @ q
        order = np.argsort(-sims)[:top_n]
        return [(cand[o], float(sims[o])) for o in order]

    def query(self, vector, top_k, project, score_item=None):
        """纯语义向量检索（无查询文本时使用）。"""
        hits = self._vector_rank(vector, top_k, project, score_item)
        out = []
        for i, sim in hits:
            m = self._metas[i]
            out.append(
                {
                    "text": m["text"],
                    "meta": {
                        "score_item": m.get("score_item", ""),
                        "source": m.get("source", ""),
                        "page": m.get("page", 0),
                    },
                    "score": round(sim, 6),
                }
            )
        return out

    def query_hybrid(self, text, vector, top_k, project, score_item=None):
        """显式混合检索：同时给出查询文本与其向量。"""
        top_n = top_k * self.RECALL_MULT
        vec_hits = self._vector_rank(vector, top_n, project, score_item)
        kw_hits = self.bm25.query(text, top_n, project, score_item)

        # BM25 返回的是文本，用 (project, text) 定位回索引下标
        pos = {(m["project"], m["text"]): i for i, m in enumerate(self._metas)}

        fused: Dict[int, float] = {}
        vscore: Dict[int, float] = {}
        kscore: Dict[int, float] = {}
        for rank, (i, sim) in enumerate(vec_hits):
            fused[i] = fused.get(i, 0.0) + 1.0 / (self.RRF_K + rank + 1)
            vscore[i] = sim
        for rank, h in enumerate(kw_hits):
            i = pos.get((project, h.get("text", "")))
            if i is None:
                continue
            fused[i] = fused.get(i, 0.0) + 1.0 / (self.RRF_K + rank + 1)
            kscore[i] = float(h.get("score", 0.0))

        out = []
        for i, score in sorted(fused.items(), key=lambda kv: -kv[1])[:top_k]:
            m = self._metas[i]
            out.append(
                {
                    "text": m["text"],
                    "meta": {
                        "score_item": m.get("score_item", ""),
                        "source": m.get("source", ""),
                        "page": m.get("page", 0),
                        "vector_score": round(vscore.get(i, 0.0), 4),
                        "bm25_score": round(kscore.get(i, 0.0), 4),
                    },
                    "score": round(score, 6),
                }
            )
        return out

    def count(self, project):
        return sum(1 for m in self._metas if m["project"] == project)


class _Chroma:
    def __init__(self, cfg: RAGConfig):
        import chromadb

        self.client = chromadb.PersistentClient(path=cfg.chroma_path)
        self.coll = self.client.get_or_create_collection(
            "bid_rag", metadata={"hnsw:space": "cosine"}
        )

    def add(self, chunks, embeddings):
        ids = [f"{c['project']}:{c['id']}" for c in chunks]
        metas = [
            {
                "project": c["project"],
                "score_item": c.get("score_item", ""),
                "source": c.get("source", ""),
                "page": c.get("page", 0),
            }
            for c in chunks
        ]
        docs = [c["text"] for c in chunks]
        self.coll.upsert(ids=ids, embeddings=embeddings, documents=docs, metadatas=metas)

    def query(self, vector, top_k, project, score_item):
        where: Dict[str, Any] = {"project": project}
        if score_item:
            where["score_item"] = score_item
        res = self.coll.query(query_embeddings=[vector], n_results=top_k, where=where)
        out = []
        for doc, meta, dist in zip(
            res["documents"][0], res["metadatas"][0], res["distances"][0]
        ):
            out.append({"text": doc, "meta": meta, "score": 1 - dist})
        return out

    def count(self, project):
        return self.coll.count()


class _PgVector:
    def __init__(self, cfg: RAGConfig):
        from sqlalchemy import create_engine, text

        if not cfg.pg_dsn:
            raise RuntimeError("pgvector 需要 BID_RAG_PG_DSN")
        self.engine = create_engine(cfg.pg_dsn)
        self._table = cfg.pg_table
        with self.engine.begin() as conn:
            conn.execute(text("CREATE EXTENSION IF NOT EXISTS vector;"))
            conn.execute(
                text(
                    f"""
                    CREATE TABLE IF NOT EXISTS {self._table} (
                        id SERIAL PRIMARY KEY,
                        project TEXT NOT NULL,
                        chunk_id TEXT NOT NULL,
                        score_item TEXT,
                        source TEXT,
                        page INT,
                        content TEXT,
                        embedding vector({_EMBED_DIM})
                    );
                    CREATE INDEX IF NOT EXISTS {self._table}_proj_idx
                        ON {self._table} (project);
                    """
                )
            )

    def add(self, chunks, embeddings):
        from sqlalchemy import text

        with self.engine.begin() as conn:
            for c, e in zip(chunks, embeddings):
                conn.execute(
                    text(
                        f"""
                        INSERT INTO {self._table}
                            (project, chunk_id, score_item, source, page, content, embedding)
                        VALUES
                            (:p, :cid, :si, :src, :pg, :ct, :emb::vector)
                        """
                    ),
                    {
                        "p": c["project"],
                        "cid": c["id"],
                        "si": c.get("score_item", ""),
                        "src": c.get("source", ""),
                        "pg": c.get("page", 0),
                        "ct": c["text"],
                        "emb": "[" + ",".join(str(x) for x in e) + "]",
                    },
                )

    def query(self, vector, top_k, project, score_item):
        from sqlalchemy import text

        emb = "[" + ",".join(str(x) for x in vector) + "]"
        sql = f"""
            SELECT content, score_item, source, page,
                   1 - (embedding <=> :emb::vector) AS score
            FROM {self._table}
            WHERE project = :p
        """
        params = {"emb": emb, "p": project}
        if score_item:
            sql += " AND score_item = :si"
            params["si"] = score_item
        sql += " ORDER BY embedding <=> :emb::vector LIMIT :k"
        params["k"] = top_k
        with self.engine.connect() as conn:
            rows = conn.execute(text(sql), params).fetchall()
        return [
            {
                "text": r[0],
                "meta": {"score_item": r[1], "source": r[2], "page": r[3]},
                "score": float(r[4]),
            }
            for r in rows
        ]

    def count(self, project):
        from sqlalchemy import text

        with self.engine.connect() as conn:
            return conn.execute(
                text(f"SELECT count(*) FROM {self._table} WHERE project=:p"),
                {"p": project},
            ).scalar()
