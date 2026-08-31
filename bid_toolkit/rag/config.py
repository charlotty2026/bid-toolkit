"""bid-toolkit RAG 配置。

设计原则（对齐飞书项目书 v0.4 第 17 章 + 云端多租户方向）：
- embedding 后端：cloud（云端 bge-m3 API，卖产品默认）/ local（本地 sentence-transformers）/ none（BM25 降级）
- 向量库：pgvector（云端多租户默认，tenant_id 行级隔离）/ chromadb（本地回退、零外部依赖）
- 双形态部署：同一套代码，靠配置切换 SaaS 云端 / 企业私有化
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Optional


def default_ov_dir() -> str:
    """OpenVINO IR 默认缓存目录（用户级，避免污染工作区与 Skill 包）。"""
    return os.path.join(
        os.path.expanduser("~"), ".cache", "bid_toolkit", "ov", "bge-small-zh-v1.5"
    )


@dataclass
class RAGConfig:
    # embedding 后端（默认 none：零依赖 BM25，克隆即用；其余为可选增强）
    embedding_backend: str = "none"            # none | openvino | local | cloud
    # 云端 embedding（OpenAI 兼容 /embeddings 接口，bge-m3）
    cloud_embed_base_url: str = ""             # 如 https://api.siliconflow.cn/v1
    cloud_embed_api_key: str = ""
    cloud_embed_model: str = "bge-m3"
    # 本地 embedding（sentence-transformers）
    local_embed_model: str = "BAAI/bge-small-zh-v1.5"   # 小模型优先，约 90MB
    # OpenVINO 本地推理（AI PC：优先 iGPU/NPU 卸载，CPU 兜底）
    ov_model_dir: str = ""                     # 空则用默认缓存目录，见 default_ov_dir()
    ov_device: str = "AUTO:GPU,CPU"            # AUTO 让 OpenVINO 挑最快的可用设备
    ov_max_length: int = 512                   # tokenizer 截断长度
    # 向量库（默认 bm25：零依赖，与 embedding_backend=none 配套）
    vector_store: str = "bm25"                 # bm25 | localvec | pgvector | chromadb
    pg_dsn: str = ""                           # PostgreSQL 连接串（含 pgvector 扩展）
    pg_table: str = "bid_rag_chunks"
    chroma_path: str = "./.rag_chroma"         # chromadb 本地持久化目录
    bm25_path: str = "./.rag_bm25.json"        # 零依赖 BM25 索引文件
    vec_path: str = "./.rag_vec.npz"           # 本地向量索引（localvec，仅需 numpy）
    # 切块
    chunk_size: int = 800
    chunk_overlap: int = 120

    @classmethod
    def load(cls, path: Optional[str] = None) -> "RAGConfig":
        """从默认值 + 环境变量加载。支持云端优先、缺省自动回退本地。

        环境变量：
          BID_RAG_EMBED_BACKEND        none|openvino|local|cloud（默认 none，零依赖 BM25）
          BID_RAG_CLOUD_EMBED_BASE_URL 云端 embedding 基址
          BID_RAG_CLOUD_EMBED_API_KEY  云端 embedding API Key
          BID_RAG_LOCAL_EMBED_MODEL    本地模型名（默认 BAAI/bge-small-zh-v1.5）
          BID_RAG_OV_MODEL_DIR         OpenVINO IR 目录（默认见 default_ov_dir）
          BID_RAG_OV_DEVICE            OpenVINO 设备（默认 AUTO:GPU,CPU）
          BID_RAG_VECTOR_STORE         bm25|pgvector|chromadb（默认 bm25）
          BID_RAG_PG_DSN               PostgreSQL 连接串
          BID_RAG_BM25_PATH            BM25 索引文件路径
        """
        cfg = cls()
        env = os.environ
        if env.get("BID_RAG_EMBED_BACKEND"):
            cfg.embedding_backend = env["BID_RAG_EMBED_BACKEND"]
        if env.get("BID_RAG_CLOUD_EMBED_BASE_URL"):
            cfg.cloud_embed_base_url = env["BID_RAG_CLOUD_EMBED_BASE_URL"]
        if env.get("BID_RAG_CLOUD_EMBED_API_KEY"):
            cfg.cloud_embed_api_key = env["BID_RAG_CLOUD_EMBED_API_KEY"]
        if env.get("BID_RAG_LOCAL_EMBED_MODEL"):
            cfg.local_embed_model = env["BID_RAG_LOCAL_EMBED_MODEL"]
        if env.get("BID_RAG_OV_MODEL_DIR"):
            cfg.ov_model_dir = env["BID_RAG_OV_MODEL_DIR"]
        if env.get("BID_RAG_OV_DEVICE"):
            cfg.ov_device = env["BID_RAG_OV_DEVICE"]
        if env.get("BID_RAG_VECTOR_STORE"):
            cfg.vector_store = env["BID_RAG_VECTOR_STORE"]
        if env.get("BID_RAG_PG_DSN"):
            cfg.pg_dsn = env["BID_RAG_PG_DSN"]
        if env.get("BID_RAG_BM25_PATH"):
            cfg.bm25_path = env["BID_RAG_BM25_PATH"]
        if env.get("BID_RAG_VEC_PATH"):
            cfg.vec_path = env["BID_RAG_VEC_PATH"]

        # 自动回退：云端 embedding 需要 Key，缺失则整体回退零依赖 BM25
        if cfg.embedding_backend == "cloud" and not cfg.cloud_embed_api_key:
            print("[RAG] 未配置 BID_RAG_CLOUD_EMBED_API_KEY，检索回退零依赖 BM25")
            cfg.embedding_backend = "none"

        if not cfg.ov_model_dir:
            cfg.ov_model_dir = default_ov_dir()

        # 一致性收敛：无 embedding 时向量检索无从谈起，强制走 BM25 索引
        if cfg.embedding_backend == "none":
            cfg.vector_store = "bm25"
        # 有 embedding 却仍是 bm25，说明用户只切了后端；升级到零依赖本地向量库，
        # 否则算出来的向量无处存放（localvec 只需 numpy，仍属纯本地）
        elif cfg.vector_store == "bm25":
            cfg.vector_store = "localvec"
        # 自动回退：pgvector 需要 DSN，缺失则回退 chromadb
        elif cfg.vector_store == "pgvector" and not cfg.pg_dsn:
            cfg.vector_store = "chromadb"
            print("[RAG] 未配置 BID_RAG_PG_DSN，向量库自动回退 chromadb（本地）")
        return cfg
