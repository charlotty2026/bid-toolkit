"""Embedding 服务：四种后端 + 逐级降级，缺任何可选依赖都不会中断主流程。

  openvino  本地 OpenVINO IR 推理，AI PC 上优先卸载到 iGPU/NPU（推荐）
  local     sentence-transformers（PyTorch）
  cloud     OpenAI 兼容 /embeddings 接口（bge-m3，1024 维）
  none      不做向量化，检索走零依赖 BM25（默认，克隆即用）

降级链：openvino -> local -> none(BM25)
"""
from __future__ import annotations

import os
from typing import List

from .config import RAGConfig


class EmbeddingService:
    def __init__(self, cfg: RAGConfig):
        self.cfg = cfg
        self._model = None
        self._ov = None          # (compiled_model, tokenizer, dim, device)
        backend = cfg.embedding_backend
        if backend == "cloud":
            if not cfg.cloud_embed_api_key:
                raise RuntimeError("云端 embedding 需要 BID_RAG_CLOUD_EMBED_API_KEY")
            self._kind = "cloud"
        elif backend in ("openvino", "local", "none"):
            self._kind = backend
        else:
            raise ValueError(f"未知 embedding 后端: {backend}")

    # ---- 公共接口 ----
    def embed(self, texts: List[str]) -> List[List[float]]:
        if self._kind == "none":
            raise RuntimeError("embedding 后端为 none，请使用 BM25-only 检索")
        if self._kind == "cloud":
            return self._embed_cloud(texts)
        if self._kind == "openvino":
            return self._embed_openvino(texts)
        return self._embed_local(texts)

    @property
    def available(self) -> bool:
        """当前 embedding 是否真正可用。

        openvino / local 后端在此延迟探测依赖与模型。任何一环缺失
        （未装 openvino、IR 未构建、未装 sentence-transformers、模型
        下载失败……）都会逐级降级，最终落到 none（BM25）并打印提示，
        保证上层永不因缺少可选依赖而崩溃。

        降级链：openvino -> local -> none(BM25)
        """
        if self._kind == "none":
            return False
        if self._kind == "cloud":
            return bool(self.cfg.cloud_embed_api_key)
        if self._kind == "openvino":
            if self._ensure_openvino():
                return True
            # OpenVINO 不可用：退一步试 sentence-transformers
            self._kind = "local"
        ok = self._ensure_local()
        if not ok:
            # 彻底降级为零依赖 BM25：同步把向量库也退回 bm25，
            # 否则上层会拿 None 向量硬塞给 localvec 而崩溃
            self._kind = "none"
            self.cfg.vector_store = "bm25"
        return ok

    @property
    def dim(self) -> int:
        if self._kind == "cloud":
            return 1024  # bge-m3
        if not self.available:
            return 0
        if self._kind == "openvino":
            return self._ov[2]
        if self._kind == "local":
            return self._local_model().get_sentence_embedding_dimension()
        return 0

    @property
    def kind(self) -> str:
        return self._kind

    @property
    def device(self) -> str:
        """OpenVINO 实际执行设备；非 openvino 后端返回空串。"""
        return self._ov[3] if self._ov else ""

    def _ensure_local(self) -> bool:
        """探测本地模型可用性；不可用时把后端降级为 none。"""
        if self._kind != "local":
            return False
        try:
            self._local_model()
            return True
        except Exception as exc:  # noqa: BLE001 - 任何加载失败都应降级而非崩溃
            print(
                f"[RAG] 本地 embedding 不可用（{type(exc).__name__}: {exc}），"
                f"已回退零依赖 BM25。\n"
                f"      如需语义检索：pip install sentence-transformers，"
                f"并确认能访问模型仓库。"
            )
            self._kind = "none"
            return False

    # ---- OpenVINO（AI PC：iGPU / NPU 卸载，CPU 兜底）----
    def _ensure_openvino(self) -> bool:
        """加载 OpenVINO IR 与 tokenizer；任一环失败返回 False 交由上层降级。

        运行期 tokenizer 仅依赖轻量 `tokenizers`（Rust 后端），不拉 torch /
        transformers，符合 AI PC 本地轻量推理定位；IR 构建期才需要 torch。
        """
        if self._ov is not None:
            return True
        model_dir = self.cfg.ov_model_dir
        xml = os.path.join(model_dir, "bge_small_zh.xml")
        try:
            if not os.path.exists(xml):
                raise FileNotFoundError(
                    f"未找到 IR：{xml}（先运行 python tools/build_ov_model.py 构建）"
                )
            import openvino as ov
            from tokenizers import Tokenizer

            core = ov.Core()
            compiled = core.compile_model(xml, self.cfg.ov_device)
            tok = Tokenizer.from_file(os.path.join(model_dir, "tokenizer.json"))
            dim = compiled.outputs[0].get_partial_shape()[-1].get_length()
            try:
                device = compiled.get_property("EXECUTION_DEVICES")
                if isinstance(device, (list, tuple)):
                    device = ",".join(str(d) for d in device)
            except Exception:  # noqa: BLE001 - 属性名随版本变化，取不到不影响推理
                device = self.cfg.ov_device
            self._ov = (compiled, tok, int(dim), str(device))
            print(f"[RAG] OpenVINO 就绪：设备={device} 维度={dim} IR={model_dir}")
            return True
        except Exception as exc:  # noqa: BLE001 - 逐级降级，不中断主流程
            print(
                f"[RAG] OpenVINO 不可用（{type(exc).__name__}: {exc}），"
                f"尝试回退 sentence-transformers。"
            )
            return False

    def _embed_openvino(self, texts: List[str]) -> List[List[float]]:
        import numpy as np

        compiled, tok, _dim, _dev = self._ov
        # tokenizers（Rust）原生加载 IR 同款 tokenizer.json，与构建期分词一致
        tok.enable_truncation(self.cfg.ov_max_length)
        tok.enable_padding()
        encs = tok.encode_batch(list(texts), add_special_tokens=True)
        input_ids, attn, ttype = [], [], []
        for e in encs:
            ids = e.ids
            input_ids.append(ids)
            attn.append(e.attention_mask)
            # bge 为单句，token_type_ids 全零；tokenizers 可能不产出 type_ids
            ttype.append(e.type_ids if e.type_ids else [0] * len(ids))
        feed = {
            "input_ids": np.array(input_ids, dtype="int64"),
            "attention_mask": np.array(attn, dtype="int64"),
            "token_type_ids": np.array(ttype, dtype="int64"),
        }
        # IR 已固化 CLS pooling + L2 归一化，输出即句向量
        return compiled(feed)[0].tolist()

    # ---- 本地 ----
    def _local_model(self):
        if self._model is None:
            from sentence_transformers import SentenceTransformer

            self._model = SentenceTransformer(self.cfg.local_embed_model)
        return self._model

    def _embed_local(self, texts: List[str]) -> List[List[float]]:
        vecs = self._local_model().encode(texts, normalize_embeddings=True)
        return vecs.tolist()

    # ---- 云端 ----
    def _embed_cloud(self, texts: List[str]) -> List[List[float]]:
        import requests  # 仅云端路径需要，惰性导入避免拖垮零依赖默认路径

        url = self.cfg.cloud_embed_base_url.rstrip("/") + "/embeddings"
        headers = {"Authorization": f"Bearer {self.cfg.cloud_embed_api_key}"}
        out: List[List[float]] = []
        for t in texts:
            payload = {"model": self.cfg.cloud_embed_model, "input": t}
            r = requests.post(url, headers=headers, json=payload, timeout=30)
            r.raise_for_status()
            out.append(r.json()["data"][0]["embedding"])
        return out
