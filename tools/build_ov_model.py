# -*- coding: utf-8 -*-
"""构建本地 OpenVINO embedding 模型（可选增强，不影响零依赖默认路径）。

把 BAAI/bge-small-zh-v1.5 转成 OpenVINO IR，并把 CLS pooling + L2 归一化
固化进计算图 —— IR 输出即句向量，推理侧无需再实现 pooling，避免实现漂移。

用法：
    python tools/build_ov_model.py                    # 构建到默认缓存目录
    python tools/build_ov_model.py --out D:/ov/bge    # 指定输出目录
    python tools/build_ov_model.py --benchmark        # 额外跑设备性能对比

前置（均为可选依赖，未装时本脚本会给出明确提示）：
    pip install openvino torch transformers

国内网络建议：
    set HF_ENDPOINT=https://hf-mirror.com
"""
from __future__ import annotations

import argparse
import os
import shutil
import sys
import time

MODEL_ID = "BAAI/bge-small-zh-v1.5"
IR_NAME = "bge_small_zh"
# tokenizer 与 sentence-transformers 元数据，随 IR 一起落盘，保证自洽
TOKENIZER_FILES = [
    "tokenizer.json",
    "tokenizer_config.json",
    "vocab.txt",
    "special_tokens_map.json",
    "config.json",
]

_SELF_TEST_TEXTS = [
    "投标保证金的退还条件与时限约定",
    "项目经理须具备一级注册建造师执业资格且无在建工程",
    "工期为合同签订后 180 个日历日内完成全部施工内容",
]


def _default_out() -> str:
    try:
        from bid_toolkit.rag.config import default_ov_dir

        return default_ov_dir()
    except Exception:  # noqa: BLE001 - 脱离包独立运行时退回等价路径
        return os.path.join(
            os.path.expanduser("~"), ".cache", "bid_toolkit", "ov", "bge-small-zh-v1.5"
        )


def _require(mod: str, hint: str):
    try:
        return __import__(mod)
    except ImportError:
        print(f"[×] 缺少依赖 {mod}。安装：{hint}")
        sys.exit(2)


def build(out_dir: str, source: str, fp32: bool) -> str:
    _require("openvino", "pip install openvino")
    _require("torch", "pip install torch --index-url https://download.pytorch.org/whl/cpu")
    _require("transformers", "pip install transformers")

    import openvino as ov
    import torch
    from transformers import AutoModel, AutoTokenizer

    class BgeEmbedder(torch.nn.Module):
        """把 CLS pooling + L2 归一化固化进图，IR 直接输出句向量。"""

        def __init__(self, backbone):
            super().__init__()
            self.backbone = backbone

        def forward(self, input_ids, attention_mask, token_type_ids):
            out = self.backbone(
                input_ids=input_ids,
                attention_mask=attention_mask,
                token_type_ids=token_type_ids,
            )
            hidden = out[0] if isinstance(out, (tuple, list)) else out.last_hidden_state
            return torch.nn.functional.normalize(hidden[:, 0], p=2, dim=1)

    print(f"[1/5] 加载源模型：{source}")
    tok = AutoTokenizer.from_pretrained(source)
    backbone = AutoModel.from_pretrained(source)
    model = BgeEmbedder(backbone)
    model.eval()

    enc = tok(
        ["招投标语义检索示例文本"],
        padding="max_length",
        truncation=True,
        max_length=128,
        return_tensors="pt",
    )
    example = {
        "input_ids": enc["input_ids"],
        "attention_mask": enc["attention_mask"],
        "token_type_ids": enc.get(
            "token_type_ids", torch.zeros_like(enc["input_ids"])
        ),
    }

    print("[2/5] 转换为 OpenVINO IR")
    t0 = time.time()
    with torch.no_grad():
        ov_model = ov.convert_model(model, example_input=dict(example))
    # batch 与 seq 均设为动态，避免固定长度带来的 padding 浪费
    for inp in ov_model.inputs:
        pshape = inp.get_partial_shape()
        pshape[0] = -1
        pshape[1] = -1
        inp.get_node().set_partial_shape(pshape)
    ov_model.validate_nodes_and_infer_types()
    print(f"      转换耗时 {time.time() - t0:.1f}s")

    os.makedirs(out_dir, exist_ok=True)
    xml = os.path.join(out_dir, f"{IR_NAME}.xml")
    ov.save_model(ov_model, xml, compress_to_fp16=not fp32)
    size = os.path.getsize(os.path.join(out_dir, f"{IR_NAME}.bin")) / 2 ** 20
    print(f"[3/5] IR 已保存：{xml}  ({size:.1f} MB, {'FP32' if fp32 else 'FP16'})")

    print("[4/5] 落盘 tokenizer")
    if os.path.isdir(source):
        for name in TOKENIZER_FILES:
            src = os.path.join(source, name)
            if os.path.exists(src):
                shutil.copy2(src, os.path.join(out_dir, name))
    else:
        tok.save_pretrained(out_dir)
        backbone.config.save_pretrained(out_dir)

    # ---- 自检：IR 与 PyTorch 必须数值一致，否则检索质量无从保证 ----
    print("[5/5] 数值一致性自检")
    import numpy as np

    core = ov.Core()
    compiled = core.compile_model(xml, "CPU")

    def ov_vec(texts):
        e = tok(list(texts), padding=True, truncation=True,
                max_length=512, return_tensors="np")
        feed = {
            "input_ids": e["input_ids"],
            "attention_mask": e["attention_mask"],
            "token_type_ids": e.get("token_type_ids", np.zeros_like(e["input_ids"])),
        }
        return compiled(feed)[0]

    def pt_vec(texts):
        e = tok(list(texts), padding=True, truncation=True,
                max_length=512, return_tensors="pt")
        with torch.no_grad():
            h = backbone(**e).last_hidden_state[:, 0]
        return torch.nn.functional.normalize(h, p=2, dim=1).numpy()

    a, b = pt_vec(_SELF_TEST_TEXTS), ov_vec(_SELF_TEST_TEXTS)
    cos = (a * b).sum(1) / (np.linalg.norm(a, axis=1) * np.linalg.norm(b, axis=1))
    print(f"      余弦相似度 min={cos.min():.6f}  最大绝对误差={np.abs(a - b).max():.2e}")
    if cos.min() < 0.999:
        print("[×] 自检未通过：IR 与 PyTorch 输出偏差过大，请勿使用该 IR")
        sys.exit(1)
    print(f"      维度={a.shape[1]}  自检通过")
    return xml


def benchmark(xml: str, source: str) -> None:
    import numpy as np
    import openvino as ov
    import torch
    from transformers import AutoModel, AutoTokenizer

    tok = AutoTokenizer.from_pretrained(os.path.dirname(xml))
    ref = AutoModel.from_pretrained(source)
    ref.eval()

    texts = [
        "投标保证金退还条件与时限",
        "项目经理须具备一级注册建造师执业资格",
        "工期为合同签订后180个日历日",
        "质量标准应符合国家现行施工验收规范",
        "评标采用综合评分法技术分60分",
    ] * 8
    e_pt = tok(texts, padding=True, truncation=True, max_length=128, return_tensors="pt")
    e_np = tok(texts, padding=True, truncation=True, max_length=128, return_tensors="np")
    feed = {
        "input_ids": e_np["input_ids"],
        "attention_mask": e_np["attention_mask"],
        "token_type_ids": e_np.get("token_type_ids", np.zeros_like(e_np["input_ids"])),
    }

    def bench(fn, n=6):
        fn()
        return min(_timeit(fn) for _ in range(n))

    def _timeit(fn):
        t0 = time.perf_counter()
        fn()
        return time.perf_counter() - t0

    print("\n" + "=" * 58)
    print(f"性能对比（{len(texts)} 条文本 / 批）")
    print("=" * 58)
    with torch.no_grad():
        base = bench(lambda: ref(**e_pt))
    print(f"PyTorch  FP32 CPU : {base * 1000:7.1f} ms   1.00x")

    core = ov.Core()
    for dev in core.available_devices:
        try:
            cm = core.compile_model(xml, dev)
            t = bench(lambda: cm(feed))
            print(f"OpenVINO IR   {dev:<4}: {t * 1000:7.1f} ms   {base / t:.2f}x")
        except Exception as exc:  # noqa: BLE001 - 设备不可用只跳过
            print(f"OpenVINO IR   {dev:<4}: 不可用（{type(exc).__name__}）")
    print("=" * 58)


def main() -> int:
    ap = argparse.ArgumentParser(description="构建 bid-toolkit 的 OpenVINO embedding IR")
    ap.add_argument("--out", default=_default_out(), help="IR 输出目录")
    ap.add_argument("--source", default=MODEL_ID,
                    help="模型来源：HuggingFace 仓库名或本地目录")
    ap.add_argument("--fp32", action="store_true",
                    help="导出 FP32（体积约 90MB；纯 CPU 机器略快，默认 FP16 约 45MB）")
    ap.add_argument("--benchmark", action="store_true", help="构建后跑设备性能对比")
    args = ap.parse_args()

    print("=" * 58)
    print("bid-toolkit · OpenVINO embedding 构建")
    print("=" * 58)
    print(f"来源: {args.source}")
    print(f"输出: {args.out}")
    if not os.environ.get("HF_ENDPOINT") and not os.path.isdir(args.source):
        print("提示: 国内网络可先设置 HF_ENDPOINT=https://hf-mirror.com")
    print("-" * 58)

    xml = build(args.out, args.source, args.fp32)
    if args.benchmark:
        benchmark(xml, args.source)

    print("\n启用方式：")
    print("  set BID_RAG_EMBED_BACKEND=openvino")
    print("  bid rag ingest 历史标书.md --project demo")
    print("  bid rag query \"投标保证金\" --project demo")
    print("\n默认设备 AUTO:GPU,CPU（优先核显/独显，自动兜底 CPU）；")
    print("可用 BID_RAG_OV_DEVICE 指定，如 GPU / CPU / NPU。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
