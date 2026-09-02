#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""bid-toolkit 一键冒烟测试（零 API 依赖）。

用法:
    python verify.py

覆盖六条主链路，全程不访问网络、不需要任何 API Key:
    1. 核心依赖完整性
    2. review   三层审标（招标文件风险扫描）
    3. render   content.json -> Word 排版
    4. check    格式自检 + 评分项覆盖
    5. desense  AI 味检测 + 敏感信息脱敏
    6. rag      零依赖 BM25 入库与检索

退出码 0 = 全部通过；1 = 存在失败项。
"""
from __future__ import annotations

import importlib
import os
import shutil
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.abspath(__file__))
PY = sys.executable

CORE_DEPS = ["docx", "markdown", "yaml", "fitz", "pdfplumber"]

# 样例文件（相对仓库根）
SAMPLE_TENDER = os.path.join("tests", "test_tender.md")
SAMPLE_CONTENT = os.path.join("tests", "sample_content.json")
SAMPLE_BID = os.path.join("examples", "demo_bid.md")
SAMPLE_RAG = os.path.join("examples", "招标文件.md")

_results: list[tuple[str, bool, str]] = []


def _env() -> dict:
    env = os.environ.copy()
    env["PYTHONPATH"] = ROOT + os.pathsep + env.get("PYTHONPATH", "")
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def run_step(name: str, args: list[str], cwd: str, expect_file: str | None = None) -> bool:
    try:
        proc = subprocess.run(
            [PY, "-m", "bid_toolkit", *args],
            cwd=cwd,
            env=_env(),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=300,
        )
    except Exception as exc:  # noqa: BLE001
        _results.append((name, False, f"执行异常: {exc}"))
        return False

    ok = proc.returncode == 0
    if ok and expect_file:
        target = expect_file if os.path.isabs(expect_file) else os.path.join(cwd, expect_file)
        ok = os.path.exists(target)
        if not ok:
            _results.append((name, False, f"未产出预期文件: {target}"))
            return False

    tail = (proc.stdout or "").strip().splitlines()
    _results.append((name, ok, tail[-1][:100] if tail else ""))
    return ok


def check_deps() -> bool:
    missing = []
    for mod in CORE_DEPS:
        try:
            importlib.import_module(mod)
        except Exception:  # noqa: BLE001
            missing.append(mod)
    ok = not missing
    _results.append(
        ("依赖检查", ok, "全部就绪" if ok else f"缺失: {', '.join(missing)}")
    )
    return ok


def main() -> int:
    print("=" * 64)
    print("  bid-toolkit 冒烟测试（零 API 依赖）")
    print("=" * 64)

    check_deps()

    work = tempfile.mkdtemp(prefix="bidtoolkit_verify_")
    try:
        # 审标样例拷进工作区，保证相对路径可用
        src_tender = os.path.join(ROOT, SAMPLE_TENDER)
        work_tender = os.path.join(work, "招标文件.md")
        if os.path.exists(src_tender):
            shutil.copy(src_tender, work_tender)

        content = os.path.join(ROOT, SAMPLE_CONTENT)
        docx = os.path.join(work, "标书.docx")

        run_step(
            "review 三层审标",
            ["review", "招标文件.md", "-o", os.path.join(work, "out", "审标报告.md")],
            cwd=work,
            expect_file=os.path.join(work, "out", "审标报告.md"),
        )
        run_step(
            "render 排版引擎",
            ["render", content, docx],
            cwd=work,
            expect_file=docx,
        )
        run_step("check 格式自检", ["check", docx, "--coverage"], cwd=work)
        run_step(
            "desense AI味/脱敏",
            ["desense", os.path.join(ROOT, SAMPLE_BID), "--mode", "bid"],
            cwd=work,
        )
        run_step(
            "orchestrate 铁律校验",
            ["orchestrate", "check", content, "--docx", docx],
            cwd=work,
        )

        rag_src = os.path.join(ROOT, SAMPLE_RAG)
        if os.path.exists(rag_src):
            run_step(
                "rag 入库（BM25）",
                ["rag", "ingest", rag_src, "--project", "verify"],
                cwd=work,
            )
            run_step(
                "rag 检索（BM25）",
                ["rag", "query", "服务方案", "--project", "verify", "--top-k", "2"],
                cwd=work,
            )
    finally:
        shutil.rmtree(work, ignore_errors=True)

    print()
    passed = sum(1 for _, ok, _ in _results if ok)
    total = len(_results)
    for name, ok, detail in _results:
        flag = "PASS" if ok else "FAIL"
        print(f"  [{flag}] {name}" + (f"  — {detail}" if detail else ""))
    print()
    print(f"  通过 {passed}/{total}")
    print("=" * 64)

    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(main())
