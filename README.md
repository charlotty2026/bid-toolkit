# 擎标 · bid-toolkit

招投标标书自动化工具链。不是「AI 写标书」，是**让 AI 写标书不翻车**。

- 纯本地运行，零 API 依赖，零网络调用
- 内容由宿主 Agent 撰写，工具链负责排版、质检、审标、脱敏与框架锁定
- 克隆即可跑，不需要配置任何 Key

## 为什么是本地

标书是企业的核心商业资产：报价策略、人员配置、技术方案、历史业绩，
任何一项外泄都可能直接影响中标结果。把这类文档上传到第三方服务风险太高。

本工具链的所有能力都在本地完成：

| 能力 | 实现方式 | 是否需要模型 |
|------|---------|-------------|
| 三层审标 | 102 条判词规则 | 否 |
| Word 排版 | python-docx + 编号引擎 | 否 |
| 格式自检 | 规则比对 | 否 |
| AI 味检测 | 句式模式库 | 否 |
| 脱敏扫描 | 敏感词库 | 否 |
| 铁律校验 | 分层规则 | 否 |
| 历史检索 | BM25（中文 bigram） | 否 |
| 语义检索 | OpenVINO IR / sentence-transformers / 云端 API | 可选 |

语义向量检索是**可选增强**，默认完全不依赖。AI PC 上推荐 `openvino` 路径：把
`bge-small-zh-v1.5` 转成 FP16 OpenVINO IR（约 45MB），`AUTO:GPU,CPU` 自动卸载到核显，
纯本地推理。实测（i7-13700H + Iris Xe，40 条/批）：PyTorch-CPU 75.1ms →
**OpenVINO-iGPU 13.1ms（5.75x）**；IR 与 PyTorch 输出余弦 1.000000，量化无损。

## 快速开始

```bash
pip install -r requirements.txt
python verify.py          # 8 项冒烟测试，约 1 分钟
```

核心依赖只有 5 个纯本地库：`python-docx` `markdown` `pyyaml` `PyMuPDF` `pdfplumber`。
语义检索可选：`openvino`+`tokenizers`（AI PC 加速）/ `sentence-transformers`（本地 PyTorch）/
云端 `/embeddings`，缺失时自动降级 BM25，默认零依赖即用。

## 典型流程

```bash
# 1. 扫风险，先知道雷在哪
python -m bid_toolkit review 招标文件.pdf -o 审标报告.md

# 2. 搭框架并锁死
python -m bid_toolkit orchestrate init
python -m bid_toolkit orchestrate parse 招标文件.pdf
python -m bid_toolkit orchestrate framework --template property
python -m bid_toolkit orchestrate lock

# 3. 撰写内容（Agent 或人工）→ content.json

# 4. 排版
python -m bid_toolkit render 内容.json 标书.docx

# 5. 门卫三连
python -m bid_toolkit orchestrate check 内容.json --docx 标书.docx
python -m bid_toolkit desense 标书.docx --mode bid
python -m bid_toolkit check 标书.docx --coverage
```

完整流程见 `references/workflow.md`。

## 目录结构

```
bid-toolkit/
├── SKILL.md                     Agent 入口（含 YAML front-matter）
├── README.md                    本文件
├── verify.py                    一键冒烟测试（8 项）
├── requirements.txt
├── bid_toolkit/                 Python 包
│   ├── render_docx.py           排版引擎（真 Heading + 自动编号）
│   ├── orchestrator/            框架锁定 / 差异检测 / 铁律校验 / 原文锚定
│   ├── review/                  三层审标管线
│   ├── rag/                     BM25 / OpenVINO 向量检索 + RRF 混合（自动降级）
│   ├── format/                  招标文件格式规则提取
│   └── scripts/                 分析、查重、脱敏等脚本
├── rules/
│   └── bid_rules.md             判词库（102 条）
├── templates/                   框架模板与排版配置
├── examples/                    示例（含成品 docx）
├── tests/                       测试样例
└── references/
    ├── command-reference.md     命令与参数
    ├── content-schema.md        content.json 结构规范
    └── workflow.md              作业流程与门卫清单
```

## 设计要点

**确定性优先。** 凡是能用规则算出来的，绝不用模型。判词库、铁律、格式自检全部是规则，
结果可复现、可解释、不会因为模型升级而漂移。

**编排即缰绳。** `orchestrate lock` 用 SHA256 锁定框架，任何改动都能被 `diff` 检出。
这在 Agent 参与撰写时尤其关键——防止多轮迭代后结构悄悄跑偏。

**降级不崩。** RAG 后端 `openvino`（AI PC 卸载）→ `local`（sentence-transformers）
→ `none`（BM25）逐级可选，本地向量库再与 BM25 做 RRF 融合；任选一环缺失都自动降级
并打印提示，不会因为缺依赖或缺 Key 直接失败。

## 环境要求

- Python 3.9+
- 无 GPU / NPU 要求（OpenVINO 路径可**选配**核显/NPU 加速，无则自动回退 CPU）
- 无网络要求（默认零依赖；语义检索的模型构建期需联网下载源模型一次）

## 许可

Apache-2.0
