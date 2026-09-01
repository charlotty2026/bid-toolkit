---
name: bid-toolkit
description: |
  招投标标书自动化工具链 · 投标侧（开源核心 / v4.1）。当用户需要：标书框架搭建与锁定、
  content.json→Word 排版、招标文件风险扫描（三层审标）、格式自检与评分项覆盖检查、
  AI 味检测与暗标脱敏、承诺链与条款映射审计、素材库整理、零依赖本地 RAG 检索时调用。

  全部能力纯本地运行：零 API 依赖、零网络调用、不上传任何标书内容，
  克隆即可用，无需配置任何 Key。

  重要分工：本工具链不接 LLM，内容由宿主 Agent 撰写；
  工具链负责排版、质检、审标、脱敏与框架锁定——即"让 AI 写标书不翻车"。

  本仓库含两个技能：本文件为「投标侧」（做投标文件）；
  招标侧（生成/合规审查招标文件）见 rfp/SKILL.md。
metadata:
  version: 4.1.0
  display_name: 擎标 · 标书自动化工具链（投标侧）
  tags: [tender, 标书, 招投标, 排版, 质检, 脱敏, 审标, RAG, AI PC, 擎标]
  license: CC BY-NC 4.0
---

# 擎标 · bid-toolkit 标书自动化工具链（投标侧 v4.1）

**分工（读这一段再动手）**

| 谁 | 负责什么 |
|---|---|
| **你（Agent 大脑）** | 读招标文件、拆解评分项、撰写各章节内容、做判断与决策 |
| **本工具链** | 排版成 Word、格式自检、废标风险扫描、去 AI 味、框架锁定、本地检索 |

**脚本不接 LLM，也不要指望脚本替你写内容。** 你写好 `content.json`，交给 `render` 排版，
再用 `review` / `check` / `desense` 给自己挑错。这套工具的价值是**约束**，不是生成。

**本仓库的两个技能**

| 技能 | 入口 | 干什么 |
|---|---|---|
| **投标侧**（本文件） | 根目录 `SKILL.md` | 拿到招标文件后，做**投标文件**：审标、搭框架、排版、质检、脱敏 |
| **招标侧** | `rfp/SKILL.md` | 站在**招标人/代理机构**一侧，生成招标文件、做合规性审查 |

两侧共用同一套规则库与排版引擎。投标为主、招标为辅时，装载本文件即可；
要做招标文件时再读 `rfp/SKILL.md`。

---

## 死规矩（违反即返工）

1. **标题一律用真实 Heading 1–5 样式**，编号由引擎自动生成（Word numbering.xml）。
   **绝不许手打编号文本**（如「第一章」「一、」），否则客户无法自由换编号皮肤。
2. **能表格化就表格化**：机制 / 时限 / 责任人 / 考核优先用 `table` 块。
3. **中文全角括号与标点**：`（一）` 而非 `(一)`；首行缩进由引擎处理，不手敲空格。

---

## 环境准备

```bash
pip install -r requirements.txt
```

核心依赖仅 5 个纯本地库：`python-docx` `markdown` `pyyaml` `PyMuPDF` `pdfplumber`。
`pypinyin` `jieba` 为可选，缺失时相关功能自动降级，**不影响主链路**。
语义检索可选：`openvino`（AI PC 加速）/ `sentence-transformers`（本地 PyTorch）/ 云端 `/embeddings`，
缺失时自动降级 BM25，**默认零依赖即用**。

首次使用先跑自检：

```bash
python verify.py
```

全部 `PASS` 再往下走。

---

## 命令速查

统一入口：`python -m bid_toolkit <命令>`（pip 安装后可直接用 `bid <命令>`）

### 审标与拆解

| 能力 | 命令 | 说明 |
|------|------|------|
| **风险扫描** | `review 招标文件.pdf -o 报告.md` | 三层审标：判词库 → 上下文 → 反向覆盖 |
| **可行性分析** | `analyze 招标文件 --profile 企业画像/` | 拆解+评分+承诺审计+时间规划 |
| **承诺链审计** | `commitments 标书.md --profile 企业画像/` | 招标要求→承诺→证据 三源追踪 |
| **条款映射** | `map-clauses 招标文件.md 方案.md` | 评分项 ↔ 方案章节自动对应 |

### 框架与排版

| 能力 | 命令 | 说明 |
|------|------|------|
| **排版出 Word** | `render 内容.json 标书.docx` | 真 Heading 样式 + 自动编号（本工具核心） |
| **结构预检** | `validate 内容.json` | 只校验不渲染，提前暴露结构错误 |
| **Markdown→Word** | `engine 标书.md 标书.docx` | 简易排版（复杂结构请用 `render`） |
| **框架初始化** | `orchestrate init --name 项目名` | 建项目骨架 |
| **解析招标文件** | `orchestrate parse 招标文件.pdf` | 抽取条款与评分项 |
| **框架搭建** | `orchestrate framework --template property` | property 物业 / security 安保 / 等 6 套 |
| **框架锁定** | `orchestrate lock` | SHA256 锁定，防后续改动跑偏 |
| **差异检测** | `orchestrate diff` | 检查实际内容是否偏离锁定框架 |
| **铁律校验** | `orchestrate check 内容.json --docx 标书.docx` | 废标级不过 = 禁止提交 |
| **原文锚定** | `orchestrate anchor 标书.md 招标文件.md` | 投标承诺 vs 招标原文逐条比对 |
| **身份卡** | `orchestrate profile` | 投标人身份信息管理 |

### 质检与交付

| 能力 | 命令 | 说明 |
|------|------|------|
| **格式自检** | `check 标书.docx --coverage` | 字体/缩进/行距 + 评分项覆盖率 |
| **去 AI 味 + 脱敏** | `desense 标书.md --mode bid` | 清 AI 痕迹句式 + 敏感信息扫描 |
| **素材库** | `materials init/analyze/apply/status 目录` | 8 类自动分类 + 必备材料清单 |
| **水印** | `watermark 标书.docx "文字"` | 加文字水印 |
| **两列表格** | `table 标书.doc` | 工程标两列表格生成（DOC/PDF 提取） |

### 检索与辅助

| 能力 | 命令 | 说明 |
|------|------|------|
| **本地检索** | `rag ingest 历史标书.md --project X` / `rag query "查询" --project X` | 零依赖 BM25；设 `BID_RAG_EMBED_BACKEND=openvino` 升级语义+混合检索 |
| **索引状态** | `rag status --project X` | 查看已入库内容 |
| **招标侧** | `rfp` | 招标文件生成，完整能力见 `rfp/SKILL.md` |
| **图形界面** | `gui` | 启动桌面 GUI |
| **命令清单** | `list` | 列出所有可用工具 |

> 陷阱提醒：风险扫描是 `review`，**不是** `engine`。`engine` 是 Markdown→Word 排版。别混。

---

## 标准工作流（写一份标书）

```bash
# 1) 读招标文件，先扫风险——动手前就知道哪些是废标红线
python -m bid_toolkit review 招标文件.pdf -o 审标报告.md

# 2) 搭框架并锁死（防止后续越改越偏）
python -m bid_toolkit orchestrate init --name 本项目
python -m bid_toolkit orchestrate parse 招标文件.pdf
python -m bid_toolkit orchestrate framework --template property   # 物业/招租用 property，安保用 security
python -m bid_toolkit orchestrate lock

# 3) 你（Agent）按框架撰写内容，写成 content.json
#    章节标题必须与 framework 一致；标题用 h1–h5，编号禁手打
#    结构规范见 references/content-schema.md

# 4) 排版出 Word
python -m bid_toolkit validate 内容.json          # 先预检结构
python -m bid_toolkit render 内容.json 标书.docx

# 5) 交付前自检（门卫三连）
python -m bid_toolkit orchestrate check 内容.json --docx 标书.docx   # 废标级铁律
python -m bid_toolkit desense 标书.docx --mode bid                   # 去 AI 味 + 脱敏
python -m bid_toolkit check 标书.docx --coverage                     # 格式 + 评分项覆盖
```

门卫三连任意一项不过，**不许交付**。

---

## 排版可配置（不改代码）

复制模板后改 YAML 即可：页面 / 字体 / 各级标题字号与对齐 / 编号格式 / 表格 / 封面 / 页码。

```bash
cp templates/format_config.example.yaml ./format_config.yaml
python -m bid_toolkit render 内容.json 标书.docx    # 自动发现同目录配置
```

编号皮肤切换（改「标题编号.格式」五项，零代码）：
中文 `第一章 / 一、/ （一）/ 1、/ 1.1` ⇄ 十进制 `1 / 1.1 / 1.1.1`。

行业模板在 `templates/`：`goods`（货物）/ `service`（服务）/ `engineering`（工程）/
`government`（政府）/ `enterprise`（企业）/ `bid_type_detection`（类型自动判别）。

---

## 零依赖本地 RAG

默认走 **BM25**，纯 Python 实现（中文 bigram 分词），不下载模型、不需要 Key：

```bash
python -m bid_toolkit rag ingest 历史标书.md --project demo
python -m bid_toolkit rag query "人员配置方案" --project demo --top-k 5
```

需要语义向量时再挂 embedding，本工具链不绑架任何云端服务。三级后端逐级降级，
**任一环缺失都会自动回退并打印提示，永不崩**：

```
openvino（AI PC：iGPU/NPU 卸载，CPU 兜底）→ local（sentence-transformers）→ none（BM25）
```

### AI PC 增强：OpenVINO 本地向量（推荐）

把 `BAAI/bge-small-zh-v1.5`（512 维）转成 FP16 OpenVINO IR（45.2MB），
**推理全程纯本地**，并把 CLS pooling + L2 归一化固化进计算图（IR 输出即句向量，
推理侧无需再 pooling）。在 Intel AI PC 上把算子卸载到核显。

**本机实测（Intel 13 代 + Iris Xe 核显，40 条文本/批，由
`python tools/build_ov_model.py --benchmark` 实跑）**：

| 后端 | 耗时 | 相对 PyTorch |
|---|---|---|
| PyTorch FP32 CPU | 63.6 ms | 1.00x |
| OpenVINO FP16 CPU | 61.9 ms | 1.03x |
| **OpenVINO FP16 iGPU** | **12.7 ms** | **5.01x** |

> 说明：FP16 在纯 CPU 上并无优势（甚至略慢），**加速来自核显卸载**。
> 本工具链的数值自检要求 IR 与 PyTorch 输出余弦 ≥ 0.999，本机实跑
> **余弦=1.000000、最大绝对误差=1.64e-07**（IR 45.2MB FP16、导出 1.3s），即量化不损失检索质量。

一键构建（仅首次，构建期需 `pip install openvino torch transformers`，约 1–2 分钟；
**运行期只需 `openvino` + `tokenizers`，不加载 torch**）：

```bash
python tools/build_ov_model.py              # 构建 FP16 IR 到 ~/.cache/bid_toolkit/ov/
python tools/build_ov_model.py --benchmark  # 额外跑 CPU / iGPU 性能对比
python tools/build_ov_model.py --fp32       # 纯 CPU 机器可选 FP32（约 90MB）
```

启用：

```bash
set BID_RAG_EMBED_BACKEND=openvino          # 等价：export ...
python -m bid_toolkit rag ingest 历史标书.md --project demo
python -m bid_toolkit rag query "投标保证金" --project demo --top-k 5
```

默认设备 `AUTO:GPU,CPU`（自动挑最快可用设备并兜底 CPU）。若想**确保吃到核显**，
可显式指定：`set BID_RAG_OV_DEVICE=GPU`（日志会打印 `设备=GPU.0`）。
NPU 机型设 `BID_RAG_OV_DEVICE=NPU`。

混合检索：本地向量库（localvec，仅 numpy）同时维护 BM25 倒排索引，检索时把
「语义向量 Top-N」与「关键词 Top-N」用 **RRF** 融合——招投标文本专有名词、条款号、
数字指标多，纯语义易漏精确匹配，纯关键词又抓不住同义表述，融合后互补更稳。

### 其他可选后端

```bash
BID_RAG_EMBED_BACKEND=local BID_RAG_LOCAL_EMBED_MODEL=BAAI/bge-small-zh \
  python -m bid_toolkit rag query "人员配置方案" --project demo
```

- `local`：sentence-transformers（PyTorch），需 `pip install sentence-transformers`，可离线
- `cloud`：OpenAI 兼容 `/embeddings`（bge-m3，1024 维），需 `BID_RAG_CLOUD_EMBED_API_KEY`
- `none`（默认）：零依赖 BM25，克隆即用

---

## 参考资料

- `references/command-reference.md` — 全部命令与参数
- `references/content-schema.md` — content.json 结构与创作规范
- `references/workflow.md` — 完整作业流程与门卫清单
- `templates/` — 6 套行业框架模板 + 排版配置模板
- `company_profile/` — 企业画像模板（`***` 占位，填自己的）
- `test_profile/` — 演示用虚构数据（张三/李四），可安全运行
- `examples/` — 示例文件
- `verify.py` — 一键冒烟测试

## 使用时机

- 收到招标文件，要拆解废标红线与评分项 → `review` + `analyze`
- 怕 Agent 自由发挥改坏已定稿的章节 → `orchestrate framework` + `lock` + `check`
- 内容写完了要出 Word → `render`
- 定稿前最后一道关 → `orchestrate check` + `desense` + `check --coverage`
- 暗标项目要去掉公司标识 → `desense`
- 要站在招标人一侧生成/审查招标文件 → 看 `rfp/SKILL.md`
