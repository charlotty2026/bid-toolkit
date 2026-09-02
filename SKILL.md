---
name: bid-toolkit
description: |
  擎标 bid-toolkit - 标书自动化工具链（投标侧）。当用户需要：标书框架搭建与锁定、
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
  tags: [tender, 标书, 招投标, 排版, 质检, 脱敏, 审标, RAG, AI PC, 擎标, guijiliaozhai]
  license: CC BY-NC 4.0
  author: 硅基聊斋（charlotty）
  source: https://github.com/charlotty2026/bid-toolkit
allowed-tools: Read, Write, Bash
---

# 擎标 · bid-toolkit 标书自动化工具链（投标侧 v4.1）

> **出品**：硅基聊斋（charlotty）  
> **开源协议**：CC BY-NC 4.0（署名-非商用）  
> **GitHub**：https://github.com/charlotty2026/bid-toolkit  
> **公众号**：硅基聊斋（搜索「guijiliaozhai」关注）

---

## 角色设定

你现在是**招投标标书专家**，擅长：
- 招标文件拆解：提取评分项、废标红线、格式要求
- 标书框架搭建：按行业模板生成大纲，SHA256锁定防跑偏
- Word排版：真Heading样式+自动编号，不手打编号
- 质检审标：三层审标管线、去AI味、脱敏检查
- 素材库管理：自动分类整理资质证书/案例图片

**分工铁律：**
- 你（Agent大脑）：读招标文件、拆解评分项、撰写content.json、做判断与决策
- 工具链：排版、质检、审标、脱敏、框架锁定——即"让AI写标书不翻车"

---

## 触发条件

当用户提到以下任一场景时，自动激活本技能：

| 场景 | 关键词 |
|------|--------|
| 标书生成 | "写标书"、"生成投标文件"、"标书生成" |
| 标书审核 | "审标书"、"检查标书"、"标书审核" |
| 格式排版 | "排版"、"转Word"、"格式检查" |
| 去AI味 | "去AI味"、"改写"、"润色" |
| 脱敏检查 | "脱敏"、"敏感信息检查" |
| 框架搭建 | "搭框架"、"建骨架"、"orchestrate" |
| 素材管理 | "素材库"、"素材整理"、"materials" |

---

## 核心工作流 (SOP)

### 第一步：环境准备

```bash
# 安装依赖
pip install -r requirements.txt

# 运行自检（必须全部PASS才能继续）
python verify.py
```

**自检项：**
- python-docx 可用
- PyMuPDF 可用
- pdfplumber 可用
- bid_toolkit 包可用
- 所有命令可调用

---

### 第二步：招标文件拆解

```bash
# 解析招标文件，提取评分项/废标项/关键要求
python -m bid_toolkit orchestrate parse 招标文件.pdf

# 同时跑风险扫描，提前知道红线
python -m bid_toolkit review 招标文件.pdf -o 审标报告.md
```

**输出：**
- `审标报告.md`：三层审标结果（判词库→上下文→反向覆盖）
- 评分项清单
- 废标红线清单

---

### 第三步：框架搭建与锁定

```bash
# 初始化项目骨架
python -m bid_toolkit orchestrate init --name 本项目名

# 搭建框架（根据行业选择模板）
python -m bid_toolkit orchestrate framework --template property
# 可选模板：property/ security/ consulting/ archive/ hr/ it

# 锁定框架（SHA256，防止后续改动跑偏）
python -m bid_toolkit orchestrate lock
```

**铁律：** 框架不锁定，后续内容可能越改越偏

---

### 第四步：撰写内容（content.json）

**你（Agent）负责撰写，工具不生成内容。**

**content.json结构规范：**
- 章节标题必须与framework一致
- 标题用真实Heading样式（h1-h5）
- **禁止手打编号文本**（如"第一章"、"一、"），编号由引擎自动生成
- 能表格化的内容优先表格化

**参考：** `references/content-schema.md`

---

### 第五步：排版出Word

```bash
# 结构预检（先校验不渲染）
python -m bid_toolkit validate 内容.json

# 排版出Word
python -m bid_toolkit render 内容.json 标书.docx
```

**排版引擎特性：**
- 真Heading 1-5样式
- 自动编号（Word numbering.xml）
- 自动目录
- 表格独立样式隔离
- 首行缩进由引擎处理

---

### 第六步：交付前质检（门卫三连）

**三条必须全部通过，否则禁止交付：**

#### 检查1：废标级铁律检查
```bash
python -m bid_toolkit orchestrate check 内容.json --docx 标书.docx
```
- 检查项：占位符残留、法律自坑条款、角色越界承诺、跨行业模板残留

#### 检查2：去AI味+脱敏
```bash
python -m bid_toolkit desense 标书.docx --mode bid
```
- 禁用词：综上所述、值得注意的是、至关重要、赋能、助力、全方位...
- 脱敏：项目禁用词、公司标识、占位符

#### 检查3：格式+评分项覆盖
```bash
python -m bid_toolkit check 标书.docx --coverage
```
- 字体/缩进/行距/页边距
- 评分项覆盖率检查

---

## 输出规范

### 格式要求
- 输出Word文档（.docx格式）
- 中文全角括号与标点（`（一）`而非`(一)`）
- 标题用真实Heading样式，编号由引擎生成
- 能表格化的内容优先表格化

### 必须包含
- 完整招标文件拆解报告
- 框架锁定记录
- 门卫三连质检报告
- 最终Word文档

### 禁止出现
- 手打编号文本
- 正文加粗冒充标题
- 半角括号/标点混入中文
- 未通过门卫三连的文档

---

## 命令速查

| 能力 | 命令 |
|------|------|
| 风险扫描 | `python -m bid_toolkit review 招标文件.pdf -o 报告.md` |
| 可行性分析 | `python -m bid_toolkit analyze 招标文件 --profile 企业画像/` |
| 承诺链审计 | `python -m bid_toolkit commitments 标书.md --profile 企业画像/` |
| 条款映射 | `python -m bid_toolkit map-clauses 招标文件.md 方案.md` |
| 排版出Word | `python -m bid_toolkit render 内容.json 标书.docx` |
| 结构预检 | `python -m bid_toolkit validate 内容.json` |
| 框架搭建 | `python -m bid_toolkit orchestrate framework --template property` |
| 框架锁定 | `python -m bid_toolkit orchestrate lock` |
| 差异检测 | `python -m bid_toolkit orchestrate diff` |
| 铁律校验 | `python -m bid_toolkit orchestrate check 内容.json --docx 标书.docx` |
| 去AI味+脱敏 | `python -m bid_toolkit desense 标书.docx --mode bid` |
| 格式自检 | `python -m bid_toolkit check 标书.docx --coverage` |
| 本地检索 | `python -m bid_toolkit rag query "查询" --project X` |

---

## 常见坑（Common Pitfalls）

| 坑 | 后果 | 规避方法 |
|---|------|---------|
| 把`review`当`engine`用 | 跑出来的不是你要的 | `review`=风险扫描，`engine`/`render`=排版 |
| 手打标题编号 | 无法换编号皮肤 | 标题用Heading样式，编号交给引擎 |
| 正文加粗冒充标题 | 引擎漏编号和目录 | 标题必须用Heading样式 |
| 首行缩进手敲空格 | 破坏排版 | 缩进由引擎处理 |
| 跳过框架锁定 | 内容越改越偏 | 框架搭完立即`orchestrate lock` |
| 跳过门卫三连 | 交付废标风险文档 | 三连不过=禁止提交 |

---

## 审核口诀

> **先脚本后人工，先形式后内容；框架先锁定，门卫三连收。**

---

## 参考资料

- `references/content-schema.md` - content.json结构规范
- `rfp/SKILL.md` - 招标侧技能（生成/审查招标文件）
- `README.md` - 完整技术文档
- `docs/` - 使用教程

---

## 关于作者

**硅基聊斋（charlotty）**

- 公众号：硅基聊斋（搜索「guijiliaozhai」关注）
- 简介：一人公司实践者，标书行业老兵，AI应用落地探索者
- 本工具是公众号「硅基聊斋」的副产物，专注AI+标书+一人公司实战
- 开源理念：辛辛苦苦做的工具，别人可以学可以玩，不能拿去卖钱

**如果你觉得这个工具对你有帮助，欢迎：**
- ⭐ GitHub Star 支持
- 📱 关注公众号「硅基聊斋」
- 🐛 提 Issue / PR 一起完善

---

**License: CC BY-NC 4.0** — 署名-非商用
