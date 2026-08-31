# 命令参考

统一入口：`python -m bid_toolkit <命令>`（`pip install -e .` 后可简写为 `bid <命令>`）。
所有命令纯本地运行，不联网、不需要 API Key。

> 陷阱：风险扫描是 `review`，**不是** `engine`。`engine` 是 Markdown→Word 排版。

---

## review — 招标文件风险扫描（三层审标）

```bash
python -m bid_toolkit review 招标文件.pdf -o 审标报告.md
```

| 参数 | 说明 |
|------|------|
| `input` | 招标文件路径（PDF / DOCX / MD / TXT） |
| `-o, --output` | 导出报告（`.md` 或 `.json`），目录不存在会自动创建 |
| `-b, --bid-file` | 投标书路径，开启 Layer 3 反向覆盖检查 |
| `--llm` | 启用 Layer 2 上下文判断（需自配 LLM，默认关闭） |

三层管线：
1. **Layer 1** 判词库逐行扫描（102 条规则，零依赖，默认必跑）
2. **Layer 2** LLM 上下文消歧（`--llm` 可选）
3. **Layer 3** 反向覆盖检查：招标要求 ↔ 投标书逐项比对（需 `-b`）

输出按 致命 / 警告 / 证明材料 / 时间节点 / 商务门槛 / 合同约束 分级。

---

## render — content.json → Word 排版（核心）

```bash
python -m bid_toolkit render 内容.json 标书.docx [--config 格式.yaml]
```

真 Heading 样式 + Word `numbering.xml` 自动编号。不传 `--config` 时自动发现同目录 `format_config.yaml`。
结构规范见 `content-schema.md`。

## engine — Markdown → Word（兼容路径）

```bash
python -m bid_toolkit engine 标书.md -o 标书.docx
```

---

## orchestrate — 编排引擎（防 Agent 跑偏）

| 子命令 | 说明 |
|--------|------|
| `init [--name]` | 初始化项目，生成 `.bidproject/` |
| `parse <招标文件>` | 解析招标文件为结构化框架基线 |
| `framework [--template security\|property\|generic] [--from-tender]` | 构建框架 |
| `lock` / `unlock` | SHA256 锁定 / 解锁框架 |
| `diff <内容文件>` | 框架增删改差异比对 |
| `check <内容文件> [--docx 文件]` | 分层铁律校验（**废标级不过 = 禁提交**） |
| `anchor <内容文件>` | 原文锚定比对（投标承诺 vs 招标原文） |
| `profile [--set-file 卡.json \| --set generic \| --show]` | 投标人身份卡管理 |
| `templates` | 列出可用框架模板 |
| `status` | 查看项目阶段 / 锁定状态 / 锚定率 |
| `prelock` | 前置锁定检查（三重锁定关卡） |

---

## check — 格式自检 + 评分项覆盖

```bash
python -m bid_toolkit check 标书.docx --coverage
```

检查字体 / 缩进 / 行距，并对照评分项算覆盖率。低于 80% 会告警并列出缺失结构。

## desense — AI 味检测 + 敏感信息脱敏

```bash
python -m bid_toolkit desense 标书.docx --mode bid
```

清 AI 痕迹句式（空洞承诺、宣传性语言、破折号滥用等），并扫描敏感信息。
暗标项目用它去除公司标识。

## analyze — 投标可行性分析

```bash
python -m bid_toolkit analyze 招标文件 --profile 企业画像/ -o 报告.md
```

一键拆解 + 评分项分析 + 承诺审计 + 时间规划。

## commitments — 承诺链三源追踪

```bash
python -m bid_toolkit commitments 标书.md --profile 企业画像
```

招标要求 → 企业承诺 → 证据材料，三源对齐审计。

## map-clauses — 条款映射审计

```bash
python -m bid_toolkit map-clauses 招标文件.md 方案.md
```

评分项自动对应方案章节，查漏补缺。

## materials — 素材库管家

```bash
python -m bid_toolkit materials init|analyze|apply|status|learn 素材库目录
```

8 类自动分类 + 必备材料清单。

## rag — 本地检索（默认零依赖 BM25）

```bash
python -m bid_toolkit rag ingest 历史标书.md --project demo [--score-json 评分项.json] [--limit N]
python -m bid_toolkit rag query "人员配置方案" --project demo --top-k 5 [--score-item 评分项名]
python -m bid_toolkit rag status --project demo
```

后端优先级与自动降级：

| 后端 | 触发条件 | 依赖 |
|------|---------|------|
| `bm25`（默认） | `BID_RAG_EMBED_BACKEND=none` 或未设置 | 无 |
| `chromadb` + 本地模型 | `BID_RAG_EMBED_BACKEND=local` | sentence-transformers + chromadb |
| `pgvector` + 云端 API | `BID_RAG_EMBED_BACKEND=cloud` | psycopg / sqlalchemy + API Key |

任一环缺失都会打印提示并自动降级，**不会崩溃**。

---

## 其他

| 命令 | 说明 |
|------|------|
| `table 源.doc` | 工程标两列表格生成 |
| `watermark 输入 输出 -t 文本` | 添加文字水印 |
| `rfp --type services --project 名称` | 招标文件生成 |
| `list` | 列出所有可用工具 |
| `gui` | 启动桌面图形界面（原型） |
