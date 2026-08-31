# 作业流程与门卫清单

## 全流程

```
阶段 0  环境自检     python verify.py
阶段 1  读标         review 扫风险 → 知道哪些是废标红线
阶段 2  搭架         orchestrate init / parse / framework / lock
阶段 3  撰写         Agent 按框架写 content.json
阶段 4  排版         render → 标书.docx
阶段 5  门卫         orchestrate check → desense → check --coverage
阶段 6  交付         导出 PDF / 加水印
```

## 阶段 1：先扫风险，再动手

```bash
python -m bid_toolkit review 招标文件.pdf -o 审标报告.md
```

目的不是生成内容，是**先知道雷在哪**。报告里的「致命风险」项（废标/否决、资格审查、
实质性偏离、逾期送达）必须在撰写阶段逐条规避。

## 阶段 2：搭框架并锁死

```bash
python -m bid_toolkit orchestrate init
python -m bid_toolkit orchestrate parse 招标文件.pdf
python -m bid_toolkit orchestrate framework --template property
python -m bid_toolkit orchestrate lock
```

`lock` 会给框架算 SHA256。后续任何改动都能用 `diff` 检出，
避免多轮迭代后章节结构悄悄跑偏——这在与 Agent 协作时尤其重要。

模板选择：`property` 物业/招租，`security` 安保，`generic` 通用。

## 阶段 3：撰写（Agent 主导）

按框架章节写 `content.json`。要点：

- 章节标题与 `framework` 严格一致
- 标题用 `h1`–`h5`，**编号禁手打**
- 机制 / 时限 / 责任人 / 考核优先表格化
- 避免空洞承诺（「全面提升」「显著优化」会被 `desense` 抓出来）
- 每个承诺尽量落到 具体措施 + 量化指标 + 时间节点

## 阶段 4：排版

```bash
cp templates/format_config.example.yaml ./format_config.yaml   # 按需改
python -m bid_toolkit render 内容.json 标书.docx
```

Word 打开后 `Ctrl+A` 然后 `F9` 更新目录与页码域。

## 阶段 5：门卫三连（交付前必跑）

```bash
python -m bid_toolkit orchestrate check 内容.json --docx 标书.docx
python -m bid_toolkit desense 标书.docx --mode bid
python -m bid_toolkit check 标书.docx --coverage
```

判定标准：

| 检查 | 通过条件 | 不通过怎么办 |
|------|---------|-------------|
| 铁律校验 | 废标级 0 项 | **禁止提交**，逐条改 |
| AI 味 / 脱敏 | 高危 0 项、敏感词 0 命中 | 按建议改写；暗标必查公司标识 |
| 格式 + 覆盖 | 覆盖率 ≥ 80% | 对照评分项补章节 |

## 常见返工原因

1. **手打编号**——客户拿到 Word 后无法自由换编号皮肤，铁律直接判 fatal。
2. **章节与框架不一致**——`orchestrate check` 会报结构偏移。
3. **空洞承诺**——`desense` 的高危项，`（全面|显著|有力|切实）+ 动词` 是典型触发模式。
4. **评分项漏覆盖**——`check --coverage` 会列出缺失结构，对着补。

## 与 Agent 协作的建议

- 让 Agent 先跑 `review`，把致命风险清单贴在上下文里再写，命中率明显更高。
- 长标书分章节写，每写完一章跑一次 `check`，别攒到最后一次性验。
- `lock` 之后不要改章节结构；要改先 `unlock`，改完重新 `lock`。
