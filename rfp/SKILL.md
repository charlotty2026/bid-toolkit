---
name: bid-toolkit-rfp
description: |
  招投标标书自动化工具链 · 招标侧（开源 / v4.1）。当用户需要：生成规范招标文件
  （货物/服务/工程三类，Markdown + Word 双格式）、审查招标文件的合规性与法律风险
  （必备章节、排他性条款、废标条款、时间节点、资格条件、评分权重、法律依据、
  价格区间、多包规则），或从零起草招标文档时调用。

  站在招标人 / 代理机构一侧工作；纯本地运行，零 API 依赖、零网络调用。

  同一仓库的投标侧技能（做投标文件）见根目录 SKILL.md。两侧共用规则库与排版引擎。
metadata:
  version: 4.1.0
  display_name: 擎标 · 招标文件工具集（招标侧）
  tags: [tender, 招标文件, RFP, 合规审查, 排他性条款, 废标条款, 政采, 招投标, 擎标]
  license: CC BY-NC 4.0
---

# 擎标 · 招标文件工具集（招标侧 v4.1）

**这一侧是给招标人 / 代理机构用的。**

如果你拿到的是招标文件、要**做投标文件**去投它 → 用投标侧技能（根目录 `SKILL.md`）。
如果你要**写一份招标文件**发出去，或**审查一份招标文件**合不合规 → 用本文件。

**分工（读这一段再动手）**

| 谁 | 负责什么 |
|---|---|
| **你（Agent 大脑）** | 理解采购需求、填项目信息、判断条款是否妥当、解读合规报告 |
| **本工具** | 生成标准章节骨架（货物/服务 8 章、工程 11 章）、出 Word、跑 9 类合规检查并出报告 |

**工具不能替你决定采购需求。** 它保证"结构不缺、条款不踩红线、分值不违法"；
具体技术规格、资格条件怎么定，仍然是你的专业判断。

---

## 死规矩（违反即返工）

1. **生成的招标文件必须过一遍合规检查**，不能生成完直接发。
   `rfp_compliance.py` 是发出前的最后一道关。
2. **价格分权重按项目类型走法定区间**，不得拍脑袋。工具会查，查出来就改。
3. **排他性条款零容忍**：指定品牌、指定产地、限定特定行政区域业绩等，
   工具标出来就删，不要"酌情保留"。
4. **法规引用要落到具体条款**：写「依据政府采购法第 X 条」，不写「依据相关法规」。

---

## 环境准备

与投标侧共用同一套依赖，在仓库根目录执行：

```bash
pip install -r requirements.txt
cd rfp    # 后续命令在 rfp/ 目录下执行
```

---

## 一、生成招标文件

标准结构由 `rfp_structure.py` 的 `get_chapters()` 定义。**货物类与服务类是 8 章**：

| 章 | 标题 | 核心内容 |
|------|------|----------|
| 1 | 投标邀请 / 招标公告 | 项目基本情况、招标范围、资格要求、获取方式、递交与开标时间 |
| 2 | 投标人须知 | 前附表、澄清与修改、文件编制、保证金、废标条款、质疑与投诉 |
| 3 | 资格审查 | 资格要求、预审/后审方式、证明文件清单 |
| 4 | 采购需求 | 项目概况、技术规格、数量清单、质量标准、验收标准、售后 |
| 5 | 评标办法 | 评标方法、评审因素及分值分配、废标条件、加分项 |
| 6 | 合同条款 | 合同标的、付款方式、违约责任、质量保证、知识产权、争议解决 |
| 7 | 投标文件格式 | 投标函、授权委托书、报价表、承诺函等模板 |
| 8 | 政府采购政策落实 | 中小企业扶持、监狱企业、残疾人福利单位、节能与环保产品 |

**工程类共 11 章**：第 1–7 章同上（第 2 章不含"质疑与投诉"），
第 8–11 章替换为：技术条件（工程建设标准）→ 图纸及勘察资料 → 工程量清单 → 最高投标限价。

第 8 章"政府采购政策落实"仅对 `goods` / `services` 生效（`condition` 字段控制），
工程类不生成。

### 方式 A：命令行快速生成

```bash
# Markdown（默认）
python rfp_generator.py --type services --project "XX后勤服务项目" --budget 500000

# 同时出 Word
python rfp_generator.py --type services --project "XX项目" --budget 500000 --docx
```

`--type` 三选一：`services`（服务）/ `goods`（货物）/ `engineering`（工程）。

通过根 CLI 也可以（在仓库根目录执行）：

```bash
python -m bid_toolkit rfp --type services --project "XX项目" --budget 500000 --docx
```

### 方式 B：从 JSON 配置生成（推荐，可复用）

先复制对应类型的模板，改里面的项目信息：

```bash
# 模板在 rfp/templates/
cp templates/services_project_template.json my_project.json
# 编辑 my_project.json：项目名、预算、采购需求、评分办法、资格条件...

python rfp_generator.py --config my_project.json -o 招标文件.md
python rfp_generator.py --config my_project.json -o 招标文件.md --docx
```

三个模板：`services_project_template.json` / `goods_project_template.json` /
`engineering_project_template.json`。

### 自定义评分办法

默认分值来自内置配置，可用 `--scoring` 覆盖：

```bash
python rfp_generator.py --config my_project.json --scoring my_scoring.json --docx
```

评分配置结构见 `rfp_generator.py` 的 `load_scoring_config()`。
生成器会校验分值（`_validate_scoring`），分项加总不等于满分会直接报错，
**不会静默生成一份分算错的招标文件**。

---

## 二、合规检查（发出前必跑）

```bash
# 检查 Markdown，文本报告
python rfp_compliance.py --rfp 招标文件.md --type services --format text

# 检查 Word，输出 JSON 报告
python rfp_compliance.py --rfp 招标文件.docx --type goods -o 合规报告.json
```

### 9 类检查项

| 检查 | 函数 | 查什么 |
|------|------|--------|
| **结构完整性** | `check_completeness` | 章节骨架是否齐全、关键字段是否缺失 |
| **必备章节** | `check_required_sections` | 按项目类型要求的必备子节 |
| **排他性条款** | `check_exclusionary` | 指定品牌/产地、限定区域业绩等 31 条模式 |
| **评分办法** | `check_scoring` | 分值分配、权重是否越界 |
| **价格区间** | `check_price_range` | 价格分是否落在法定区间内 |
| **废标条款** | `check_rejection_clauses` | 废标情形是否表述规范、无滥用 |
| **时间节点** | `check_time_nodes` | 公告期、答疑期、投标截止等是否合法定最短期限 |
| **资格条件** | `check_qualification` | 资格要求是否构成不合理限制 |
| **法律依据** | `check_legal_basis` | 法规引用是否存在、是否具体 |
| **多包规则** | `check_multi_package_rule` | 分包/多包项目的规则是否自洽 |

一次跑全部：`run_all_checks(text, project_type)`；报告格式化：`format_text_report(report)`。

### 规则可配置（不改代码）

规则全部外置在 `compliance_rules/`，编辑 JSON 即可增删：

| 文件 | 内容 | 规则数 |
|------|------|--------|
| `exclusionary_patterns.json` | 排他性条款检测正则 | 31 条 |
| `rejection_keywords.json` | 废标条款关键词 | 16 个 |
| `required_sections.json` | 必备子节清单（按类型） | 3 类 |
| `time_keywords.json` | 时间节点关键词 | 5 个 |

各机构有自己的合规口径，改这里比改代码安全。

---

## 三、标准工作流（起草一份招标文件）

```bash
cd rfp

# 1) 选模板，填项目信息
cp templates/services_project_template.json my_project.json

# 2) 生成骨架
python rfp_generator.py --config my_project.json -o 招标文件.md --docx

# 3) 你（Agent）按实际采购需求补内容
#    技术规格、资格条件、评分细则这些必须由人确认，工具只是给骨架

# 4) 合规检查（发出前的最后一道关）
python rfp_compliance.py --rfp 招标文件.md --type services --format text

# 5) 按报告逐条修，修完重跑第 4 步，直到无 fatal
```

有 `fatal` 级问题**不许发出**。`warn` 级要逐条确认是否需要处理。

---

## 交付前验收清单（一条不过不许发出）

- [ ] 生成的招标文件**必跑合规检查**（`rfp_compliance.py`），不能生成完直接发
- [ ] 价格分权重合法定区间（工具查过，查出来就改）
- [ ] **排他性条款零容忍**：指定品牌/产地/限定区域业绩等，工具标出即删，不"酌情保留"
- [ ] 法规引用落到具体条款（「依据政府采购法第 X 条」，不写「依据相关法规」）
- [ ] 结构完整、必备章节齐全（货物/服务 8 章、工程 11 章）
- [ ] 评分办法经校验，分项加总 = 满分（不静默生成分算错的招标文件）

---

## 常见坑（Common Pitfalls）

1. **生成完就发，跳过合规检查**：`rfp_compliance.py` 是发出前最后一道关，fatal 必须清零。
2. **排他性条款"酌情保留"**：指定品牌/产地/限定行政区域业绩等，违规就删，没有商量余地。
3. **法规引用泛泛**：写「依据相关法规」不合格，必须落到具体条款号。
4. **价格分权重拍脑袋**：按项目类型走法定区间，工具会查，查到就改。
5. **在 rfp/ 外跑命令报错**：多数命令要在 `rfp/` 目录下执行（`cd rfp`），或从仓库根用 `python -m bid_toolkit rfp ...`。
6. **把范本当案例库**：`samples/` 是骨架参考，不是可复制的真实项目；内容须由专业人员按实际采购需求填。

---

## 四、标准结构与字段查询

不生成、只查结构定义时用 `rfp_structure.py`：

| 函数 | 返回 |
|------|------|
| `get_chapters(project_type)` | 该类型的章节清单 |
| `get_all_key_fields(project_type)` | 全部关键字段 |
| `get_compliance_checklist(project_type)` | 合规检查清单 |

结构化版本见 `standard_structure.json`。

---

## 五、法规依据

- `laws/财政部87号令.txt` / `.pdf` — 政府采购货物和服务招标投标管理办法
- `legal/regulations_reference.md` — 法规索引
- `references/招标文件法律法规清单.md` — 完整清单
- `docs/招标文件工具设计思路.md` — 设计说明

生成时会自动写入政府采购政策相关章节（`generate_gov_policy_content()`），
包括中小企业扶持、节能环保等法定要求。

---

## 六、参考范本

`samples/` 下保留 3 份**官方公开范本**，用于对照结构：

- `发改委标准设计招标文件.txt` — 国家发改委《标准设计招标文件》（2017 年版）
- `示例工程类项目_工程类.txt` — 示例市住建局施工招标范本（合成示例，EXAMPLE-ZB-2018-2）
- `goods/北京市政府采购公开招标示范文本_parsed.json` — 北京市政采示范文本

> 范本是"骨架参考"，不是"案例库"。本工具开源的是**生成与审查能力**，
> 不含任何真实项目招标文件。

---

## 参考资料

- `rfp_structure.py` — 标准章节结构定义（货物/服务 8 章、工程 11 章）
- `rfp_generator.py` — 生成器（Markdown + Word）
- `rfp_compliance.py` — 合规检查器（9 类检查）
- `compliance_rules/` — 可配置规则库
- `templates/` — 三类项目配置模板
- `standard_structure.json` — 标准结构 JSON
- `tests/test_rfp_suite.py` — 测试用例

## 使用时机

- 要起草一份招标文件（货物/服务/工程）→ `rfp_generator.py`
- 招标文件写完了要审合规 → `rfp_compliance.py`
- 想查某类项目的标准章节/必备字段 → `rfp_structure.py`
- 要改排他性/废标检测规则 → 编辑 `compliance_rules/*.json`
