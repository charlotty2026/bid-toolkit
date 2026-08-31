# RFP 招标文件工具集

从公开招标文件范本中归纳标准结构，生成规范招标文件 + 检查招标文件合规性。

> **Agent 使用者请看 `SKILL.md`**（本目录内），那是给 AI 读的完整操作手册。
> 本文件是人类阅读的项目说明。

## 目录结构

```
rfp/
├── SKILL.md                # 招标侧技能手册（Agent 入口）
├── rfp_structure.py        # 标准结构定义（货物/服务8章、工程11章）
├── rfp_generator.py        # 招标文件生成器
├── rfp_compliance.py       # 合规检查器（9类检查）
├── samples/                # 官方公开范本（3份：设计/工程/货物）
├── laws/                   # 相关法律法规（87号令PDF+TXT）
├── legal/                  # 法规索引
├── docs/                   # 设计思路
├── references/             # 法律法规清单
├── templates/              # 三类项目配置模板
├── tests/                  # 测试用例
├── standard_structure.json # 标准结构框架（JSON）
└── compliance_rules/       # 合规规则库（可配置JSON）
```

## 标准结构（由 `rfp_structure.py` 定义）

**货物类 / 服务类 = 8 章**

| 章 | 标题 | 核心内容 |
|------|------|----------|
| 1 | 投标邀请/招标公告 | 项目基本情况、招标范围、资格要求、获取方式、递交与开标时间 |
| 2 | 投标人须知 | 前附表、澄清与修改、文件编制、保证金、废标条款、质疑与投诉 |
| 3 | 资格审查 | 资格要求、预审/后审方式、证明文件清单 |
| 4 | 采购需求 | 项目概况、技术规格、数量清单、质量标准、验收标准、售后 |
| 5 | 评标办法 | 评标方法、评审因素及分值分配、废标条件、加分项 |
| 6 | 合同条款 | 合同标的、付款方式、违约责任、质量保证、知识产权、争议解决 |
| 7 | 投标文件格式 | 投标函、授权委托书、报价表、承诺函等模板 |
| 8 | 政府采购政策落实 | 中小企业扶持、监狱企业、残疾人福利单位、节能与环保产品 |

**工程类 = 11 章**：1–7 章同上（第 2 章不含"质疑与投诉"），
第 8 章起为技术条件 → 图纸及勘察资料 → 工程量清单 → 最高投标限价。

第 8 章政府采购政策仅对 `goods`/`services` 生效，工程类不生成。

## 用法

### 生成招标文件

```bash
# 交互式生成（服务类）
python rfp_generator.py --type services --project "XX后勤服务项目" --budget 500000

# 同时生成Word文档
python rfp_generator.py --type services --project "XX项目" --budget 500000 --docx

# 从JSON配置生成
python rfp_generator.py --config project.json -o 招标文件.md
```

### 合规检查招标文件

```bash
# 检查Markdown招标文件
python rfp_compliance.py --rfp 招标文件.md --type services --format text

# 检查Word招标文件，输出JSON报告
python rfp_compliance.py --rfp 招标文件.docx --type goods -o 合规报告.json
```

## 数据来源

- 结构来源：发改委标准文件、地方住建局范本、政府采购示范文本——**均为官方公开范本**
- 法规依据：政府采购法、招标投标法、财政部87号令

> 本工具开源的是**生成与审查能力**。仓库内不含任何真实项目的招标文件，
> 也不含真实标书案例库。
