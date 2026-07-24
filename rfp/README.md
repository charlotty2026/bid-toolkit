# RFP 招标文件工具集

从公开招标文件样本中归纳标准结构，生成规范招标文件 + 检查投标文件合规性。

## 目录结构

```
rfp/
├── rfp_structure.py        # 标准结构定义（六章框架）
├── rfp_generator.py        # 招标文件生成器
├── rfp_compliance.py       # 合规检查器
├── samples/                # 公开招标文件样本（5份，覆盖货物/服务/工程）
├── laws/                   # 相关法律法规（87号令PDF+TXT）
├── standard_structure.json # 标准结构框架（JSON）
└── compliance_rules/       # 合规规则库
```

## 标准六章结构（从5份样本归纳）

| 章节 | 标题 | 核心内容 |
|------|------|----------|
| 第一章 | 投标邀请/招标公告 | 项目基本信息、招标范围、资格要求、截止时间 |
| 第二章 | 投标人须知 | 投标规则、文件编制要求、废标条款 |
| 第三章 | 采购需求/招标需求 | 技术规格、服务要求、数量清单、质量标准 |
| 第四章 | 评标办法 | 评分方法、评审因素、分值分配 |
| 第五章 | 合同条款 | 付款方式、违约责任、质量保证 |
| 第六章 | 投标文件格式 | 投标函、授权委托书、报价表等模板 |

工程类额外增加：技术条件、图纸、工程量清单、最高投标限价。

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

- 样本来源：政府采购网、各省公共资源交易中心、发改委标准文件
- 法规依据：政府采购法、招标投标法、财政部87号令
