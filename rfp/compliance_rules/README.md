# 合规规则配置目录

本目录将 rfp_compliance.py 中硬编码的规则外置为可配置 JSON 文件，便于：
- 非开发人员通过编辑 JSON 增删规则
- 不同项目类型使用不同规则集
- 规则版本化管理

## 文件说明

| 文件 | 说明 | 规则数 |
|------|------|--------|
| exclusionary_patterns.json | 排他性条款检测正则 | 31条 |
| rejection_keywords.json | 废标条款关键词 | 16个 |
| required_sections.json | 必备子节清单（按项目类型） | 3类 |
| time_keywords.json | 时间节点关键词 | 5个 |

## 添加新规则

直接编辑对应 JSON 文件，添加新条目即可。
rfp_compliance.py 会在下次运行时自动加载（需代码支持，当前为规划状态）。
