# 示例：校园快递服务站运营服务方案

开源版端到端样例。演示 `bid render` 排版引擎如何把结构化的 `content.json` 渲染成
**带中文自动编号、真 Heading 样式**的 Word 文档。

## 运行

```bash
cd bid-toolkit
python -m bid_toolkit.render_docx \
  examples/sample_express_station/content.json \
  examples/sample_express_station/output.docx
```

用 Word 打开 `output.docx`，按 `Ctrl+A` 再按 `F9` 更新目录与页码域。

## 覆盖的块类型

| type | 含义 | 本例用法 |
|------|------|----------|
| `h1`–`h5` | 标题（真 Heading 1–5 + 自动编号） | 项目总体理解 / 服务定位 / 服务目标… |
| `p` | 正文段落（首行缩进 2 字符） | 需求理解陈述 |
| `table` | 表格（表 x-y 自动编号，表头加粗） | 服务目标、三级架构、人员配置 |
| `list` | 列表 | 培训与考核 |
| `figure` | 图片（图 x-y 自动编号） | 平面布局示意（`img` 留空→占位不崩） |

## 关键约定（写 content.json 必须守）

1. **标题 `text` 只写标题文字，不带编号前缀**。
   写 `{"type":"h1","text":"项目总体理解"}` → 引擎渲染成「**第一章 项目总体理解**」。
   若手打 `第一章` 前缀，会与自动编号叠加成「第一章 第一章 …」，`orchestrate check`
   的「⑩禁止手打编号」铁律会判不通过。
2. **编号皮肤可切换**：复制 `templates/format_config.example.yaml` 为 `format_config.yaml`，
   改「标题编号.格式」五项即可在 `第一章/一、/（一）/1、/1.1` 与 `1/1.1/1.1.1` 间切换，零代码。
3. **表格**：`{"type":"table","title":...,"header":[...],"rows":[[...]]}`。
4. **图片**：`{"type":"figure","img":"相对路径或绝对路径","title":...,"width_cm":12}`；
   `img` 为空或文件缺失时自动显示「【此处为占位图】」，不中断渲染。

## 已知缺口（开源版待补，详见仓库 ISSUE/待办）

- `orchestrate framework` 暂未导出 `content.json` 骨架，作者需按章节标题手工对齐内容；
  后续计划加 `orchestrate export-content`。
- `content.json` 若 JSON 非法会中途报错；写完后建议先 `python -m json.tool content.json` 校验。
