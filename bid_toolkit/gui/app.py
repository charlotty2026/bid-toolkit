#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bid-toolkit GUI 主界面
=====================
Gradio 4 Tab 界面，调用 bid-toolkit 现有脚本，不重复实现核心逻辑。

Tab 1: 招标文件解析 — parse_bid.py + rfp_generator.py
Tab 2: 格式检查 — fix_bid_format.py + coverage_check.py
Tab 3: 脱敏扫描 — desensitization_scan.py
Tab 4: 审标扫描 — review/scanner.py + review/report.py

设计原则：
    - GUI 是薄壳，只负责传参数 + 展示结果
    - 所有核心逻辑在 scripts/ 里，GUI 通过 import 调用
    - 耗时操作用 yield 生成器模式，实时展示进度
    - 新增技能只需在 TAB_REGISTRY 注册一个 Tab 配置
"""

import os
import sys
import io
import json
import tempfile
import importlib.util
from pathlib import Path

# ─── 路径定位 ──────────────────────────────────────────────


def _package_dir():
    """定位 bid_toolkit 包目录。"""
    return Path(__file__).resolve().parent.parent


def _scripts_dir():
    """定位 scripts 目录。"""
    return _package_dir() / "scripts"


def _rfp_dir():
    """定位 rfp 目录。"""
    return _package_dir() / "rfp"


def _load_script_module(name, script_dir=None):
    """从指定目录动态加载 Python 脚本为模块。

    Args:
        name: 脚本文件名（如 'parse_bid.py'）
        script_dir: 脚本所在目录，默认为 scripts/

    Returns:
        加载后的模块对象
    """
    if script_dir is None:
        script_dir = _scripts_dir()

    script_path = Path(script_dir) / name
    if not script_path.is_file():
        raise FileNotFoundError(f"找不到脚本: {script_path}")

    mod_name = name.replace(".py", "")
    spec = importlib.util.spec_from_file_location(mod_name, str(script_path))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[mod_name] = mod
    spec.loader.exec_module(mod)
    return mod


# ─── Tab 1: 招标文件解析 ───────────────────────────────────


def parse_tender_file(file_obj, output_format="json"):
    """解析招标文件，提取结构化信息。

    调用 parse_bid.py 的 parse_bid_document() 函数。

    Args:
        file_obj: Gradio File 组件返回的文件对象
        output_format: 输出格式（json/markdown）

    Yields:
        str: 进度信息
    """
    from bid_toolkit.gui.gradio_compat import normalize_file_input

    file_path = normalize_file_input(file_obj)
    if not file_path:
        yield "请上传招标文件"
        return

    yield f"正在解析: {Path(file_path).name}\n"

    try:
        mod = _load_script_module("parse_bid.py")
        result = mod.parse_bid_document(file_path)

        stats = result.get("_统计", {})
        yield f"解析完成!\n\n"
        yield f"**摘要:**\n"
        yield f"- 废标红线: {stats.get('废标项', 0)} 条\n"
        yield f"- 文件清单: {stats.get('文件清单', 0)} 项\n"
        yield f"- 评分项: {stats.get('评分项', 0)} 个\n"
        yield f"- 资质要求: {stats.get('资质要求', 0)} 条\n"
        yield f"- 表格数: {stats.get('表格数', 0)} 个\n\n"

        if output_format == "json":
            output = json.dumps(result, ensure_ascii=False, indent=2)

            # 保存到临时文件供下载
            tmp_path = os.path.join(tempfile.gettempdir(), "bid_parsed.json")
            with open(tmp_path, "w", encoding="utf-8") as f:
                f.write(output)

            yield f"完整结果已保存，可点击下方下载。\n\n"
            yield (
                f"```json\n{output[:3000]}\n```"
                if len(output) > 3000
                else f"```json\n{output}\n```"
            )
        else:
            # Markdown 格式输出
            yield _format_parse_result_md(result)

    except Exception as e:
        yield f"解析失败: {e}\n请检查文件格式是否正确。"


def _format_parse_result_md(result):
    """将解析结果格式化为 Markdown。"""
    lines = ["## 招标文件解析结果\n"]

    fmt = result.get("格式要求", {})
    if fmt:
        lines.append("### 格式要求")
        for k, v in fmt.items():
            if v:
                lines.append(f"- **{k}**: {v}")

    disq = result.get("废标红线", [])
    if disq:
        lines.append(f"\n### 废标红线 ({len(disq)} 条)")
        for item in disq[:10]:
            lines.append(f"- {item}")
        if len(disq) > 10:
            lines.append(f"- ... 共 {len(disq)} 条")

    checklist = result.get("文件清单", [])
    if checklist:
        lines.append(f"\n### 文件清单 ({len(checklist)} 项)")
        for item in checklist[:10]:
            lines.append(f"- {item}")

    budget = result.get("预算", {})
    if budget:
        lines.append(f"\n### 预算")
        for k, v in budget.items():
            if v:
                lines.append(f"- **{k}**: {v}")

    timeline = result.get("时间节点", {})
    if timeline:
        lines.append(f"\n### 时间节点")
        for k, v in timeline.items():
            if v:
                lines.append(f"- **{k}**: {v}")

    return "\n".join(lines)


# ─── Tab 2: 格式检查 ───────────────────────────────────────


def check_bid_format(file_obj, scan_only=True):
    """标书格式检查。

    调用 fix_bid_format.py，通过重定向 stdout 捕获扫描报告。

    Args:
        file_obj: Gradio File 组件返回的文件对象
        scan_only: True 时仅扫描不修复

    Yields:
        str: 进度信息和扫描报告
    """
    from bid_toolkit.gui.gradio_compat import normalize_file_input

    file_path = normalize_file_input(file_obj)
    if not file_path:
        yield "请上传投标文件 (.docx)"
        return

    yield f"正在检查: {Path(file_path).name}\n\n"

    try:
        # 捕获 stdout
        old_argv = sys.argv
        old_stdout = sys.stdout
        sys.stdout = io.StringIO()

        sys.argv = ["fix_bid_format.py", file_path]
        if scan_only:
            sys.argv.append("--scan-only")

        mod = _load_script_module("fix_bid_format.py")
        try:
            mod.main()
        except SystemExit:
            pass
        except TypeError as e:
            # fix_bid_format 的自动修复阶段可能有小问题，不影响扫描结果
            pass

        output = sys.stdout.getvalue()
        sys.stdout = old_stdout
        sys.argv = old_argv

        if output.strip():
            yield f"```\n{output}\n```"
        else:
            yield "扫描完成，未发现问题。"

    except Exception as e:
        sys.stdout = old_stdout
        sys.argv = old_argv
        yield f"检查失败: {e}\n请确认文件是有效的 .docx 格式。"


# ─── Tab 3: 脱敏扫描 ───────────────────────────────────────


def scan_desensitization(file_obj):
    """敏感信息脱敏扫描。

    调用 desensitization_scan.py 的扫描功能。
    注意：现有脚本是针对代码仓库扫描设计的，GUI 模式下针对用户上传的文件扫描。

    Args:
        file_obj: Gradio File 组件返回的文件对象

    Yields:
        str: 进度信息和扫描报告
    """
    from bid_toolkit.gui.gradio_compat import normalize_file_input

    file_path = normalize_file_input(file_obj)
    if not file_path:
        yield "请上传需要扫描的文件"
        return

    yield f"正在扫描: {Path(file_path).name}\n\n"

    try:
        # 读取文件内容
        suffix = Path(file_path).suffix.lower()

        if suffix in (".docx", ".doc"):
            yield _scan_docx_sensitive(file_path)
        elif suffix == ".pdf":
            yield _scan_pdf_sensitive(file_path)
        elif suffix in (".md", ".txt"):
            yield _scan_text_sensitive(file_path)
        else:
            yield f"不支持的文件格式: {suffix}\n支持: .docx / .pdf / .md / .txt"

    except Exception as e:
        yield f"扫描失败: {e}"


def _scan_text_sensitive(file_path):
    """扫描文本文件中的敏感信息。"""
    import re

    with open(file_path, "r", encoding="utf-8", errors="replace") as f:
        text = f.read()

    patterns = {
        "手机号": re.compile(r"1[3-9]\d{9}"),
        "身份证号": re.compile(
            r"\b[1-9]\d{5}(?:19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\d{3}[\dXx]\b"
        ),
        "银行卡号": re.compile(r"\b[1-9]\d{15,18}\b"),
        "邮箱": re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+"),
        "IP地址": re.compile(r"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"),
    }

    lines = text.split("\n")
    findings = []

    for i, line in enumerate(lines, 1):
        for name, pattern in patterns.items():
            matches = pattern.findall(line)
            for match in matches:
                # 银行卡号做 Luhn 校验（alder-1 方案）
                if name == "银行卡号" and not _luhn_check(match):
                    continue
                # 掩码处理
                masked = _mask_value(match, name)
                context = line.strip()[:60]
                findings.append(f"| {name} | 第{i}行 | `{masked}` | ...{context}... |")

    if not findings:
        return "未发现敏感信息。"

    result = f"发现 {len(findings)} 处敏感信息:\n\n"
    result += "| 类型 | 位置 | 内容 | 上下文 |\n"
    result += "|------|------|------|--------|\n"
    result += "\n".join(findings)
    return result


def _scan_docx_sensitive(file_path):
    """扫描 Word 文档中的敏感信息。"""
    try:
        from docx import Document
    except ImportError:
        return "需要 python-docx 库: pip install python-docx"

    doc = Document(file_path)
    text = "\n".join(p.text for p in doc.paragraphs)

    # 写入临时文本文件，复用文本扫描逻辑
    tmp_path = os.path.join(tempfile.gettempdir(), "_bid_scan_tmp.txt")
    with open(tmp_path, "w", encoding="utf-8") as f:
        f.write(text)

    return _scan_text_sensitive(tmp_path)


def _scan_pdf_sensitive(file_path):
    """扫描 PDF 文件中的敏感信息。"""
    try:
        import fitz  # PyMuPDF
    except ImportError:
        return "需要 PyMuPDF 库: pip install PyMuPDF"

    text = ""
    with fitz.open(file_path) as doc:
        for page in doc:
            text += page.get_text()

    if not text.strip():
        return "PDF 内容为空（可能是扫描件，无文字层）。"

    tmp_path = os.path.join(tempfile.gettempdir(), "_bid_scan_tmp.txt")
    with open(tmp_path, "w", encoding="utf-8") as f:
        f.write(text)

    return _scan_text_sensitive(tmp_path)


def _luhn_check(card_number):
    """Luhn 算法校验银行卡号真伪。

    从 alder-1 产出借鉴，过滤掉不是真正银行卡号的数字串。
    """
    digits = [int(d) for d in card_number if d.isdigit()]
    if len(digits) < 12:
        return False
    checksum = 0
    parity = len(digits) % 2
    for i, d in enumerate(digits):
        if i % 2 == parity:
            d *= 2
            if d > 9:
                d -= 9
        checksum += d
    return checksum % 10 == 0


def _mask_value(value, value_type):
    """对敏感值进行掩码处理。"""
    if value_type == "手机号" and len(value) >= 11:
        return value[:3] + "****" + value[-4:]
    elif value_type == "身份证号" and len(value) >= 18:
        return value[:6] + "********" + value[-4:]
    elif value_type == "银行卡号" and len(value) >= 16:
        return value[:4] + "********" + value[-4:]
    elif value_type == "邮箱":
        parts = value.split("@")
        if len(parts) == 2:
            return parts[0][:2] + "***@" + parts[1]
    elif value_type == "IP地址":
        parts = value.split(".")
        return parts[0] + "." + parts[1] + ".*.*"
    return value[:2] + "***" + value[-2:] if len(value) > 4 else "***"


# ─── Tab 4: 审标扫描 ───────────────────────────────────────


def scan_tender_risk(file_obj, enable_llm=False):
    """招标文件风险扫描（三层审标管线）。

    调用 review/scanner.py 的 scan_tender() 函数。

    Args:
        file_obj: Gradio File 组件返回的文件对象
        enable_llm: 是否启用 LLM 上下文判断

    Yields:
        str: 进度信息和审标报告
    """
    from bid_toolkit.gui.gradio_compat import normalize_file_input

    file_path = normalize_file_input(file_obj)
    if not file_path:
        yield "请上传招标文件"
        return

    yield f"正在扫描: {Path(file_path).name}\n\n"

    try:
        # 添加 review 模块路径
        pkg_dir = _package_dir()
        review_dir = str(pkg_dir / "review")
        if review_dir not in sys.path:
            sys.path.insert(0, str(pkg_dir))

        from bid_toolkit.review import scanner, report as rpt

        yield "Layer 1: 判词库逐行扫描...\n"
        result = scanner.scan_tender(file_path, with_llm=enable_llm)
        yield f"Layer 1 完成: {len(result.hits)} 处命中\n\n"

        # 生成报告
        old_stdout = sys.stdout
        sys.stdout = io.StringIO()
        rpt.format_report(result)
        output = sys.stdout.getvalue()
        sys.stdout = old_stdout

        if output.strip():
            yield f"```\n{output}\n```"
        else:
            yield _format_scan_result_md(result)

    except ImportError as e:
        yield f"审标模块加载失败: {e}\n请确认 bid-toolkit 完整安装。"
    except Exception as e:
        yield f"扫描失败: {e}"


def _format_scan_result_md(result):
    """将扫描结果格式化为 Markdown。"""
    lines = ["## 审标扫描结果\n"]

    lines.append(f"- 总字符数: {result.total_chars:,}")
    lines.append(f"- 总行数: {result.total_lines:,}")
    lines.append(f"- 命中数: {len(result.hits)}")
    lines.append(f"- 致命项: {len(result.fatals)}")
    lines.append(f"- 警告项: {len(result.warnings)}")
    lines.append(f"- 信息项: {len(result.info_items)}")

    if result.fatals:
        lines.append(f"\n### 致命项 ({len(result.fatals)} 条)")
        for h in result.fatals[:20]:
            lines.append(f"- **第{h.line_num}行** `{h.keyword}` [{h.category}]")
            if h.context:
                lines.append(f"  > {h.context[:80]}")

    if result.warnings:
        lines.append(f"\n### 警告项 ({len(result.warnings)} 条)")
        for h in result.warnings[:20]:
            lines.append(f"- **第{h.line_num}行** `{h.keyword}` [{h.category}]")

    return "\n".join(lines)


# ─── Tab 注册表（预留扩展接口） ────────────────────────────

TAB_REGISTRY = [
    {
        "id": "parse",
        "label": "招标文件解析",
        "icon": "📋",
        "desc": "上传招标文件（PDF/DOCX/MD），自动提取废标红线、格式要求、评分项等结构化信息",
        "handler": parse_tender_file,
    },
    {
        "id": "format",
        "label": "格式检查",
        "icon": "🔍",
        "desc": "上传投标文件（DOCX），自动检查格式问题：字体/字号/占位符/编号/标题层级等17项",
        "handler": check_bid_format,
    },
    {
        "id": "desense",
        "label": "脱敏扫描",
        "icon": "🔒",
        "desc": "上传文件（DOCX/PDF/TXT），扫描手机号/身份证/银行卡号/邮箱/IP地址等敏感信息",
        "handler": scan_desensitization,
    },
    {
        "id": "review",
        "label": "审标扫描",
        "icon": "⚠️",
        "desc": "上传招标文件，三层审标管线扫描风险项：判词库匹配 -> 规则判断 -> 反向覆盖",
        "handler": scan_tender_risk,
    },
]


# ─── 构建 Gradio 界面 ──────────────────────────────────────


def build_app():
    """构建 Gradio Blocks 界面。

    使用 Blocks API 构建 4 Tab 界面，每个 Tab 对应一个核心功能。
    支持后续通过 TAB_REGISTRY 注册新 Tab。

    Returns:
        gr.Blocks 实例
    """
    import gradio as gr
    from bid_toolkit.gui.gradio_compat import get_gradio_version

    ver = get_gradio_version()

    # 自定义 CSS
    custom_css = """
    .gradio-container { max-width: 1200px !important; }
    .tab-desc { color: #666; font-size: 0.9em; margin-bottom: 10px; }
    """

    # Gradio 6.x: theme/css 移到 launch() 参数，不在 Blocks 构造函数中传
    blocks_kwargs = {"title": "bid-toolkit"}
    launch_kwargs = {}
    if ver < 6:
        blocks_kwargs["theme"] = gr.themes.Soft() if ver >= 4 else gr.themes.Default()
        blocks_kwargs["css"] = custom_css
    else:
        launch_kwargs["theme"] = gr.themes.Soft()
        launch_kwargs["css"] = custom_css

    with gr.Blocks(**blocks_kwargs) as app:
        gr.Markdown("# bid-toolkit 桌面版")
        gr.Markdown("招投标全流程工具链 — 离线运行，数据不泄露")

        with gr.Tabs():
            # Tab 1: 招标文件解析
            with gr.Tab("📋 招标文件解析"):
                gr.Markdown(
                    "> 上传招标文件（PDF/DOCX/MD/TXT），自动提取"
                    "废标红线、格式要求、评分项、资质要求等结构化信息"
                )
                with gr.Row():
                    parse_input = gr.File(
                        label="招标文件",
                        file_types=[".pdf", ".docx", ".doc", ".md", ".txt"],
                    )
                    parse_format = gr.Radio(
                        ["json", "markdown"],
                        value="markdown",
                        label="输出格式",
                    )
                parse_btn = gr.Button("开始解析", variant="primary")
                parse_output = gr.Markdown(label="解析结果")
                parse_btn.click(
                    fn=parse_tender_file,
                    inputs=[parse_input, parse_format],
                    outputs=parse_output,
                )

            # Tab 2: 格式检查
            with gr.Tab("🔍 格式检查"):
                gr.Markdown(
                    "> 上传投标文件（DOCX），自动检查17项格式问题："
                    "字体/字号/占位符/编号/标题层级/全角半角/表格样式等"
                )
                with gr.Row():
                    check_input = gr.File(
                        label="投标文件",
                        file_types=[".docx", ".doc"],
                    )
                    check_scan_only = gr.Checkbox(
                        value=True,
                        label="仅扫描不修复",
                    )
                check_btn = gr.Button("开始检查", variant="primary")
                check_output = gr.Markdown(label="检查报告")
                check_btn.click(
                    fn=check_bid_format,
                    inputs=[check_input, check_scan_only],
                    outputs=check_output,
                )

            # Tab 3: 脱敏扫描
            with gr.Tab("🔒 脱敏扫描"):
                gr.Markdown(
                    "> 上传文件，扫描敏感信息："
                    "手机号/身份证号/银行卡号（Luhn校验）/邮箱/IP地址"
                )
                desense_input = gr.File(
                    label="待扫描文件",
                    file_types=[".docx", ".doc", ".pdf", ".md", ".txt"],
                )
                desense_btn = gr.Button("开始扫描", variant="primary")
                desense_output = gr.Markdown(label="扫描报告")
                desense_btn.click(
                    fn=scan_desensitization,
                    inputs=[desense_input],
                    outputs=desense_output,
                )

            # Tab 4: 审标扫描
            with gr.Tab("⚠️ 审标扫描"):
                gr.Markdown(
                    "> 上传招标文件，三层审标管线："
                    "判词库逐行扫描 -> 规则上下文判断 -> 致命/警告/信息分级"
                )
                with gr.Row():
                    review_input = gr.File(
                        label="招标文件",
                        file_types=[".pdf", ".docx", ".doc", ".md", ".txt"],
                    )
                    review_llm = gr.Checkbox(
                        value=False,
                        label="启用 LLM 判断（需配置 API）",
                    )
                review_btn = gr.Button("开始审标", variant="primary")
                review_output = gr.Markdown(label="审标报告")
                review_btn.click(
                    fn=scan_tender_risk,
                    inputs=[review_input, review_llm],
                    outputs=review_output,
                )

        # 底部信息
        gr.Markdown("---")
        gr.Markdown(f"bid-toolkit GUI | Gradio {ver}.x | CLI: `bid list` 查看所有命令")

    # Gradio 6.x 需要 theme/css 在 launch() 传，挂到 app 上供 desktop.py 读取
    if launch_kwargs:
        app._bid_launch_kwargs = launch_kwargs

    return app
