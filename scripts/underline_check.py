#!/usr/bin/env python3
"""
下划线格式检查工具
检查Word(.docx)文档中的下划线格式是否符合配置规则。

功能：
- 检测哪些应该有下划线但没有（ERROR）
- 检测下划线类型（单线/双线/波浪线等）是否正确（WARN）
- 检测非预期的下划线（WARN）

用法：
    python underline_check.py check 投标文件.docx
    python underline_check.py check 投标文件.docx --config user_config.yaml
    python underline_check.py check 投标文件.docx --json
"""

import sys
import re
import json
import argparse
from pathlib import Path
from dataclasses import dataclass, asdict
from typing import Optional, Any

try:
    import yaml
except ImportError:
    print("需要安装pyyaml: pip install pyyaml", file=sys.stderr)
    sys.exit(1)

try:
    from docx import Document
    from docx.oxml.ns import qn
except ImportError:
    print("需要安装python-docx: pip install python-docx", file=sys.stderr)
    sys.exit(1)


# ============================================================
#  数据结构
# ============================================================

@dataclass
class CheckIssue:
    """检查问题记录"""
    level: str           # 'ERROR' 或 'WARN'
    page: int            # 近似页码
    paragraph: int       # 段落序号（1-based）
    text_preview: str    # 问题文本预览
    message: str         # 问题描述
    expected: str = ''   # 期望值
    actual: str = ''     # 实际值


# ============================================================
#  配置加载
# ============================================================

def load_config(config_path: Optional[str] = None) -> dict:
    """加载user_config.yaml配置文件。

    查找优先级：
    1. 显式传入的config_path
    2. 当前目录下的 user_config.yaml
    3. 项目根目录 templates/user_config.yaml
    """
    if config_path:
        p = Path(config_path)
        if not p.exists():
            print(f"配置文件不存在: {config_path}", file=sys.stderr)
            return {}
        with open(p, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f) or {}

    # 自动查找
    candidates = [
        Path.cwd() / 'user_config.yaml',
        Path(__file__).parent.parent / 'templates' / 'user_config.yaml',
    ]
    for p in candidates:
        if p.exists():
            with open(p, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f) or {}

    print("未找到user_config.yaml，已尝试: " + ', '.join(str(c) for c in candidates),
          file=sys.stderr)
    return {}


# ============================================================
#  Word文档下划线检测
# ============================================================

def get_underline_type(run) -> Optional[str]:
    """获取run的下划线类型。

    返回值：
        None  — 无下划线
        'single'  — 单线
        'double'  — 双线
        'wave'    — 波浪线
        'dotted'  — 点线
        'dash'    — 虚线
        其他OOXML定义的类型
    """
    rPr = run._element.find(qn('w:rPr'))
    if rPr is None:
        return None
    u_elem = rPr.find(qn('w:u'))
    if u_elem is None:
        return None
    val = u_elem.get(qn('w:val'))
    if val is None or val == 'none':
        return None
    return val


def _count_page_breaks(paragraph) -> int:
    """检测段落中的分页符数量（用于近似页码计算）。

    检测两种分页标记：
    - w:lastRenderedPageBreak: Word上次渲染时的分页位置
    - w:br type="page": 手动插入的分页符
    """
    count = 0
    for run in paragraph.runs:
        # lastRenderedPageBreak
        lrpb_list = run._element.findall(qn('w:lastRenderedPageBreak'))
        count += len(lrpb_list)
        # 手动分页符
        brs = run._element.findall(qn('w:br'))
        for br in brs:
            if br.get(qn('w:type')) == 'page':
                count += 1
    return count


def _build_run_map(paragraph) -> list[tuple[int, int, Any]]:
    """构建段落中run的字符位置映射。

    返回: [(start_pos, end_pos, run), ...]
    用于根据文本中的字符位置定位到对应的run对象。
    """
    run_map: list[tuple[int, int, Any]] = []
    pos = 0
    for run in paragraph.runs:
        text = run.text or ''
        if text:
            start = pos
            end = pos + len(text)
            run_map.append((start, end, run))
            pos = end
    return run_map


def _find_runs_for_range(
    run_map: list[tuple[int, int, Any]],
    match_start: int,
    match_end: int,
) -> list[Any]:
    """找到覆盖指定字符范围[start, end)的所有runs。"""
    result: list[Any] = []
    for start, end, run in run_map:
        if start >= match_end:
            break
        if end <= match_start:
            continue
        result.append(run)
    return result


def _extract_label_from_pattern(pattern_str: str) -> Optional[str]:
    """从正则模式字符串中提取人类可读的标签文本。

    例如从 "(致[：:]\\\\s*)(_{3,})" 中提取 "致"
    用于在完整模式未匹配时（占位符已填写）做标签级搜索。
    """
    m = re.search(r'([\u4e00-\u9fff]{2,})', pattern_str)
    if m:
        return m.group(1)
    return None


def check_underlines(docx_path: str, config: dict) -> list[CheckIssue]:
    """检查Word文档中的下划线格式。

    参数：
        docx_path: .docx文件路径
        config: 包含underline_patterns的配置字典

    返回：
        CheckIssue列表
    """
    issues: list[CheckIssue] = []

    patterns_raw = config.get('underline_patterns', [])
    if not patterns_raw:
        print("配置中没有underline_patterns", file=sys.stderr)
        return issues

    expected_underline_type: str = config.get('expected_underline_type', 'single')

    # 编译正则
    compiled_patterns: list[tuple[re.Pattern, str, Optional[str]]] = []
    for p_str in patterns_raw:
        try:
            compiled = re.compile(p_str)
            label = _extract_label_from_pattern(p_str)
            compiled_patterns.append((compiled, p_str, label))
        except re.error as e:
            print(f"正则编译失败，跳过: {p_str} ({e})", file=sys.stderr)

    pattern_count = len(compiled_patterns)
    if pattern_count == 0:
        return issues

    # 读取文档
    try:
        doc = Document(docx_path)
    except Exception as e:
        print(f"无法读取文档: {e}", file=sys.stderr)
        sys.exit(1)

    current_page = 1
    para_num = 0

    for para in doc.paragraphs:
        para_num += 1

        # 更新页码（分页符标记出现在段落开头表示新页开始）
        page_breaks = _count_page_breaks(para)
        if page_breaks > 0:
            current_page += page_breaks

        text = para.text
        if not text or not text.strip():
            continue

        run_map = _build_run_map(para)

        # 记录本段落中已被模式匹配覆盖的run（用于后续非预期下划线检测）
        matched_run_ids: set[int] = set()

        # ── 检查1: 模式匹配 → 下划线是否存在且类型正确 ──
        for compiled, p_str, label in compiled_patterns:
            matches = list(compiled.finditer(text))

            if matches:
                # 完整模式匹配（占位符未填写，下划线字符仍在文本中）
                for match in matches:
                    match_start = match.start()
                    match_end = match.end()

                    # 优先检查下划线部分（group 2 如果存在）
                    if match.lastindex and match.lastindex >= 2:
                        check_start = match.start(2)
                        check_end = match.end(2)
                    else:
                        check_start = match_start
                        check_end = match_end

                    matched_runs = _find_runs_for_range(run_map, check_start, check_end)
                    for r in matched_runs:
                        matched_run_ids.add(id(r))

                    preview = text[match_start:match_end][:30].replace('\n', ' ').strip()

                    if not matched_runs:
                        issues.append(CheckIssue(
                            level='WARN',
                            page=current_page,
                            paragraph=para_num,
                            text_preview=preview,
                            message=f"'{preview}' 无法定位到具体run",
                            expected=expected_underline_type,
                            actual='unknown',
                        ))
                        continue

                    # 检查下划线状态
                    has_underline = False
                    underline_types: set[str] = set()
                    for run in matched_runs:
                        ul_type = get_underline_type(run)
                        if ul_type is not None:
                            has_underline = True
                            underline_types.add(ul_type)

                    if not has_underline:
                        issues.append(CheckIssue(
                            level='ERROR',
                            page=current_page,
                            paragraph=para_num,
                            text_preview=preview,
                            message=f"'{preview}' 应有下划线但未找到",
                            expected=expected_underline_type,
                            actual='none',
                        ))
                    else:
                        for ul_type in underline_types:
                            if ul_type != expected_underline_type:
                                issues.append(CheckIssue(
                                    level='WARN',
                                    page=current_page,
                                    paragraph=para_num,
                                    text_preview=preview,
                                    message=(f"'{preview}' 下划线类型不符"
                                             f"（期望: {expected_underline_type},"
                                             f" 实际: {ul_type}）"),
                                    expected=expected_underline_type,
                                    actual=ul_type,
                                ))
            else:
                # 完整模式未匹配 → 尝试标签级匹配（占位符可能已被填写）
                if not label:
                    continue
                label_pos = text.find(label)
                if label_pos < 0:
                    continue

                # 检查标签之后的文本是否有下划线
                after_label_start = label_pos + len(label)
                # 取标签后最多50个字符作为检查范围
                check_end = min(after_label_start + 50, len(text))
                matched_runs = _find_runs_for_range(run_map, after_label_start, check_end)

                # 过滤掉纯空格的runs
                content_runs = [r for r in matched_runs if (r.text or '').strip()]
                if not content_runs:
                    continue

                for r in content_runs:
                    matched_run_ids.add(id(r))

                # 检查这些runs是否有下划线
                has_underline = False
                underline_types: set[str] = set()
                for run in content_runs:
                    ul_type = get_underline_type(run)
                    if ul_type is not None:
                        has_underline = True
                        underline_types.add(ul_type)

                preview = label + '...' + text[after_label_start:after_label_start + 20].strip()
                preview = preview[:30]

                if not has_underline:
                    # 标签存在但后续无下划线 — 可能需要下划线
                    issues.append(CheckIssue(
                        level='ERROR',
                        page=current_page,
                        paragraph=para_num,
                        text_preview=preview,
                        message=f"'{preview}' 应有下划线但未找到",
                        expected=expected_underline_type,
                        actual='none',
                    ))
                else:
                    for ul_type in underline_types:
                        if ul_type != expected_underline_type:
                            issues.append(CheckIssue(
                                level='WARN',
                                page=current_page,
                                paragraph=para_num,
                                text_preview=preview,
                                message=(f"'{preview}' 下划线类型不符"
                                         f"（期望: {expected_underline_type},"
                                         f" 实际: {ul_type}）"),
                                expected=expected_underline_type,
                                actual=ul_type,
                            ))

        # ── 检查2: 非预期下划线（有下划线但不在任何模式范围内） ──
        for run in para.runs:
            run_text = (run.text or '').strip()
            if not run_text or len(run_text) < 2:
                continue
            ul_type = get_underline_type(run)
            if ul_type is None:
                continue
            if id(run) in matched_run_ids:
                continue  # 已被模式覆盖

            # 检查是否是超链接的一部分（超链接的下划线通常不需要报告）
            parent = run._element.getparent()
            if parent is not None and parent.tag == qn('w:hyperlink'):
                continue

            preview = run_text[:30]
            issues.append(CheckIssue(
                level='WARN',
                page=current_page,
                paragraph=para_num,
                text_preview=preview,
                message=f"'{preview}' 存在非预期下划线（未匹配任何配置模式）",
                expected='none',
                actual=ul_type,
            ))

    return issues


# ============================================================
#  报告生成
# ============================================================

def format_text_report(
    docx_path: str,
    issues: list[CheckIssue],
    pattern_count: int,
) -> str:
    """生成文本格式检查报告。"""
    lines: list[str] = []
    lines.append("下划线检查报告")
    lines.append("===")
    lines.append(f"检查文件: {docx_path}")
    lines.append(f"检查规则: {pattern_count}条")
    lines.append("")

    if issues:
        lines.append("问题列表:")
        for issue in issues:
            lines.append(
                f"[{issue.level}] 第{issue.page}页 第{issue.paragraph}段: "
                f"{issue.message}"
            )
    else:
        lines.append("问题列表:")
        lines.append("（无问题）")

    error_count = sum(1 for i in issues if i.level == 'ERROR')
    warn_count = sum(1 for i in issues if i.level == 'WARN')
    lines.append("")
    lines.append(f"汇总: {len(issues)}个问题 ({error_count}个ERROR, {warn_count}个WARN)")

    return '\n'.join(lines)


def format_json_report(
    docx_path: str,
    issues: list[CheckIssue],
    pattern_count: int,
) -> str:
    """生成JSON格式检查报告。"""
    data = {
        "file": docx_path,
        "pattern_count": pattern_count,
        "issues": [asdict(i) for i in issues],
        "summary": {
            "total": len(issues),
            "errors": sum(1 for i in issues if i.level == 'ERROR'),
            "warnings": sum(1 for i in issues if i.level == 'WARN'),
        },
    }
    return json.dumps(data, ensure_ascii=False, indent=2)


# ============================================================
#  CLI入口
# ============================================================

def main() -> None:
    parser = argparse.ArgumentParser(
        description='下划线格式检查工具 — 检查Word文档中的下划线格式是否符合配置规则'
    )
    subparsers = parser.add_subparsers(dest='command', help='子命令')

    check_parser = subparsers.add_parser('check', help='检查文档下划线')
    check_parser.add_argument('docx', help='Word文档路径(.docx)')
    check_parser.add_argument(
        '--config', default=None,
        help='配置文件路径(user_config.yaml)，默认自动查找',
    )
    check_parser.add_argument(
        '--json', action='store_true', dest='json_output',
        help='输出JSON格式报告',
    )

    args = parser.parse_args()

    if args.command != 'check':
        parser.print_help()
        sys.exit(1)

    docx_path = args.docx
    if not Path(docx_path).exists():
        print(f"文件不存在: {docx_path}", file=sys.stderr)
        sys.exit(1)

    if not docx_path.lower().endswith('.docx'):
        print(f"警告: 文件不是.docx格式: {docx_path}", file=sys.stderr)

    # 加载配置
    config = load_config(args.config)
    patterns_raw = config.get('underline_patterns', [])

    if not patterns_raw:
        print("配置中没有underline_patterns，请检查配置文件", file=sys.stderr)
        sys.exit(1)

    # 执行检查
    issues = check_underlines(docx_path, config)
    pattern_count = len(patterns_raw)

    # 输出报告
    if args.json_output:
        print(format_json_report(docx_path, issues, pattern_count))
    else:
        print(format_text_report(docx_path, issues, pattern_count))

    # 有ERROR时返回非零退出码（便于CI/CD集成）
    if any(i.level == 'ERROR' for i in issues):
        sys.exit(1)


if __name__ == '__main__':
    main()
