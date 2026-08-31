#!/usr/bin/env python3
"""Post-Audit Scanner - Layer 3 of WeChat Publishing Pipeline.

v1.0 (2026-08-12): Initial release.
  - Parse generated HTML with regex fallback (BeautifulSoup optional)
  - Verify compliance against column template rules
  - Check card borders, colors, typography, required elements, images
  - Output structured audit report with fix suggestions
"""

import json
import re
import sys
import os
from pathlib import Path
from typing import Dict, List, Tuple, Any, Optional

# Try to import BeautifulSoup, fallback to regex if not available
try:
    from bs4 import BeautifulSoup

    HAS_BS4 = True
except ImportError:
    HAS_BS4 = False

SCRIPT_DIR = Path(__file__).parent.resolve()
CONFIG_PATH = SCRIPT_DIR / "column_templates.json"


class AuditResult:
    CRITICAL = "CRITICAL"
    WARNING = "WARNING"
    INFO = "INFO"
    PASS = "PASS"
    SKIP = "SKIP"

    def __init__(
        self,
        category: str,
        name: str,
        status: str,
        message: str = "",
        position: str = "",
        fix_suggestion: str = "",
        details: List[str] = None,
    ):
        self.category = category
        self.name = name
        self.status = status
        self.message = message
        self.position = position
        self.fix_suggestion = fix_suggestion
        self.details = details or []

    def to_dict(self) -> Dict[str, Any]:
        return {
            "category": self.category,
            "name": self.name,
            "status": self.status,
            "message": self.message,
            "position": self.position,
            "fix_suggestion": self.fix_suggestion,
            "details": self.details,
        }


class AuditReport:
    def __init__(self):
        self.results: List[AuditResult] = []
        self.blocked = False
        self.block_reasons: List[str] = []

    def add(self, result: AuditResult):
        self.results.append(result)
        if result.status == AuditResult.CRITICAL:
            self.blocked = True
            self.block_reasons.append(
                f"[{result.category}] {result.name}: {result.message}"
            )

    def get_summary(self) -> Dict[str, Any]:
        critical = sum(1 for r in self.results if r.status == AuditResult.CRITICAL)
        warnings = sum(1 for r in self.results if r.status == AuditResult.WARNING)
        infos = sum(1 for r in self.results if r.status == AuditResult.INFO)
        passed = sum(1 for r in self.results if r.status == AuditResult.PASS)
        return {
            "total": len(self.results),
            "critical": critical,
            "warning": warnings,
            "info": infos,
            "pass": passed,
            "blocked": self.blocked,
            "block_reasons": self.block_reasons,
        }

    def to_dict(self) -> Dict[str, Any]:
        return {
            "summary": self.get_summary(),
            "results": [r.to_dict() for r in self.results],
        }

    def print_report(self):
        summary = self.get_summary()
        print(f"\n{'=' * 60}")
        print(f"  后置审计报告")
        print(f"{'=' * 60}")
        print(f"  总计: {summary['total']} 项")
        print(
            f"  通过: {summary['pass']}  |  警告: {summary['warning']}  |  严重: {summary['critical']}"
        )
        print(f"  信息: {summary['info']}")
        if summary["blocked"]:
            print(f"\n  [BLOCKED] 以下严重问题需要修复:")
            for reason in summary["block_reasons"]:
                print(f"    - {reason}")
        print(f"{'=' * 60}\n")

        for r in self.results:
            icon = {
                AuditResult.PASS: "[PASS]",
                AuditResult.WARNING: "[WARN]",
                AuditResult.CRITICAL: "[FAIL]",
                AuditResult.INFO: "[INFO]",
            }.get(r.status, "[?]")
            print(f"  {icon} [{r.category}] {r.name}: {r.message}")
            if r.position:
                print(f"      位置: {r.position}")
            if r.fix_suggestion:
                print(f"      建议: {r.fix_suggestion}")
            for d in r.details[:3]:
                print(f"      -> {d}")
            if len(r.details) > 3:
                print(f"      ... 还有 {len(r.details) - 3} 项")


def load_template(column: str) -> Optional[Dict[str, Any]]:
    """Load column template configuration."""
    if not CONFIG_PATH.exists():
        return None
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("templates", {}).get(column)


class HTMLAuditor:
    def __init__(self, html_content: str, column: str):
        self.html = html_content
        self.column = column
        self.template = load_template(column)
        self.report = AuditReport()

        if HAS_BS4:
            self.soup = BeautifulSoup(html_content, "html.parser")
        else:
            self.soup = None

    def run_audit(self) -> AuditReport:
        """Run full audit against template rules."""
        if not self.template:
            self.report.add(
                AuditResult(
                    category="模板校验",
                    name="模板加载",
                    status=AuditResult.CRITICAL,
                    message=f"栏目「{self.column}」模板未找到",
                    fix_suggestion="请在 column_templates.json 中配置该栏目模板",
                )
            )
            return self.report

        # Run all audit modules
        self._audit_card_borders()
        self._audit_color_scheme()
        self._audit_typography()
        self._audit_required_elements()
        self._audit_images()
        self._audit_structure()

        return self.report

    def _audit_card_borders(self):
        """Check card border compliance."""
        layout = self.template.get("layout_rules", {})
        requires_border = layout.get("card_border", False)

        if not requires_border:
            self.report.add(
                AuditResult(
                    category="排版规范",
                    name="卡片边框检查",
                    status=AuditResult.PASS,
                    message="该栏目不要求卡片边框",
                )
            )
            return

        if not HAS_BS4:
            # Fallback: regex check for common border patterns
            border_patterns = [
                r"border\s*:\s*\d+px\s+solid",
                r"border-radius\s*:\s*\d+px",
                r"box-shadow\s*:",
                r"background\s*:\s*#(?:fff|[Ff]{6})",
            ]
            found_borders = sum(1 for p in border_patterns if re.search(p, self.html))
            if found_borders >= 2:
                self.report.add(
                    AuditResult(
                        category="排版规范",
                        name="卡片边框检查",
                        status=AuditResult.PASS,
                        message=f"检测到卡片化排版特征({found_borders}/4项)",
                    )
                )
            else:
                self.report.add(
                    AuditResult(
                        category="排版规范",
                        name="卡片边框检查",
                        status=AuditResult.CRITICAL,
                        message="未检测到卡片化排版特征, 边框/卡片化铁律未执行",
                        fix_suggestion="确保HTML中包含 border-radius, border, box-shadow 等卡片样式",
                        details=[f"仅匹配到 {found_borders}/4 个边框特征"],
                    )
                )
            return

        # With BeautifulSoup: check for card-like containers
        card_elements = []
        for elem in self.soup.find_all(["section", "div", "article"]):
            style = elem.get("style", "")
            if any(
                s in style
                for s in [
                    "border",
                    "border-radius",
                    "box-shadow",
                    "background:#fff",
                    "background:#FFF",
                ]
            ):
                card_elements.append(elem)

        if len(card_elements) >= 3:
            self.report.add(
                AuditResult(
                    category="排版规范",
                    name="卡片边框检查",
                    status=AuditResult.PASS,
                    message=f"发现 {len(card_elements)} 个卡片容器, 卡片化排版已启用",
                )
            )
        else:
            self.report.add(
                AuditResult(
                    category="排版规范",
                    name="卡片边框检查",
                    status=AuditResult.CRITICAL,
                    message=f"仅发现 {len(card_elements)} 个卡片容器, 卡片化排版未充分执行",
                    fix_suggestion="为文章各段落/章节添加 section/ div 容器, 设置 border, border-radius, box-shadow 样式",
                    details=[f"需要至少3个卡片容器, 当前仅 {len(card_elements)} 个"],
                )
            )

    def _audit_color_scheme(self):
        """Check color scheme compliance."""
        colors = self.template.get("color_scheme", {})
        primary = colors.get("primary", "")

        if not primary:
            self.report.add(
                AuditResult(
                    category="排版规范",
                    name="配色方案检查",
                    status=AuditResult.INFO,
                    message="模板未指定主色调, 跳过检查",
                )
            )
            return

        primary_uses = len(re.findall(re.escape(primary), self.html, re.IGNORECASE))
        if primary_uses >= 3:
            self.report.add(
                AuditResult(
                    category="排版规范",
                    name="配色方案检查",
                    status=AuditResult.PASS,
                    message=f"主色调「{primary}」已使用 {primary_uses} 次",
                )
            )
        else:
            self.report.add(
                AuditResult(
                    category="排版规范",
                    name="配色方案检查",
                    status=AuditResult.WARNING,
                    message=f"主色调「{primary}」仅使用 {primary_uses} 次(建议 >= 3次)",
                    fix_suggestion=f"在标题, 强调文字, 边框等位置增加主色调「{primary}」的使用",
                )
            )

    def _audit_typography(self):
        """Check typography compliance."""
        typo = self.template.get("typography", {})
        expected_body = typo.get("body_size", "15px")

        body_sizes = re.findall(r"font-size\s*:\s*(\d+px)", self.html)
        if body_sizes:
            sizes = [int(s.replace("px", "")) for s in body_sizes]
            most_common = max(set(sizes), key=sizes.count)
            expected_int = int(expected_body.replace("px", ""))
            if abs(most_common - expected_int) <= 1:
                self.report.add(
                    AuditResult(
                        category="排版规范",
                        name="正文字号检查",
                        status=AuditResult.PASS,
                        message=f"正文字号符合要求({most_common}px, 期望 {expected_body})",
                    )
                )
            else:
                self.report.add(
                    AuditResult(
                        category="排版规范",
                        name="正文字号检查",
                        status=AuditResult.WARNING,
                        message=f"正文字号偏差({most_common}px, 期望 {expected_body})",
                        fix_suggestion=f"调整段落字号为 {expected_body}",
                    )
                )
        else:
            self.report.add(
                AuditResult(
                    category="排版规范",
                    name="正文字号检查",
                    status=AuditResult.WARNING,
                    message="未检测到明确的正文字号设置",
                    fix_suggestion=f"为正文章节设置 font-size: {expected_body}",
                )
            )

    def _audit_required_elements(self):
        """Check required structural elements."""
        required = self.template.get("required_elements", [])

        if not required:
            self.report.add(
                AuditResult(
                    category="结构要素",
                    name="必需元素检查",
                    status=AuditResult.PASS,
                    message="该栏目无特殊必需元素",
                )
            )
            return

        missing = []
        found = []

        if "h2_left_bar" in required:
            h2_bars = re.findall(r'<h2[^>]*style="[^"]*border-left[^"]*"', self.html)
            if h2_bars:
                found.append("h2_left_bar")
            else:
                missing.append("H2标题左竖线装饰(border-left)")

        if "quote_block_styled" in required:
            quotes = re.findall(
                r'<blockquote[^>]*style="[^"]*(?:background|rgba\(|#(?:f0f7ff|[Ff]0[Ff]7[Ff][Ff]))',
                self.html,
            )
            if quotes:
                found.append("quote_block_styled")
            else:
                missing.append("样式化引用块(背景色块)")

        if "code_block_styled" in required:
            code_blocks = re.findall(
                r'<section[^>]*style="[^"]*(?:#1e1e1e|[Cc]onsolas|[Mm]onaco)', self.html
            )
            if code_blocks:
                found.append("code_block_styled")
            else:
                missing.append("样式化代码块(深色背景 + 等宽字体)")

        if "table_styled" in required:
            tables = re.findall(r'<table[^>]*style="[^"]*border-collapse', self.html)
            if tables:
                found.append("table_styled")
            else:
                missing.append("样式化表格(border-collapse)")

        if missing:
            self.report.add(
                AuditResult(
                    category="结构要素",
                    name="必需元素检查",
                    status=AuditResult.WARNING,
                    message=f"缺少 {len(missing)} 个必需元素",
                    fix_suggestion="为缺失元素添加对应的HTML结构和样式",
                    details=missing,
                )
            )
        else:
            self.report.add(
                AuditResult(
                    category="结构要素",
                    name="必需元素检查",
                    status=AuditResult.PASS,
                    message=f"所有必需元素已满足: {', '.join(found)}",
                )
            )

    def _audit_images(self):
        """Check image compliance."""
        if HAS_BS4:
            images = self.soup.find_all("img")
            image_count = len(images)
        else:
            image_count = len(re.findall(r"<img[^>]*>", self.html))

        min_images = self.template.get("min_image_count", 1)
        long_threshold = self.template.get("long_article_threshold", 1500)

        text_only = re.sub(r"<[^>]+>", " ", self.html)
        article_length = len(text_only.strip())

        if article_length >= long_threshold:
            min_required = self.template.get("long_article_min_images", 2)
            if image_count >= min_required:
                self.report.add(
                    AuditResult(
                        category="图文规范",
                        name="长文配图检查",
                        status=AuditResult.PASS,
                        message=f"长文({article_length}字) 包含 {image_count} 张图片(要求 >= {min_required})",
                    )
                )
            else:
                self.report.add(
                    AuditResult(
                        category="图文规范",
                        name="长文配图检查",
                        status=AuditResult.CRITICAL,
                        message=f"长文({article_length}字) 仅 {image_count} 张图片(要求 >= {min_required})",
                        fix_suggestion=f"在文章中插入至少 {min_required - image_count} 张配图",
                    )
                )
        else:
            if image_count >= min_images:
                self.report.add(
                    AuditResult(
                        category="图文规范",
                        name="配图检查",
                        status=AuditResult.PASS,
                        message=f"文章包含 {image_count} 张图片(要求 >= {min_images})",
                    )
                )
            else:
                self.report.add(
                    AuditResult(
                        category="图文规范",
                        name="配图检查",
                        status=AuditResult.WARNING,
                        message=f"文章仅 {image_count} 张图片(建议 >= {min_images})",
                        fix_suggestion=f"在关键段落插入配图增强阅读体验",
                    )
                )

    def _audit_structure(self):
        """Check overall document structure."""
        has_h1 = bool(re.search(r"<h1", self.html))
        has_h2 = bool(re.search(r"<h2", self.html))
        has_p = bool(re.search(r"<p", self.html))

        issues = []
        if not has_h1:
            issues.append("缺少H1标题")
        if not has_p:
            issues.append("缺少正文段落")

        if issues:
            self.report.add(
                AuditResult(
                    category="文档结构",
                    name="基础结构检查",
                    status=AuditResult.CRITICAL,
                    message="文档结构不完整",
                    fix_suggestion="确保文档包含基础的标题和段落结构",
                    details=issues,
                )
            )
        else:
            self.report.add(
                AuditResult(
                    category="文档结构",
                    name="基础结构检查",
                    status=AuditResult.PASS,
                    message=f"基础结构完整(H1={'是' if has_h1 else '否'}, H2={'是' if has_h2 else '否'}, P={'是' if has_p else '否'})",
                )
            )

        empty_sections = re.findall(r"<section[^>]*>\s*</section>", self.html)
        if empty_sections:
            self.report.add(
                AuditResult(
                    category="文档结构",
                    name="空容器检查",
                    status=AuditResult.WARNING,
                    message=f"发现 {len(empty_sections)} 个空 section 容器",
                    fix_suggestion="移除空容器或填充内容",
                )
            )
        else:
            self.report.add(
                AuditResult(
                    category="文档结构",
                    name="空容器检查",
                    status=AuditResult.PASS,
                    message="未发现空容器",
                )
            )


def run_audit(html_content: str, column: str) -> AuditReport:
    """Run HTML audit against column template."""
    auditor = HTMLAuditor(html_content, column)
    return auditor.run_audit()


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Post-audit scanner for WeChat HTML")
    parser.add_argument("html_path", help="Path to generated HTML file")
    parser.add_argument("--column", default="江湖夜话", help="WeChat column name")
    parser.add_argument("--json", action="store_true", help="Output as JSON")
    parser.add_argument("--bypass", action="store_true", help="Bypass blocking")
    args = parser.parse_args()

    if not os.path.exists(args.html_path):
        print(f"ERROR: File not found: {args.html_path}")
        sys.exit(1)

    with open(args.html_path, "r", encoding="utf-8") as f:
        html_content = f.read()

    report = run_audit(html_content, args.column)

    if args.json:
        import json as _json

        print(_json.dumps(report.to_dict(), ensure_ascii=False, indent=2))
    else:
        report.print_report()

    if report.blocked and not args.bypass:
        print("\n[BLOCKED] 发现严重排版问题, 请在修复后重新生成HTML.")
        print("如需强制继续, 请添加 --bypass 参数.")
        sys.exit(2)
    elif report.blocked and args.bypass:
        print("\n[WARNING] 已绕过严重问题(此操作已被记录).")
        sys.exit(0)
    else:
        print("\n[PASS] 后置审计通过, 可以发布.")
        sys.exit(0)


if __name__ == "__main__":
    main()
