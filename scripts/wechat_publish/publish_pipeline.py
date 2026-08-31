#!/usr/bin/env python3
"""WeChat Article Publishing Pipeline - Three-Layer Defense System.

v1.0 (2026-08-12): Initial release.
  Layer 1: Pre-publish Checklist (pre_publish_checklist.py)
  Layer 2: Column Template Binding (column_templates.json + publish_bidding_law.py)
  Layer 3: Post-Audit Scanner (post_audit_scanner.py)

Usage:
  python publish_pipeline.py <md_path> <cover_path> --column <column_name>
  python publish_pipeline.py <md_path> <cover_path> --column <column_name> --skip-checklist  # Bypass layer 1
  python publish_pipeline.py <md_path> <cover_path> --column <column_name> --skip-audit      # Bypass layer 3

Exit codes:
  0 - Success, draft created
  1 - System error (file not found, API error, etc.)
  2 - Layer 1 blocked (checklist failure)
  3 - Layer 3 blocked (audit failure)
"""

import json
import os
import sys
import subprocess
from pathlib import Path

# ============================================================
# Configuration
# ============================================================
SCRIPT_DIR = Path(__file__).parent.resolve()
CHECKLIST_SCRIPT = SCRIPT_DIR / "pre_publish_checklist.py"
AUDIT_SCRIPT = SCRIPT_DIR / "post_audit_scanner.py"
PUBLISH_SCRIPT = SCRIPT_DIR / "publish_bidding_law_v3.py"
CONFIG_PATH = SCRIPT_DIR / "column_templates.json"


# ============================================================
# Pipeline Orchestrator
# ============================================================
class PublishPipeline:
    def __init__(
        self,
        md_path: str,
        cover_path: str,
        column: str,
        skip_checklist: bool = False,
        skip_audit: bool = False,
        bypass_checklist: bool = False,
        bypass_audit: bool = False,
        verbose: bool = True,
        dry_run: bool = False,
    ):
        self.md_path = md_path
        self.cover_path = cover_path
        self.column = column
        self.skip_checklist = skip_checklist
        self.skip_audit = skip_audit
        self.bypass_checklist = bypass_checklist
        self.bypass_audit = bypass_audit
        self.verbose = verbose
        self.dry_run = dry_run
        self.checklist_passed = False
        self.audit_passed = False
        self.draft_id = None

    def log(self, msg: str):
        if self.verbose:
            print(f"[PIPELINE] {msg}")

    def run(self) -> int:
        """Execute the full three-layer pipeline."""
        self.log(f"=" * 60)
        self.log(f"微信公众号发布流水线 - 三层防御")
        self.log(f"=" * 60)
        self.log(f"文件: {self.md_path}")
        self.log(f"封面: {self.cover_path}")
        self.log(f"栏目: {self.column}")
        self.log(f"=" * 60)

        # Verify files exist
        if not os.path.exists(self.md_path):
            self.log(f"错误: Markdown 文件不存在: {self.md_path}")
            return 1
        if not os.path.exists(self.cover_path):
            self.log(f"错误: 封面图片不存在: {self.cover_path}")
            return 1
        if not CHECKLIST_SCRIPT.exists():
            self.log(f"错误: 检查清单脚本不存在: {CHECKLIST_SCRIPT}")
            return 1
        if not PUBLISH_SCRIPT.exists():
            self.log(f"错误: 排版脚本不存在: {PUBLISH_SCRIPT}")
            return 1

        # Layer 1: Pre-publish Checklist
        if not self.skip_checklist:
            self.log("\n[LAYER 1] 正在运行排版前检查清单...")
            exit_code = self._run_checklist()
            if exit_code != 0:
                if self.bypass_checklist:
                    self.log("[警告] 已绕过检查清单未通过(日志已记录)")
                    self.checklist_passed = True
                else:
                    self.log("[阻断] 第一层防御激活 - 检查清单未通过")
                    return 2
            else:
                self.checklist_passed = True
                self.log("[通过] 第一层防御通过 - 检查清单全部合格")
        else:
            self.log("\n[LAYER 1] 跳过检查清单( --skip-checklist )")
            self.checklist_passed = True

        # Layer 2: Column Template Binding + HTML Generation
        self.log("\n[LAYER 2] 正在运行栏目模板绑定与排版...")
        html_path = self._run_layout_generation()
        if not html_path:
            return 1
        self.log(f"[通过] 第二层防御通过 - HTML生成完成: {html_path}")

        # Layer 3: Post-Audit Scanner
        if not self.skip_audit:
            self.log("\n[LAYER 3] 正在运行后置审计...")
            audit_code = self._run_audit(html_path)
            if audit_code != 0:
                if self.bypass_audit:
                    self.log("[警告] 已绕过审计未通过(日志已记录)")
                    self.audit_passed = True
                else:
                    self.log("[阻断] 第三层防御激活 - 审计未通过")
                    return 3
            else:
                self.audit_passed = True
                self.log("[通过] 第三层防御通过 - 审计全部合格")
        else:
            self.log("\n[LAYER 3] 跳过审计( --skip-audit )")
            self.audit_passed = True

        # Final: Push to WeChat Draft
        if self.dry_run:
            self.log("\n[DRY-RUN] 跳过推送微信草稿箱(--dry-run)")
            self.log(f"  HTML 已生成: {html_path}")
            self.log(f"  三层防御: 全部通过")
            self.log(f"{'=' * 60}\n")
            return 0

        self.log("\n[最终] 正在推送到微信草稿箱...")
        draft_id = self._run_publish(html_path)
        if not draft_id:
            return 1

        self.draft_id = draft_id
        self.log(f"\n{'=' * 60}")
        self.log(f"  成功! 草稿已创建")
        self.log(f"  MEDIA_ID: {draft_id}")
        self.log(f"  栏目: {self.column}")
        self.log(f"  三层防御: 全部通过")
        self.log(f"{'=' * 60}\n")

        # Cleanup intermediate HTML file
        try:
            os.remove(html_path)
            self.log(f"[清理] 已删除中间HTML文件")
        except OSError:
            pass

        return 0

    def _run_checklist(self) -> int:
        """Run Layer 1 checklist."""
        cmd = [
            sys.executable,
            str(CHECKLIST_SCRIPT),
            self.md_path,
            "--column",
            self.column,
        ]
        if self.bypass_checklist:
            cmd.append("--bypass")

        self.log(f"  执行: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")

        if result.stdout:
            for line in result.stdout.strip().split("\n"):
                self.log(f"  {line}")
        if result.stderr:
            for line in result.stderr.strip().split("\n"):
                self.log(f"  [STDERR] {line}")

        return result.returncode

    def _run_layout_generation(self) -> str:
        """Run Layer 2: Generate HTML from markdown with column template."""
        # Generate HTML to temp file
        html_path = Path(self.md_path).with_suffix(".pipeline.html")

        cmd = [
            sys.executable,
            str(PUBLISH_SCRIPT),
            self.md_path,
            self.cover_path,
            "--column",
            self.column,
            "--output-html",
            str(html_path),
        ]

        self.log(f"  执行: 排版引擎生成HTML")
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")

        if result.stdout:
            for line in result.stdout.strip().split("\n"):
                self.log(f"  {line}")
        if result.stderr:
            for line in result.stderr.strip().split("\n"):
                self.log(f"  [STDERR] {line}")

        if result.returncode != 0:
            self.log(f"  排版失败! 退出码: {result.returncode}")
            return ""

        # Verify HTML was generated
        if html_path.exists():
            return str(html_path)
        else:
            self.log("  HTML文件未生成")
            return ""

    def _run_audit(self, html_path: str) -> int:
        """Run Layer 3 audit on generated HTML."""
        cmd = [
            sys.executable,
            str(AUDIT_SCRIPT),
            html_path,
            "--column",
            self.column,
        ]
        if self.bypass_audit:
            cmd.append("--bypass")

        self.log(f"  执行: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")

        if result.stdout:
            for line in result.stdout.strip().split("\n"):
                self.log(f"  {line}")
        if result.stderr:
            for line in result.stderr.strip().split("\n"):
                self.log(f"  [STDERR] {line}")

        return result.returncode

    def _run_publish(self, html_path: str) -> str:
        """Run final publish step to push to WeChat draft."""
        # Re-run publish script (it will generate HTML internally and push to WeChat)
        cmd = [
            sys.executable,
            str(PUBLISH_SCRIPT),
            self.md_path,
            self.cover_path,
            "--column",
            self.column,
        ]

        self.log(f"  执行: 推送草稿箱")
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")

        if result.stdout:
            for line in result.stdout.strip().split("\n"):
                self.log(f"  {line}")
        if result.stderr:
            for line in result.stderr.strip().split("\n"):
                self.log(f"  [STDERR] {line}")

        if result.returncode != 0:
            self.log(f"  推送失败! 退出码: {result.returncode}")
            return ""

        # Extract draft_id from stdout
        for line in result.stdout.split("\n"):
            if "Draft media_id:" in line or "media_id:" in line:
                parts = line.split(":")
                if len(parts) >= 2:
                    return parts[-1].strip()

        return "unknown"


# ============================================================
# Main Entry Point
# ============================================================
def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="WeChat Article Publishing Pipeline - Three-Layer Defense",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Full pipeline (recommended)
  python publish_pipeline.py article.md cover.jpg --column 江湖夜话

  # Skip checklist (emergency)
  python publish_pipeline.py article.md cover.jpg --column 江湖夜话 --skip-checklist

  # Bypass with warning logged
  python publish_pipeline.py article.md cover.jpg --column 江湖夜话 --bypass-checklist --bypass-audit
""",
    )
    parser.add_argument("md_path", help="Path to markdown article")
    parser.add_argument("cover_path", help="Path to cover image (max 64KB)")
    parser.add_argument(
        "--column", default="江湖夜话", help="WeChat column name (default: 江湖夜话)"
    )
    parser.add_argument(
        "--skip-checklist",
        action="store_true",
        help="Skip Layer 1 pre-publish checklist (NOT RECOMMENDED)",
    )
    parser.add_argument(
        "--skip-audit",
        action="store_true",
        help="Skip Layer 3 post-audit scanner (NOT RECOMMENDED)",
    )
    parser.add_argument(
        "--bypass-checklist",
        action="store_true",
        help="Bypass checklist blocking (logs warning)",
    )
    parser.add_argument(
        "--bypass-audit",
        action="store_true",
        help="Bypass audit blocking (logs warning)",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        default=True,
        help="Verbose output (default: True)",
    )
    parser.add_argument(
        "--quiet", dest="verbose", action="store_false", help="Minimal output"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Run all 3 layers but skip final WeChat push (for testing)",
    )
    args = parser.parse_args()

    pipeline = PublishPipeline(
        md_path=args.md_path,
        cover_path=args.cover_path,
        column=args.column,
        skip_checklist=args.skip_checklist,
        skip_audit=args.skip_audit,
        bypass_checklist=args.bypass_checklist,
        bypass_audit=args.bypass_audit,
        verbose=args.verbose,
        dry_run=args.dry_run,
    )

    exit_code = pipeline.run()
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
