"""bid-toolkit CLI — 一行命令跑全套标书工具链"""

import argparse
import os
import sys
import importlib.util


def _scripts_dir():
    """定位 scripts 目录（兼容 editable install 和正式 install）"""
    # editable install: __file__ 在源代码目录
    pkg_dir = os.path.dirname(os.path.abspath(__file__))
    scripts_dir = os.path.join(pkg_dir, "scripts")
    if os.path.isdir(scripts_dir):
        return scripts_dir
    # 正式 install: 从 site-packages 找
    import bid_toolkit
    scripts_dir = os.path.join(os.path.dirname(bid_toolkit.__file__), "scripts")
    return scripts_dir


def _load_script(name):
    """从 scripts 目录动态加载模块"""
    scripts_dir = _scripts_dir()
    path = os.path.join(scripts_dir, name)
    if not os.path.isfile(path):
        print(f"❌ 找不到脚本: {path}")
        sys.exit(1)
    spec = importlib.util.spec_from_file_location(name.replace(".py", ""), path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[mod.__name__] = mod
    spec.loader.exec_module(mod)
    return mod


def cmd_engine(args):
    """Markdown → Word 标书排版"""
    mod = _load_script("bid_engine.py")
    sys.argv = ["bid_engine.py", args.input]
    if args.output:
        sys.argv.extend(["-o", args.output])
    mod.main()


def cmd_check(args):
    """标书格式自检 + 评分项覆盖检查"""
    mod = _load_script("fix_bid_format.py")
    print(f"🔍 格式检查: {args.file}")
    sys.argv = ["fix_bid_format.py", args.file]
    try:
        mod.main()
    except TypeError as e:
        print(f"\n⚠️  扫描完成，但自动修复阶段遇到小问题: {e}")
        print("   可手动查看生成的扫描报告")

    # 评分项覆盖检查
    if args.coverage:
        print(f"\n{'='*60}")
        print(f"📋 评分项覆盖检查")
        print(f"{'='*60}")
        cov_mod = _load_script("coverage_check.py")
        sys.argv = ["coverage_check.py", args.file]
        if args.coverage_type:
            sys.argv.extend(["--type", args.coverage_type])
        if args.coverage_output:
            sys.argv.extend(["--output", args.coverage_output])
        cov_mod.main()


def cmd_rfp(args):
    """招标文件生成"""
    rfp_dir = os.path.join(os.path.dirname(_scripts_dir()), "rfp")
    sys.path.insert(0, rfp_dir)
    spec = importlib.util.spec_from_file_location(
        "rfp_generator", os.path.join(rfp_dir, "rfp_generator.py")
    )
    mod = importlib.util.module_from_spec(spec)
    sys.modules["rfp_generator"] = mod
    spec.loader.exec_module(mod)

    sys.argv = ["rfp_generator.py"]
    if args.type:
        sys.argv.extend(["--type", args.type])
    if args.project:
        sys.argv.extend(["--project", args.project])
    if args.budget:
        sys.argv.extend(["--budget", str(args.budget)])
    if args.docx:
        sys.argv.append("--docx")
    if args.output:
        sys.argv.extend(["--output", args.output])
    mod.main()


def cmd_review(args):
    """招标文件风险扫描 — 三层审标管线"""
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
    from bid_toolkit.review import scanner, report as rpt, reverse_check

    print(f"📋 审标扫描: {args.input}")
    print(f"{'='*60}")
    print(f"  Layer 1: 判词库逐行扫描 ...")
    result = scanner.scan_tender(args.input, with_llm=args.llm)
    print(f"  ✅ 完成: {len(result.hits)} 处命中")
    print()

    # Layer 3: 反向覆盖检查（需指定投标书）
    coverage_result = None
    if args.bid_file:
        print(f"  Layer 3: 反向覆盖检查（vs 投标书）...")
        coverage_result = reverse_check.reverse_coverage_check(result, args.bid_file)
        print(f"  ✅ 完成: {coverage_result['coverage_rate']}")

    # 输出报告
    rpt.format_report(result)

    if coverage_result:
        reverse_check.print_coverage_report(coverage_result)

    # 导出报告
    if args.output:
        is_json = args.output.endswith('.json')
        if is_json:
            # JSON 输出
            fatals_list = [{'keyword': h.keyword, 'category': h.category, 'line_num': h.line_num,
                           'context': h.context[:80], 'llm_label': h.llm_label, 'llm_reason': h.llm_reason}
                          for h in result.hits if h.llm_label == 'fatal']
            warns_list = [{'keyword': h.keyword, 'category': h.category, 'line_num': h.line_num,
                          'context': h.context[:80], 'llm_label': h.llm_label, 'llm_reason': h.llm_reason}
                         for h in result.hits if h.llm_label == 'warn']
            json_out = {
                'file': args.input,
                'total_hits': len(result.hits),
                'summary': {
                    'fatal': len([h for h in result.hits if h.llm_label == 'fatal']),
                    'warn': len([h for h in result.hits if h.llm_label == 'warn']),
                    'info': len([h for h in result.hits if h.llm_label == 'info']),
                },
                'fatals': fatals_list,
                'warns': warns_list,
                'coverage': coverage_result,
            }
            import json
            with open(args.output, 'w', encoding='utf-8') as f:
                json.dump(json_out, f, ensure_ascii=False, indent=2)
            print(f"📁 JSON报告已导出: {args.output} ({len(fatals_list)}fatal/{len(warns_list)}warn)")
        else:
            # Markdown 输出
            md = rpt.format_checklist_md(result)
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(md)
            print(f"📁 完整报告已导出: {args.output}")

            if coverage_result:
                cov_path = args.output.replace('.md', '_覆盖检查.md') if not is_json else args.output.replace('.json', '_覆盖检查.json')
                with open(cov_path, 'w', encoding='utf-8') as f:
                    f.write(f"# 反向覆盖检查报告\n\n")
                    f.write(f"总风险项: {coverage_result['total']} | 已回应: {coverage_result['covered']} | 未回应: {coverage_result['missing']}\n\n")
                    for item in coverage_result['items']:
                        icon = '✅' if item['status'] == 'covered' else '❌'
                        f.write(f"- {icon} `{item['keyword']}` L{item['line_num']} — {item['suggestion']}\n")
                print(f"📁 覆盖检查已导出: {cov_path}")


def cmd_desense(args):
    """敏感信息脱敏扫描"""
    mod = _load_script("desensitization_scan.py")
    print(f"🔍 脱敏扫描: {args.file}")
    # 这个脚本没有 main()，直接调 generate_report
    report = mod.generate_report()
    mod.print_report(report)


def cmd_watermark(args):
    """Word文档添加文字水印"""
    mod = _load_script("watermark.py")
    print(f"💧 添加水印: {args.input} → {args.output}")
    mod.add_watermark(args.input, args.output,
                      text=args.text, font=args.font,
                      color=args.color, opacity=args.opacity)
    print(f"✅ 水印 \"{args.text}\" 已添加到 {args.output}")


def cmd_materials(args):
    """标书素材库管家 — 自动分类整理 + 必备材料清单状态"""
    mod = _load_script("bid_materials.py")
    sys.argv = ["bid_materials.py", args.action, args.dir]
    if getattr(args, "min_conf", None) is not None:
        sys.argv.extend(["--min-conf", str(args.min_conf)])
    if getattr(args, "force", False):
        sys.argv.append("--force")
    if getattr(args, "verbose", False):
        sys.argv.append("-v")
    if getattr(args, "llm", False):
        sys.argv.append("--llm")
    mod.main()


def cmd_gui(args):
    """启动桌面图形界面（Gradio + pywebview）"""
    try:
        from bid_toolkit.gui import launch
        launch(port=args.port, browser_only=args.browser)
    except ImportError:
        print("缺少桌面端依赖: gradio")
        print('请运行: pip install "bid-toolkit[desktop]"')
        sys.exit(1)

def cmd_list(args=None):
    """列出所有可用工具"""
    tools = [
        ("engine",   "Markdown转Word标书排版",   "bid engine input.md -o output.docx"),
        ("check",    "标书格式自检 + 评分项覆盖", "bid check input.docx --coverage"),
        ("review",   "招标文件风险扫描（三层审标）", "bid review 招标文件.pdf -o report.md"),
        ("rfp",      "招标文件生成器",           "bid rfp --type services --project XX项目"),
        ("desense",  "敏感信息脱敏扫描",         "bid desense input.docx"),
        ("watermark","Word文档添加文字水印",      "bid watermark input.docx output.docx -t 仅供参考"),
        ("materials","标书素材库管家（自动分类整理+材料清单）", "bid materials analyze 素材库目录"),
        ("gui",      "启动桌面图形界面",           "bid gui"),
    ]
    print("📋 bid-toolkit 工具清单\n")
    for name, desc, example in tools:
        print(f"  bid {name:<8} {desc}")
        print(f"    {'':8} 示例: {example}")
        print()


def main():
    parser = argparse.ArgumentParser(
        prog="bid",
        description="📄 bid-toolkit — 招投标全流程工具链",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="command", help="子命令")

    # engine
    p = sub.add_parser("engine", help="Markdown → Word 标书排版")
    p.add_argument("input", help="输入的 Markdown 文件路径")
    p.add_argument("-o", "--output", default=None, help="输出 Word 路径")

    # check
    p = sub.add_parser("check", help="标书格式自检 + 评分项覆盖检查")
    p.add_argument("file", help="Word 文件路径")
    p.add_argument("--coverage", action="store_true", help="同时检查评分项覆盖情况")
    p.add_argument("--coverage-type", "-t", choices=["工程", "服务", "货物"], help="标书类型（不指定则自动识别）")
    p.add_argument("--coverage-output", help="覆盖检查报告导出路径（.txt / .md）")

    # review
    p = sub.add_parser("review", help="招标文件风险扫描（三层审标管线）")
    p.add_argument("input", help="输入的招标文件路径（PDF/DOCX/MD）")
    p.add_argument("--bid-file", "-b", help="投标书路径（可选，开启Layer 3反向覆盖检查）")
    p.add_argument("--output", "-o", help="导出审标报告到文件（.md / .json）")
    p.add_argument("--llm", action="store_true", help="启用LLM上下文判断（Layer 2）")

    # rfp
    p = sub.add_parser("rfp", help="招标文件生成")
    p.add_argument("--type", choices=["services", "goods", "engineering"], help="项目类型")
    p.add_argument("--project", help="项目名称")
    p.add_argument("--budget", type=float, help="预算金额（元）")
    p.add_argument("--docx", action="store_true", help="同时输出 Word 格式")
    p.add_argument("--output", help="输出目录")

    # desense
    p = sub.add_parser("desense", help="敏感信息脱敏扫描")
    p.add_argument("file", help="输入文件路径")

    # watermark
    p = sub.add_parser("watermark", help="Word文档添加文字水印")
    p.add_argument("input", help="输入 Word 文件路径")
    p.add_argument("output", help="输出 Word 文件路径")
    p.add_argument("-t", "--text", default="仅供参考", help="水印文字 (默认: 仅供参考)")
    p.add_argument("-f", "--font", default="微软雅黑", help="字体 (默认: 微软雅黑)")
    p.add_argument("-c", "--color", default="#808080", help="颜色十六进制 (默认: #808080)")
    p.add_argument("-o", "--opacity", default="0.5", help="透明度 0-1 (默认: 0.5)")

    # materials
    p = sub.add_parser("materials", help="标书素材库管家（自动分类整理+材料清单）")
    p.add_argument("action", choices=["init", "analyze", "apply", "status", "learn"],
                   help="操作: init初始化 / analyze扫描 / apply执行 / status状态 / learn学习")
    p.add_argument("dir", nargs="?", default=".", help="素材库目录 (默认: 当前目录)")
    p.add_argument("--min-conf", type=float, default=None, help="置信度阈值 (默认0.7)")
    p.add_argument("--force", action="store_true", help="apply时跳过needs_review强制执行")
    p.add_argument("-v", "--verbose", action="store_true", help="apply时显示移动明细")
    p.add_argument("--llm", action="store_true", help="learn时开启AI建议")

    # list
    sub.add_parser("list", help="列出所有可用工具")

    # gui
    p = sub.add_parser("gui", help="启动桌面图形界面")
    p.add_argument("--port", type=int, default=7860, help="端口号 (默认: 7860)")
    p.add_argument("--browser", action="store_true", help="浏览器模式（不启动桌面窗口）")

    args = parser.parse_args()

    if args.command == "engine":
        cmd_engine(args)
    elif args.command == "check":
        cmd_check(args)
    elif args.command == "review":
        cmd_review(args)
    elif args.command == "rfp":
        cmd_rfp(args)
    elif args.command == "desense":
        cmd_desense(args)
    elif args.command == "watermark":
        cmd_watermark(args)
    elif args.command == "materials":
        cmd_materials(args)
    elif args.command == "list":
        cmd_list()
    elif args.command == "gui":
        cmd_gui(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
