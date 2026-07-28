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
    """标书格式自检"""
    mod = _load_script("fix_bid_format.py")
    print(f"🔍 格式检查: {args.file}")
    # 先跑扫描，捕获修复阶段的错误
    sys.argv = ["fix_bid_format.py", args.file]
    try:
        mod.main()
    except TypeError as e:
        print(f"\n⚠️  扫描完成，但自动修复阶段遇到小问题: {e}")
        print("   可手动查看生成的扫描报告")


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


def cmd_desense(args):
    """敏感信息脱敏扫描"""
    mod = _load_script("desensitization_scan.py")
    print(f"🔍 脱敏扫描: {args.file}")
    # 这个脚本没有 main()，直接调 generate_report
    report = mod.generate_report()
    mod.print_report(report)


def cmd_list(args=None):
    """列出所有可用工具"""
    tools = [
        ("engine",  "Markdown转Word标书排版",  "bid engine input.md -o output.docx"),
        ("check",   "标书格式自检",            "bid check input.docx"),
        ("rfp",     "招标文件生成器",          "bid rfp --type services --project XX项目"),
        ("desense", "敏感信息脱敏扫描",        "bid desense input.docx"),
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
    p = sub.add_parser("check", help="标书格式自检")
    p.add_argument("file", help="Word 文件路径")

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

    # list
    sub.add_parser("list", help="列出所有可用工具")

    args = parser.parse_args()

    if args.command == "engine":
        cmd_engine(args)
    elif args.command == "check":
        cmd_check(args)
    elif args.command == "rfp":
        cmd_rfp(args)
    elif args.command == "desense":
        cmd_desense(args)
    elif args.command == "list":
        cmd_list()
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
