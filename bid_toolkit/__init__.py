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

    # 引擎铁律校验
    if args.iron:
        print(f"\n{'='*60}")
        print(f"⚖️  引擎铁律校验（8项铁律）")
        print(f"{'='*60}")
        from bid_toolkit.orchestrator.engine import main as engine_main

        # 尝试从docx提取文本用于校验
        content_file = args.file
        if args.file.endswith(".docx"):
            try:
                import importlib.util as _ilu
                _scripts = _scripts_dir()
                _spec = _ilu.spec_from_file_location(
                    "bid_table_generator",
                    os.path.join(_scripts, "bid_table_generator.py"))
                _mod = _ilu.module_from_spec(_spec)
                _spec.loader.exec_module(_mod)
                text = _mod.extract_text_from_docx(args.file)
                tmp_md = args.file.rsplit(".", 1)[0] + "_iron_check.md"
                with open(tmp_md, "w", encoding="utf-8") as f:
                    f.write(text)
                content_file = tmp_md
            except Exception:
                pass  # 直接用docx路径，引擎内部会处理

        old_argv = sys.argv
        sys.argv = ["bid-orchestrate", "check", content_file, "--docx", args.file]
        try:
            engine_main()
        except Exception as e:
            print(f"\n⚠️  铁律校验遇到问题: {e}")
            print("   可单独运行: bid orchestrate check <内容文件>")
        finally:
            sys.argv = old_argv


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
        _out_dir = os.path.dirname(os.path.abspath(args.output))
        if _out_dir:
            os.makedirs(_out_dir, exist_ok=True)
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
    """AI味检测 + 敏感信息脱敏扫描"""
    # 优先调用 AI味雷达（检测AI文风）
    mod = _load_script("ai_flavor_radar.py")
    print(f"🔍 AI味雷达扫描: {args.file}")
    print(f"{'='*60}")
    extra_args = []
    if getattr(args, "no_color", False):
        extra_args.append("--no-color")
    if getattr(args, "no_examples", False):
        extra_args.append("--no-examples")
    mod.run_scan(args.file, mode=args.mode, fmt=args.format, output=args.output, extra_args=extra_args)
    print()

    # 敏感信息脱敏扫描（原功能）
    desense_mod = _load_script("desensitization_scan.py")
    print(f"🔍 敏感信息扫描: {args.file}")
    print(f"{'='*60}")
    report = desense_mod.generate_report()
    desense_mod.print_report(report)


def cmd_watermark(args):
    """Word文档添加文字水印"""
    mod = _load_script("watermark.py")
    print(f"💧 添加水印: {args.input} → {args.output}")
    mod.add_watermark(args.input, args.output,
                      text=args.text, font=args.font,
                      color=args.color, opacity=args.opacity)
    print(f"✅ 水印 \"{args.text}\" 已添加到 {args.output}")


def cmd_commitments(args):
    """承诺链三源追踪：扫描标书中的承诺并对照企业资料库验证"""
    # 承诺链脚本在 scripts/ 根目录（不是 bid_toolkit/scripts/）
    scripts_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "scripts")
    path = os.path.join(scripts_dir, "commitment_scanner.py")
    if not os.path.isfile(path):
        print(f"❌ 找不到脚本: {path}")
        sys.exit(1)
    spec = importlib.util.spec_from_file_location("commitment_scanner", path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[mod.__name__] = mod
    spec.loader.exec_module(mod)
    sys.argv = ["commitment_scanner.py", args.file]
    if args.profile:
        sys.argv.extend(["--profile", args.profile])
    if args.output:
        sys.argv.extend(["--report", args.output])
    mod.main()


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


def cmd_analyze(args):
    """投标可行性分析 — 一键拆解招标文件+评分分析+承诺审计+时间规划"""
    mod = _load_script("bid_analyzer.py")
    sys.argv = ["bid_analyzer.py", args.input]
    if args.profile:
        sys.argv += ["--profile", args.profile]
    if args.output:
        sys.argv += ["--output", args.output]
    if args.config:
        sys.argv += ["--config", args.config]
    if args.list_only:
        sys.argv.append("--list-only")
    if args.gen_config:
        sys.argv.append("--gen-config")
    if args.json:
        sys.argv.append("--json")
    if args.open_date:
        sys.argv += ["--open-date", args.open_date]
    mod.main()


def cmd_table(args):
    """工程标两列表格生成器 — 从DOC/PDF/MD提取文字自动生成两列表格"""
    mod = _load_script("bid_table_generator.py")
    sys.argv = ["bid_table_generator.py", args.input]
    if args.output:
        sys.argv += ["-o", args.output]
    if args.config:
        sys.argv += ["--config", args.config]
    if args.left_width is not None:
        sys.argv += ["--left-width", str(args.left_width)]
    if args.right_width is not None:
        sys.argv += ["--right-width", str(args.right_width)]
    if args.left_color:
        sys.argv += ["--left-color", args.left_color]
    if args.right_color:
        sys.argv += ["--right-color", args.right_color]
    if args.border_color:
        sys.argv += ["--border-color", args.border_color]
    if args.font_size is not None:
        sys.argv += ["--font-size", str(args.font_size)]
    if args.font_name:
        sys.argv += ["--font-name", args.font_name]
    if args.no_merge:
        sys.argv.append("--no-merge")
    if args.list:
        sys.argv.append("--list")
    if args.gen_config:
        sys.argv += ["--gen-config", args.gen_config]
    mod.main()


def cmd_map_clauses(args):
    """条款-方案映射审计：招标文件评分条款 → 方案章节覆盖分析"""
    scripts_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "scripts")
    path = os.path.join(scripts_dir, "clause_mapper.py")
    if not os.path.isfile(path):
        print(f"❌ 找不到脚本: {path}")
        sys.exit(1)
    spec = importlib.util.spec_from_file_location("clause_mapper", path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[mod.__name__] = mod
    spec.loader.exec_module(mod)
    sys.argv = ["clause_mapper.py", args.tender_file]
    if args.bid_file:
        sys.argv.extend(["-b", args.bid_file])
    if args.report:
        sys.argv.extend(["--report", args.report])
    if getattr(args, "json_output", None):
        sys.argv.extend(["--json", args.json_output])
    mod.main()


def cmd_render(args):
    """content.json 转 Word 排版"""
    from bid_toolkit.render_docx import build
    try:
        build(args.content, args.output, format_config=args.config)
    except ValueError as e:
        print(str(e))
        sys.exit(1)


def cmd_validate(args):
    """校验 content.json 结构（不渲染，提前暴露错误）"""
    import json
    from bid_toolkit.render_docx import validate_content_json
    path = args.content
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        print(f"❌ JSON 语法错误（第 {e.lineno} 行第 {e.colno} 列）：{e.msg}")
        sys.exit(1)
    except FileNotFoundError:
        print(f"❌ 找不到文件: {path}")
        sys.exit(1)
    ok, errs = validate_content_json(data, path)
    if ok:
        n = len(data.get('body', []))
        print(f"✅ content.json 校验通过，共 {n} 个内容块，可安全渲染。")
    else:
        print("❌ content.json 结构校验未通过：")
        for e in errs:
            print(f"   • {e}")
        sys.exit(1)


def cmd_orchestrate(args):
    """编排引擎 - 框架锁定/差异检测/铁律校验"""
    from bid_toolkit.orchestrator.engine import main as engine_main

    # 把 bid orchestrate <sub> 转成引擎CLI的参数
    engine_argv = ["bid-orchestrate"]
    if args.sub == "init":
        engine_argv.append("init")
        if args.name:
            engine_argv.extend(["--name", args.name])
    elif args.sub == "parse":
        engine_argv.extend(["parse", args.tender_file])
    elif args.sub == "framework":
        engine_argv.append("framework")
        if args.template:
            engine_argv.extend(["--template", args.template])
        if args.from_tender:
            engine_argv.append("--from-tender")
    elif args.sub == "lock":
        engine_argv.append("lock")
    elif args.sub == "unlock":
        engine_argv.extend(["unlock", "--confirm"])
    elif args.sub == "diff":
        engine_argv.extend(["diff", args.content_file])
    elif args.sub == "check":
        engine_argv.extend(["check", args.content_file])
        if args.docx:
            engine_argv.extend(["--docx", args.docx])
    elif args.sub == "status":
        engine_argv.append("status")
    elif args.sub == "templates":
        engine_argv.append("templates")
    elif args.sub == "prelock":
        engine_argv.append("prelock")
    elif args.sub == "anchor":
        engine_argv.extend(["anchor", args.content_file])
    elif args.sub == "profile":
        engine_argv.append("profile")
        if getattr(args, "set_file", None):
            engine_argv.extend(["--set-file", args.set_file])
        if getattr(args, "set", None):
            engine_argv.extend(["--set", args.set])
        if getattr(args, "show", False):
            engine_argv.append("--show")
    else:
        engine_argv.append("--help")

    old_argv = sys.argv
    sys.argv = engine_argv
    try:
        engine_main()
    finally:
        sys.argv = old_argv


def cmd_list(args=None):
    """列出所有可用工具"""
    tools = [
        ("render",   "content.json 转 Word 排版", "bid render content.json output.docx"),
        ("orchestrate", "编排引擎（框架锁定/差异检测/铁律校验）", "bid orchestrate status"),
        ("analyze",  "投标可行性分析（拆解+评分+承诺审计+时间规划）",
         "bid analyze 招标文件.md --profile company_profile/ -o 报告.md"),
        ("engine",   "Markdown转Word标书排版",   "bid engine input.md -o output.docx"),
        ("check",    "标书格式自检 + 评分项覆盖", "bid check input.docx --coverage"),
        ("review",   "招标文件风险扫描（三层审标）", "bid review 招标文件.pdf -o report.md"),
        ("rfp",      "招标文件生成器",           "bid rfp --type services --project XX项目"),
        ("desense",  "AI味检测 + 敏感信息脱敏扫描", "bid desense 投标方案.md --mode bid"),
        ("commitments", "承诺链三源追踪审计", "bid commitments 标书.md --profile company_profile"),
        ("map-clauses", "条款-方案映射审计", "bid map-clauses 招标文件.md 方案.md"),
        ("watermark", "Word文档添加文字水印", "bid watermark input.docx output.docx -t 仅供参考"),
        ("materials","标书素材库管家（自动分类整理+材料清单）", "bid materials analyze 素材库目录"),
        ("table",    "工程标两列表格生成器", "bid table 技术方案.docx -o 输出.docx --left-color D9E2F3"),
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
    p.add_argument("--iron", action="store_true", help="同时运行引擎铁律校验（8项铁律）")
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
    p = sub.add_parser("desense", help="AI味检测 + 敏感信息脱敏扫描")
    p.add_argument("file", help="输入文件路径（.md/.txt/.docx）")
    p.add_argument("--mode", choices=["bid", "social", "xiaohongshu", "email", "paper"], default="bid",
                   help="检测模式: bid=投标方案(默认), social=自媒体, xiaohongshu=小红书, email=邮件, paper=学术论文")
    p.add_argument("--format", choices=["text", "json", "markdown", "docx"], default="text",
                   help="输出格式: text(默认), json, markdown, docx(Word修订标记)")
    p.add_argument("-o", "--output", help="输出文件路径（docx格式必填）")
    p.add_argument("--no-color", action="store_true", help="禁用终端颜色")
    p.add_argument("--no-examples", action="store_true", help="不显示改前→改后示范")

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

    # commitments
    p = sub.add_parser("commitments", help="承诺链三源追踪审计")
    p.add_argument("file", help="标书文件路径（.md/.txt）")
    p.add_argument("--profile", "-p", default="company_profile", help="企业资料库目录 (默认: company_profile/)")
    p.add_argument("--output", "-o", help="输出审计报告到文件")

    # map-clauses
    p = sub.add_parser("map-clauses", help="条款-方案映射审计")
    p.add_argument("tender_file", help="招标文件路径（.md/.txt）")
    p.add_argument("bid_file", nargs="?", default=None, help="方案文件路径（可选，不指定则只解析评分表）")
    p.add_argument("--report", "-r", help="导出报告路径")
    p.add_argument("--json", dest="json_output", help="导出 JSON 路径")

    # analyze
    p = sub.add_parser("analyze", help="投标可行性分析 — 一键拆解招标文件+评分分析+承诺审计+时间规划")
    p.add_argument("input", help="招标文件路径（.md）")
    p.add_argument("--profile", "-p", default="", help="企业资料库目录路径")
    p.add_argument("--output", "-o", default="", help="输出报告路径（.md）")
    p.add_argument("--config", "-c", default="", help="配置文件路径")
    p.add_argument("--list-only", action="store_true", help="仅拆解招标文件，不做深度分析")
    p.add_argument("--gen-config", action="store_true", help="生成默认配置文件")
    p.add_argument("--json", action="store_true", help="同时输出JSON格式")
    p.add_argument("--open-date", default="", help="开标日期（YYYY-MM-DD）")

    # table
    p = sub.add_parser("table", help="工程标两列表格生成器 — 从DOC/PDF提取文字自动生成两列表格")
    p.add_argument("input", help="输入文件路径（DOC/DOCX/PDF/MD）")
    p.add_argument("-o", "--output", default=None, help="输出Word文件路径")
    p.add_argument("--config", default=None, help="YAML配置文件路径")
    p.add_argument("--left-width", type=float, default=None, help="左列宽度（cm）")
    p.add_argument("--right-width", type=float, default=None, help="右列宽度（cm）")
    p.add_argument("--left-color", default=None, help="左列背景色（十六进制，如D9E2F3）")
    p.add_argument("--right-color", default=None, help="右列背景色（十六进制，如FFFFFF）")
    p.add_argument("--border-color", default=None, help="边框颜色（十六进制，如000000）")
    p.add_argument("--font-size", type=int, default=None, help="正文字号（pt）")
    p.add_argument("--font-name", default=None, help="正文字体名称")
    p.add_argument("--no-merge", action="store_true", help="不分组合并相同标签")
    p.add_argument("--list", action="store_true", help="仅列出识别的标签-内容对，不生成表格")
    p.add_argument("--gen-config", default=None, help="生成默认配置文件到指定路径")

    # render
    p = sub.add_parser("render", help="content.json 转 Word 排版（擎标排版引擎）")
    p.add_argument("content", help="content.json 路径")
    p.add_argument("output", help="输出 .docx 路径")
    p.add_argument("--config", "-c", default=None, help="排版配置 YAML 路径")

    # validate
    p = sub.add_parser("validate", help="校验 content.json 结构（不渲染，提前暴露错误）")
    p.add_argument("content", help="content.json 路径")

    # orchestrate
    p = sub.add_parser("orchestrate", help="编排引擎（框架锁定/差异检测/铁律校验）")
    orch_sub = p.add_subparsers(dest="sub", help="引擎子命令")
    orch_sub.add_parser("init", help="初始化项目").add_argument("--name", help="项目名称")
    p_parse = orch_sub.add_parser("parse", help="解析招标文件")
    p_parse.add_argument("tender_file", help="招标文件路径")
    p_fw = orch_sub.add_parser("framework", help="构建投标文件框架")
    p_fw.add_argument("--template", help="使用模板 (security/property/generic)")
    p_fw.add_argument("--from-tender", action="store_true", help="从招标文件提取结果构建")
    orch_sub.add_parser("lock", help="锁定框架")
    orch_sub.add_parser("unlock", help="解锁框架")
    p_diff = orch_sub.add_parser("diff", help="检查框架差异")
    p_diff.add_argument("content_file", help="内容文件路径 (md/docx)")
    p_check = orch_sub.add_parser("check", help="铁律校验")
    p_check.add_argument("content_file", help="内容文件路径 (md/txt)")
    p_check.add_argument("--docx", help="Word文档路径（用于排版格式检查）")
    orch_sub.add_parser("status", help="查看项目状态")
    orch_sub.add_parser("templates", help="列出可用框架模板")
    orch_sub.add_parser("prelock", help="前置锁定检查（三重锁定关卡）")
    p_anchor = orch_sub.add_parser("anchor", help="原文锚定比对（投标承诺 vs 招标文件原文）")
    p_anchor.add_argument("content_file", help="内容文件路径 (md/txt)")
    p_profile = orch_sub.add_parser("profile", help="投标人身份卡管理")
    p_profile.add_argument("--set-file", dest="set_file", help="从JSON文件加载企业身份卡")
    p_profile.add_argument("--set", help="使用内置预设 (generic)")
    p_profile.add_argument("--show", action="store_true", help="查看当前身份卡")

    # list
    sub.add_parser("list", help="列出所有可用工具")

    # rag
    p_rag = sub.add_parser("rag", help="历史标书 RAG 检索（默认零依赖 BM25，可选向量增强）")
    rag_sub = p_rag.add_subparsers(dest="rag_sub", help="RAG 子命令")
    rag_in = rag_sub.add_parser("ingest", help="入库：切块历史标书+评分项打标+建索引")
    rag_in.add_argument("path", help="历史标书路径（PDF/DOCX/MD）")
    rag_in.add_argument("--project", default="default", help="租户/项目名（多租户隔离）")
    rag_in.add_argument("--score-json", dest="score_json", help="评标办法评分项 JSON（评分项打标锚点）")
    rag_in.add_argument("--limit", type=int, default=0, help="仅处理前 N 页（0=全部，大文件调试）")
    rag_q = rag_sub.add_parser("query", help="检索：输入评分项→召回历史标杆章节")
    rag_q.add_argument("query", help="查询（评分项名称或自然语言）")
    rag_q.add_argument("--project", default="default", help="租户/项目名")
    rag_q.add_argument("--top-k", dest="top_k", type=int, default=5)
    rag_q.add_argument("--score-item", dest="score_item", help="按评分项标签过滤")
    rag_st = rag_sub.add_parser("status", help="查看 RAG 索引状态")
    rag_st.add_argument("--project", default="default")

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
    elif args.command == "analyze":
        cmd_analyze(args)
    elif args.command == "desense":
        cmd_desense(args)
    elif args.command == "watermark":
        cmd_watermark(args)
    elif args.command == "materials":
        cmd_materials(args)
    elif args.command == "commitments":
        cmd_commitments(args)
    elif args.command == "map-clauses":
        cmd_map_clauses(args)
    elif args.command == "render":
        cmd_render(args)
    elif args.command == "validate":
        cmd_validate(args)
    elif args.command == "orchestrate":
        cmd_orchestrate(args)
    elif args.command == "list":
        cmd_list()
    elif args.command == "gui":
        cmd_gui(args)
    elif args.command == "table":
        cmd_table(args)
    elif args.command == "rag":
        from bid_toolkit.rag import run
        run(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
