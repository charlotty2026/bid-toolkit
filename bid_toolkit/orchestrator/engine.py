#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标书编排引擎 (Bid Orchestration Engine) v1.0
=============================================

把铁律从"靠Agent自觉"变成"引擎强制校验"。

工作流：拆招标文件 -> 搭框架 -> 锁框架 -> 填内容 -> 铁律校验

用法:
  # 初始化项目
  python -m bid_toolkit.orchestrator init --name "南大物业管理标书"

  # 解析招标文件（PDF/DOCX -> JSON）
  python -m bid_toolkit.orchestrator parse 招标文件.pdf

  # 构建框架（从解析结果或模板）
  python -m bid_toolkit.orchestrator framework --template property
  python -m bid_toolkit.orchestrator framework --from-tender

  # 锁定框架
  python -m bid_toolkit.orchestrator lock

  # 检查框架差异（检查内容是否有未经授权的新增标题）
  python -m bid_toolkit.orchestrator diff 投标文件/正文.md

  # 设置投标人身份卡（前置锁定）
  python -m bid_toolkit.orchestrator profile --set-file /path/to/profile.json
  python -m bid_toolkit.orchestrator profile --show

  # 铁律校验
  python -m bid_toolkit.orchestrator check 投标文件/正文.md
  python -m bid_toolkit.orchestrator check 投标文件/正文.md --docx 输出.docx

  # 原文锚定比对（投标文件承诺 vs 招标文件原文）
  python -m bid_toolkit.orchestrator anchor 投标文件/正文.md

  # 查看项目状态
  python -m bid_toolkit.orchestrator status

  # 解锁框架（需确认）
  python -m bid_toolkit.orchestrator unlock --confirm
"""

import sys
import os
import json
import argparse
from pathlib import Path

# 确保能导入bid_toolkit
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from bid_toolkit.orchestrator.project import ProjectManager
from bid_toolkit.orchestrator.framework import (
    Framework, FrameworkLock,
    build_framework_from_template, build_framework_from_tender_extract,
    extract_titles_from_markdown, extract_titles_from_docx,
    STANDARD_FRAMEWORK_TEMPLATES
)
from bid_toolkit.orchestrator.iron_rules import IronRuleChecker
from bid_toolkit.orchestrator.bidder_profile import BidderProfile, PRESET_PROFILES
from bid_toolkit.orchestrator.source_anchoring import SourceAnchor

try:
    sys.stdout.reconfigure(encoding='utf-8')
except (AttributeError, Exception):
    pass


class BidOrchestrator:
    """标书编排引擎主控制器"""

    def __init__(self, project_root=None):
        self.pm = ProjectManager(project_root)
        self.root = self.pm.root

    # ===== init =====

    def cmd_init(self, args):
        """初始化项目"""
        meta = self.pm.load_meta()
        if args.name:
            meta["project_name"] = args.name
            self.pm.save_meta(meta)

        print(f"📁 项目已初始化: {self.root}")
        print(f"   项目名称: {meta.get('project_name', self.root.name)}")
        print(f"   状态目录: {self.pm.state_dir}")
        print(f"   目录结构:")
        print(f"     招标文件/  - 放招标文件")
        print(f"     投标文件/  - 投标文件工作区")
        print(f"     附件/      - 资质附件")
        print(f"     .bidproject/ - 引擎状态")
        print(f"\n下一步: 把招标文件放到 招标文件/ 目录，然后执行 parse")

    # ===== parse =====

    def cmd_parse(self, args):
        """解析招标文件"""
        tender_path = Path(args.tender_file)
        if not tender_path.exists():
            # 尝试在招标文件目录下找
            tender_path = self.root / "招标文件" / args.tender_file
            if not tender_path.exists():
                print(f"❌ 招标文件不存在: {args.tender_file}")
                sys.exit(1)

        print(f"📄 正在解析招标文件: {tender_path.name}")

        # 提取文本
        text = self._extract_document_text(tender_path)
        if not text:
            print(f"❌ 无法提取文本内容")
            sys.exit(1)

        print(f"   提取文本: {len(text)} 字符")

        # 解析招标文件结构
        extract = self._parse_tender_content(text, tender_path.name)

        # 保存提取结果
        self.pm.save_tender_extract(extract)
        # 保存原文（用于锚定比对）
        self.pm.save_tender_raw(text)

        # 更新项目元数据
        meta = self.pm.load_meta()
        meta["tender_file"] = str(tender_path)
        self.pm.save_meta(meta)
        self.pm.update_phase("parsed", f"解析招标文件: {tender_path.name}")

        # 输出摘要
        print(f"\n📊 解析结果:")
        print(f"   项目名称: {extract.get('project_name', '未知')}")
        print(f"   投标文件章节: {len(extract.get('bid_sections', []))} 个")
        print(f"   评分项: {len(extract.get('scoring_items', []))} 个")
        print(f"   资格要求: {len(extract.get('qualification_reqs', []))} 个")
        print(f"   废标条款: {len(extract.get('disqualify_rules', []))} 个")
        print(f"   要求列表: {len(extract.get('requirements', []))} 条")

        print(f"\n提取结果已保存: {self.pm.tender_extract_path}")
        print(f"下一步: 执行 framework --from-tender 构建投标文件框架")

    def _extract_document_text(self, path: Path) -> str:
        """提取PDF/DOCX文本"""
        suffix = path.suffix.lower()

        if suffix == ".pdf":
            try:
                import fitz
                doc = fitz.open(str(path))
                text = ""
                for page in doc:
                    text += page.get_text()
                return text
            except ImportError:
                print("❌ 需要PyMuPDF: pip install pymupdf")
                return ""
        elif suffix in (".docx", ".doc"):
            try:
                from docx import Document
                doc = Document(str(path))
                return "\n".join(para.text for para in doc.paragraphs)
            except ImportError:
                print("❌ 需要python-docx: pip install python-docx")
                return ""
        elif suffix in (".txt", ".md"):
            return path.read_text(encoding="utf-8")
        else:
            print(f"❌ 不支持的文件格式: {suffix}")
            return ""

    def _parse_tender_content(self, text: str, filename: str) -> dict:
        """解析招标文件内容，提取关键信息"""
        import re

        extract = {
            "project_name": "",
            "tender_file": filename,
            "bid_sections": [],
            "scoring_items": [],
            "qualification_reqs": [],
            "disqualify_rules": [],
            "requirements": [],
            "raw_text_length": len(text)
        }

        # 提取项目名称（通常在第一页）
        lines = text.split("\n")
        for line in lines[:20]:
            line = line.strip()
            if len(line) > 5 and ("项目" in line or "采购" in line or "招标" in line):
                if not line.startswith("第") and "编号" not in line:
                    extract["project_name"] = line[:100]
                    break

        # 提取投标文件章节（查找"投标文件组成"/"投标文件格式"等关键词）
        section_patterns = [
            r'第[一二三四五六七八九十]+部分[：:]\s*(.+)',
            r'(\d+)[\.、]\s*投标文件(?:的)?(?:组成|格式)',
            r'投标文件(?:应由|包括|包含|由以下)[：:]\s*(.+)',
        ]

        for pattern in section_patterns:
            for m in re.finditer(pattern, text):
                section_title = m.group(1).strip() if m.groups() else m.group(0)
                extract["bid_sections"].append({
                    "title": section_title[:80],
                    "level": 1,
                    "source": f"招标文件 - {filename}",
                    "required": True
                })

        # 如果没提取到章节，用通用模板
        if not extract["bid_sections"]:
            # 尝试从"投标文件格式"章节提取
            format_start = text.find("投标文件格式")
            if format_start > 0:
                format_text = text[format_start:format_start + 5000]
                for m in re.finditer(r'(投标函|开标一览表|投标报价|法定代表人|授权委托书|资格证明|服务方案|技术响应|商务响应|服务承诺|其他材料)', format_text):
                    title = m.group(1)
                    # 去重
                    if not any(s["title"] == title for s in extract["bid_sections"]):
                        extract["bid_sections"].append({
                            "title": title,
                            "level": 1,
                            "source": f"招标文件-投标文件格式",
                            "required": True
                        })

        # 提取评分项
        scoring_start = text.find("评分")
        if scoring_start > 0:
            scoring_text = text[scoring_start:scoring_start + 3000]
            for m in re.finditer(r'(\d+(?:\.\d+)?)\s*分', scoring_text):
                score = m.group(1)
                context = scoring_text[max(0, m.start() - 20):m.end() + 20]
                extract["scoring_items"].append({
                    "score": float(score),
                    "context": context.strip()[:100]
                })

        # 提取资格要求
        qual_keywords = ["资格要求", "投标人资格", "申请人资格", "投标人应具备"]
        for kw in qual_keywords:
            idx = text.find(kw)
            if idx > 0:
                qual_text = text[idx:idx + 2000]
                for m in re.finditer(r'(\d+)[\.、）)]\s*([^。\n]{5,50})', qual_text):
                    req = m.group(2).strip()
                    if any(w in req for w in ["营业", "许可", "资质", "注册", "资本", "信用", "财务", "业绩", "社保", "纳税"]):
                        extract["qualification_reqs"].append(req[:80])
                break

        # 提取废标条款
        disqualify_keywords = ["废标", "无效投标", "否决", "按照无效投标处理"]
        for kw in disqualify_keywords:
            idx = text.find(kw)
            if idx > 0:
                disq_text = text[max(0, idx - 200):idx + 2000]
                for m in re.finditer(r'(\d+)[\.、）)]\s*([^。\n]{5,80})', disq_text):
                    rule = m.group(2).strip()
                    if any(w in rule for w in ["未", "不", "缺少", "逾期", "虚假", "不符合", "超过", "低于"]):
                        extract["disqualify_rules"].append(rule[:100])
                break

        # 提取服务需求（作为requirements）
        req_keywords = ["服务需求", "项目需求", "技术需求", "服务要求"]
        for kw in req_keywords:
            idx = text.find(kw)
            if idx > 0:
                req_text = text[idx:idx + 5000]
                for m in re.finditer(r'(\d+)[\.、）)]\s*([^。\n]{10,100})', req_text):
                    req = m.group(2).strip()
                    if len(req) > 10:
                        extract["requirements"].append({
                            "text": req[:100],
                            "keywords": [req[:20]],
                            "required": True
                        })
                break

        return extract

    # ===== framework =====

    def cmd_framework(self, args):
        """构建投标文件框架"""
        if args.from_tender:
            tender_extract = self.pm.load_tender_extract()
            if not tender_extract:
                print("❌ 未找到招标文件提取结果，请先执行 parse")
                sys.exit(1)

            meta = self.pm.load_meta()
            framework = build_framework_from_tender_extract(
                tender_extract,
                project_name=meta.get("project_name", ""),
                tender_file=meta.get("tender_file", "")
            )
        elif args.template:
            if args.template not in STANDARD_FRAMEWORK_TEMPLATES:
                print(f"❌ 未知模板: {args.template}")
                print(f"   可用模板: {', '.join(STANDARD_FRAMEWORK_TEMPLATES.keys())}")
                sys.exit(1)

            meta = self.pm.load_meta()
            framework = build_framework_from_template(
                args.template,
                project_name=meta.get("project_name", ""),
                tender_file=meta.get("tender_file", "")
            )
        else:
            print("❌ 请指定 --template 或 --from-tender")
            sys.exit(1)

        # 保存框架
        framework_data = framework.to_dict()
        self.pm.save_framework(framework_data)
        self.pm.update_phase("frameworked", f"构建框架: {len(framework.sections)}个章节")

        # 输出框架大纲
        print(framework.to_markdown())
        print(f"\n📁 框架已保存: {self.pm.framework_path}")
        print(f"📊 共 {len(framework.sections)} 个章节")
        print(f"\n下一步: 检查框架是否正确，然后执行 lock 锁定框架")

        # 同时输出框架大纲到投标文件目录
        outline_path = self.root / "投标文件" / "框架.md"
        outline_path.write_text(framework.to_markdown(), encoding="utf-8")
        print(f"📝 框架大纲已输出: {outline_path}")

        # 导出 content.json 骨架（缝合 framework -> render，避免两张皮）
        try:
            skeleton = framework.to_content_json_skeleton()
            skeleton_path = self.root / "投标文件" / "content.json"
            skeleton_path.parent.mkdir(parents=True, exist_ok=True)
            skeleton_path.write_text(
                json.dumps(skeleton, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            print(f"📄 content.json 骨架已导出: {skeleton_path}")
            print(f"   继续: bid render {skeleton_path} 投标文件.docx")
        except Exception as e:  # 骨架导出是增强，失败不阻断框架主流程
            print(f"⚠️  content.json 骨架导出跳过: {e}")

    # ===== lock =====

    def cmd_lock(self, args):
        """锁定框架"""
        framework_data = self.pm.load_framework()
        if not framework_data:
            print("❌ 未找到框架，请先执行 framework")
            sys.exit(1)

        if self.pm.is_framework_locked():
            print("⚠️  框架已锁定！")
            print("   如需修改框架，请先执行 unlock --confirm")
            return

        framework = Framework.from_dict(framework_data)
        lock_data = FrameworkLock.lock(framework)

        # 更新框架数据（添加locked_at）
        self.pm.save_framework(framework.to_dict())
        self.pm.lock_framework(framework.to_dict())
        self.pm.update_phase("locked", f"锁定框架: {lock_data['section_count']}个章节")

        print(f"🔒 框架已锁定！")
        print(f"   锁定时间: {lock_data['locked_at']}")
        print(f"   章节数: {lock_data['section_count']}")
        print(f"   锁定哈希: {lock_data['lock_hash']}")
        print(f"\n⚠️  框架锁定后，任何增删改标题都需要人工确认！")
        print(f"下一步: 在框架内填充内容，然后执行 check 进行铁律校验")

    # ===== unlock =====

    def cmd_unlock(self, args):
        """解锁框架"""
        if not self.pm.is_framework_locked():
            print("框架未锁定，无需解锁")
            return

        if not args.confirm:
            print("⚠️  解锁框架将允许修改框架结构！")
            print("   确认解锁请执行: unlock --confirm")
            return

        self.pm.lock_path.unlink()
        self.pm.update_phase("frameworked", "框架已解锁")
        print("🔓 框架已解锁，可以修改框架结构")

    # ===== diff =====

    def cmd_diff(self, args):
        """检查框架差异"""
        framework_data = self.pm.load_framework()
        if not framework_data:
            print("❌ 未找到框架")
            sys.exit(1)

        if not self.pm.is_framework_locked():
            print("⚠️  框架未锁定，无法检测差异")
            print("   请先执行 lock 锁定框架")
            sys.exit(1)

        framework = Framework.from_dict(framework_data)

        # 从内容文件提取标题
        content_path = Path(args.content_file)
        if not content_path.exists():
            content_path = self.root / args.content_file
            if not content_path.exists():
                print(f"❌ 文件不存在: {args.content_file}")
                sys.exit(1)

        if content_path.suffix.lower() in (".docx", ".doc"):
            current_titles = extract_titles_from_docx(str(content_path))
        else:
            text = content_path.read_text(encoding="utf-8")
            current_titles = extract_titles_from_markdown(text)

        if not current_titles:
            print("⚠️  未从内容文件中提取到标题")
            return

        print(f"📋 锁定框架: {len(framework.sections)} 个标题")
        print(f"📋 当前内容: {len(current_titles)} 个标题")
        print()

        # 检测差异
        diff_result = FrameworkLock.diff(framework, current_titles)
        self.pm.save_diff_report(diff_result)

        print(diff_result["summary"])
        print()

        if diff_result["added"]:
            print("🔴 新增标题（未经授权）:")
            for a in diff_result["added"]:
                print(f"   位置{a['position']}: {a['title']}")
            print()

        if diff_result["removed"]:
            print("🟡 删除标题:")
            for r in diff_result["removed"]:
                print(f"   {r['title']}")
            print()

        if diff_result["renamed"]:
            print("🟡 重命名标题:")
            for r in diff_result["renamed"]:
                print(f"   {r['old_title']} -> {r['new_title']}")
            print()

        if diff_result["clean"]:
            print("✅ 框架一致，内容在框架范围内！")
        else:
            print("⚠️  发现框架差异！")
            print("   新增标题需人工确认是否保留")
            print("   确认后如需更新框架，执行 unlock --confirm -> 修改 -> lock")

    # ===== check =====

    def cmd_check(self, args):
        """铁律校验（9项·分层）"""
        content_path = Path(args.content_file)
        if not content_path.exists():
            content_path = self.root / args.content_file
            if not content_path.exists():
                print(f"❌ 文件不存在: {args.content_file}")
                sys.exit(1)

        # 前置锁定守卫
        self._warn_if_not_prelocked("执行 check")

        content = content_path.read_text(encoding="utf-8")

        # 加载招标文件提取结果、锁定框架、身份卡
        tender_extract = self.pm.load_tender_extract() or {}
        framework_data = self.pm.load_framework()
        bidder_profile = self.pm.load_bidder_profile()

        locked_titles = []
        if framework_data:
            framework = Framework.from_dict(framework_data)
            locked_titles = framework.get_titles()

        # 运行铁律校验
        checker = IronRuleChecker(
            tender_extract=tender_extract,
            locked_titles=locked_titles,
            bidder_profile=bidder_profile
        )

        report = checker.check_all(content, docx_path=args.docx)
        self.pm.save_iron_rule_report(report)
        self.pm.update_phase("checked", f"铁律校验: {'通过' if report['all_passed'] else '未通过'}")

        # 输出报告
        print(IronRuleChecker.format_report(report))

        # 只有废标级未通过才拦截
        if not report.get("fatal_passed", False):
            print("\n" + "=" * 60)
            print("⛔ 废标级铁律未通过，内容被打回！禁止提交！")
            print("   请修复上述 🔴 问题后重新执行 check")
            sys.exit(1)
        elif not report["all_passed"]:
            print("\n" + "=" * 60)
            print("⚠️ 质量级铁律有问题，建议修复后提交（废标级已通过）")

    # ===== profile =====

    def cmd_profile(self, args):
        """投标人身份卡管理"""
        if args.set_file:
            # 从JSON文件加载身份卡（企业专属，不入库）
            file_path = Path(args.set_file)
            if not file_path.exists():
                print(f"❌ 文件不存在: {args.set_file}")
                sys.exit(1)

            with open(file_path, "r", encoding="utf-8") as f:
                profile_data = json.load(f)

            self.pm.save_bidder_profile(profile_data)

            bp = BidderProfile(profile_data)
            print(f"✅ 投标人身份卡已设置: {bp.data.get('bidder_name', '?')}")
            print(f"   角色: {bp.data.get('role', '?')} / 招标方: {bp.data.get('client_role', '?')}")
            print()
            print("   身份锚定铁律(⑨)已激活，check 时将自动校验")

        elif args.set:
            # 设置内置预设身份卡
            preset_name = args.set.lower()
            if preset_name not in PRESET_PROFILES:
                print(f"❌ 未知预设: {args.set}")
                print(f"   可用预设: {', '.join(PRESET_PROFILES.keys())}")
                sys.exit(1)

            profile_data = PRESET_PROFILES[preset_name]
            self.pm.save_bidder_profile(profile_data)

            bp = BidderProfile(profile_data)
            print(f"✅ 投标人身份卡已设置: {bp.data.get('bidder_name', '?')}")
            print(f"   角色: {bp.data.get('role', '?')} / 招标方: {bp.data.get('client_role', '?')}")
            print()
            print("   身份锚定铁律(⑨)已激活，check 时将自动校验")

        elif args.show:
            # 查看当前身份卡
            profile_data = self.pm.load_bidder_profile()
            if not profile_data:
                print("⚠️ 未设置投标人身份卡")
                print("   使用 profile --set-file <json路径> 加载企业身份卡")
                return

            bp = BidderProfile(profile_data)
            print(bp.summary())

    # ===== anchor =====

    def cmd_anchor(self, args):
        """原文锚定比对 - 投标文件承诺 vs 招标文件原文"""
        bid_file = Path(args.bid_file)

        if not bid_file.is_absolute():
            bid_file = self.root / bid_file

        if not bid_file.exists():
            print(f"❌ 投标文件不存在: {bid_file}")
            sys.exit(1)

        # 前置锁定守卫
        self._warn_if_not_prelocked("执行 anchor")

        # 读取投标文件
        content = bid_file.read_text(encoding="utf-8")
        if not content.strip():
            print("❌ 投标文件内容为空")
            sys.exit(1)

        # 加载招标文件原文
        tender_raw = self.pm.load_tender_raw()
        if not tender_raw:
            print("⚠️  未找到招标文件原文，请先执行 parse 解析招标文件")
            print("   （parse 会自动保存原文用于锚定比对）")
            sys.exit(1)

        print(f"📄 投标文件: {bid_file.name} ({len(content)} 字符)")
        print(f"📄 招标文件原文: {len(tender_raw)} 字符")
        print()

        # 执行锚定比对
        anchor = SourceAnchor(tender_raw)
        report = anchor.anchor_all(content)

        # 保存报告
        self.pm.save_anchor_report(report)

        # 输出报告
        print(SourceAnchor.format_report(report))

        print(f"\n报告已保存: {self.pm.anchor_report_path}")

    # ===== prelock =====

    def cmd_prelock(self, args):
        """前置锁定 - 三重锁定关卡，全过才允许Agent写内容

        三重锁定：
        1. 招标文件已解析 + 原文锚定基线已保存
        2. 框架已构建 + 已锁定（哈希校验通过）
        3. 投标人身份卡已设置

        全部通过后，引擎进入 prelocked 阶段，Agent可以在框架内填写内容。
        """
        checks = []

        # --- 锁定1：招标文件解析 + 原文基线 ---
        tender_raw = self.pm.load_tender_raw()
        tender_extract = self.pm.load_tender_extract()
        if tender_raw and tender_extract:
            checks.append((
                "招标文件解析 + 原文锚定基线",
                True,
                f"原文 {len(tender_raw)} 字符，提取 {len(tender_extract.get('bid_sections', []))} 章节"
            ))
        elif tender_extract and not tender_raw:
            checks.append(("招标文件解析 + 原文锚定基线", False, "招标文件已解析但原文未保存，请重新 parse"))
        else:
            checks.append(("招标文件解析 + 原文锚定基线", False, "请先执行 parse <招标文件>"))

        # --- 锁定2：框架构建 + 锁定 ---
        framework = self.pm.load_framework()
        locked = self.pm.is_framework_locked()
        if not framework:
            checks.append(("框架构建 + 锁定", False, "请先执行 framework --from-tender 构建框架"))
        elif not locked:
            checks.append(("框架构建 + 锁定", False, "框架已构建但未锁定，请执行 lock"))
        else:
            integrity = self.pm.verify_framework_integrity(framework)
            if integrity.get("tampered"):
                checks.append(("框架构建 + 锁定", False, "⚠️ 框架被篡改！哈希校验失败，请重新 lock 或检查框架"))
            else:
                section_count = len(framework.get("sections", []))
                checks.append(("框架构建 + 锁定", True, f"{section_count} 个章节已锁定，哈希校验通过"))

        # --- 锁定3：投标人身份卡 ---
        profile = self.pm.load_bidder_profile()
        if profile:
            bidder_name = profile.get("bidder_name") or profile.get("name") or "未知"
            checks.append(("投标人身份卡", True, f"投标人: {bidder_name}"))
        else:
            checks.append(("投标人身份卡", False, "请先执行 profile --set-file <json路径>"))

        # --- 判定 ---
        all_pass = all(c[1] for c in checks)

        print("🔒 前置锁定检查")
        print("=" * 50)
        for name, passed, detail in checks:
            status = "✅" if passed else "❌"
            print(f"  {status} {name}")
            print(f"     {detail}")

        if all_pass:
            assert tender_raw is not None
            assert framework is not None
            assert profile is not None
            self.pm.update_phase("prelocked", "三重锁定完成，允许Agent写内容")
            print()
            print("=" * 50)
            print("✅ 三重锁定完成！Agent可以在框架内填写内容。")
            print()
            print("📋 锁定基线:")
            print(f"   ① 原文锚定基线: {len(tender_raw)} 字符（承诺反查依据）")
            print(f"   ② 框架锁定: {len(framework.get('sections', []))} 个章节（不可增删）")
            print(f"   ③ 身份卡: {profile.get('bidder_name', '?')}（身份铁律⑨）")
            print()
            print("下一步:")
            print("  1. 在 投标文件/ 目录下按框架填写内容")
            print("  2. 执行 diff <内容文件> 检查框架差异")
            print("  3. 执行 check <内容文件> 铁律校验")
            print("  4. 执行 anchor <内容文件> 原文锚定比对")
        else:
            failed = [c[0] for c in checks if not c[1]]
            print()
            print(f"❌ 前置锁定未通过（{len(failed)}/{len(checks)} 项未就绪）")
            print("   引擎不允许在未锁定状态下开始写内容。")
            print("   请先完成上述 ❌ 项。")

    def _is_prelocked(self):
        """检查是否已通过前置锁定（供其他命令调用做守卫）"""
        checks = [
            self.pm.load_tender_raw() is not None,
            self.pm.load_framework() is not None,
            self.pm.is_framework_locked(),
            self.pm.load_bidder_profile() is not None,
        ]
        return all(checks)

    def _warn_if_not_prelocked(self, command_name):
        """如果未通过前置锁定，打印警告（不阻断，仅提醒）"""
        if not self._is_prelocked():
            print(f"⚠️  注意：尚未通过前置锁定（prelock）。")
            print(f"   建议先执行 prelock 完成三重锁定后再{command_name}。")
            print(f"   未锁定状态下写内容，后续校验可能大量报错。")
            print()

    # ===== status =====

    def cmd_status(self, args):
        """查看项目状态"""
        status = self.pm.status()

        print(f"📊 项目状态")
        print(f"=" * 40)
        print(f"项目名称: {status['project_name']}")
        print(f"当前阶段: {status['phase']}")
        print(f"创建时间: {status['created'][:19]}")
        print(f"更新时间: {status['updated'][:19]}")
        print()
        print(f"招标文件: {status.get('tender_file', '未指定')}")
        print(f"招标文件已解析: {'✅' if status['tender_parsed'] else '❌'}")
        print(f"框架已构建: {'✅' if status['framework_built'] else '❌'}")
        print(f"框架已锁定: {'✅' if status['framework_locked'] else '❌'}")
        print(f"框架章节数: {status['framework_sections']}")
        print(f"身份卡已设置: {'✅' if status.get('bidder_profile_set') else '❌ (建议 profile --set-file <json>)'}")
        print(f"铁律已校验: {'✅' if status['iron_rules_checked'] else '❌'}")
        print(f"铁律全部通过: {'✅' if status['iron_rules_passed'] else '❌'}")
        if status.get('iron_rules_checked') and not status['iron_rules_passed']:
            print(f"废标级通过: {'✅' if status.get('fatal_passed') else '⛔'}")

        if status.get("anchor_report"):
            rate = status.get("anchor_rate")
            rate_str = f"{rate:.0%}" if rate is not None else "?"
            print(f"锚定比对已执行: ✅ (锚定率 {rate_str})")
        else:
            print(f"锚定比对: ❌ (建议 anchor <投标文件>)")

        # 前置锁定状态
        prelocked = self._is_prelocked()
        if prelocked:
            print(f"前置锁定: ✅ (三重锁定已通过)")
        elif status["phase"] in ("locked", "filling", "checked"):
            print(f"前置锁定: ⚠️  (建议执行 prelock 完成三重检查)")

        if status.get("framework_tampered"):
            print(f"⚠️  框架已被篡改！")

        # 输出下一步建议
        print()
        print("📋 下一步:")
        if status["phase"] == "init":
            print("  1. 把招标文件放到 招标文件/ 目录")
            print("  2. 执行 parse <招标文件路径>")
        elif status["phase"] == "parsed":
            print("  1. 执行 framework --from-tender 构建框架")
            print("  2. 或 framework --template <模板名> 用模板构建")
        elif status["phase"] == "frameworked":
            print("  1. 检查框架是否正确")
            print("  2. 执行 lock 锁定框架")
        elif status["phase"] == "locked":
            print("  1. 执行 profile --set-file <json> 设置身份卡（如未设置）")
            print("  2. 执行 prelock 三重锁定检查")
            print("  3. 锁定通过后，在框架内填充内容")
        elif status["phase"] == "prelocked":
            print("  1. 在 投标文件/ 目录下按框架填写内容")
            print("  2. 执行 diff <内容文件> 检查框架差异")
            print("  3. 执行 check <内容文件> 铁律校验")
            print("  4. 执行 anchor <内容文件> 原文锚定比对")
        elif status["phase"] == "checked":
            print("  1. 执行 anchor <内容文件> 原文锚定比对")
            print("  2. 检查无原文依据的承诺，删除或补充依据")
            if status["iron_rules_passed"]:
                print("  ✅ 铁律全部通过，可以生成Word文档并提交！")
            else:
                print("  1. 修复铁律校验报告中的问题")
                print("  2. 重新执行 check")

    # ===== templates =====

    def cmd_templates(self, args):
        """列出可用模板"""
        print("📋 可用框架模板:")
        print()
        for name, template in STANDARD_FRAMEWORK_TEMPLATES.items():
            print(f"  {name}: {template['name']} ({len(template['sections'])}个章节)")


def main():
    parser = argparse.ArgumentParser(
        description="标书编排引擎 - 拆招标文件->搭框架->锁定->填内容->铁律校验",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )

    subparsers = parser.add_subparsers(dest="command", help="子命令")

    # init
    p_init = subparsers.add_parser("init", help="初始化项目")
    p_init.add_argument("--name", help="项目名称")

    # parse
    p_parse = subparsers.add_parser("parse", help="解析招标文件")
    p_parse.add_argument("tender_file", help="招标文件路径")

    # framework
    p_fw = subparsers.add_parser("framework", help="构建投标文件框架")
    p_fw.add_argument("--template", help="使用模板 (security/property/generic)")
    p_fw.add_argument("--from-tender", action="store_true", help="从招标文件提取结果构建")

    # lock
    subparsers.add_parser("lock", help="锁定框架")

    # unlock
    p_unlock = subparsers.add_parser("unlock", help="解锁框架")
    p_unlock.add_argument("--confirm", action="store_true", help="确认解锁")

    # diff
    p_diff = subparsers.add_parser("diff", help="检查框架差异")
    p_diff.add_argument("content_file", help="内容文件路径 (md/docx)")

    # check
    p_check = subparsers.add_parser("check", help="铁律校验（9项·分层）")
    p_check.add_argument("content_file", help="内容文件路径 (md/txt)")
    p_check.add_argument("--docx", help="Word文档路径（用于排版格式检查）")

    # anchor
    p_anchor = subparsers.add_parser("anchor", help="原文锚定比对（投标承诺 vs 招标文件原文）")
    p_anchor.add_argument("bid_file", help="投标文件路径 (md/txt)")

    # prelock
    subparsers.add_parser("prelock", help="前置锁定检查（三重锁定关卡）")

    # profile
    p_profile = subparsers.add_parser("profile", help="投标人身份卡管理")
    p_profile.add_argument("--set-file", dest="set_file", help="从JSON文件加载企业身份卡")
    p_profile.add_argument("--set", help="使用内置预设 (generic)")
    p_profile.add_argument("--show", action="store_true", help="查看当前身份卡")

    # status
    subparsers.add_parser("status", help="查看项目状态")

    # templates
    subparsers.add_parser("templates", help="列出可用框架模板")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(0)

    orchestrator = BidOrchestrator()

    commands = {
        "init": orchestrator.cmd_init,
        "parse": orchestrator.cmd_parse,
        "framework": orchestrator.cmd_framework,
        "lock": orchestrator.cmd_lock,
        "unlock": orchestrator.cmd_unlock,
        "diff": orchestrator.cmd_diff,
        "check": orchestrator.cmd_check,
        "anchor": orchestrator.cmd_anchor,
        "prelock": orchestrator.cmd_prelock,
        "profile": orchestrator.cmd_profile,
        "status": orchestrator.cmd_status,
        "templates": orchestrator.cmd_templates,
    }

    if args.command in commands:
        commands[args.command](args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
