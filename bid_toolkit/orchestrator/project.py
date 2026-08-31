# -*- coding: utf-8 -*-
"""
项目管理器 - 管理标书项目的状态和文件
=====================================

项目目录结构：
  项目根目录/
  ├── .bidproject/              # 引擎状态目录
  │   ├── project.json          # 项目元数据
  │   ├── framework.json        # 锁定的框架
  │   ├── framework.lock        # 框架锁定标记 + 哈希
  │   ├── tender_extract.json   # 招标文件提取结果
  │   ├── iron_rule_report.json # 铁律校验报告
  │   └── diff_report.json      # 框架差异报告
  ├── 招标文件/                  # 原始招标文件
  ├── 投标文件/                  # 投标文件工作区
  │   ├── 框架.md               # 框架大纲（锁定后生成）
  │   ├── 正文.md               # 正文内容
  │   └── 输出.docx             # 最终Word输出
  └── 附件/                      # 资质附件等
"""

import json
import os
import hashlib
from datetime import datetime
from pathlib import Path


class ProjectManager:
    """管理标书项目的生命周期"""

    BIDPROJECT_DIR = ".bidproject"

    def __init__(self, project_root=None):
        self.root = Path(project_root or os.getcwd()).resolve()
        self.state_dir = self.root / self.BIDPROJECT_DIR
        self.state_dir.mkdir(parents=True, exist_ok=True)

        # 确保子目录存在
        (self.root / "招标文件").mkdir(exist_ok=True)
        (self.root / "投标文件").mkdir(exist_ok=True)
        (self.root / "附件").mkdir(exist_ok=True)

    # ===== 项目元数据 =====

    @property
    def meta_path(self):
        return self.state_dir / "project.json"

    def load_meta(self):
        """加载项目元数据"""
        if self.meta_path.exists():
            with open(self.meta_path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {
            "created": datetime.now().isoformat(),
            "project_name": self.root.name,
            "phase": "init",  # init -> parsed -> frameworked -> locked -> prelocked -> filling -> checked -> done
            "tender_file": None,
            "bid_packages": [],
            "history": []
        }

    def save_meta(self, meta):
        """保存项目元数据"""
        meta["updated"] = datetime.now().isoformat()
        with open(self.meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)

    def update_phase(self, phase, note=""):
        """更新项目阶段"""
        meta = self.load_meta()
        old_phase = meta.get("phase", "init")
        meta["phase"] = phase
        meta["history"].append({
            "time": datetime.now().isoformat(),
            "from": old_phase,
            "to": phase,
            "note": note
        })
        self.save_meta(meta)
        return meta

    # ===== 招标文件提取结果 =====

    @property
    def tender_extract_path(self):
        return self.state_dir / "tender_extract.json"

    def save_tender_extract(self, extract):
        """保存招标文件提取结果"""
        with open(self.tender_extract_path, "w", encoding="utf-8") as f:
            json.dump(extract, f, ensure_ascii=False, indent=2)

    def load_tender_extract(self):
        """加载招标文件提取结果"""
        if self.tender_extract_path.exists():
            with open(self.tender_extract_path, "r", encoding="utf-8") as f:
                return json.load(f)
        return None

    # ===== 招标文件原文 =====

    @property
    def tender_raw_path(self):
        return self.state_dir / "tender_raw.txt"

    def save_tender_raw(self, text: str):
        """保存招标文件原文（用于锚定比对）"""
        with open(self.tender_raw_path, "w", encoding="utf-8") as f:
            f.write(text)

    def load_tender_raw(self):
        """加载招标文件原文"""
        if self.tender_raw_path.exists():
            return self.tender_raw_path.read_text(encoding="utf-8")
        return None

    # ===== 框架文件 =====

    @property
    def framework_path(self):
        return self.state_dir / "framework.json"

    @property
    def lock_path(self):
        return self.state_dir / "framework.lock"

    def save_framework(self, framework_data):
        """保存框架数据"""
        with open(self.framework_path, "w", encoding="utf-8") as f:
            json.dump(framework_data, f, ensure_ascii=False, indent=2)

    def load_framework(self):
        """加载框架数据"""
        if self.framework_path.exists():
            with open(self.framework_path, "r", encoding="utf-8") as f:
                return json.load(f)
        return None

    def is_framework_locked(self):
        """检查框架是否已锁定"""
        return self.lock_path.exists()

    def lock_framework(self, framework_data):
        """锁定框架 - 写入锁定标记和哈希"""
        framework_json = json.dumps(framework_data, ensure_ascii=False, sort_keys=True)
        framework_hash = hashlib.sha256(framework_json.encode("utf-8")).hexdigest()

        lock_data = {
            "locked_at": datetime.now().isoformat(),
            "framework_hash": framework_hash,
            "section_count": len(framework_data.get("sections", [])),
            "locked_by": "engine"
        }
        with open(self.lock_path, "w", encoding="utf-8") as f:
            json.dump(lock_data, f, ensure_ascii=False, indent=2)

        return lock_data

    def verify_framework_integrity(self, framework_data):
        """验证框架完整性 - 检查是否被篡改"""
        if not self.lock_path.exists():
            return {"locked": False, "tampered": False}

        with open(self.lock_path, "r", encoding="utf-8") as f:
            lock_data = json.load(f)

        framework_json = json.dumps(framework_data, ensure_ascii=False, sort_keys=True)
        current_hash = hashlib.sha256(framework_json.encode("utf-8")).hexdigest()

        return {
            "locked": True,
            "tampered": current_hash != lock_data["framework_hash"],
            "locked_at": lock_data["locked_at"],
            "locked_section_count": lock_data["section_count"],
            "current_section_count": len(framework_data.get("sections", []))
        }

    # ===== 铁律校验报告 =====

    @property
    def iron_rule_report_path(self):
        return self.state_dir / "iron_rule_report.json"

    def save_iron_rule_report(self, report):
        """保存铁律校验报告"""
        with open(self.iron_rule_report_path, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)

    def load_iron_rule_report(self):
        """加载最近一次铁律校验报告"""
        if self.iron_rule_report_path.exists():
            with open(self.iron_rule_report_path, "r", encoding="utf-8") as f:
                return json.load(f)
        return None

    # ===== 投标人身份卡 =====

    @property
    def bidder_profile_path(self):
        return self.state_dir / "bidder_profile.json"

    def save_bidder_profile(self, profile_data):
        """保存投标人身份卡"""
        with open(self.bidder_profile_path, "w", encoding="utf-8") as f:
            json.dump(profile_data, f, ensure_ascii=False, indent=2)

    def load_bidder_profile(self):
        """加载投标人身份卡"""
        if self.bidder_profile_path.exists():
            with open(self.bidder_profile_path, "r", encoding="utf-8") as f:
                return json.load(f)
        return None

    # ===== 差异报告 =====

    @property
    def diff_report_path(self):
        return self.state_dir / "diff_report.json"

    def save_diff_report(self, report):
        """保存框架差异报告"""
        with open(self.diff_report_path, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)

    def load_diff_report(self):
        """加载最近一次差异报告"""
        if self.diff_report_path.exists():
            with open(self.diff_report_path, "r", encoding="utf-8") as f:
                return json.load(f)
        return None

    # ===== 锚定比对报告 =====

    @property
    def anchor_report_path(self):
        return self.state_dir / "anchor_report.json"

    def save_anchor_report(self, report):
        """保存原文锚定比对报告"""
        from datetime import datetime
        report["timestamp"] = datetime.now().isoformat()
        with open(self.anchor_report_path, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)

    def load_anchor_report(self):
        """加载最近一次锚定比对报告"""
        if self.anchor_report_path.exists():
            with open(self.anchor_report_path, "r", encoding="utf-8") as f:
                return json.load(f)
        return None

    # ===== 项目状态摘要 =====

    def status(self):
        """返回项目状态摘要"""
        meta = self.load_meta()
        tender = self.load_tender_extract()
        framework = self.load_framework()
        locked = self.is_framework_locked()
        iron_report = self.load_iron_rule_report()
        diff_report = self.load_diff_report()

        summary = {
            "project_name": meta.get("project_name", self.root.name),
            "phase": meta.get("phase", "init"),
            "created": meta.get("created", ""),
            "updated": meta.get("updated", ""),
            "tender_file": meta.get("tender_file"),
            "tender_parsed": tender is not None,
            "framework_built": framework is not None,
            "framework_locked": locked,
            "framework_sections": len(framework.get("sections", [])) if framework else 0,
            "bidder_profile_set": self.load_bidder_profile() is not None,
            "iron_rules_checked": iron_report is not None,
            "iron_rules_passed": iron_report.get("all_passed", False) if iron_report else False,
            "fatal_passed": iron_report.get("fatal_passed", False) if iron_report else False,
            "last_diff": diff_report.get("timestamp") if diff_report else None,
            "anchor_report": self.load_anchor_report() is not None,
            "anchor_rate": (self.load_anchor_report() or {}).get("anchor_rate") if self.load_anchor_report() else None,
            "history_count": len(meta.get("history", []))
        }

        if framework and locked:
            integrity = self.verify_framework_integrity(framework)
            summary["framework_tampered"] = integrity.get("tampered", False)

        return summary
