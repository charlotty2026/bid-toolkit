# -*- coding: utf-8 -*-
"""
框架管理器 - 框架搭建、锁定、差异检测
=====================================

核心设计：
  1. 框架 = 有序的章节列表，每个章节有稳定ID
  2. 框架一旦锁定，任何增删改都需要人工确认
  3. 差异检测：对比当前文档结构 vs 锁定框架
     - 新增标题 -> 🔴 红色警报（最严重）
     - 删除标题 -> 🟡 黄色警告
     - 重命名标题 -> 🟡 黄色警告
     - 顺序变化 -> 🟡 黄色警告

框架数据结构：
  {
    "project_name": "xxx",
    "tender_file": "xxx.pdf",
    "locked_at": "2026-08-12T...",
    "sections": [
      {
        "id": "s1",
        "title": "投标函",
        "level": 1,
        "source": "招标文件第五章投标文件格式",
        "required": true,
        "status": "empty"  # empty -> filling -> filled -> checked
      },
      ...
    ]
  }
"""

import json
import hashlib
import re
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional, Tuple


class Section:
    """框架章节"""

    def __init__(self, sid: str, title: str, level: int = 1,
                 source: str = "", required: bool = True, status: str = "empty"):
        self.id = sid
        self.title = title
        self.level = level
        self.source = source  # 来自招标文件的哪个部分
        self.required = required  # 是否招标文件要求的
        self.status = status  # empty / filling / filled / checked

    def to_dict(self):
        return {
            "id": self.id,
            "title": self.title,
            "level": self.level,
            "source": self.source,
            "required": self.required,
            "status": self.status
        }

    @classmethod
    def from_dict(cls, d):
        return cls(
            sid=d["id"],
            title=d["title"],
            level=d.get("level", 1),
            source=d.get("source", ""),
            required=d.get("required", True),
            status=d.get("status", "empty")
        )

    def title_match(self, other_title: str, threshold: float = 0.8) -> bool:
        """模糊匹配标题（允许微小差异）"""
        if self.title == other_title:
            return True
        # 简单的相似度：共同字符比例
        if not self.title or not other_title:
            return False
        s1 = set(self.title)
        s2 = set(other_title)
        overlap = len(s1 & s2)
        return overlap / max(len(s1), len(s2)) >= threshold


class Framework:
    """投标文件框架"""

    def __init__(self, project_name: str = "", tender_file: str = ""):
        self.project_name = project_name
        self.tender_file = tender_file
        self.sections: List[Section] = []
        self.created_at = datetime.now().isoformat()
        self.locked_at: Optional[str] = None

    def add_section(self, title: str, level: int = 1, source: str = "",
                    required: bool = True) -> Section:
        """添加章节"""
        sid = f"s{len(self.sections) + 1}"
        section = Section(sid, title, level, source, required)
        self.sections.append(section)
        return section

    def to_dict(self) -> dict:
        return {
            "project_name": self.project_name,
            "tender_file": self.tender_file,
            "created_at": self.created_at,
            "locked_at": self.locked_at,
            "sections": [s.to_dict() for s in self.sections]
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Framework":
        fw = cls(
            project_name=d.get("project_name", ""),
            tender_file=d.get("tender_file", "")
        )
        fw.created_at = d.get("created_at", "")
        fw.locked_at = d.get("locked_at")
        fw.sections = [Section.from_dict(s) for s in d.get("sections", [])]
        return fw

    def to_markdown(self) -> str:
        """输出框架大纲为Markdown"""
        lines = [f"# {self.project_name} - 投标文件框架", ""]
        if self.locked_at:
            lines.append(f"> 框架已锁定 | 锁定时间: {self.locked_at} | 章节数: {len(self.sections)}")
        else:
            lines.append(f"> 框架未锁定 | 章节数: {len(self.sections)}")
        lines.append("")

        for s in self.sections:
            indent = "  " * (s.level - 1)
            marker = "🔒" if s.status == "checked" else ("📝" if s.status == "filled" else "⬜")
            req = "" if s.required else " (非必须)"
            lines.append(f"{indent}- {marker} {s.title}{req}")
            if s.source:
                lines.append(f"{indent}  _来源: {s.source}_")

        return "\n".join(lines)

    def to_content_json_skeleton(self) -> dict:
        """导出 content.json 骨架，供 render 引擎直接消费（缝合 framework -> render）。"""
        body = []
        for s in self.sections:
            htype = f"h{min(int(s.level), 5)}"
            body.append({"type": htype, "text": s.title})
            body.append({"type": "p", "text": f"请在此处填写「{s.title}」内容……"})
        return {
            "meta": {
                "project_name": self.project_name or "投标文件",
                "bidder": "",
                "tender_name": "",
                "date": "",
            },
            "body": body,
        }

    def get_titles(self) -> List[str]:
        """获取所有标题列表"""
        return [s.title for s in self.sections]

    def find_section_by_title(self, title: str) -> Optional[Section]:
        """按标题查找章节"""
        for s in self.sections:
            if s.title == title or s.title_match(title):
                return s
        return None


class FrameworkLock:
    """框架锁定与差异检测"""

    @staticmethod
    def compute_hash(framework: Framework) -> str:
        """计算框架哈希（基于章节标题和顺序）"""
        titles = [f"{s.id}:{s.title}:{s.level}" for s in framework.sections]
        content = "|".join(titles)
        return hashlib.sha256(content.encode("utf-8")).hexdigest()[:16]

    @staticmethod
    def lock(framework: Framework) -> dict:
        """锁定框架"""
        framework.locked_at = datetime.now().isoformat()
        lock_hash = FrameworkLock.compute_hash(framework)
        return {
            "locked_at": framework.locked_at,
            "lock_hash": lock_hash,
            "section_count": len(framework.sections),
            "sections_snapshot": [
                {"id": s.id, "title": s.title, "level": s.level}
                for s in framework.sections
            ]
        }

    @staticmethod
    def diff(locked_framework: Framework, current_titles: List[Dict]) -> dict:
        """
        检测框架差异

        Args:
            locked_framework: 锁定的框架
            current_titles: 当前文档的标题列表
                [{"title": "投标函", "level": 1}, ...]

        Returns:
            {
                "added": [...],      # 新增的标题（最严重）
                "removed": [...],    # 删除的标题
                "renamed": [...],    # 重命名的标题
                "reordered": bool,   # 顺序是否变化
                "clean": bool,       # 是否无差异
                "summary": str       # 人类可读摘要
            }
        """
        locked_sections = locked_framework.sections
        locked_titles = [(s.title, s.level) for s in locked_sections]
        current_pairs = [(t.get("title", ""), t.get("level", 1)) for t in current_titles]

        added = []
        removed = []
        renamed = []

        # 检测新增：当前文档中有但锁定框架中没有的
        for i, (title, level) in enumerate(current_pairs):
            found = False
            for lt, ll in locked_titles:
                if title == lt:
                    found = True
                    break
            if not found:
                # 模糊匹配：可能是重命名
                match_found = False
                for s in locked_sections:
                    if s.title_match(title, threshold=0.6) and s.title != title:
                        renamed.append({
                            "section_id": s.id,
                            "old_title": s.title,
                            "new_title": title,
                            "position": i
                        })
                        match_found = True
                        break
                if not match_found:
                    added.append({
                        "title": title,
                        "level": level,
                        "position": i
                    })

        # 检测删除：锁定框架中有但当前文档中没有的
        for s in locked_sections:
            found = False
            for title, level in current_pairs:
                if s.title == title or s.title_match(title, threshold=0.6):
                    found = True
                    break
            if not found:
                removed.append({
                    "section_id": s.id,
                    "title": s.title,
                    "level": s.level
                })

        # 检测顺序变化
        reordered = False
        if len(current_pairs) == len(locked_titles):
            for i, (title, _) in enumerate(current_pairs):
                if i < len(locked_titles) and title != locked_titles[i][0]:
                    reordered = True
                    break

        clean = len(added) == 0 and len(removed) == 0 and len(renamed) == 0 and not reordered

        # 生成人类可读摘要
        parts = []
        if added:
            parts.append(f"🔴 新增{len(added)}个标题: {', '.join(a['title'] for a in added)}")
        if removed:
            parts.append(f"🟡 删除{len(removed)}个标题: {', '.join(r['title'] for r in removed)}")
        if renamed:
            parts.append(f"🟡 重命名{len(renamed)}个标题")
        if reordered:
            parts.append("🟡 章节顺序变化")
        if clean:
            parts.append("✅ 框架一致，无变化")

        return {
            "added": added,
            "removed": removed,
            "renamed": renamed,
            "reordered": reordered,
            "clean": clean,
            "summary": " | ".join(parts),
            "timestamp": datetime.now().isoformat()
        }


# ===== 标准投标文件框架模板 =====

STANDARD_FRAMEWORK_TEMPLATES = {
    "security": {
        "name": "安保服务类标书框架",
        "sections": [
            ("投标函", 1, "招标文件第五章投标文件格式"),
            ("开标一览表", 1, "招标文件第五章投标文件格式"),
            ("投标报价明细表", 1, "招标文件第五章投标文件格式"),
            ("法定代表人身份证明", 1, "招标文件第五章投标文件格式"),
            ("法定代表人授权委托书", 1, "招标文件第五章投标文件格式"),
            ("资格证明文件", 1, "招标文件第二章投标人须知"),
            ("保安服务许可证", 2, "资格要求第5条"),
            ("营业执照", 2, "资格要求第1条"),
            ("信用承诺函", 2, "资格要求第2条"),
            ("财务状况报告", 2, "资格要求第3条"),
            ("服务方案", 1, "招标文件第二章服务需求书"),
            ("项目理解与需求分析", 2, "服务需求书"),
            ("安保服务实施方案", 2, "服务需求书"),
            ("人员配置方案", 2, "服务需求书-人力配置"),
            ("设备投入方案", 2, "服务需求书-设备投入比例"),
            ("管理制度", 2, "服务需求书"),
            ("应急预案", 2, "服务需求书"),
            ("培训方案", 2, "服务需求书"),
            ("质量保证与考核承诺", 2, "服务需求书-季度考核"),
            ("技术响应表", 1, "招标文件第二章服务需求书"),
            ("商务响应表", 1, "招标文件第二章服务需求书"),
            ("服务承诺书", 1, "招标文件第五章投标文件格式"),
            ("其他材料", 1, "招标文件第五章投标文件格式"),
        ]
    },
    "property": {
        "name": "物业管理类标书框架",
        "sections": [
            ("投标函", 1, "招标文件第五章投标文件格式"),
            ("开标一览表", 1, "招标文件第五章投标文件格式"),
            ("投标报价明细表", 1, "招标文件第五章投标文件格式"),
            ("法定代表人身份证明", 1, "招标文件第五章投标文件格式"),
            ("法定代表人授权委托书", 1, "招标文件第五章投标文件格式"),
            ("资格证明文件", 1, "招标文件第二章投标人须知"),
            ("营业执照", 2, "资格要求"),
            ("物业服务相关资质", 2, "资格要求"),
            ("信用承诺函", 2, "资格要求"),
            ("财务状况报告", 2, "资格要求"),
            ("类似项目业绩", 2, "资格要求-业绩"),
            ("服务方案", 1, "招标文件第二章服务需求书"),
            ("项目理解与需求分析", 2, "服务需求书"),
            ("物业管理实施方案", 2, "服务需求书"),
            ("人员配置方案", 2, "服务需求书"),
            ("设备设施投入方案", 2, "服务需求书"),
            ("管理制度", 2, "服务需求书"),
            ("应急预案", 2, "服务需求书"),
            ("培训方案", 2, "服务需求书"),
            ("质量保证体系", 2, "服务需求书"),
            ("技术响应表", 1, "招标文件服务需求书"),
            ("商务响应表", 1, "招标文件服务需求书"),
            ("服务承诺书", 1, "招标文件第五章投标文件格式"),
            ("其他材料", 1, "招标文件第五章投标文件格式"),
        ]
    },
    "generic": {
        "name": "通用服务类标书框架",
        "sections": [
            ("投标函", 1, "招标文件第五章投标文件格式"),
            ("开标一览表", 1, "招标文件第五章投标文件格式"),
            ("投标报价明细表", 1, "招标文件第五章投标文件格式"),
            ("法定代表人身份证明", 1, "招标文件第五章投标文件格式"),
            ("法定代表人授权委托书", 1, "招标文件第五章投标文件格式"),
            ("资格证明文件", 1, "招标文件第二章投标人须知"),
            ("服务方案", 1, "招标文件第二章服务需求书"),
            ("技术响应表", 1, "招标文件服务需求书"),
            ("商务响应表", 1, "招标文件服务需求书"),
            ("服务承诺书", 1, "招标文件第五章投标文件格式"),
            ("其他材料", 1, "招标文件第五章投标文件格式"),
        ]
    }
}


def build_framework_from_template(template_name: str, project_name: str = "",
                                   tender_file: str = "") -> Framework:
    """从模板构建框架"""
    template = STANDARD_FRAMEWORK_TEMPLATES.get(template_name, STANDARD_FRAMEWORK_TEMPLATES["generic"])
    fw = Framework(project_name=project_name or template["name"],
                   tender_file=tender_file)
    for title, level, source in template["sections"]:
        fw.add_section(title=title, level=level, source=source)
    return fw


def build_framework_from_tender_extract(tender_extract: dict, project_name: str = "",
                                         tender_file: str = "") -> Framework:
    """
    从招标文件提取结果构建框架

    Args:
        tender_extract: 招标文件提取结果（来自TenderParser）
            {
                "project_name": "...",
                "bid_sections": [...],     # 招标文件要求的投标文件章节
                "scoring_items": [...],    # 评分项
                "qualification_reqs": [...], # 资格要求
                "disqualify_rules": [...]   # 废标条款
            }
    """
    fw = Framework(
        project_name=project_name or tender_extract.get("project_name", ""),
        tender_file=tender_file
    )

    # 从招标文件要求的投标文件章节构建
    for section_info in tender_extract.get("bid_sections", []):
        title = section_info.get("title", "")
        level = section_info.get("level", 1)
        source = section_info.get("source", "招标文件")
        required = section_info.get("required", True)
        if title:
            fw.add_section(title=title, level=level, source=source, required=required)

    # 如果招标文件提取结果为空，使用通用模板兜底
    if not fw.sections:
        return build_framework_from_template("generic", project_name, tender_file)

    return fw


def extract_titles_from_markdown(md_text: str) -> List[Dict]:
    """
    从Markdown文本提取标题结构

    Returns:
        [{"title": "...", "level": 1}, ...]
    """
    titles = []
    for line in md_text.split("\n"):
        line = line.strip()
        if not line:
            continue
        # Markdown标题: # / ## / ### / ####
        m = re.match(r'^(#{1,4})\s+(.+)', line)
        if m:
            level = len(m.group(1))
            title = m.group(2).strip()
            # 去掉Markdown残留标记
            title = re.sub(r'\*+', '', title)
            title = re.sub(r'`+', '', title)
            titles.append({"title": title, "level": level})
        else:
            # 数字编号标题: 一、/ 1./ (一)/ 1.1 等
            m2 = re.match(r'^(第[一二三四五六七八九十]+[章节篇]|[一二三四五六七八九十]+[、.]|\(\d+\)|\d+\.\d*\s+)(.+)', line)
            if m2:
                title = line.strip()
                # 推断层级
                prefix = m2.group(1)
                if prefix.startswith("第") or re.match(r'^[一二三四五六七八九十]+[、.]', prefix):
                    level = 1
                elif re.match(r'^\(\d+\)', prefix):
                    level = 2
                else:
                    level = 2
                titles.append({"title": title, "level": level})

    return titles


def extract_titles_from_docx(docx_path: str) -> List[Dict]:
    """从Word文档提取标题结构"""
    try:
        from docx import Document
        doc = Document(docx_path)
        titles = []
        for para in doc.paragraphs:
            style_name = (para.style.name if para.style else "") or ""
            text = para.text.strip()
            if not text:
                continue
            # 通过样式判断标题
            if style_name.startswith("Heading"):
                try:
                    level = int(style_name.replace("Heading ", "").replace("Heading", "1"))
                except ValueError:
                    level = 1
                titles.append({"title": text, "level": level})
            elif style_name == "Title":
                titles.append({"title": text, "level": 1})
        return titles
    except ImportError:
        print("❌ 需要 python-docx: pip install python-docx")
        return []
    except Exception as e:
        print(f"❌ 读取Word文档失败: {e}")
        return []
