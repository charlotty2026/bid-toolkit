# -*- coding: utf-8 -*-
"""
标书编排引擎 (Bid Orchestration Engine)
=======================================
把铁律从"靠Agent自觉"变成"引擎强制校验"。

工作流：拆招标文件 -> 搭框架 -> 锁框架 -> 填内容 -> 铁律校验
"""

from .engine import BidOrchestrator
from .framework import Framework, FrameworkLock
from .iron_rules import IronRuleChecker
from .project import ProjectManager

__version__ = "1.0.0"
__all__ = ["BidOrchestrator", "Framework", "FrameworkLock", "IronRuleChecker", "ProjectManager"]
