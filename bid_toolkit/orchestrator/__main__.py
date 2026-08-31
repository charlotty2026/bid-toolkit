#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""让编排引擎可以通过 python -m bid_toolkit.orchestrator 运行"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from bid_toolkit.orchestrator.engine import main
main()
