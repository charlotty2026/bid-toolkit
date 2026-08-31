#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
bid-toolkit GUI 模块入口
========================
提供 launch() 函数，供 CLI 的 `bid gui` 子命令调用。

使用方式：
    from bid_toolkit.gui import launch
    launch(port=7860, browser_only=False)
"""

from bid_toolkit.gui.app import build_app
from bid_toolkit.gui.desktop import start_gradio_then_desktop


def launch(port=7860, browser_only=False):
    """启动 bid-toolkit GUI。

    构建 Gradio 界面，然后通过 desktop 模块启动桌面窗口或浏览器。

    Args:
        port: 端口号（默认 7860，被占用时自动探测可用端口）
        browser_only: True 时只用浏览器，不用 pywebview 桌面窗口

    Raises:
        ImportError: 当 gradio 未安装时，提示用户安装 desktop 可选依赖
    """
    print("=" * 50)
    print("  bid-toolkit GUI 启动中...")
    print(f"  模式: {'浏览器' if browser_only else '桌面窗口'}")
    print(f"  端口: {port}（被占用时自动探测）")
    print("=" * 50)

    try:
        app = build_app()
    except ImportError:
        raise ImportError(
            "gradio is not installed. "
            'Install desktop dependencies with: pip install "bid-toolkit[desktop]"'
        )

    start_gradio_then_desktop(app, port=port, browser_only=browser_only)


__all__ = ["launch"]
