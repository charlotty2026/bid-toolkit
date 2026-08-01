#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
桌面窗口封装
============
pywebview 封装 + 端口自动探测 + 浏览器回退。

功能：
    1. 端口探测：从指定端口开始扫描，找到可用端口
    2. pywebview 窗口：将 Gradio Web 服务包装成原生桌面窗口
    3. 浏览器回退：pywebview 未安装或导入失败时，自动用系统浏览器打开
    4. 窗口居中：基于 screeninfo 获取屏幕尺寸，自动居中（可选依赖）

来源：dirge-1（pywebview 回退方案）+ bison-1（端口探测方案）
"""

import socket
import sys
import threading
import time
import webbrowser


def find_available_port(start=7860, max_tries=20):
    """从指定端口开始扫描，找到可用端口。

    投标人员电脑上可能装了 WPS、CAD 等软件，端口冲突常见。
    不要让用户手动改端口，工具自己解决。

    Args:
        start: 起始端口号（默认 7860，Gradio 默认端口）
        max_tries: 最多尝试多少个端口

    Returns:
        int: 可用端口号

    Raises:
        RuntimeError: 连续 max_tries 个端口都被占用
    """
    for port in range(start, start + max_tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except socket.OSError:
                continue
    raise RuntimeError(
        f"端口 {start}-{start + max_tries - 1} 全部被占用，"
        f"请检查是否有其他程序占用，或手动指定端口：bid gui --port 8080"
    )


def _get_window_geometry():
    """获取窗口尺寸和位置（居中）。

    使用 screeninfo 检测屏幕分辨率，计算居中位置。
    screeninfo 不可用时返回默认值。

    Returns:
        tuple: (width, height, x, y)
    """
    default_size = (1280, 860, None, None)

    try:
        import screeninfo

        monitors = screeninfo.get_monitors()
        if monitors:
            monitor = monitors[0]
            width = min(1280, int(monitor.width * 0.8))
            height = min(860, int(monitor.height * 0.8))
            x = int((monitor.width - width) / 2)
            y = int((monitor.height - height) / 2)
            return (width, height, x, y)
    except (ImportError, Exception):
        pass

    return default_size


def launch_desktop(url, title="bid-toolkit"):
    """启动桌面窗口或回退到浏览器。

    尝试用 pywebview 启动原生桌面窗口，如果 pywebview 不可用，
    自动回退到系统默认浏览器打开。

    Args:
        url: Gradio 服务的 URL（如 http://127.0.0.1:7860）
        title: 窗口标题
    """
    try:
        import webview

        width, height, x, y = _get_window_geometry()

        window_kwargs = {
            "width": width,
            "height": height,
            "title": title,
        }
        if x is not None and y is not None:
            window_kwargs["x"] = x
            window_kwargs["y"] = y

        print(f"启动桌面窗口: {title} ({width}x{height})")
        webview.create_window(title=title, url=url, **window_kwargs)
        webview.start()

    except ImportError:
        print("pywebview 未安装，使用系统浏览器打开...")
        webbrowser.open(url)
        print(f"已在浏览器中打开: {url}")
        print("如需桌面窗口体验，请运行: pip install pywebview")
        # 保持进程不退出，让 Gradio 服务继续运行
        print("按 Ctrl+C 退出...")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n已退出")
            sys.exit(0)

    except Exception as e:
        print(f"桌面窗口启动失败: {e}")
        print("回退到系统浏览器...")
        webbrowser.open(url)
        print(f"已在浏览器中打开: {url}")
        print("按 Ctrl+C 退出...")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\n已退出")
            sys.exit(0)


def start_gradio_then_desktop(gradio_app, port=7860, browser_only=False):
    """启动 Gradio 服务，然后打开桌面窗口。

    在子线程中启动 Gradio，主线程打开 pywebview 窗口。
    关闭窗口时自动退出整个程序。

    Args:
        gradio_app: Gradio Blocks 实例
        port: 端口号
        browser_only: True 时只用浏览器，不用 pywebview
    """
    from bid_toolkit.gui.gradio_compat import safe_launch

    actual_port = find_available_port(port)
    if actual_port != port:
        print(f"端口 {port} 被占用，自动切换到 {actual_port}")

    url = f"http://127.0.0.1:{actual_port}"

    if browser_only:
        # 浏览器模式：Gradio 自己启动 + 打开浏览器
        print(f"启动 Web 服务: {url}")
        safe_launch(gradio_app, server_port=actual_port, inbrowser=True)
        return

    # 桌面模式：子线程跑 Gradio，主线程跑 pywebview
    def _run_gradio():
        safe_launch(gradio_app, server_port=actual_port, inbrowser=False)

    gradio_thread = threading.Thread(target=_run_gradio, daemon=True)
    gradio_thread.start()

    # 等待 Gradio 服务就绪
    print(f"正在启动服务...", end="")
    for _ in range(30):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", actual_port)) == 0:
                break
        print(".", end="", flush=True)
        time.sleep(0.5)
    print(f" ready")

    launch_desktop(url, title="bid-toolkit 桌面版")
