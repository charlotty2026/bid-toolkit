#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gradio 4/5/6 跨版本兼容层
=========================
解决不同 Gradio 版本之间 API 差异，保证 bid-toolkit GUI 在
Gradio 4.x / 5.x / 6.x 三个大版本下都能正常运行。

使用方式：
    from bid_toolkit.gui.gradio_compat import get_gradio_version, safe_launch

主要兼容点：
    1. launch() 方法参数差异（server_port -> server_name/port）
    2. Blocks API 变化（css 参数传递方式）
    3. Tab/Tabs 组件参数差异
    4. File 组件返回值类型差异（4.x 返回路径字符串，5.x+ 返回文件对象）

来源：方舟36期 dirge-1 终局产出验证方案
"""

import functools


def get_gradio_version():
    """获取当前安装的 Gradio 主版本号。

    Returns:
        int: 主版本号（4/5/6），未安装时返回 0
    """
    try:
        import gradio as gr

        return int(gr.__version__.split(".")[0])
    except (ImportError, ValueError, AttributeError):
        return 0


def safe_launch(app, **kwargs):
    """跨版本安全启动 Gradio 应用。

    不同版本的 launch() 参数名和默认行为有差异，本函数统一处理。

    Args:
        app: Gradio Blocks 或 Interface 实例
        **kwargs: 传递给 launch() 的参数

    Returns:
        app.launch() 的返回值
    """
    ver = get_gradio_version()

    # 所有版本（4/5/6）统一用 server_port，不转换
    # （实测 Gradio 6.22 仍用 server_port，早期分析有误）

    # Gradio 6.x 的 prevent_thread_lock 默认行为变化
    if ver >= 6:
        kwargs.setdefault("prevent_thread_lock", False)

    # 统一设置：不在浏览器自动打开（由 desktop.py 控制）
    kwargs.setdefault("inbrowser", False)
    kwargs.setdefault("show_error", True)

    # Gradio 6.x: theme/css 从 Blocks 构造函数移到 launch() 参数
    launch_extra = getattr(app, "_bid_launch_kwargs", None)
    if launch_extra:
        kwargs.setdefault("theme", launch_extra.get("theme"))
        kwargs.setdefault("css", launch_extra.get("css"))

    return app.launch(**kwargs)


def normalize_file_input(file_obj):
    """统一文件输入的返回值类型。

    Gradio 4.x 的 File 组件返回文件路径字符串，
    5.x+ 可能返回 tempfile 对象或 NamedString。

    Args:
        file_obj: Gradio File 组件的返回值

    Returns:
        str: 文件路径字符串
    """
    if file_obj is None:
        return None

    # Gradio 4.x: 直接是路径字符串
    if isinstance(file_obj, str):
        return file_obj

    # Gradio 5.x+: NamedString 或 tempfile 对象
    # NamedString 有 .path 或 .name 属性
    if hasattr(file_obj, "path"):
        return file_obj.path
    if hasattr(file_obj, "name"):
        return file_obj.name

    # 列表情况：取第一个
    if isinstance(file_obj, (list, tuple)) and len(file_obj) > 0:
        return normalize_file_input(file_obj[0])

    return str(file_obj)


def compat_tab(label, **kwargs):
    """创建跨版本兼容的 Tab 组件。

    Args:
        label: Tab 标签文本
        **kwargs: 传递给 gr.Tab 的参数

    Returns:
        gr.Tab 实例
    """
    import gradio as gr

    ver = get_gradio_version()

    # Gradio 4.x 用 label 参数，5.x+ 也兼容
    # 但 6.x 的 selected 参数行为不同，不传
    kwargs.pop("selected", None)

    return gr.Tab(label=label, **kwargs)


def compat_markdown(value):
    """创建跨版本兼容的 Markdown 组件。

    Args:
        value: 初始 Markdown 内容

    Returns:
        gr.Markdown 实例
    """
    import gradio as gr

    return gr.Markdown(value=value)


def compat_file(label, file_types=None):
    """创建跨版本兼容的 File 组件。

    Args:
        label: 文件输入标签
        file_types: 允许的文件类型列表

    Returns:
        gr.File 实例
    """
    import gradio as gr

    kwargs = {"label": label}
    if file_types:
        kwargs["file_types"] = file_types

    return gr.File(**kwargs)
