#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LLM客户端模块 v1.0
支持OpenAI兼容接口（DeepSeek/千帆/Ollama/通义千问等）
内置重试机制、超时处理、rate limit退避

用法：
    from llm_client import LLMClient
    client = LLMClient(config_path='templates/user_config.yaml')
    response = client.chat("请写一段公司简介")
"""

import os
import sys
import json
import time
import yaml
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

# 默认配置
DEFAULT_LLM_CONFIG = {
    'api_key': '',           # 用户的API Key
    'base_url': '',          # API地址，如 https://api.deepseek.com/v1
    'model': 'deepseek-chat', # 模型名
    'temperature': 0.7,      # 温度
    'max_tokens': 4096,      # 单次最大token
    'timeout': 60,           # 单次请求超时（秒）
    'retry': {
        'max_attempts': 3,   # 最大重试次数（含首次）
        'base_delay': 2,     # 基础退避延迟（秒）
        'max_delay': 30,     # 最大退避延迟（秒）
    },
}


class LLMClient:
    """OpenAI兼容接口的LLM客户端，内置重试和rate limit处理"""

    def __init__(self, config_path=None, config_dict=None):
        """
        初始化LLM客户端
        
        参数：
            config_path: user_config.yaml路径
            config_dict: 直接传入配置字典（优先于config_path）
        """
        self.config = self._load_config(config_path, config_dict)
        self._client = None
        self._init_client()

    def _load_config(self, config_path, config_dict):
        """加载配置，优先级：config_dict > config_path > 默认"""
        config = DEFAULT_LLM_CONFIG.copy()
        config['retry'] = DEFAULT_LLM_CONFIG['retry'].copy()

        if config_dict:
            llm_cfg = config_dict.get('llm', {})
        elif config_path:
            llm_cfg = self._load_from_yaml(config_path)
        else:
            # 自动查找 templates/user_config.yaml
            auto_path = Path(__file__).parent.parent / 'templates' / 'user_config.yaml'
            if auto_path.exists():
                llm_cfg = self._load_from_yaml(str(auto_path))
            else:
                llm_cfg = {}

        # 合并配置
        for k, v in llm_cfg.items():
            if k == 'retry' and isinstance(v, dict):
                config['retry'].update(v)
            else:
                config[k] = v

        return config

    def _load_from_yaml(self, config_path):
        """从YAML文件加载llm配置段"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                full_config = yaml.safe_load(f) or {}
            return full_config.get('llm', {})
        except Exception as e:
            print(f"⚠️  加载LLM配置失败: {e}", file=sys.stderr)
            return {}

    def _init_client(self):
        """初始化OpenAI客户端"""
        api_key = self.config.get('api_key', '')
        base_url = self.config.get('base_url', '')

        if not api_key:
            print("⚠️  未配置API Key，LLM功能不可用。请在templates/user_config.yaml中配置llm.api_key", file=sys.stderr)
            self._client = None
            return

        try:
            from openai import OpenAI
            client_kwargs = {'api_key': api_key, 'timeout': self.config.get('timeout', 60)}
            if base_url:
                client_kwargs['base_url'] = base_url
            self._client = OpenAI(**client_kwargs)
        except ImportError:
            print("❌ 需要安装openai库: pip install openai", file=sys.stderr)
            self._client = None

    def chat(self, messages, **kwargs):
        """
        发送聊天请求，内置重试机制
        
        参数：
            messages: OpenAI格式消息列表
            kwargs: 覆盖默认参数（temperature/max_tokens等）
        
        返回：
            str: 模型回复文本，失败时返回空字符串
        """
        if not self._client:
            print("❌ LLM客户端未初始化", file=sys.stderr)
            return ''

        # 合并参数
        params = {
            'model': self.config.get('model', 'deepseek-chat'),
            'temperature': self.config.get('temperature', 0.7),
            'max_tokens': self.config.get('max_tokens', 4096),
        }
        params.update(kwargs)

        retry_cfg = self.config.get('retry', {})
        max_attempts = retry_cfg.get('max_attempts', 3)
        base_delay = retry_cfg.get('base_delay', 2)
        max_delay = retry_cfg.get('max_delay', 30)

        last_error = None
        for attempt in range(1, max_attempts + 1):
            try:
                response = self._client.chat.completions.create(
                    messages=messages,
                    **params
                )
                return response.choices[0].message.content or ''

            except Exception as e:
                last_error = e
                error_str = str(e)

                # 判断错误类型
                is_rate_limit = 'rate' in error_str.lower() or '429' in error_str
                is_timeout = 'timeout' in error_str.lower() or 'timed out' in error_str.lower()
                is_server_error = '500' in error_str or '502' in error_str or '503' in error_str

                if attempt < max_attempts and (is_rate_limit or is_timeout or is_server_error):
                    # 指数退避
                    delay = min(base_delay * (2 ** (attempt - 1)), max_delay)
                    if is_rate_limit:
                        delay = delay * 2  # rate limit多等一会
                    print(f"⚠️  请求失败（第{attempt}次），{delay}秒后重试: {error_str[:100]}", file=sys.stderr)
                    time.sleep(delay)
                    continue
                else:
                    # 不可重试的错误或重试用完
                    print(f"❌ LLM请求失败（第{attempt}/{max_attempts}次）: {error_str[:200]}", file=sys.stderr)
                    break

        return ''

    def chat_simple(self, system_prompt, user_prompt, **kwargs):
        """
        简化版调用：传入system和user文本
        
        参数：
            system_prompt: 系统提示词
            user_prompt: 用户提示词
            kwargs: 覆盖默认参数
        
        返回：
            str: 模型回复文本
        """
        messages = []
        if system_prompt:
            messages.append({'role': 'system', 'content': system_prompt})
        messages.append({'role': 'user', 'content': user_prompt})
        return self.chat(messages, **kwargs)

    def is_available(self):
        """检查LLM客户端是否可用"""
        return self._client is not None


def load_user_config(config_path=None):
    """
    加载完整的user_config.yaml配置
    
    返回：
        dict: 完整配置字典
    """
    if config_path is None:
        config_path = Path(__file__).parent.parent / 'templates' / 'user_config.yaml'
    config_path = Path(config_path)

    if not config_path.exists():
        print(f"⚠️  配置文件不存在: {config_path}", file=sys.stderr)
        return {}

    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f) or {}
    except Exception as e:
        print(f"❌ 加载配置失败: {e}", file=sys.stderr)
        return {}


if __name__ == '__main__':
    # 自测
    print("LLM客户端模块 v1.0")
    print("=" * 40)

    config = load_user_config()
    if config:
        llm_cfg = config.get('llm', {})
        print(f"模型: {llm_cfg.get('model', '未配置')}")
        print(f"API地址: {llm_cfg.get('base_url', '未配置')}")
        print(f"API Key: {'已配置' if llm_cfg.get('api_key') else '未配置'}")
        print(f"超时: {llm_cfg.get('timeout', 60)}秒")
        retry = llm_cfg.get('retry', {})
        print(f"重试: 最多{retry.get('max_attempts', 3)}次, 基础延迟{retry.get('base_delay', 2)}秒")
    else:
        print("未找到user_config.yaml，使用默认配置")

    client = LLMClient()
    if client.is_available():
        print("\n✅ 客户端已就绪")
        # 简单测试
        resp = client.chat_simple("你是测试助手", "请回复'OK'")
        if resp:
            print(f"测试回复: {resp[:50]}")
        else:
            print("⚠️  测试请求失败，请检查API配置")
    else:
        print("\n⚠️  客户端未就绪，请配置API Key")
