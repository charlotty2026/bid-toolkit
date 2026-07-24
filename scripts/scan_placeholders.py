#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""占位符扫描工具"""
import re, sys
sys.stdout.reconfigure(encoding='utf-8')

def scan_placeholders(text):
    """扫描[XX]、[400-]等占位符"""
    patterns = [
        r'\[XX\]',
        r'\[\d{3}-\]',
        r'\[.*?XX.*?\]',
        r'\d{4}年\d{1,2}月\s+日',
    ]
    
    found = []
    for pattern in patterns:
        matches = re.findall(pattern, text)
        found.extend(matches)
    
    if found:
        print(f'⚠️ 发现{len(found)}处占位符/未填日期:')
        for f in set(found):
            print(f'  - {f}')
    else:
        print('✅ 无占位符')
    
    return found

if __name__ == '__main__':
    print('占位符扫描工具 v1.0')
