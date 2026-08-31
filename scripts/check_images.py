#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""docx图片状态检查工具"""
import zipfile, os, sys
sys.stdout.reconfigure(encoding='utf-8')

def check_docx_images(docx_path):
    """检查docx中的图片渲染状态"""
    print(f'检查: {docx_path}')
    
    with zipfile.ZipFile(docx_path, 'r') as z:
        # 列出所有媒体文件
        media_files = [f for f in z.namelist() if f.startswith('word/media/')]
        print(f'  媒体文件: {len(media_files)}个')
        
        # 检查document.xml引用
        doc_xml = z.read('word/document.xml').decode('utf-8')
        for media in media_files:
            filename = os.path.basename(media)
            if filename in doc_xml:
                print(f'  ✅ {filename} - 已引用')
            else:
                print(f'  ⚠️ {filename} - 未引用（幽灵图片）')

if __name__ == '__main__':
    print('docx图片检查工具 v1.0')
