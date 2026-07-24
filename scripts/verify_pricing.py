#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""报价逐项核验工具"""
import json, sys
sys.stdout.reconfigure(encoding='utf-8')

def verify_pricing(excel_data):
    """核验单价×数量=小计"""
    errors = []
    for i, item in enumerate(excel_data, 1):
        name = item.get('name', f'第{i}项')
        price = float(item.get('price', 0))
        qty = float(item.get('quantity', 0))
        subtotal = float(item.get('subtotal', 0))
        
        calc = price * qty
        if abs(calc - subtotal) > 0.01:
            errors.append(f'{name}: {price}×{qty}={calc} ≠ {subtotal}')
    
    if errors:
        print('❌ 报价错误:')
        for e in errors:
            print(f'  {e}')
    else:
        print('✅ 报价核验通过，零差异')
    return len(errors) == 0

if __name__ == '__main__':
    print('报价核验工具 v1.0')
