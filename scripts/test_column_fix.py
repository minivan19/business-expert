#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试列查找修复
"""

import pandas as pd
import os

def test_column_finding():
    data_path = r"C:\Users\mingh\client-data\raw\客户档案\CBD\订阅合同收款情况\订阅合同收款情况.xlsx"
    df = pd.read_excel(data_path)
    
    print("测试改进的列查找逻辑...")
    
    # 查找应收金额列（计划收款金额）
    due_column = None
    for col in df.columns:
        col_str = str(col)
        if '计划' in col_str and pd.api.types.is_numeric_dtype(df[col]):
            due_column = col
            print(f'找到应收金额列: {col}')
            break
    
    if not due_column:
        for col in df.columns:
            col_str = str(col)
            if '应收' in col_str and pd.api.types.is_numeric_dtype(df[col]):
                due_column = col
                print(f'找到应收金额列: {col}')
                break
    
    # 查找实收金额列（已收款金额）
    received_column = None
    for col in df.columns:
        col_str = str(col)
        if ('已收款' in col_str or '实收' in col_str) and pd.api.types.is_numeric_dtype(df[col]):
            received_column = col
            print(f'找到实收金额列: {col}')
            break
    
    if not received_column:
        # 搜索包含'收款'但不包含'计划'的列
        for col in df.columns:
            col_str = str(col)
            if '收款' in col_str and '计划' not in col_str and pd.api.types.is_numeric_dtype(df[col]):
                received_column = col
                print(f'找到实收金额列: {col}')
                break
    
    print()
    if due_column:
        print(f'✅ 应收金额列: "{due_column}"')
        print(f'   总和: {df[due_column].sum():,.0f}元')
    else:
        print('❌ 未找到应收金额列')
    
    if received_column:
        print(f'✅ 实收金额列: "{received_column}"')
        print(f'   总和: {df[received_column].sum():,.0f}元')
    else:
        print('❌ 未找到实收金额列')
        print('  现有包含"收款"的列:')
        for col in df.columns:
            if '收款' in str(col):
                print(f'    - "{col}": {df[col].dtype}')
    
    # 计算收款率
    if due_column and received_column and df[due_column].sum() > 0:
        collection_rate = df[received_column].sum() / df[due_column].sum() * 100
        print(f'\n📊 收款率: {collection_rate:.1f}%')
        print(f'   应收总额: {df[due_column].sum():,.0f}元')
        print(f'   实收总额: {df[received_column].sum():,.0f}元')
        print(f'   未收金额: {df[due_column].sum() - df[received_column].sum():,.0f}元')
    
    return due_column, received_column

if __name__ == "__main__":
    due_col, received_col = test_column_finding()
    
    print("\n" + "="*60)
    print("数据验证")
    print("="*60)
    
    if due_col and received_col:
        df = pd.read_excel(r"C:\Users\mingh\client-data\raw\客户档案\CBD\订阅合同收款情况\订阅合同收款情况.xlsx")
        
        print(f"应收金额 ({due_col}): {df[due_col].sum():,.0f}元")
        print(f"实收金额 ({received_col}): {df[received_col].sum():,.0f}元")
        
        # 检查其他相关列
        print("\n其他相关金额列:")
        for col in df.columns:
            if '金额' in str(col) and pd.api.types.is_numeric_dtype(df[col]):
                total = df[col].sum()
                if total > 0:
                    print(f'  "{col}": {total:,.0f}元')