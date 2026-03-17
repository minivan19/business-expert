#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简单测试列查找
"""

import pandas as pd
import os

def main():
    data_path = r"C:\Users\mingh\client-data\raw\客户档案\CBD\订阅合同收款情况\订阅合同收款情况.xlsx"
    df = pd.read_excel(data_path)
    
    print("测试列查找...")
    
    # 查找应收金额列
    due_column = None
    for col in df.columns:
        col_str = str(col)
        if '计划' in col_str and pd.api.types.is_numeric_dtype(df[col]):
            due_column = col
            print(f'应收金额列: {col}')
            break
    
    # 查找实收金额列
    received_column = None
    for col in df.columns:
        col_str = str(col)
        if '已收款' in col_str and pd.api.types.is_numeric_dtype(df[col]):
            received_column = col
            print(f'实收金额列: {col}')
            break
    
    if not received_column:
        for col in df.columns:
            col_str = str(col)
            if '收款' in col_str and '计划' not in col_str and pd.api.types.is_numeric_dtype(df[col]):
                received_column = col
                print(f'实收金额列: {col}')
                break
    
    print()
    if due_column:
        print(f'应收金额列: "{due_column}"')
        print(f'总和: {df[due_column].sum():,.0f}元')
    
    if received_column:
        print(f'实收金额列: "{received_column}"')
        print(f'总和: {df[received_column].sum():,.0f}元')
    
    # 计算收款率
    if due_column and received_column and df[due_column].sum() > 0:
        collection_rate = df[received_column].sum() / df[due_column].sum() * 100
        print(f'\n收款率: {collection_rate:.1f}%')
        print(f'应收总额: {df[due_column].sum():,.0f}元')
        print(f'实收总额: {df[received_column].sum():,.0f}元')
        print(f'未收金额: {df[due_column].sum() - df[received_column].sum():,.0f}元')

if __name__ == "__main__":
    main()