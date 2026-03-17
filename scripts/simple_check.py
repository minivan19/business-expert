#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简单检查收款数据
"""

import pandas as pd
import os

def main():
    data_path = r"C:\Users\mingh\client-data\raw\客户档案\CBD\订阅合同收款情况\订阅合同收款情况.xlsx"
    
    print("检查收款数据...")
    print(f"文件路径: {data_path}")
    
    if not os.path.exists(data_path):
        print("文件不存在!")
        return
    
    try:
        df = pd.read_excel(data_path)
        print(f"\n数据形状: {df.shape}")
        print(f"列数: {len(df.columns)}")
        
        print("\n前5个列名:")
        for i, col in enumerate(df.columns[:5]):
            print(f'{i}: "{col}"')
        
        print("\n所有数值列:")
        numeric_cols = df.select_dtypes(include=['number']).columns
        for col in numeric_cols:
            total = df[col].sum()
            non_null = df[col].notna().sum()
            print(f'  "{col}": 总和={total:,.0f}, 非空值={non_null}/{len(df)}')
        
        print("\n检查特定列:")
        # 检查计划收款金额
        plan_col = None
        for col in df.columns:
            if '计划收款金额' in str(col):
                plan_col = col
                break
        if plan_col:
            print(f'  "{plan_col}": {df[plan_col].sum():,.0f}')
        else:
            print('  "计划收款金额"列不存在')
            
        # 检查实际收款金额
        actual_col = None
        for col in df.columns:
            if '实际收款金额' in str(col):
                actual_col = col
                break
        if actual_col:
            print(f'  "{actual_col}": {df[actual_col].sum():,.0f}')
        else:
            print('  "实际收款金额"列不存在')
            
        # 检查已开票金额
        invoice_col = None
        for col in df.columns:
            if '已开票金额' in str(col):
                invoice_col = col
                break
        if invoice_col:
            print(f'  "{invoice_col}": {df[invoice_col].sum():,.0f}')
        else:
            print('  "已开票金额"列不存在')
            
        # 检查实际收款回款
        received_col = None
        for col in df.columns:
            if '实际收款回款' in str(col):
                received_col = col
                break
        if received_col:
            print(f'  "{received_col}": {df[received_col].sum():,.0f}')
        else:
            print('  "实际收款回款"列不存在')
            
        print("\n搜索包含'收款'的列:")
        for col in df.columns:
            if '收款' in str(col):
                print(f'  "{col}": {df[col].dtype}')
                if pd.api.types.is_numeric_dtype(df[col]):
                    print(f'    总和: {df[col].sum():,.0f}')
                    
        print("\n搜索包含'金额'的列:")
        for col in df.columns:
            if '金额' in str(col):
                print(f'  "{col}": {df[col].dtype}')
                if pd.api.types.is_numeric_dtype(df[col]):
                    print(f'    总和: {df[col].sum():,.0f}')
                    
    except Exception as e:
        print(f"错误: {str(e)}")

if __name__ == "__main__":
    main()