#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查收款数据列名和内容
"""

import pandas as pd
import os

def check_collection_data():
    """检查收款数据"""
    data_path = r"C:\Users\mingh\client-data\raw\客户档案\CBD\订阅合同收款情况\订阅合同收款情况.xlsx"
    
    print("=" * 60)
    print("收款数据详细检查")
    print("=" * 60)
    print(f"文件路径: {data_path}")
    print(f"文件存在: {os.path.exists(data_path)}")
    
    if not os.path.exists(data_path):
        print("文件不存在！")
        return
    
    try:
        # 读取Excel文件
        df = pd.read_excel(data_path)
        print(f"\n数据形状: {df.shape}")
        print(f"列数: {len(df.columns)}")
        
        print("\n所有列名:")
        for i, col in enumerate(df.columns):
            print(f'{i:3d}: "{col}"')
        
        print("\n数据类型:")
        print(df.dtypes)
        
        print("\n前3行数据（完整显示）:")
        print(df.head(3).to_string())
        
        # 检查所有数值列
        print("\n" + "=" * 60)
        print("数值列详细分析")
        print("=" * 60)
        
        numeric_cols = df.select_dtypes(include=['number']).columns
        print(f"找到 {len(numeric_cols)} 个数值列:")
        
        for col in numeric_cols:
            non_null = df[col].notna().sum()
            total = df[col].sum()
            mean = df[col].mean()
            min_val = df[col].min()
            max_val = df[col].max()
            
            print(f"\n{col}:")
            print(f"  数据类型: {df[col].dtype}")
            print(f"  非空值数: {non_null}/{len(df)} ({non_null/len(df)*100:.1f}%)")
            print(f"  总和: {total:,.0f}")
            print(f"  平均值: {mean:,.0f}")
            print(f"  最小值: {min_val:,.0f}")
            print(f"  最大值: {max_val:,.0f}")
            
            # 检查是否为金额列
            if total > 0 and max_val > 1000:  # 假设金额大于1000
                print(f"  可能是金额列!")
        
        # 检查文本列中可能包含金额的列
        print("\n" + "=" * 60)
        print("文本列中的金额信息检查")
        print("=" * 60)
        
        text_cols = df.select_dtypes(include=['object']).columns
        amount_keywords = ['金额', '收款', '支付', '付费', '费用', '计划', '实际', '应收', '实收']
        
        for col in text_cols:
            col_name = str(col)
            for keyword in amount_keywords:
                if keyword in col_name:
                    print(f"\n{col}:")
                    print(f"  非空值数: {df[col].notna().sum()}/{len(df)}")
                    print(f"  前3个值:")
                    for val in df[col].head(3):
                        print(f"    - {val}")
                    break
        
        # 检查特定的列
        print("\n" + "=" * 60)
        print("特定列检查")
        print("=" * 60)
        
        # 检查可能包含实收金额的列
        possible_received_cols = [
            '实际收款金额', '实收金额', '已收款', '已付款', 
            '实际收款', '实收', '收款金额', '收款额'
        ]
        
        found_cols = []
        for possible_col in possible_received_cols:
            for actual_col in df.columns:
                if possible_col in str(actual_col):
                    found_cols.append(actual_col)
        
        if found_cols:
            print(f"找到 {len(found_cols)} 个可能包含实收金额的列:")
            for col in found_cols:
                print(f"\n{col}:")
                print(f"  数据类型: {df[col].dtype}")
                if pd.api.types.is_numeric_dtype(df[col]):
                    total = df[col].sum()
                    print(f"  总和: {total:,.0f}")
                    print(f"  非空值: {df[col].notna().sum()}/{len(df)}")
                else:
                    print(f"  前3个值:")
                    for val in df[col].head(3):
                        print(f"    - {val}")
        else:
            print("未找到明显的实收金额列")
            
        # 检查计划收款金额列
        print("\n" + "=" * 60)
        print("计划收款金额检查")
        print("=" * 60)
        
        possible_due_cols = ['计划收款金额', '应收金额', '应付款', '应收款']
        for possible_col in possible_due_cols:
            for actual_col in df.columns:
                if possible_col in str(actual_col):
                    print(f"\n{actual_col}:")
                    print(f"  数据类型: {df[actual_col].dtype}")
                    if pd.api.types.is_numeric_dtype(df[actual_col]):
                        total = df[actual_col].sum()
                        print(f"  总和: {total:,.0f}")
                        print(f"  非空值: {df[actual_col].notna().sum()}/{len(df)}")
                    break
        
    except Exception as e:
        print(f"读取文件失败: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_collection_data()