#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
检查实际数据列名
"""

import pandas as pd
import os
import sys

def check_subscription_columns():
    """检查订阅数据列名"""
    data_path = r"C:\Users\mingh\client-data\raw\客户档案\CBD\订阅合同行\订阅明细.xlsx"
    
    if not os.path.exists(data_path):
        print(f"文件不存在: {data_path}")
        return
    
    try:
        df = pd.read_excel(data_path)
        print("=" * 60)
        print("订阅数据列名检查")
        print("=" * 60)
        print(f"文件: {data_path}")
        print(f"数据形状: {df.shape}")
        print(f"列数: {len(df.columns)}")
        print("\n列名列表:")
        for i, col in enumerate(df.columns):
            print(f"{i:3d}: '{col}'")
        
        print("\n数据类型:")
        print(df.dtypes)
        
        print("\n前3行数据:")
        print(df.head(3).to_string())
        
    except Exception as e:
        print(f"读取文件失败: {str(e)}")

def check_collection_columns():
    """检查收款数据列名"""
    data_path = r"C:\Users\mingh\client-data\raw\客户档案\CBD\订阅合同收款情况\订阅合同收款情况.xlsx"
    
    if not os.path.exists(data_path):
        print(f"文件不存在: {data_path}")
        return
    
    try:
        df = pd.read_excel(data_path)
        print("\n" + "=" * 60)
        print("收款数据列名检查")
        print("=" * 60)
        print(f"文件: {data_path}")
        print(f"数据形状: {df.shape}")
        print(f"列数: {len(df.columns)}")
        print("\n列名列表:")
        for i, col in enumerate(df.columns):
            print(f"{i:3d}: '{col}'")
        
        print("\n数据类型:")
        print(df.dtypes)
        
        print("\n前3行数据:")
        print(df.head(3).to_string())
        
    except Exception as e:
        print(f"读取文件失败: {str(e)}")

if __name__ == "__main__":
    check_subscription_columns()
    check_collection_columns()