#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将生成的Word文档移动到client_data文件夹
"""

import os
import shutil
from pathlib import Path

def main():
    print("开始移动文件到client_data文件夹...")
    
    # 源文件路径
    source_file = r"C:\Users\mingh\.openclaw\workspace\skills\business-expert\scripts\reports_fixed\CBD_经营分析报告_20260317_final.docx"
    
    # 目标文件夹
    client_data_root = r"C:\Users\mingh\client_data"
    client_name = "CBD"
    target_dir = Path(client_data_root) / client_name
    
    print(f"源文件: {source_file}")
    print(f"目标目录: {target_dir}")
    
    # 检查源文件
    if not os.path.exists(source_file):
        print("错误: 源文件不存在")
        return 1
    
    source_size = os.path.getsize(source_file)
    print(f"源文件大小: {source_size:,} 字节")
    
    # 创建目标目录
    os.makedirs(target_dir, exist_ok=True)
    print(f"目标目录已创建: {target_dir}")
    
    # 目标文件路径
    target_file = target_dir / "CBD_经营分析报告_20260317_final.docx"
    print(f"目标文件: {target_file}")
    
    # 复制文件
    try:
        shutil.copy2(source_file, target_file)
        print("文件复制成功")
        
        # 验证
        if os.path.exists(target_file):
            target_size = os.path.getsize(target_file)
            print(f"验证成功")
            print(f"目标文件大小: {target_size:,} 字节")
            
            # 列出目录内容
            print(f"\n目标目录内容:")
            for item in os.listdir(target_dir):
                item_path = target_dir / item
                if os.path.isfile(item_path):
                    size = os.path.getsize(item_path)
                    print(f"  {item} ({size:,} 字节)")
                else:
                    print(f"  {item}/")
            
            return 0
        else:
            print("错误: 目标文件创建失败")
            return 1
            
    except Exception as e:
        print(f"文件移动失败: {e}")
        return 1

if __name__ == "__main__":
    exit(main())