#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
分析具体数据源表结构
根据用户指出的具体数据来源：
1. 第1部分：客户主数据表
2. 第3部分：固定金额合同台账、人天框架合同台帐
3. 第4部分：运维工单明细表
"""

import os
import sys
import pandas as pd
import json
from datetime import datetime

def analyze_cbd_data_sources():
    """分析CBD客户的具体数据源表结构"""
    
    print("分析CBD客户具体数据源表结构")
    print("=" * 60)
    
    # CBD客户数据目录
    cbd_dir = r"C:\Users\mingh\client-data\raw\客户档案\CBD"
    
    if not os.path.exists(cbd_dir):
        print(f"错误: CBD客户目录不存在: {cbd_dir}")
        return None
    
    print(f"CBD客户目录: {cbd_dir}")
    
    analysis_results = {
        "客户主数据": None,
        "固定金额合同台账": None,
        "人天框架合同台帐": None,
        "运维工单明细": None,
        "分析时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    
    # 1. 分析客户主数据表
    print(f"\n1. 分析客户主数据表...")
    customer_main_files = [
        os.path.join(cbd_dir, "客户主数据.xlsx"),
        os.path.join(cbd_dir, "客户主数据.xls"),
        os.path.join(cbd_dir, "客户基础档案.xlsx"),
        os.path.join(cbd_dir, "客户信息.xlsx")
    ]
    
    for file_path in customer_main_files:
        if os.path.exists(file_path):
            print(f"  找到文件: {file_path}")
            analysis_results["客户主数据"] = analyze_excel_file(file_path)
            break
    
    if not analysis_results["客户主数据"]:
        print(f"  未找到客户主数据表")
    
    # 2. 分析固定金额合同台账
    print(f"\n2. 分析固定金额合同台账...")
    fixed_contract_files = [
        os.path.join(cbd_dir, "固定金额合同台账.xlsx"),
        os.path.join(cbd_dir, "固定金额合同台账.xls"),
        os.path.join(cbd_dir, "固定金额实施.xlsx"),
        os.path.join(cbd_dir, "实施优化情况.xlsx")
    ]
    
    for file_path in fixed_contract_files:
        if os.path.exists(file_path):
            print(f"  找到文件: {file_path}")
            analysis_results["固定金额合同台账"] = analyze_excel_file(file_path)
            break
    
    if not analysis_results["固定金额合同台账"]:
        print(f"  未找到固定金额合同台账")
    
    # 3. 分析人天框架合同台帐
    print(f"\n3. 分析人天框架合同台帐...")
    manday_contract_files = [
        os.path.join(cbd_dir, "人天框架合同台帐.xlsx"),
        os.path.join(cbd_dir, "人天框架合同台帐.xls"),
        os.path.join(cbd_dir, "人天框架实施.xlsx"),
        os.path.join(cbd_dir, "人天合同.xlsx")
    ]
    
    for file_path in manday_contract_files:
        if os.path.exists(file_path):
            print(f"  找到文件: {file_path}")
            analysis_results["人天框架合同台帐"] = analyze_excel_file(file_path)
            break
    
    if not analysis_results["人天框架合同台帐"]:
        print(f"  未找到人天框架合同台帐")
    
    # 4. 分析运维工单明细
    print(f"\n4. 分析运维工单明细...")
    operation_dir = os.path.join(cbd_dir, "运维工单")
    if os.path.exists(operation_dir):
        print(f"  找到运维工单目录: {operation_dir}")
        
        operation_files = [
            os.path.join(operation_dir, "运维工单明细.xlsx"),
            os.path.join(operation_dir, "运维工单明细.xls"),
            os.path.join(operation_dir, "工单明细.xlsx"),
            os.path.join(operation_dir, "运维工单.xlsx")
        ]
        
        for file_path in operation_files:
            if os.path.exists(file_path):
                print(f"  找到文件: {file_path}")
                analysis_results["运维工单明细"] = analyze_excel_file(file_path)
                break
        
        if not analysis_results["运维工单明细"]:
            print(f"  在运维工单目录中未找到明细文件")
            
            # 尝试列出目录中的所有文件
            print(f"  目录内容:")
            for item in os.listdir(operation_dir)[:10]:  # 最多显示10个
                print(f"    - {item}")
    else:
        print(f"  未找到运维工单目录")
    
    # 5. 检查订阅合同收款情况（第2部分）
    print(f"\n5. 检查订阅合同收款情况（第2部分）...")
    subscription_dir = os.path.join(cbd_dir, "订阅合同收款情况")
    if os.path.exists(subscription_dir):
        print(f"  找到订阅合同收款情况目录: {subscription_dir}")
        
        subscription_files = [
            os.path.join(subscription_dir, "订阅合同收款情况.xlsx"),
            os.path.join(subscription_dir, "订阅合同.xlsx"),
            os.path.join(subscription_dir, "收款情况.xlsx")
        ]
        
        for file_path in subscription_files:
            if os.path.exists(file_path):
                print(f"  找到文件: {file_path}")
                # 这里只分析，不存储到结果中
                analyze_excel_file(file_path, show_details=False)
                break
    else:
        print(f"  未找到订阅合同收款情况目录")
    
    # 保存分析结果
    output_file = r"C:\Users\mingh\.openclaw\workspace\data_source_analysis.json"
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(analysis_results, f, ensure_ascii=False, indent=2)
    
    print(f"\n分析结果已保存到: {output_file}")
    
    # 生成分析报告
    generate_analysis_report(analysis_results)
    
    return analysis_results

def analyze_excel_file(file_path, show_details=True):
    """分析Excel文件结构"""
    try:
        print(f"    读取文件: {os.path.basename(file_path)}")
        
        # 尝试读取Excel文件
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            print(f"    读取失败: {e}")
            return None
        
        file_info = {
            "file_path": file_path,
            "file_size": os.path.getsize(file_path),
            "sheet_count": len(pd.ExcelFile(file_path).sheet_names),
            "row_count": len(df),
            "column_count": len(df.columns),
            "columns": list(df.columns),
            "sample_data": {}
        }
        
        if show_details:
            print(f"    工作表数: {file_info['sheet_count']}")
            print(f"    行数: {file_info['row_count']}")
            print(f"    列数: {file_info['column_count']}")
            print(f"    列名: {file_info['columns']}")
        
        # 获取每列的数据类型和示例值
        for col in df.columns[:10]:  # 最多分析10列
            col_data = df[col].dropna()
            if len(col_data) > 0:
                sample_value = col_data.iloc[0]
                data_type = type(sample_value).__name__
                
                file_info["sample_data"][col] = {
                    "data_type": data_type,
                    "sample_value": str(sample_value)[:100],  # 限制长度
                    "unique_count": df[col].nunique() if len(col_data) > 0 else 0,
                    "null_count": df[col].isnull().sum()
                }
        
        # 显示关键字段（尝试识别）
        if show_details:
            identify_key_fields(df, file_path)
        
        return file_info
        
    except Exception as e:
        print(f"    分析文件时出错: {e}")
        return None

def identify_key_fields(df, file_path):
    """识别关键字段"""
    filename = os.path.basename(file_path)
    
    print(f"    关键字段识别:")
    
    # 根据文件名猜测关键字段
    field_patterns = {
        "客户": ["客户", "名称", "公司", "企业", "client", "customer"],
        "金额": ["金额", "价格", "费用", "收入", "amount", "price", "fee"],
        "日期": ["日期", "时间", "day", "date", "time", "创建时间", "更新时间"],
        "合同": ["合同", "协议", "contract", "agreement"],
        "产品": ["产品", "服务", "商品", "product", "service", "item"],
        "状态": ["状态", "status", "state", "阶段"],
        "模块": ["模块", "module", "功能", "function"],
        "工单": ["工单", "问题", "故障", "ticket", "issue", "problem"],
        "负责人": ["负责人", "经办人", "处理人", "owner", "assignee"],
        "部门": ["部门", "科室", "单位", "department", "division"]
    }
    
    identified_fields = {}
    
    for field_type, patterns in field_patterns.items():
        for col in df.columns:
            col_lower = str(col).lower()
            for pattern in patterns:
                if pattern.lower() in col_lower:
                    if field_type not in identified_fields:
                        identified_fields[field_type] = []
                    identified_fields[field_type].append(col)
                    break
    
    # 显示识别结果
    for field_type, columns in identified_fields.items():
        if columns:
            print(f"      {field_type}: {', '.join(columns)}")
    
    # 特别关注购买模块字段（根据用户要求）
    if "客户主数据" in filename:
        print(f"    寻找'购买模块'字段...")
        purchase_module_cols = [col for col in df.columns if "购买" in str(col) or "模块" in str(col) or "module" in str(col).lower()]
        if purchase_module_cols:
            print(f"      找到相关字段: {purchase_module_cols}")
            # 显示示例值
            for col in purchase_module_cols[:3]:  # 最多显示3个
                if len(df[col].dropna()) > 0:
                    sample = df[col].dropna().iloc[0]
                    print(f"        {col}: {str(sample)[:50]}...")
        else:
            print(f"      未找到明确的购买模块字段")

def generate_analysis_report(analysis_results):
    """生成数据分析报告"""
    print(f"\n" + "=" * 60)
    print(f"数据分析报告")
    print(f"=" * 60)
    
    report = []
    report.append("# CBD客户数据源分析报告")
    report.append(f"**分析时间**: {analysis_results['分析时间']}")
    report.append("")
    
    # 1. 客户主数据表
    report.append("## 1. 客户主数据表")
    customer_data = analysis_results["客户主数据"]
    if customer_data:
        report.append(f"- **文件**: {os.path.basename(customer_data['file_path'])}")
        report.append(f"- **大小**: {customer_data['file_size']:,} 字节")
        report.append(f"- **工作表**: {customer_data['sheet_count']} 个")
        report.append(f"- **数据量**: {customer_data['row_count']} 行 × {customer_data['column_count']} 列")
        report.append("")
        report.append("### 字段列表")
        for i, col in enumerate(customer_data['columns'], 1):
            report.append(f"{i}. `{col}`")
        
        # 特别关注购买模块字段
        report.append("")
        report.append("### 购买模块字段分析")
        purchase_cols = [col for col in customer_data['columns'] 
                        if any(keyword in str(col) for keyword in ['购买', '模块', 'Module', 'module'])]
        if purchase_cols:
            report.append("找到以下可能相关的字段:")
            for col in purchase_cols:
                if col in customer_data['sample_data']:
                    sample = customer_data['sample_data'][col]['sample_value']
                    report.append(f"- `{col}`: {sample}")
        else:
            report.append("未找到明确的'购买模块'字段")
    else:
        report.append("*未找到客户主数据表*")
    report.append("")
    
    # 2. 固定金额合同台账
    report.append("## 2. 固定金额合同台账")
    fixed_data = analysis_results["固定金额合同台账"]
    if fixed_data:
        report.append(f"- **文件**: {os.path.basename(fixed_data['file_path'])}")
        report.append(f"- **数据量**: {fixed_data['row_count']} 行 × {fixed_data['column_count']} 列")
        report.append("")
        report.append("### 关键字段")
        # 识别关键字段
        key_fields = []
        for col in fixed_data['columns']:
            col_str = str(col)
            if any(keyword in col_str for keyword in ['金额', '合同', '实施', '固定', '费用', 'Amount', 'Contract']):
                key_fields.append(col)
        
        for col in key_fields[:10]:  # 最多显示10个
            report.append(f"- `{col}`")
    else:
        report.append("*未找到固定金额合同台账*")
    report.append("")
    
    # 3. 人天框架合同台帐
    report.append("## 3. 人天框架合同台帐")
    manday_data = analysis_results["人天框架合同台帐"]
    if manday_data:
        report.append(f"- **文件**: {os.path.basename(manday_data['file_path'])}")
        report.append(f"- **数据量**: {manday_data['row_count']} 行 × {manday_data['column_count']} 列")
        report.append("")
        report.append("### 关键字段")
        # 识别关键字段
        key_fields = []
        for col in manday_data['columns']:
            col_str = str(col)
            if any(keyword in col_str for keyword in ['人天', '框架', '合同', '实施', 'Manday', 'man-day', 'Frame']):
                key_fields.append(col)
        
        for col in key_fields[:10]:  # 最多显示10个
            report.append(f"- `{col}`")
    else:
        report.append("*未找到人天框架合同台帐*")
    report.append("")
    
    # 4. 运维工单明细
    report.append("## 4. 运维工单明细")
    operation_data = analysis_results["运维工单明细"]
    if operation_data:
        report.append(f"- **文件**: {os.path.basename(operation_data['file_path'])}")
        report.append(f"- **数据量**: {operation_data['row_count']} 行 × {operation_data['column_count']} 列")
        report.append("")
        report.append("### 关键字段")
        # 识别关键字段
        key_fields = []
        for col in operation_data['columns']:
            col_str = str(col)
            if any(keyword in col_str for keyword in ['工单', '问题', '状态', '处理', 'Ticket', 'Issue', 'Status']):
                key_fields.append(col)
        
        for col in key_fields[:10]:  # 最多显示10个
            report.append(f"- `{col}`")
    else:
        report.append("*未找到运维工单明细表*")
    report.append("")
    
    # 5. 数据完整性评估
    report.append("## 5. 数据完整性评估")
    found_tables = sum(1 for k, v in analysis_results.items() 
                      if k != "分析时间" and v is not None)
    total_tables = 4  # 客户主数据、固定金额、人天框架、运维工单
    
    report.append(f"- **已找到表**: {found_tables}/{total_tables}")
    report.append(f"- **完整度**: {found_tables/total_tables*100:.0f}%")
    report.append("")
    
    # 6. 下一步建议
    report.append("## 6. 下一步建议")
    report.append("1. **完善缺失表**: 补充未找到的数据表")
    report.append("2. **字段映射**: 根据分析结果建立字段映射关系")
    report.append("3. **数据清洗**: 清理数据中的空值和异常值")
    report.append("4. **模板创建**: 创建标准数据模板确保一致性")
    report.append("5. **自动化集成**: 将分析结果集成到报告生成器中")
    
    # 保存报告
    report_file = r"C:\Users\mingh\.openclaw\workspace\data_source_analysis_report.md"
    with open(report_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(report))
    
    print(f"分析报告已保存到: {report_file}")
    
    # 在控制台显示摘要
    print(f"\n数据源分析摘要:")
    print(f"- 客户主数据表: {'✅ 已找到' if customer_data else '❌ 未找到'}")
    print(f"- 固定金额合同台账: {'✅ 已找到' if fixed_data else '❌ 未找到'}")
    print(f"- 人天框架合同台帐: {'✅ 已找到' if manday_data else '❌ 未找到'}")
    print(f"- 运维工单明细: {'✅ 已找到' if operation_data else '❌ 未找到'}")
    print(f"- 数据完整度: {found_tables}/{total_tables} ({found_tables/total_tables*100:.0f}%)")