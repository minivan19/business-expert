#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Part 1: 客户基础档案分析模块
基于客户主数据生成详细的客户基础档案
"""

import pandas as pd
import logging
from datetime import datetime

logger = logging.getLogger(__name__)


class BasicProfileAnalyzer:
    """客户基础档案分析器"""
    
    def __init__(self):
        self.field_mapping = {
            '客户简称': '客户简称',
            '真实服务对象': '客户全称',
            '计费ARR': '计费ARR',
            '服务阶段': '服务阶段',
            '客户状态': '客户状态',
            '客户所属区域': '所属区域',
            '行业': '行业',
            '主要产品': '主要产品',
            '营收规模': '营收规模',
            '首次购买时间': '首次购买时间',
            '最近购买时间': '最近购买时间',
            '累计购买金额': '累计购买金额',
            '购买次数': '购买次数'
        }
    
    def analyze(self, df_basic):
        """
        分析客户基础档案
        
        Args:
            df_basic: 客户主数据DataFrame
            
        Returns:
            str: Markdown格式的分析报告
        """
        if df_basic is None or df_basic.empty:
            logger.warning("客户主数据为空")
            return "## 1. 客户基础档案\n\n暂无数据\n"
        
        logger.info(f"开始分析客户基础档案，数据行数: {len(df_basic)}")
        
        try:
            # 获取第一行数据（通常只有一行）
            if len(df_basic) > 0:
                row = df_basic.iloc[0]
            else:
                return "## 1. 客户基础档案\n\n数据为空\n"
            
            content = "## 1. 客户基础档案\n\n"
            
            # 1.1 基本信息
            content += self._generate_basic_info(row)
            
            # 1.2 业务概况
            content += self._generate_business_profile(row)
            
            # 1.3 购买信息
            content += self._generate_purchase_info(row)
            
            # 1.4 联系信息（如果存在）
            content += self._generate_contact_info(row)
            
            # 1.5 决策信息（如果存在）
            content += self._generate_decision_info(row)
            
            logger.info("客户基础档案分析完成")
            return content
            
        except Exception as e:
            error_msg = f"客户基础档案分析失败: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return f"## 1. 客户基础档案\n\n分析失败: {error_msg}\n"
    
    def _generate_basic_info(self, row):
        """生成基本信息"""
        content = "### 1.1 基本信息\n\n"
        
        # 使用表格格式
        content += "| 指标 | 内容 |\n"
        content += "|------|------|\n"
        
        basic_fields = ['客户简称', '客户全称', '计费ARR', '服务阶段', '客户状态', '所属区域']
        
        for field in basic_fields:
            display_name = self.field_mapping.get(field, field)
            if field in ['客户全称', '真实服务对象']:
                # 处理客户全称字段
                value = row.get('真实服务对象', row.get('客户全称', 'N/A'))
            elif field == '行业':
                # 尝试获取行业信息
                value = self._get_industry_value(row)
            else:
                value = row.get(field, 'N/A')
            
            # 格式化ARR金额
            if field == '计费ARR' and isinstance(value, (int, float)):
                value = f"{value:,.0f} 元"
            
            content += f"| {display_name} | {value} |\n"
        
        content += "\n"
        return content
    
    def _generate_business_profile(self, row):
        """生成业务概况"""
        content = "### 1.2 业务概况\n\n"
        
        # 使用表格格式
        content += "| 指标 | 内容 |\n"
        content += "|------|------|\n"
        
        business_fields = ['行业', '主要产品', '营收规模']
        
        for field in business_fields:
            display_name = self.field_mapping.get(field, field)
            
            if field == '行业':
                value = self._get_industry_value(row)
            else:
                value = row.get(field, 'N/A')
            
            content += f"| {display_name} | {value} |\n"
        
        # 添加其他业务信息（如果存在）
        additional_fields = ['员工规模', '总部地点', '上市情况', '成立时间']
        for field in additional_fields:
            if field in row and pd.notna(row[field]) and str(row[field]).strip():
                value = row[field]
                content += f"| {field} | {value} |\n"
        
        content += "\n"
        return content
    
    def _generate_purchase_info(self, row):
        """生成购买信息"""
        content = "### 1.3 购买信息\n\n"
        
        # 使用表格格式
        content += "| 指标 | 内容 |\n"
        content += "|------|------|\n"
        
        purchase_fields = ['首次购买时间', '最近购买时间', '累计购买金额', '购买次数']
        
        for field in purchase_fields:
            display_name = self.field_mapping.get(field, field)
            value = row.get(field, 'N/A')
            
            # 格式化时间和金额
            if '时间' in field and pd.notna(value):
                try:
                    if isinstance(value, pd.Timestamp):
                        value = value.strftime('%Y-%m-%d')
                    else:
                        # 尝试解析为日期
                        value = pd.to_datetime(value).strftime('%Y-%m-%d')
                except:
                    pass
            
            if '金额' in field and isinstance(value, (int, float)):
                value = f"{value:,.0f} 元"
            
            if field == '购买次数' and isinstance(value, (int, float)):
                value = f"{value:.0f} 次"
                # 计算平均购买金额
                if '累计购买金额' in row and pd.notna(row['累计购买金额']) and row['累计购买金额'] > 0:
                    avg_amount = row['累计购买金额'] / value if value > 0 else 0
                    content += f"| 平均购买金额 | {avg_amount:,.0f} 元 |\n"
            
            content += f"| {display_name} | {value} |\n"
        
        # 计算购买频率（如果数据足够）
        if '首次购买时间' in row and '最近购买时间' in row and '购买次数' in row:
            try:
                first_purchase = pd.to_datetime(row['首次购买时间'])
                last_purchase = pd.to_datetime(row['最近购买时间'])
                purchase_count = float(row['购买次数']) if pd.notna(row['购买次数']) else 0
                
                if purchase_count > 1 and pd.notna(first_purchase) and pd.notna(last_purchase):
                    days_diff = (last_purchase - first_purchase).days
                    if days_diff > 0:
                        frequency = days_diff / (purchase_count - 1) if purchase_count > 1 else 0
                        content += f"| 平均购买频率 | 每 {frequency:.0f} 天 |\n"
            except:
                pass
        
        content += "\n"
        return content
    
    def _generate_contact_info(self, row):
        """生成联系信息"""
        content = ""
        
        # 检查是否存在联系信息字段
        contact_fields = ['联系人', '联系电话', '联系邮箱', '联系地址']
        has_contact_info = any(field in row and pd.notna(row[field]) and str(row[field]).strip() for field in contact_fields)
        
        if has_contact_info:
            content += "### 1.4 联系信息\n\n"
            content += "| 指标 | 内容 |\n"
            content += "|------|------|\n"
            
            for field in contact_fields:
                if field in row and pd.notna(row[field]) and str(row[field]).strip():
                    value = row[field]
                    content += f"| {field} | {value} |\n"
            
            content += "\n"
        
        return content
    
    def _generate_decision_info(self, row):
        """生成决策信息"""
        content = ""
        
        # 检查是否存在决策信息字段
        decision_fields = ['决策人', '决策人职位', 'IT负责人', '采购负责人']
        has_decision_info = any(field in row and pd.notna(row[field]) and str(row[field]).strip() for field in decision_fields)
        
        if has_decision_info:
            content += "### 1.5 决策信息\n\n"
            content += "| 角色 | 姓名 | 职位 |\n"
            content += "|------|------|------|\n"
            
            # 决策人
            if '决策人' in row and pd.notna(row['决策人']) and str(row['决策人']).strip():
                name = row['决策人']
                position = row.get('决策人职位', '')
                content += f"| 决策人 | {name} | {position} |\n"
            
            # IT负责人
            if 'IT负责人' in row and pd.notna(row['IT负责人']) and str(row['IT负责人']).strip():
                name = row['IT负责人']
                content += f"| IT负责人 | {name} | |\n"
            
            # 采购负责人
            if '采购负责人' in row and pd.notna(row['采购负责人']) and str(row['采购负责人']).strip():
                name = row['采购负责人']
                content += f"| 采购负责人 | {name} | |\n"
            
            content += "\n"
        
        return content
    
    def _get_industry_value(self, row):
        """获取行业信息"""
        # 尝试多个可能的字段
        industry_fields = ['行业_clean', '行业', '所属行业', 'industry']
        
        for field in industry_fields:
            if field in row and pd.notna(row[field]) and str(row[field]).strip():
                return row[field]
        
        # 如果都没有，尝试从第18列获取（原始数据中的位置）
        if hasattr(row, 'iloc') and len(row) > 18:
            try:
                industry_value = row.iloc[18]
                if pd.notna(industry_value) and str(industry_value).strip():
                    return industry_value
            except:
                pass
        
        return 'N/A'
    
    def extract_key_metrics(self, row):
        """提取关键指标"""
        metrics = {}
        
        # 基础指标
        if '计费ARR' in row and pd.notna(row['计费ARR']):
            metrics['计费ARR'] = float(row['计费ARR'])
        
        if '累计购买金额' in row and pd.notna(row['累计购买金额']):
            metrics['累计购买金额'] = float(row['累计购买金额'])
        
        if '购买次数' in row and pd.notna(row['购买次数']):
            metrics['购买次数'] = float(row['购买次数'])
        
        # 计算平均购买金额
        if '累计购买金额' in metrics and '购买次数' in metrics and metrics['购买次数'] > 0:
            metrics['平均购买金额'] = metrics['累计购买金额'] / metrics['购买次数']
        
        # 服务阶段评分
        service_stage = row.get('服务阶段', '')
        stage_scores = {
            '实施中': 1,
            '运维中': 2,
            '成熟期': 3,
            '衰退期': 0
        }
        metrics['服务阶段评分'] = stage_scores.get(service_stage, 0)
        
        # 客户状态评分
        customer_status = row.get('客户状态', '')
        status_scores = {
            '绿色': 3,
            '黄色': 2,
            '红色': 1,
            '黑色': 0
        }
        metrics['客户状态评分'] = status_scores.get(customer_status, 0)
        
        return metrics


# 测试代码
if __name__ == "__main__":
    # 设置日志
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    
    # 测试数据
    test_data = {
        '客户简称': '测试客户',
        '真实服务对象': '测试有限公司',
        '计费ARR': 100000,
        '服务阶段': '运维中',
        '客户状态': '绿色',
        '客户所属区域': '华东区',
        '行业': '科技',
        '主要产品': '软件服务',
        '营收规模': '10-50亿元',
        '首次购买时间': '2023-01-15',
        '最近购买时间': '2025-01-20',
        '累计购买金额': 300000,
        '购买次数': 3,
        '联系人': '张三',
        '联系电话': '13800138000',
        '决策人': '李四',
        '决策人职位': 'CEO'
    }
    
    df_test = pd.DataFrame([test_data])
    
    analyzer = BasicProfileAnalyzer()
    report = analyzer.analyze(df_test)
    
    print("测试输出:")
    print(report)
    
    metrics = analyzer.extract_key_metrics(df_test.iloc[0])
    print("\n关键指标:")
    for key, value in metrics.items():
        print(f"{key}: {value}")