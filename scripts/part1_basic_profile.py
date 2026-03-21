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
        # 字段映射：报表字段名 -> Excel列名
        self.field_mapping = {
            # 1.1 基本信息
            '客户简称': '客户简称',
            '客户全称': '真实服务对象',
            '计费ARR': '计费ARR',
            '服务阶段': '服务阶段',
            '客户状态': '客户状态',
            '所属区域': '客户所属区域',
            # 1.2 业务概况
            '行业': '所属行业',
            '主要产品': '主要产品',
            '营收规模': '营收规模',
            '总部地点': '客户位置',
            # 1.3 购买模块
            '购买模块': '购买模块',
            # 1.4 项目团队
            '客户成功经理': '客户成功经理',
            '项目经理': '项目经理',
            '运维经理': '运维主责',
            '交付顾问': '交付经理',
            '销售经理': '销售',
            # 1.5 决策地图
            '采购总': '采购高层（姓名-职位）',
            '采购经理': '采购中层（姓名-职位）',
            'IT总': 'IT高层（姓名-职位）',
            'IT经理': 'IT中层（姓名-职位）',
            '对接人': '客户对接人',
            '决策链': '决策链',
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
            
            # 1.4 项目团队
            content += self._generate_project_team(row)
            
            # 1.5 决策地图
            content += self._generate_decision_map(row)
            
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
            excel_col = self.field_mapping.get(field, field)
            value = row.get(excel_col, 'N/A')
            
            if pd.isna(value) or str(value).strip() == '':
                value = 'N/A'
            elif field == '计费ARR' and isinstance(value, (int, float)):
                value = f"{value:,.0f} 元"
            
            content += f"| {field} | {value} |\n"
        
        content += "\n"
        return content
    
    def _generate_business_profile(self, row):
        """生成业务概况"""
        content = "### 1.2 业务概况\n\n"
        
        # 使用表格格式
        content += "| 指标 | 内容 |\n"
        content += "|------|------|\n"
        
        business_fields = ['行业', '主要产品', '营收规模', '总部地点']
        
        for field in business_fields:
            display_name = self.field_mapping.get(field, field)
            value = row.get(display_name, 'N/A')
            
            if pd.isna(value) or str(value).strip() == '':
                value = 'N/A'
            
            content += f"| {display_name} | {value} |\n"
        
        content += "\n"
        return content
    
    def _generate_purchase_info(self, row):
        """生成购买信息"""
        content = "### 1.3 购买模块\n\n"
        
        # 只显示购买模块，不使用表格
        purchase_module = row.get(self.field_mapping.get('购买模块', '购买模块'), None)
        
        if pd.notna(purchase_module) and str(purchase_module).strip():
            content += f"- {purchase_module}\n"
        else:
            content += "- 暂无数据\n"
        
        content += "\n"
        return content
    
    def _generate_project_team(self, row):
        """生成项目团队信息"""
        content = "### 1.4 项目团队\n\n"
        
        # 使用表格格式
        content += "| 角色 | 姓名 |\n"
        content += "|------|------|\n"
        
        team_fields = ['客户成功经理', '项目经理', '运维经理', '交付顾问', '销售经理']
        
        for role in team_fields:
            excel_col = self.field_mapping.get(role, role)
            value = row.get(excel_col, 'N/A')
            
            if pd.isna(value) or str(value).strip() == '':
                value = '-'
            
            content += f"| {role} | {value} |\n"
        
        content += "\n"
        return content
    
    def _generate_decision_map(self, row):
        """生成决策地图"""
        content = "### 1.5 决策地图\n\n"
        
        # 使用表格格式
        content += "| 角色 | 信息 |\n"
        content += "|------|------|\n"
        
        decision_fields = ['采购总', '采购经理', 'IT总', 'IT经理', '对接人', '决策链']
        
        for role in decision_fields:
            excel_col = self.field_mapping.get(role, role)
            value = row.get(excel_col, 'N/A')
            
            if pd.isna(value) or str(value).strip() == '':
                value = '-'
            
            # 清理敏感信息（换行、电话、邮箱）
            value = self._clean_decision_value(value)
            
            content += f"| {role} | {value} |\n"
        
        content += "\n"
        return content
    
    def _clean_decision_value(self, value):
        """清理决策地图中的敏感信息"""
        value = str(value)
        import re
        # 替换换行符为空格
        value = value.replace('\n', ' ').replace('\r', ' ')
        # 移除手机号（1开头的11位数字）→ 替换为脱敏
        value = re.sub(r'1\d{10}', '***', value)
        # 移除邮箱 → 替换为脱敏
        value = re.sub(r'[\w\.-]+@[\w\.-]+', '***', value)
        return value
    
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