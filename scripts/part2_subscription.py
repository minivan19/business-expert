#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Part 2: 订阅续约与续费分析模块
分析客户的订阅情况、续约趋势和收款进度
"""

import pandas as pd
import numpy as np
import logging
from datetime import datetime, timedelta

logger = logging.getLogger(__name__)


class SubscriptionAnalyzer:
    """订阅续约与续费分析器"""
    
    def __init__(self):
        # 根据实际数据列名映射
        self.column_mapping = {
            'subscription': {
                '合同号': '合同编号',  # 实际列名可能是'合同编号'或'合同号'
                '产品名称': '产品名称',
                '开始时间': '合同开始时间',
                '结束时间': '合同结束时间', 
                '金额': '年订阅费金额',
                '用户数': '每用户数量'
            },
            'collection': {
                '合同号': '合同编号',
                '应收金额': '计划收款金额',
                '实收金额': '实际收款金额',
                '应收日期': '计划收款日期',
                '实收日期': '实际收款日期',
                '状态': '收款状态'
            }
        }
    
    def analyze(self, df_subscription, df_collection):
        """
        分析订阅续约与续费情况
        
        Args:
            df_subscription: 订阅数据DataFrame
            df_collection: 收款数据DataFrame
            
        Returns:
            str: Markdown格式的分析报告
        """
        logger.info("开始分析订阅续约与续费情况")
        
        content = "## 2. 订阅续约与续费\n\n"
        
        try:
            # 2.1 订阅统计
            content += self._generate_subscription_stats(df_subscription)
            
            # 2.2 续约统计
            content += self._generate_renewal_stats(df_subscription)
            
            # 2.3 收款统计
            content += self._generate_collection_stats(df_collection)
            
            # 2.4 智能分析（占位符，将由LLM填充）
            content += self._generate_analysis_placeholder(df_subscription, df_collection)
            
            logger.info("订阅续约与续费分析完成")
            return content
            
        except Exception as e:
            error_msg = f"订阅续约与续费分析失败: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return content + f"分析失败: {error_msg}\n"
    
    def _generate_subscription_stats(self, df):
        """生成订阅统计"""
        content = "### 2.1 订阅统计\n\n"
        
        if df is None or df.empty:
            content += "暂无订阅数据\n\n"
            return content
        
        # 基础统计
        total_subscriptions = len(df)
        
        # 尝试查找金额列
        amount_column = None
        possible_amount_columns = ['年订阅费金额', '总订阅金额', '金额', '合同金额', '合同总金额']
        for col in possible_amount_columns:
            if col in df.columns and pd.api.types.is_numeric_dtype(df[col]):
                amount_column = col
                break
        
        total_amount = 0
        if amount_column:
            total_amount = df[amount_column].sum()
            content += f"- 总订阅合同数: **{total_subscriptions}** 个\n"
            content += f"- 总订阅金额: **{total_amount:,.0f}** 元\n"
        else:
            content += f"- 总订阅合同数: **{total_subscriptions}** 个\n"
            content += f"- 总订阅金额: 无法计算（金额列未找到）\n"
        
        # 按产品统计
        product_column = None
        possible_product_columns = ['产品名称', '产品', '服务名称', '产品描述']
        for col in possible_product_columns:
            if col in df.columns:
                product_column = col
                break
        
        if product_column:
            try:
                product_stats = df.groupby(product_column).size().reset_index(name='合同数')
                
                if amount_column:
                    amount_stats = df.groupby(product_column)[amount_column].sum().reset_index(name='总金额')
                    product_stats = pd.merge(product_stats, amount_stats, on=product_column)
                
                content += "\n#### 按产品统计\n"
                content += "| 产品名称 | 合同数 |"
                if amount_column:
                    content += " 总金额(元) |"
                content += "\n"
                
                content += "|----------|--------|"
                if amount_column:
                    content += "------------|"
                content += "\n"
                
                for _, row in product_stats.iterrows():
                    content += f"| {row[product_column]} | {row['合同数']} |"
                    if amount_column and '总金额' in row:
                        content += f" {row['总金额']:,.0f} |"
                    elif amount_column:
                        content += " - |"
                    content += "\n"
            except Exception as e:
                logger.warning(f"按产品统计失败: {str(e)}")
        
        # 按年度统计
        date_column = None
        possible_date_columns = ['合同开始时间', '开始时间', '生效时间', '签约时间', '合同签订时间']
        for col in possible_date_columns:
            if col in df.columns:
                date_column = col
                break
        
        if date_column:
            try:
                df['开始年份'] = pd.to_datetime(df[date_column], errors='coerce').dt.year
                yearly_stats = df.groupby('开始年份').size().reset_index(name='合同数')
                
                if amount_column:
                    yearly_amount = df.groupby('开始年份')[amount_column].sum().reset_index(name='总金额')
                    yearly_stats = pd.merge(yearly_stats, yearly_amount, on='开始年份')
                
                yearly_stats = yearly_stats.sort_values('开始年份')
                
                content += "\n#### 按年度统计\n"
                content += "| 年度 | 合同数 |"
                if amount_column:
                    content += " 总金额(元) | 平均金额(元) |"
                content += "\n"
                
                content += "|------|--------|"
                if amount_column:
                    content += "------------|-------------|"
                content += "\n"
                
                for _, row in yearly_stats.iterrows():
                    year = int(row['开始年份']) if pd.notna(row['开始年份']) else '未知'
                    content += f"| {year} | {row['合同数']} |"
                    if amount_column and '总金额' in row and pd.notna(row['总金额']):
                        avg_amount = row['总金额'] / row['合同数'] if row['合同数'] > 0 else 0
                        content += f" {row['总金额']:,.0f} | {avg_amount:,.0f} |"
                    elif amount_column:
                        content += " - | - |"
                    content += "\n"
            except Exception as e:
                logger.warning(f"按年度统计失败: {str(e)}")
        
        # 显示关键字段信息
        content += "\n#### 数据字段信息\n"
        content += f"- 数据行数: {len(df)}\n"
        content += f"- 数据列数: {len(df.columns)}\n"
        
        # 显示前几个列名
        if len(df.columns) > 0:
            content += f"- 前5个字段: {', '.join(df.columns[:5])}\n"
        
        content += "\n"
        return content
    
    def _generate_renewal_stats(self, df):
        """生成续约统计"""
        content = "### 2.2 续约统计\n\n"
        
        if df is None or df.empty or '合同号' not in df.columns or '开始时间' not in df.columns:
            content += "暂无续约数据\n\n"
            return content
        
        try:
            # 按合同号分组，查看续约情况
            df_sorted = df.sort_values(['合同号', '开始时间'])
            df_sorted['合同序号'] = df_sorted.groupby('合同号').cumcount() + 1
            
            # 识别续约合同
            renewal_stats = df_sorted[df_sorted['合同序号'] > 1]
            total_renewals = len(renewal_stats)
            
            # 计算续约率
            unique_contracts = df_sorted['合同号'].nunique()
            renewed_contracts = renewal_stats['合同号'].nunique()
            renewal_rate = (renewed_contracts / unique_contracts * 100) if unique_contracts > 0 else 0
            
            content += f"- 总续约次数: **{total_renewals}** 次\n"
            content += f"- 续约合同数: **{renewed_contracts}** 个\n"
            content += f"- 总合同数: **{unique_contracts}** 个\n"
            content += f"- 续约率: **{renewal_rate:.1f}%**\n"
            
            # 续约价格变化分析
            if total_renewals > 0 and '金额' in renewal_stats.columns and pd.api.types.is_numeric_dtype(renewal_stats['金额']):
                # 获取每个合同的首次和最近价格
                contract_prices = df_sorted.groupby('合同号').agg({
                    '金额': ['first', 'last'],
                    '开始时间': ['first', 'last']
                }).reset_index()
                
                contract_prices.columns = ['合同号', '首次金额', '最近金额', '首次时间', '最近时间']
                
                # 计算价格变化
                contract_prices['价格变化'] = contract_prices['最近金额'] - contract_prices['首次金额']
                contract_prices['价格变化率'] = (contract_prices['价格变化'] / contract_prices['首次金额'] * 100) if contract_prices['首次金额'] > 0 else 0
                
                price_increase = len(contract_prices[contract_prices['价格变化'] > 0])
                price_decrease = len(contract_prices[contract_prices['价格变化'] < 0])
                price_stable = len(contract_prices[contract_prices['价格变化'] == 0])
                
                content += "\n#### 续约价格变化\n"
                content += f"- 价格上涨: **{price_increase}** 个合同\n"
                content += f"- 价格下降: **{price_decrease}** 个合同\n"
                content += f"- 价格不变: **{price_stable}** 个合同\n"
                
                if price_increase + price_decrease > 0:
                    avg_change_rate = contract_prices['价格变化率'].mean()
                    content += f"- 平均价格变化率: **{avg_change_rate:+.1f}%**\n"
            
            # 续约时间间隔分析
            if total_renewals > 0 and '开始时间' in df_sorted.columns:
                try:
                    df_sorted['开始时间_dt'] = pd.to_datetime(df_sorted['开始时间'])
                    
                    # 计算续约间隔
                    renewal_intervals = []
                    for contract in df_sorted['合同号'].unique():
                        contract_data = df_sorted[df_sorted['合同号'] == contract].sort_values('开始时间_dt')
                        if len(contract_data) > 1:
                            for i in range(1, len(contract_data)):
                                interval = (contract_data.iloc[i]['开始时间_dt'] - contract_data.iloc[i-1]['开始时间_dt']).days
                                renewal_intervals.append(interval)
                    
                    if renewal_intervals:
                        avg_interval = np.mean(renewal_intervals)
                        content += f"- 平均续约间隔: **{avg_interval:.0f}** 天\n"
                except Exception as e:
                    logger.warning(f"续约时间间隔分析失败: {str(e)}")
            
        except Exception as e:
            logger.warning(f"续约统计生成失败: {str(e)}")
            content += f"续约统计生成失败: {str(e)}\n"
        
        content += "\n"
        return content
    
    def _generate_collection_stats(self, df):
        """生成收款统计"""
        content = "### 2.3 收款统计\n\n"
        
        if df is None or df.empty:
            content += "暂无收款数据\n\n"
            return content
        
        try:
            # 基础统计
            total_records = len(df)
            content += f"- 总收款记录数: **{total_records}** 条\n"
            
            # 尝试查找金额列
            due_column = None
            received_column = None
            possible_due_columns = ['计划收款金额', '应收金额', '应收款', '应付款']
            possible_received_columns = ['实际收款金额', '实收金额', '已收款', '已付款']
            
            for col in possible_due_columns:
                if col in df.columns and pd.api.types.is_numeric_dtype(df[col]):
                    due_column = col
                    break
            
            # 如果没找到，尝试搜索包含"计划"或"应收"的列
            if not due_column:
                for col in df.columns:
                    col_str = str(col)
                    if ('计划' in col_str or '应收' in col_str) and pd.api.types.is_numeric_dtype(df[col]):
                        due_column = col
                        print(f"找到应收金额列: {col}")
                        break
            
            for col in possible_received_columns:
                if col in df.columns and pd.api.types.is_numeric_dtype(df[col]):
                    received_column = col
                    break
            
            # 如果没找到，尝试搜索包含"收款"但不包含"计划"的列
            if not received_column:
                for col in df.columns:
                    col_str = str(col)
                    if (('收款' in col_str or '实收' in col_str or '已收' in col_str) and 
                        '计划' not in col_str and 
                        pd.api.types.is_numeric_dtype(df[col])):
                        received_column = col
                        logger.info(f"找到实收金额列: {col}")
                        break
            
            total_due = 0
            total_received = 0
            
            if due_column:
                total_due = df[due_column].sum()
                content += f"- 总应收金额: **{total_due:,.0f}** 元\n"
            else:
                content += f"- 总应收金额: 无法计算（应收金额列未找到）\n"
            
            if received_column:
                total_received = df[received_column].sum()
                content += f"- 总实收金额: **{total_received:,.0f}** 元\n"
                
                # 计算收款率
                if total_due > 0:
                    collection_rate = (total_received / total_due * 100)
                    content += f"- 总体收款率: **{collection_rate:.1f}%**\n"
            else:
                content += f"- 总实收金额: 无法计算（实收金额列未找到）\n"
            
            # 按状态统计
            status_column = None
            possible_status_columns = ['收款状态', '状态', '付款状态', '结算状态']
            for col in possible_status_columns:
                if col in df.columns:
                    status_column = col
                    break
            
            if status_column:
                try:
                    status_stats = df[status_column].value_counts()
                    content += "\n#### 按状态统计\n"
                    content += "| 状态 | 记录数 | 占比 |\n"
                    content += "|------|--------|------|\n"
                    
                    for status, count in status_stats.items():
                        percentage = (count / total_records * 100)
                        content += f"| {status} | {count} | {percentage:.1f}% |\n"
                except Exception as e:
                    logger.warning(f"按状态统计失败: {str(e)}")
            
            # 日期分析
            due_date_column = None
            received_date_column = None
            possible_due_date_columns = ['计划收款日期', '应收日期', '应付款日期']
            possible_received_date_columns = ['实际收款日期', '实收日期', '付款日期']
            
            for col in possible_due_date_columns:
                if col in df.columns:
                    due_date_column = col
                    break
            
            for col in possible_received_date_columns:
                if col in df.columns:
                    received_date_column = col
                    break
            
            # 月度收款趋势
            if received_date_column and received_column:
                try:
                    df['实收月份'] = pd.to_datetime(df[received_date_column], errors='coerce').dt.to_period('M')
                    monthly_stats = df.groupby('实收月份')[received_column].sum().reset_index()
                    monthly_stats = monthly_stats.sort_values('实收月份')
                    
                    if len(monthly_stats) > 0:
                        content += "\n#### 月度收款趋势\n"
                        content += "| 月份 | 实收金额(元) |\n"
                        content += "|------|-------------|\n"
                        
                        for _, row in monthly_stats.iterrows():
                            if pd.notna(row['实收月份']):
                                month_str = row['实收月份'].strftime('%Y-%m')
                                amount = row[received_column]
                                content += f"| {month_str} | {amount:,.0f} |\n"
                except Exception as e:
                    logger.warning(f"月度收款趋势分析失败: {str(e)}")
            
            # 显示关键字段信息
            content += "\n#### 数据字段信息\n"
            content += f"- 数据行数: {len(df)}\n"
            content += f"- 数据列数: {len(df.columns)}\n"
            
            # 显示前几个列名
            if len(df.columns) > 0:
                content += f"- 前5个字段: {', '.join(df.columns[:5])}\n"
            
        except Exception as e:
            logger.warning(f"收款统计生成失败: {str(e)}")
            content += f"收款统计生成失败: {str(e)}\n"
        
        content += "\n"
        return content
    
    def _generate_analysis_placeholder(self, df_subscription, df_collection):
        """生成智能分析占位符"""
        content = "### 2.4 智能分析\n\n"
        content += "*此部分将由LLM进行智能分析，包括：*\n"
        content += "1) 订阅费用时间阶段变化趋势分析\n"
        content += "2) 续约/降价原因分析\n"
        content += "3) 收款进度评估和逾期风险分析\n"
        content += "4) 续费策略建议\n\n"
        
        # 提供数据摘要供LLM使用
        data_summary = self._prepare_data_summary(df_subscription, df_collection)
        content += f"*数据摘要（供LLM参考）:*\n```\n{data_summary}\n```\n\n"
        
        return content
    
    def _prepare_data_summary(self, df_subscription, df_collection):
        """准备数据摘要供LLM分析使用"""
        summary = []
        
        # 订阅数据摘要
        if df_subscription is not None and not df_subscription.empty:
            summary.append("订阅数据摘要:")
            summary.append(f"- 总订阅合同数: {len(df_subscription)}")
            
            if '金额' in df_subscription.columns and pd.api.types.is_numeric_dtype(df_subscription['金额']):
                total_amount = df_subscription['金额'].sum()
                avg_amount = df_subscription['金额'].mean()
                summary.append(f"- 总订阅金额: {total_amount:,.0f}元")
                summary.append(f"- 平均合同金额: {avg_amount:,.0f}元")
            
            if '开始时间' in df_subscription.columns:
                try:
                    df_subscription['开始年份'] = pd.to_datetime(df_subscription['开始时间']).dt.year
                    yearly_counts = df_subscription['开始年份'].value_counts().sort_index()
                    summary.append("- 年度分布:")
                    for year, count in yearly_counts.items():
                        summary.append(f"  {year}年: {count}个合同")
                except:
                    pass
        
        # 收款数据摘要
        if df_collection is not None and not df_collection.empty:
            summary.append("\n收款数据摘要:")
            summary.append(f"- 总收款记录数: {len(df_collection)}")
            
            if '应收金额' in df_collection.columns and pd.api.types.is_numeric_dtype(df_collection['应收金额']):
                total_due = df_collection['应收金额'].sum()
                summary.append(f"- 总应收金额: {total_due:,.0f}元")
            
            if '实收金额' in df_collection.columns and pd.api.types.is_numeric_dtype(df_collection['实收金额']):
                total_received = df_collection['实收金额'].sum()
                summary.append(f"- 总实收金额: {total_received:,.0f}元")
                
                if 'total_due' in locals() and total_due > 0:
                    collection_rate = (total_received / total_due * 100)
                    summary.append(f"- 收款率: {collection_rate:.1f}%")
            
            if '状态' in df_collection.columns:
                status_counts = df_collection['状态'].value_counts()
                summary.append("- 状态分布:")
                for status, count in status_counts.head(5).items():
                    summary.append(f"  {status}: {count}条")
        
        return "\n".join(summary)