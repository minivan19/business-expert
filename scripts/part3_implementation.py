#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Part 3: 实施优化情况分析模块
分析客户的实施费用、固定合同明细、人天框架明细和优化趋势
"""

import pandas as pd
import numpy as np
import logging
from datetime import datetime

# 导入LLM客户端
try:
    from llm_client import get_llm_client
    LLM_AVAILABLE = True
except ImportError:
    LLM_AVAILABLE = False
    logger.warning("LLM客户端不可用，智能分析将使用默认内容")

logger = logging.getLogger(__name__)


class ImplementationAnalyzer:
    """实施优化情况分析器"""
    
    def __init__(self):
        # 固定金额合同字段映射
        self.fixed_fields = {
            '签约日期': '合同签订时间',
            '合同行号': '合同行号',
            '归属部门': '项目归属部门',
            '合同金额': '固定金额',
            '人天数量': '总人天',
        }
        
        # 人天框架合同字段映射
        self.dayspan_fields = {
            '签约日期': '合同签约日期',
            '合同行号': '合同行号',
            '归属部门': '项目归属部门',
            '合同金额': '应收金额',
            '人天数量': '总人天',
        }
    
    def analyze(self, df_fixed, df_dayspan):
        """
        分析实施优化情况
        
        Args:
            df_fixed: 固定金额台账DataFrame
            df_dayspan: 人天框架台账DataFrame
            
        Returns:
            str: Markdown格式的分析报告
        """
        logger.info("开始分析实施优化情况")
        
        content = "## 3. 实施优化情况\n\n"
        
        try:
            # 3.1 实施概览
            overview_content = self._generate_overview(df_fixed, df_dayspan)
            content += overview_content
            
            # 3.2 固定金额合同明细
            fixed_content = self._generate_fixed_details(df_fixed)
            content += fixed_content
            
            # 3.3 人天框架合同明细
            dayspan_content = self._generate_dayspan_details(df_dayspan)
            content += dayspan_content
            
            # 保存已生成的分析内容，供3.4使用
            self._generated_analysis = overview_content + "\n" + fixed_content + "\n" + dayspan_content
            
            # 3.4 智能分析（基于已生成的分析内容）
            content += self._generate_intelligent_analysis(df_fixed, df_dayspan)
            
            logger.info("实施优化情况分析完成")
            return content
            
        except Exception as e:
            error_msg = f"实施优化情况分析失败: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return content + f"分析失败: {error_msg}\n"
    
    def _generate_overview(self, df_fixed, df_dayspan):
        """生成3.1实施概览"""
        content = "### 3.1 实施概览\n\n"
        
        # 按年份汇总
        yearly_data = {}
        
        # 处理固定金额合同
        if df_fixed is not None and not df_fixed.empty:
            if '合同签订时间' in df_fixed.columns and '固定金额' in df_fixed.columns:
                df_fixed['签约年份'] = pd.to_datetime(df_fixed['合同签订时间'], errors='coerce').dt.year
                for year, group in df_fixed.groupby('签约年份'):
                    year_key = int(year) if pd.notna(year) else '未知'
                    if year_key not in yearly_data:
                        yearly_data[year_key] = {'固定合同金额': 0, '人天框架金额': 0}
                    yearly_data[year_key]['固定合同金额'] += group['固定金额'].sum()
        
        # 处理人天框架合同
        if df_dayspan is not None and not df_dayspan.empty:
            if '合同签约日期' in df_dayspan.columns and '应收金额' in df_dayspan.columns:
                df_dayspan['签约年份'] = pd.to_datetime(df_dayspan['合同签约日期'], errors='coerce').dt.year
                for year, group in df_dayspan.groupby('签约年份'):
                    year_key = int(year) if pd.notna(year) else '未知'
                    if year_key not in yearly_data:
                        yearly_data[year_key] = {'固定合同金额': 0, '人天框架金额': 0}
                    yearly_data[year_key]['人天框架金额'] += group['应收金额'].sum()
        
        if yearly_data:
            # 生成表格
            content += "| 年份 | 固定合同金额 | 人天框架金额 | 汇总金额 |\n"
            content += "|------|-------------|-------------|----------|\n"
            
            total_fixed = 0
            total_dayspan = 0
            
            for year in sorted(yearly_data.keys()):
                fixed = yearly_data[year]['固定合同金额']
                dayspan = yearly_data[year]['人天框架金额']
                total = fixed + dayspan
                total_fixed += fixed
                total_dayspan += dayspan
                
                content += f"| {year} | {fixed:,.0f}元 | {dayspan:,.0f}元 | {total:,.0f}元 |\n"
            
            # 汇总行
            total_all = total_fixed + total_dayspan
            content += f"| **汇总** | **{total_fixed:,.0f}元** | **{total_dayspan:,.0f}元** | **{total_all:,.0f}元** |\n\n"
        else:
            content += "| 年份 | 固定合同金额 | 人天框架金额 | 汇总金额 |\n"
            content += "|------|-------------|-------------|----------|\n"
            content += "| - | - | - | - |\n\n"
        
        return content
    
    def _generate_fixed_details(self, df):
        """生成3.2固定金额合同明细"""
        content = "### 3.2 固定金额合同明细\n\n"
        
        if df is None or df.empty:
            content += "暂无固定金额合同数据\n\n"
            return content
        
        # 筛选需要的字段
        available_fields = {}
        for field_name, col_name in self.fixed_fields.items():
            if col_name in df.columns:
                available_fields[field_name] = col_name
        
        if not available_fields:
            content += "固定金额合同数据格式不正确\n\n"
            return content
        
        # 选择字段
        df_select = df[list(available_fields.values())].copy()
        df_select.columns = list(available_fields.keys())
        
        # 转换为数值（处理字符串格式的金额）
        if '合同金额' in df_select.columns:
            df_select['合同金额'] = pd.to_numeric(df_select['合同金额'], errors='coerce').fillna(0)
        if '人天数量' in df_select.columns:
            df_select['人天数量'] = pd.to_numeric(df_select['人天数量'], errors='coerce').fillna(0)
        
        # 计算折合人天单价
        if '合同金额' in df_select.columns and '人天数量' in df_select.columns:
            df_select['折合人天单价'] = df_select.apply(
                lambda x: x['合同金额'] / x['人天数量'] if x['人天数量'] and x['人天数量'] > 0 else 0,
                axis=1
            )
        
        # 格式化日期
        if '签约日期' in df_select.columns:
            df_select['签约日期'] = pd.to_datetime(df_select['签约日期'], errors='coerce')
            df_select['签约日期'] = df_select['签约日期'].apply(
                lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else '-'
            )
        
        # 格式化金额
        for col in ['合同金额', '折合人天单价']:
            if col in df_select.columns:
                df_select[col] = df_select[col].apply(
                    lambda x: f"{x:,.2f}元" if pd.notna(x) and x > 0 else '-'
                )
        
        # 人天数量格式化
        if '人天数量' in df_select.columns:
            df_select['人天数量'] = df_select['人天数量'].apply(
                lambda x: f"{x:,.0f}人天" if pd.notna(x) else '-'
            )
        
        # 生成表格
        content += "| " + " | ".join(df_select.columns) + " |\n"
        content += "|" + "|".join([" --- " for _ in df_select.columns]) + "|\n"
        
        for _, row in df_select.iterrows():
            content += "| " + " | ".join(str(v) for v in row.values) + " |\n"
        
        content += "\n"
        return content
    
    def _generate_dayspan_details(self, df):
        """生成3.3人天框架合同明细"""
        content = "### 3.3 人天框架合同明细\n\n"
        
        if df is None or df.empty:
            content += "暂无人天框架合同数据\n\n"
            return content
        
        # 筛选需要的字段
        available_fields = {}
        for field_name, col_name in self.dayspan_fields.items():
            if col_name in df.columns:
                available_fields[field_name] = col_name
        
        if not available_fields:
            content += "人天框架合同数据格式不正确\n\n"
            return content
        
        # 选择字段
        df_select = df[list(available_fields.values())].copy()
        df_select.columns = list(available_fields.keys())
        
        # 转换为数值
        if '合同金额' in df_select.columns:
            df_select['合同金额'] = pd.to_numeric(df_select['合同金额'], errors='coerce').fillna(0)
        if '人天数量' in df_select.columns:
            df_select['人天数量'] = pd.to_numeric(df_select['人天数量'], errors='coerce').fillna(0)
        
        # 计算折合人天单价
        if '合同金额' in df_select.columns and '人天数量' in df_select.columns:
            df_select['折合人天单价'] = df_select.apply(
                lambda x: x['合同金额'] / x['人天数量'] if x['人天数量'] and x['人天数量'] > 0 else 0,
                axis=1
            )
        
        # 格式化日期
        if '签约日期' in df_select.columns:
            df_select['签约日期'] = pd.to_datetime(df_select['签约日期'], errors='coerce')
            df_select['签约日期'] = df_select['签约日期'].apply(
                lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else '-'
            )
        
        # 格式化金额
        for col in ['合同金额', '折合人天单价']:
            if col in df_select.columns:
                df_select[col] = df_select[col].apply(
                    lambda x: f"{x:,.2f}元" if pd.notna(x) and x > 0 else '-'
                )
        
        # 人天数量格式化
        if '人天数量' in df_select.columns:
            df_select['人天数量'] = df_select['人天数量'].apply(
                lambda x: f"{x:,.0f}人天" if pd.notna(x) else '-'
            )
        
        # 生成表格
        content += "| " + " | ".join(df_select.columns) + " |\n"
        content += "|" + "|".join([" --- " for _ in df_select.columns]) + "|\n"
        
        for _, row in df_select.iterrows():
            content += "| " + " | ".join(str(v) for v in row.values) + " |\n"
        
        content += "\n"
        return content
    
    def _generate_intelligent_analysis(self, df_fixed, df_dayspan):
        """生成3.4智能分析（基于已生成的分析内容）"""
        content = "### 3.4 智能分析\n\n"
        
        # 使用已生成的分析内容供LLM分析
        generated_content = getattr(self, '_generated_analysis', '') or ''
        
        # 调用LLM进行智能分析（使用已生成的分析内容）
        if LLM_AVAILABLE:
            try:
                logger.info("开始调用LLM进行实施优化智能分析（基于3.1-3.3分析内容）...")
                llm_client = get_llm_client()
                llm_result = llm_client.analyze_implementation_from_content(generated_content)
                
                if llm_result:
                    # 清理LLM输出
                    cleaned = self._clean_llm_output(llm_result)
                    content += f"{cleaned}\n\n"
                else:
                    content += "- LLM调用失败\n\n"
                    
            except Exception as e:
                logger.error(f"LLM调用失败: {e}")
                content += f"- LLM调用异常: {str(e)}\n\n"
        else:
            content += "- LLM不可用\n\n"
        
        return content
    
    def _clean_llm_output(self, text):
        """清理LLM输出，去除标题序号和多余符号"""
        import re
        # 去除以"2.4"、"3.4"、"4.4"、"5."、"6."等开头的标题
        text = re.sub(r'^\d+\.\d+\s+', '', text, flags=re.MULTILINE)
        # 去除"**"加粗标记
        text = text.replace('**', '')
        # 去除多余的*号
        text = re.sub(r'^\s*[-*]\s+', '- ', text, flags=re.MULTILINE)
        return text.strip()
    
    def _prepare_chapter_summary(self, df_fixed, df_dayspan):
        """准备本章节统计好的数据"""
        lines = []
        
        # 3.1 实施概览统计
        lines.append("【3.1 实施概览】")
        total_fixed = 0
        total_dayspan = 0
        
        if df_fixed is not None and not df_fixed.empty:
            if '固定金额' in df_fixed.columns:
                total_fixed = pd.to_numeric(df_fixed['固定金额'], errors='coerce').fillna(0).sum()
            if '总人天' in df_fixed.columns:
                total_fixed_days = pd.to_numeric(df_fixed['总人天'], errors='coerce').fillna(0).sum()
                lines.append(f"- 固定合同：{len(df_fixed)}个，总金额{total_fixed:,.0f}元，总人天{total_fixed_days:,.0f}")
        
        if df_dayspan is not None and not df_dayspan.empty:
            if '应收金额' in df_dayspan.columns:
                total_dayspan = pd.to_numeric(df_dayspan['应收金额'], errors='coerce').fillna(0).sum()
            if '人天数量' in df_dayspan.columns:
                total_dayspan_days = pd.to_numeric(df_dayspan['人天数量'], errors='coerce').fillna(0).sum()
                lines.append(f"- 人天框架：{len(df_dayspan)}个，总金额{total_dayspan:,.0f}元，总人天{total_dayspan_days:,.0f}")
        
        lines.append(f"- 实施费用总计：{total_fixed + total_dayspan:,.0f}元")
        
        # 3.2 固定合同明细统计
        lines.append("\n【3.2 固定合同明细统计】")
        if df_fixed is not None and not df_fixed.empty:
            key_cols = ['合同编号', '合同签订时间', '固定金额', '总人天']
            available = [c for c in key_cols if c in df_fixed.columns]
            if available:
                df_show = df_fixed[available].head(5)
                lines.append("| " + " | ".join(available) + " |")
                lines.append("|---" * len(available) + "|")
                for _, row in df_show.iterrows():
                    vals = []
                    for c in available:
                        v = row.get(c, '')
                        if pd.isna(v): vals.append('-')
                        elif isinstance(v, (int,float)): vals.append(f"{v:,.0f}")
                        else: vals.append(str(v)[:20])
                    lines.append("| " + " | ".join(vals) + " |")
        else:
            lines.append("- 无固定合同数据")
        
        # 3.3 人天框架明细统计
        lines.append("\n【3.3 人天框架明细统计】")
        if df_dayspan is not None and not df_dayspan.empty:
            key_cols = ['合同编号', '签订时间', '合同金额', '人天数量']
            available = [c for c in key_cols if c in df_dayspan.columns]
            if available:
                df_show = df_dayspan[available].head(5)
                lines.append("| " + " | ".join(available) + " |")
                lines.append("|---" * len(available) + "|")
                for _, row in df_show.iterrows():
                    vals = []
                    for c in available:
                        v = row.get(c, '')
                        if pd.isna(v): vals.append('-')
                        elif isinstance(v, (int,float)): vals.append(f"{v:,.0f}")
                        else: vals.append(str(v)[:20])
                    lines.append("| " + " | ".join(vals) + " |")
        else:
            lines.append("- 无人天框架数据")
        
        return "\n".join(lines)
    
    def _prepare_full_data(self, df):
        """准备完整数据供LLM分析使用"""
        if df is None or df.empty:
            return "无数据"
        
        # 选择关键字段（取前10列）
        available_fields = list(df.columns)[:10]
        
        # 选择数据
        df_select = df[available_fields].head(20).copy()
        
        # 格式化
        lines = []
        lines.append("列名: " + " | ".join(available_fields))
        lines.append("--- " * len(available_fields))
        
        for _, row in df_select.iterrows():
            values = []
            for col in available_fields:
                val = row.get(col, '')
                if pd.isna(val):
                    values.append('-')
                elif isinstance(val, (int, float)):
                    values.append(f"{val:,.0f}")
                else:
                    values.append(str(val)[:30])
            lines.append(" | ".join(values))
        
        if len(df) > 20:
            lines.append(f"... (共 {len(df)} 条记录)")
        
        return "\n".join(lines)
    
    def _prepare_data_summary(self, df_fixed, df_dayspan):
        """准备数据摘要供LLM分析使用"""
        summary = []
        
        # 固定金额合同摘要
        if df_fixed is not None and not df_fixed.empty:
            summary.append("固定金额合同摘要:")
            summary.append(f"- 总记录数: {len(df_fixed)}")
            
            if '固定金额' in df_fixed.columns:
                total = pd.to_numeric(df_fixed['固定金额'], errors='coerce').fillna(0).sum()
                summary.append(f"- 总实施金额: {total:,.0f}元")
            
            if '总人天' in df_fixed.columns:
                total_days = pd.to_numeric(df_fixed['总人天'], errors='coerce').fillna(0).sum()
                summary.append(f"- 总人天: {total_days:,.0f}人天")
        
        # 人天框架合同摘要
        if df_dayspan is not None and not df_dayspan.empty:
            summary.append("\n人天框架合同摘要:")
            summary.append(f"- 总记录数: {len(df_dayspan)}")
            
            if '应收金额' in df_dayspan.columns:
                total = pd.to_numeric(df_dayspan['应收金额'], errors='coerce').fillna(0).sum()
                summary.append(f"- 总实施金额: {total:,.0f}元")
            
            if '总人天' in df_dayspan.columns:
                total_days = pd.to_numeric(df_dayspan['总人天'], errors='coerce').fillna(0).sum()
                summary.append(f"- 总人天: {total_days:,.0f}人天")
        
        return "\n".join(summary)


# 测试代码
if __name__ == "__main__":
    import os
    
    # 测试加载数据
    client_path = r"C:\Users\mingh\client-data\raw\客户档案\CBD"
    fixed_path = os.path.join(client_path, "实施合同行", "固定金额台账.xlsx")
    dayspan_path = os.path.join(client_path, "实施合同行", "人天框架台账.xlsx")
    
    df_fixed = pd.read_excel(fixed_path) if os.path.exists(fixed_path) else None
    df_dayspan = pd.read_excel(dayspan_path) if os.path.exists(dayspan_path) else None
    
    # 分析
    analyzer = ImplementationAnalyzer()
    result = analyzer.analyze(df_fixed, df_dayspan)
    
    print(result)
