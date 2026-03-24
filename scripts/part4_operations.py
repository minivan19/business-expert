#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Part 4: 运维情况分析模块
分析客户的运维工单情况、工时统计、问题分类和模块分布
"""

import pandas as pd
import numpy as np
import logging
import os
from datetime import datetime
from glob import glob

# 导入LLM客户端
try:
    from llm_client import get_llm_client
    LLM_AVAILABLE = True
except ImportError:
    LLM_AVAILABLE = False
    logger.warning("LLM客户端不可用，智能分析将使用默认内容")

logger = logging.getLogger(__name__)


class OperationsAnalyzer:
    """运维情况分析器"""
    
    def __init__(self):
        pass
    
    def analyze(self, df_ops):
        """
        分析运维情况
        
        Args:
            df_ops: 运维工单DataFrame（已合并所有文件）
            
        Returns:
            str: Markdown格式的分析报告
        """
        logger.info("开始分析运维情况")
        
        content = "## 4. 运维情况\n\n"
        
        try:
            # 4.1 运维概览
            overview_content = self._generate_overview(df_ops)
            content += overview_content
            
            # 4.2 模块分布
            module_content = self._generate_module_distribution(df_ops)
            content += module_content
            
            # 4.3 类型分布
            type_content = self._generate_type_distribution(df_ops)
            content += type_content
            
            # 保存已生成的分析内容，供4.4使用
            self._generated_analysis = overview_content + "\n" + module_content + "\n" + type_content
            
            # 4.4 智能分析（基于已生成的分析内容）
            content += self._generate_intelligent_analysis(df_ops)
            
            logger.info("运维情况分析完成")
            return content
            
        except Exception as e:
            error_msg = f"运维情况分析失败: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return content + f"分析失败: {error_msg}\n"
    
    def _generate_overview(self, df):
        """生成4.1运维概览"""
        content = "### 4.1 运维概览\n\n"
        
        if df is None or df.empty:
            content += "| 年份 | 工单数量 | 工时总和 | 单均工时 |\n"
            content += "|------|----------|---------|----------|\n"
            content += "| - | - | - | - |\n\n"
            return content
        
        # 提取年份
        if '创建时间' in df.columns:
            df['年份'] = pd.to_datetime(df['创建时间'], errors='coerce').dt.year
        elif '提单时间' in df.columns:
            df['年份'] = pd.to_datetime(df['提单时间'], errors='coerce').dt.year
        else:
            df['年份'] = '未知'
        
        # 按年份统计
        # 先将总工时转为数值（处理"X小时Y分"格式）
        if '总工时' in df.columns:
            def parse_hours(val):
                if pd.isna(val):
                    return 0
                val = str(val)
                hours = 0
                minutes = 0
                if '小时' in val:
                    parts = val.split('小时')
                    hours = float(parts[0]) if parts[0] else 0
                    if len(parts) > 1 and '分' in parts[1]:
                        minutes = float(parts[1].replace('分', '')) if parts[1].replace('分', '') else 0
                elif '分' in val:
                    minutes = float(val.replace('分', ''))
                return hours + minutes / 60
            
            df['总工时_小时'] = df['总工时'].apply(parse_hours)
        else:
            df['总工时_小时'] = 0
        
        yearly_stats = df.groupby('年份').agg(
            工单数量=('年份', 'count'),
            工时总和=('总工时_小时', 'sum')
        ).reset_index()
        
        # 计算单均工时
        yearly_stats['单均工时'] = yearly_stats['工时总和'] / yearly_stats['工单数量']
        
        # 格式化
        yearly_stats['工时总和'] = yearly_stats['工时总和'].apply(lambda x: f"{x:,.1f}小时" if pd.notna(x) else '-')
        yearly_stats['单均工时'] = yearly_stats['单均工时'].apply(lambda x: f"{x:,.1f}小时" if pd.notna(x) and x > 0 else '-')
        
        # 生成表格
        content += "| 年份 | 工单数量 | 工时总和 | 单均工时 |\n"
        content += "|------|----------|---------|----------|\n"
        
        for _, row in yearly_stats.iterrows():
            # 处理空值
            year = int(row['年份']) if pd.notna(row['年份']) else '-'
            工单数量 = int(row['工单数量']) if pd.notna(row['工单数量']) else '-'
            工时总和 = row['工时总和'] if pd.notna(row['工时总和']) else '-'
            单均工时 = row['单均工时'] if pd.notna(row['单均工时']) else '-'
            content += f"| {year} | {工单数量} | {工时总和} | {单均工时} |\n"
        
        content += "\n"
        return content
    
    def _generate_module_distribution(self, df):
        """生成4.2模块分布"""
        content = "### 4.2 模块分布\n\n"
        
        if df is None or df.empty or '模块' not in df.columns:
            content += "暂无模块数据\n\n"
            return content
        
        # 提取年份
        if '创建时间' in df.columns:
            df['年份'] = pd.to_datetime(df['创建时间'], errors='coerce').dt.year
        elif '提单时间' in df.columns:
            df['年份'] = pd.to_datetime(df['提单时间'], errors='coerce').dt.year
        else:
            content += "无法提取年份信息\n\n"
            return content
        
        # 透视表：模块 x 年份
        pivot = df.pivot_table(
            index='模块',
            columns='年份',
            values='编号',  # 用编号计数
            aggfunc='count',
            fill_value=0
        )
        
        # 按年份从旧到新排列
        years = sorted(pivot.columns.tolist())
        pivot = pivot[years]
        
        # 添加总计列
        pivot['总计'] = pivot.sum(axis=1)
        
        # 按总计降序排列
        pivot = pivot.sort_values('总计', ascending=False)
        
        # 生成表格
        content += "| 模块 | " + " | ".join([str(int(y)) for y in years]) + " | 总计 |\n"
        content += "|------|" + "|".join(["---" for _ in years]) + "|------|\n"
        
        for module, row in pivot.iterrows():
            # 处理空值为"-"
            module = "-" if pd.isna(module) or str(module).strip() == "" else str(module)
            values = [str(int(row[y])) for y in years]
            values.append(str(int(row['总计'])))
            content += f"| {module} | " + " | ".join(values) + " |\n"
        
        content += "\n"
        return content
    
    def _generate_type_distribution(self, df):
        """生成4.3类型分布"""
        content = "### 4.3 类型分布\n\n"
        
        if df is None or df.empty:
            content += "暂无类型数据\n\n"
            return content
        
        # 确定使用哪个字段作为分类（优先使用"分类"字段）
        type_col = None
        for col in ['分类', '工单类型', '系统问题分类']:
            if col in df.columns:
                type_col = col
                break
        
        if type_col is None:
            content += "未找到分类字段\n\n"
            return content
        
        # 提取年份
        if '创建时间' in df.columns:
            df['年份'] = pd.to_datetime(df['创建时间'], errors='coerce').dt.year
        elif '提单时间' in df.columns:
            df['年份'] = pd.to_datetime(df['提单时间'], errors='coerce').dt.year
        else:
            content += "无法提取年份信息\n\n"
            return content
        
        # 透视表：分类 x 年份
        pivot = df.pivot_table(
            index=type_col,
            columns='年份',
            values='编号',
            aggfunc='count',
            fill_value=0
        )
        
        # 按年份从旧到新排列
        years = sorted(pivot.columns.tolist())
        pivot = pivot[years]
        
        # 添加总计列
        pivot['总计'] = pivot.sum(axis=1)
        
        # 按总计降序排列
        pivot = pivot.sort_values('总计', ascending=False)
        
        # 生成表格
        content += f"| {type_col} | " + " | ".join([str(int(y)) for y in years]) + " | 总计 |\n"
        content += "|------|" + "|".join(["---" for _ in years]) + "|------|\n"
        
        for type_name, row in pivot.iterrows():
            # 处理空值为"-"
            type_name = "-" if pd.isna(type_name) or str(type_name).strip() == "" else str(type_name)
            values = [str(int(row[y])) for y in years]
            values.append(str(int(row['总计'])))
            content += f"| {type_name} | " + " | ".join(values) + " |\n"
        
        content += "\n"
        return content
    
    def _generate_intelligent_analysis(self, df):
        """生成4.4智能分析（基于已生成的分析内容）"""
        content = "### 4.4 智能分析\n\n"
        
        # 使用已生成的分析内容供LLM分析
        generated_content = getattr(self, '_generated_analysis', '') or ''
        
        # 调用LLM进行智能分析（使用已生成的分析内容）
        if LLM_AVAILABLE:
            try:
                logger.info("开始调用LLM进行运维智能分析（基于4.1-4.3分析内容）...")
                llm_client = get_llm_client()
                llm_result = llm_client.analyze_operations_from_content(generated_content)
                
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
    
    def _prepare_chapter_summary(self, df):
        """准备本章节统计好的数据"""
        lines = []
        
        # 4.1 运维概览统计
        lines.append("【4.1 运维概览】")
        if df is not None and not df.empty:
            lines.append(f"- 总工单数：{len(df)}")
            
            if '工单时长（小时）' in df.columns:
                total_hours = pd.to_numeric(df['工单时长（小时）'], errors='coerce').fillna(0).sum()
                lines.append(f"- 总工时：{total_hours:.1f}小时")
                avg_hours = total_hours / len(df)
                lines.append(f"- 平均工单时长：{avg_hours:.1f}小时")
        else:
            lines.append("- 无运维数据")
        
        # 4.2 模块分布统计
        lines.append("\n【4.2 模块问题分布】")
        if df is not None and not df.empty and '模块' in df.columns:
            module_counts = df['模块'].value_counts().head(10)
            for module, count in module_counts.items():
                lines.append(f"- {module}: {count}个工单")
        else:
            lines.append("- 无模块数据")
        
        # 4.3 类型分布统计
        lines.append("\n【4.3 问题类型分布】")
        if df is not None and not df.empty and '问题类型' in df.columns:
            type_counts = df['问题类型'].value_counts().head(10)
            for qtype, count in type_counts.items():
                lines.append(f"- {qtype}: {count}个工单")
        else:
            lines.append("- 无类型数据")
        
        return "\n".join(lines)
    
    def _prepare_full_data(self, df):
        """准备完整数据供LLM分析使用"""
        if df is None or df.empty:
            return "无数据"
        
        # 选择关键字段（取前10列）
        available_fields = list(df.columns)[:10]
        
        # 选择数据
        df_select = df[available_fields].head(30).copy()
        
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
        
        if len(df) > 30:
            lines.append(f"... (共 {len(df)} 条记录)")
        
        return "\n".join(lines)
    
    def _prepare_data_summary(self, df):
        """准备数据摘要"""
        summary = []
        
        if df is None or df.empty:
            summary.append("无运维数据")
            return "\n".join(summary)
        
        summary.append(f"总工单数: {len(df)}")
        
        # 年份分布
        if '创建时间' in df.columns:
            df['年份'] = pd.to_datetime(df['创建时间'], errors='coerce').dt.year
        elif '提单时间' in df.columns:
            df['年份'] = pd.to_datetime(df['提单时间'], errors='coerce').dt.year
        
        if '年份' in df.columns:
            year_counts = df['年份'].value_counts().sort_index()
            summary.append("年份分布:")
            for year, count in year_counts.items():
                summary.append(f"  {int(year)}年: {count}单")
        
        # 工时统计
        if '总工时_小时' in df.columns:
            total_hours = df['总工时_小时'].sum()
            avg_hours = df['总工时_小时'].mean()
            summary.append(f"\n工时统计:")
            summary.append(f"  总工时: {total_hours:,.1f}小时")
            summary.append(f"  平均单均工时: {avg_hours:,.1f}小时")
        elif '总工时' in df.columns:
            # 尝试解析
            def parse_hours(val):
                if pd.isna(val):
                    return 0
                val = str(val)
                hours = 0
                minutes = 0
                if '小时' in val:
                    parts = val.split('小时')
                    hours = float(parts[0]) if parts[0] else 0
                    if len(parts) > 1 and '分' in parts[1]:
                        minutes = float(parts[1].replace('分', '')) if parts[1].replace('分', '') else 0
                elif '分' in val:
                    minutes = float(val.replace('分', ''))
                return hours + minutes / 60
            
            hours_numeric = df['总工时'].apply(parse_hours)
            total_hours = hours_numeric.sum()
            avg_hours = hours_numeric.mean()
            summary.append(f"\n工时统计:")
            summary.append(f"  总工时: {total_hours:,.1f}小时")
            summary.append(f"  平均单均工时: {avg_hours:,.1f}小时")
        
        # 模块分布
        if '模块' in df.columns:
            module_counts = df['模块'].value_counts().head(5)
            summary.append("\n模块分布(Top5):")
            for module, count in module_counts.items():
                summary.append(f"  {module}: {count}单")
        
        return "\n".join(summary)


# 测试代码
if __name__ == "__main__":
    # 加载所有运维工单文件
    ops_dir = r"/Users/limingheng/AI/client-data/客户档案/CBD/运维工单"
    
    if os.path.exists(ops_dir):
        files = glob(os.path.join(ops_dir, "*.xlsx"))
        dfs = []
        for f in files:
            df = pd.read_excel(f)
            dfs.append(df)
        
        if dfs:
            df_all = pd.concat(dfs, ignore_index=True)
            print(f"共加载 {len(df_all)} 条工单")
            
            analyzer = OperationsAnalyzer()
            result = analyzer.analyze(df_all)
            print(result)
    else:
        print("运维目录不存在")
