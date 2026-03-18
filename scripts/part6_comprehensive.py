#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Part 6: 综合经营分析模块
基于Part1-5的分析结果，进行综合评估和战略建议
"""

import pandas as pd
import numpy as np
import logging

# 延迟导入llm_client以避免循环依赖
_llm_client = None

logger = logging.getLogger(__name__)


def _get_llm_client():
    """获取LLM客户端"""
    global _llm_client
    if _llm_client is None:
        from llm_client import get_llm_client as _get
        _llm_client = _get()
    return _llm_client


class ComprehensiveAnalyzer:
    """综合经营分析器"""
    
    def __init__(self):
        pass
    
    def analyze(self, part1_data, part2_summary, part3_summary, part4_summary, 
                part1_5_full=None):
        """
        分析综合经营情况
        
        Args:
            part1_data: Part1客户基础档案数据
            part2_summary: Part2订阅续约数据摘要
            part3_summary: Part3实施优化数据摘要
            part4_summary: Part4运维数据摘要
            part1_5_full: Part1-5已生成的完整markdown内容
            
        Returns:
            str: Markdown格式的分析报告
        """
        logger.info("开始综合经营分析")
        
        content = "## 6. 综合经营分析\n\n"
        
        try:
            # 直接使用LLM进行综合分析（基于已生成的Part1-5内容）
            llm_result = None
            try:
                # 优先使用已生成的Part1-5内容
                if part1_5_full:
                    logger.info("开始调用LLM进行Part6综合分析（基于Part1-5已生成内容）...")
                    llm_client = _get_llm_client()
                    llm_result = llm_client.analyze_comprehensive_from_content(part1_5_full)
                    
                    if llm_result:
                        logger.info(f"Part6 LLM分析完成，长度: {len(llm_result)} 字")
                else:
                    # 如果没有已生成内容，回退到原始数据方式
                    logger.warning("Part1-5内容为空，回退到原始数据分析")
                    has_part1 = part1_data is not None
                    
                    if has_part1:
                        part1_str = self._format_part1_full(part1_data)
                        
                        logger.info("开始调用LLM进行Part6综合分析（原始数据）...")
                        llm_client = _get_llm_client()
                        llm_result = llm_client.analyze_comprehensive_full(
                            part1_str, "无数据", "无数据", "无数据", "无数据"
                        )
                        
            except Exception as e:
                logger.error(f"Part6 LLM调用失败: {e}")
            
            # 直接输出LLM分析结果，不分小点
            if llm_result:
                # 清理LLM输出
                cleaned = self._clean_llm_output(llm_result)
                content += f"{cleaned}\n\n"
            else:
                # 如果LLM失败，至少显示客户价值分级
                content += self._generate_customer_value(part2_summary, part3_summary)
                content += "\n*注：LLM综合分析生成失败*\n\n"
            
            logger.info("综合经营分析完成")
            return content
            
        except Exception as e:
            error_msg = f"综合经营分析失败: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return content + f"分析失败: {error_msg}\n"
    
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
    
    def _generate_customer_value(self, part2_summary, part3_summary):
        """生成6.1客户价值分级"""
        content = "### 6.1 客户价值分级\n\n"
        
        # 计算年综合费用
        # 当前ARR（来自Part2）+ 实施金额（来自Part3，按最新年份）
        current_arr = part2_summary.get('current_arr', 0) if part2_summary else 0
        
        # Part3的实施金额汇总
        implementation_amount = 0
        if part3_summary and 'yearly_data' in part3_summary:
            # 取报告生成时间的上一年（比如现在是2026年，就取2025年）
            import datetime
            current_year = datetime.datetime.now().year
            last_year = current_year - 1
            
            # 只取上一年的数据，如果没有则视为0
            if last_year in part3_summary['yearly_data']:
                impl_data = part3_summary['yearly_data'][last_year]
                implementation_amount = impl_data.get('固定合同金额', 0) + impl_data.get('人天框架金额', 0)
        
        annual_expense = current_arr + implementation_amount
        
        # 判断客户等级
        if annual_expense >= 500000:
            level = "战略客户"
        elif annual_expense >= 200000:
            level = "重点客户"
        else:
            level = "普通客户"
        
        content += f"| 年综合费用 | 客户等级 |\n"
        content += f"|-----------|----------|\n"
        content += f"| {annual_expense:,.0f}元 | **{level}** |\n\n"
        
        # 详细说明
        content += f"- 订阅费用(ARR): {current_arr:,.0f}元\n"
        content += f"- 实施费用: {implementation_amount:,.0f}元\n"
        content += f"- 年综合费用: {annual_expense:,.0f}元\n\n"
        
        return content
    
    def _generate_health_assessment(self):
        """生成6.2经营健康度评估"""
        content = "### 6.2 经营健康度评估\n\n"
        content += "- 订阅健康度：待LLM分析\n"
        content += "- 收款健康度：待LLM分析\n"
        content += "- 运维健康度：待LLM分析\n\n"
        return content
    
    def _generate_opportunity_analysis(self):
        """生成6.3机会分析"""
        content = "### 6.3 机会分析\n\n"
        content += "- 增购机会：待LLM分析\n"
        content += "- 交叉销售机会：待LLM分析\n"
        content += "- 服务升级机会：待LLM分析\n\n"
        return content
    
    def _generate_risk_warning(self):
        """生成6.4风险预警"""
        content = "### 6.4 风险预警\n\n"
        content += "- 流失风险：待LLM分析\n"
        content += "- 收款风险：待LLM分析\n"
        content += "- 服务风险：待LLM分析\n\n"
        return content
    
    def _generate_action_suggestions(self):
        """生成6.5下一步行动建议"""
        content = "### 6.5 下一步行动建议\n\n"
        content += "**短期（1-3个月）**：\n"
        content += "- 待LLM分析\n\n"
        content += "**中期（3-12个月）**：\n"
        content += "- 待LLM分析\n\n"
        content += "**长期（1年以上）**：\n"
        content += "- 待LLM分析\n\n"
        return content
    
    def _prepare_data_summary(self, part1_data, part2_summary, part3_summary, part4_summary):
        """准备数据摘要"""
        summary = []
        
        summary.append("综合经营分析数据摘要:")
        
        # Part2摘要
        if part2_summary:
            summary.append("\n订阅续约数据:")
            if 'current_arr' in part2_summary:
                summary.append(f"  当前ARR: {part2_summary['current_arr']:,.0f}元")
            if 'total_planned' in part2_summary:
                summary.append(f"  计划收款总额: {part2_summary['total_planned']:,.0f}元")
            if 'total_received' in part2_summary:
                summary.append(f"  已收款金额: {part2_summary['total_received']:,.0f}元")
        
        # Part3摘要
        if part3_summary and 'yearly_data' in part3_summary:
            summary.append("\n实施优化数据:")
            years = sorted(part3_summary['yearly_data'].keys())
            if years:
                latest_year = years[-1]
                impl_data = part3_summary['yearly_data'][latest_year]
                total = impl_data.get('固定合同金额', 0) + impl_data.get('人天框架金额', 0)
                summary.append(f"  {latest_year}年实施金额: {total:,.0f}元")
        
        # Part4摘要
        if part4_summary:
            summary.append("\n运维数据:")
            if 'total_tickets' in part4_summary:
                summary.append(f"  总工单数: {part4_summary['total_tickets']}")
            if 'total_hours' in part4_summary:
                summary.append(f"  总工时: {part4_summary['total_hours']:.1f}小时")
        
        return "\n".join(summary)
    
    def _format_part1_full(self, df):
        """格式化Part1完整数据"""
        if df is None or (hasattr(df, 'empty') and df.empty):
            return "无数据"
        
        lines = []
        for _, row in df.head(5).iterrows():
            for col in df.columns[:8]:
                val = row.get(col, '')
                if pd.notna(val):
                    lines.append(f"{col}: {val}")
            lines.append("---")
        
        return "\n".join(lines) if lines else "无有效数据"
    
    def _format_part2_full(self, data_dict):
        """格式化Part2完整数据"""
        if not data_dict:
            return "无数据"
        
        lines = []
        for key, df in data_dict.items():
            if df is not None and not df.empty:
                lines.append(f"\n### {key}")
                for _, row in df.head(5).iterrows():
                    row_data = []
                    for col in df.columns[:6]:
                        val = row.get(col, '')
                        if pd.notna(val):
                            row_data.append(f"{col}={val}")
                    if row_data:
                        lines.append(", ".join(row_data))
        
        return "\n".join(lines) if lines else "无数据"
    
    def _format_part3_full(self, data_dict):
        """格式化Part3完整数据"""
        if not data_dict:
            return "无数据"
        
        lines = []
        for key, df in data_dict.items():
            if df is not None and not df.empty:
                lines.append(f"\n### {key}")
                for _, row in df.head(5).iterrows():
                    row_data = []
                    for col in df.columns[:6]:
                        val = row.get(col, '')
                        if pd.notna(val):
                            row_data.append(f"{col}={val}")
                    if row_data:
                        lines.append(", ".join(row_data))
        
        return "\n".join(lines) if lines else "无数据"
    
    def _format_part4_full(self, df):
        """格式化Part4完整数据"""
        if df is None or (hasattr(df, 'empty') and df.empty):
            return "无数据"
        
        lines = []
        for _, row in df.head(10).iterrows():
            row_data = []
            for col in df.columns[:6]:
                val = row.get(col, '')
                if pd.notna(val):
                    row_data.append(f"{col}={val}")
            if row_data:
                lines.append(", ".join(row_data))
        
        return "\n".join(lines) if lines else "无有效数据"


# 测试代码
if __name__ == "__main__":
    # 模拟数据
    part2_summary = {
        'current_arr': 155000,
        'total_planned': 990000,
        'total_received': 680000,
    }
    
    part3_summary = {
        'yearly_data': {
            2023: {'固定合同金额': 25000, '人天框架金额': 0},
        }
    }
    
    part4_summary = {
        'total_tickets': 21,
        'total_hours': 13.7,
    }
    
    analyzer = ComprehensiveAnalyzer()
    result = analyzer.analyze(None, part2_summary, part3_summary, part4_summary)
    print(result)
