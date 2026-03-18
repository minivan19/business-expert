#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
报告生成器 - 主协调模块
负责协调各个分析模块，生成完整的经营分析报告
"""

import os
import logging
from datetime import datetime
import sys

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 导入各分析模块
from data_loader import get_data_loader
from part1_basic_profile import BasicProfileAnalyzer
from part2_subscription import SubscriptionAnalyzer
from part3_implementation import ImplementationAnalyzer
from part4_operations import OperationsAnalyzer
from part6_comprehensive import ComprehensiveAnalyzer

logger = logging.getLogger(__name__)


class ReportGenerator:
    """报告生成器类"""
    
    def __init__(self, output_dir=None, skip_llm=False):
        """
        初始化报告生成器
        
        Args:
            output_dir: 输出目录
            skip_llm: 是否跳过LLM分析
        """
        if output_dir is None:
            self.output_dir = r"C:\Users\mingh\client-data\_temp"
        else:
            self.output_dir = output_dir
        
        self.skip_llm = skip_llm
        
        # 创建输出目录
        os.makedirs(self.output_dir, exist_ok=True)
        
        # 创建日志目录
        self.log_dir = os.path.join(self.output_dir, "logs")
        os.makedirs(self.log_dir, exist_ok=True)
        
        logger.info(f"报告生成器初始化，输出目录: {self.output_dir}")
        logger.info(f"跳过LLM分析: {self.skip_llm}")
    
    def generate_report(self, client_id, data=None):
        """
        生成客户经营分析报告
        
        Args:
            client_id: 客户ID
            data: 加载的数据字典（如果为None则自动加载）
            
        Returns:
            tuple: (report_path, error_message)
        """
        logger.info(f"开始生成客户报告: {client_id}")
        start_time = datetime.now()
        
        # 自动加载数据
        if data is None:
            loader = get_data_loader()
            data, error = loader.load_client_data(client_id)
            if error:
                logger.warning(f"数据加载部分失败: {error}")
        
        try:
            # 生成报告内容
            report_content = self._generate_report_content(client_id, data)
            
            if not report_content:
                error_msg = "报告内容生成失败"
                logger.error(error_msg)
                return None, error_msg
            
            # 生成报告文件路径
            report_filename = f"{client_id}_经营分析报告_{datetime.now().strftime('%Y%m%d')}.md"
            report_path = os.path.join(self.output_dir, report_filename)
            
            # 保存报告
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(report_content)
            
            elapsed_time = (datetime.now() - start_time).total_seconds()
            logger.info(f"报告生成成功: {report_path}")
            logger.info(f"处理耗时: {elapsed_time:.2f}秒")
            
            return report_path, None
            
        except Exception as e:
            error_msg = f"报告生成过程中发生异常: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return None, error_msg
    
    def _generate_report_content(self, client_id, data):
        """生成报告内容"""
        # 报告头部
        content = f"# {client_id}经营分析报告\n\n"
        content += f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        content += f"**分析工具**: 商务专家Skill v2.0\n"
        content += f"**跳过LLM分析**: {'是' if self.skip_llm else '否'}\n\n"
        
        # 检查数据完整性
        missing_parts = []
        if "part1" not in data or data["part1"].empty:
            missing_parts.append("客户基础档案")
        if "part2" not in data or data["part2"].empty:
            missing_parts.append("订阅明细")
        if "part4" not in data or data["part4"].empty:
            missing_parts.append("收款数据")
        if "part3_fixed" not in data and "part3_dayspan" not in data:
            missing_parts.append("实施数据")
        if "part5_ops" not in data or data["part5_ops"].empty:
            missing_parts.append("运维工单")
        
        if missing_parts:
            content += f"**数据缺失警告**: 以下部分数据缺失: {', '.join(missing_parts)}\n\n"
        
        # 生成各部分内容
        content += self._generate_all_parts(data)
        
        # 报告尾部
        content += "\n---\n"
        content += "*报告生成工具: 商务专家Skill v2.0*\n"
        content += f"*生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*\n"
        
        return content
    
    def _generate_all_parts(self, data):
        """生成所有分析部分"""
        content = ""
        
        # ===== Part 1: 客户基础档案 =====
        if "part1" in data and not data["part1"].empty:
            analyzer1 = BasicProfileAnalyzer()
            content += analyzer1.analyze(data["part1"])
        else:
            content += "## 1. 客户基础档案\n\n暂无数据\n\n"
        
        # ===== Part 2: 订阅续约与续费 =====
        df_sub = data.get("part2")
        df_coll = data.get("part4")
        
        if df_sub is not None and not df_sub.empty:
            analyzer2 = SubscriptionAnalyzer()
            content += analyzer2.analyze(df_sub, df_coll)
        else:
            content += "## 2. 订阅续约与续费\n\n暂无数据\n\n"
        
        # ===== Part 3: 实施优化情况 =====
        df_fixed = data.get("part3_fixed")
        df_dayspan = data.get("part3_dayspan")
        
        if df_fixed is not None or df_dayspan is not None:
            analyzer3 = ImplementationAnalyzer()
            content += analyzer3.analyze(df_fixed, df_dayspan)
        else:
            content += "## 3. 实施优化情况\n\n暂无数据\n\n"
        
        # ===== Part 4: 运维情况 =====
        df_ops = data.get("part5_ops")
        
        if df_ops is not None and not df_ops.empty:
            analyzer4 = OperationsAnalyzer()
            content += analyzer4.analyze(df_ops)
        else:
            content += "## 4. 运维情况\n\n暂无数据\n\n"
        
        # ===== Part 6: 综合经营分析 =====
        # 准备各部分摘要数据
        part1_data = data.get("part1")
        part2_summary = self._extract_part2_summary(df_sub, df_coll)
        part3_summary = self._extract_part3_summary(df_fixed, df_dayspan)
        part4_summary = self._extract_part4_summary(df_ops)
        
        analyzer6 = ComprehensiveAnalyzer()
        content += analyzer6.analyze(part1_data, part2_summary, part3_summary, part4_summary)
        
        return content
    
    def _extract_part2_summary(self, df_sub, df_coll):
        """提取Part2摘要数据"""
        summary = {}
        
        if df_sub is not None and not df_sub.empty:
            # 当前ARR
            active_subs = df_sub[df_sub['订阅状态'] == '订阅中']
            summary['current_arr'] = active_subs['年订阅费金额'].sum() if not active_subs.empty else 0
        
        if df_coll is not None and not df_coll.empty:
            summary['total_planned'] = df_coll['计划收款金额'].sum()
            summary['total_received'] = df_coll['已收款金额'].sum()
            summary['total_unreceived'] = df_coll['未收款金额'].sum()
        
        return summary
    
    def _extract_part3_summary(self, df_fixed, df_dayspan):
        """提取Part3摘要数据"""
        summary = {'yearly_data': {}}
        
        # 固定金额合同
        if df_fixed is not None and not df_fixed.empty:
            if '合同签订时间' in df_fixed.columns and '固定金额' in df_fixed.columns:
                df_fixed['签约年份'] = pd.to_datetime(df_fixed['合同签订时间'], errors='coerce').dt.year
                for year, group in df_fixed.groupby('签约年份'):
                    year_key = int(year) if pd.notna(year) else 0
                    if year_key not in summary['yearly_data']:
                        summary['yearly_data'][year_key] = {'固定合同金额': 0, '人天框架金额': 0}
                    summary['yearly_data'][year_key]['固定合同金额'] += group['固定金额'].sum()
        
        # 人天框架合同
        if df_dayspan is not None and not df_dayspan.empty:
            if '合同签约日期' in df_dayspan.columns and '应收金额' in df_dayspan.columns:
                df_dayspan['签约年份'] = pd.to_datetime(df_dayspan['合同签约日期'], errors='coerce').dt.year
                for year, group in df_dayspan.groupby('签约年份'):
                    year_key = int(year) if pd.notna(year) else 0
                    if year_key not in summary['yearly_data']:
                        summary['yearly_data'][year_key] = {'固定合同金额': 0, '人天框架金额': 0}
                    summary['yearly_data'][year_key]['人天框架金额'] += group['应收金额'].sum()
        
        return summary
    
    def _extract_part4_summary(self, df_ops):
        """提取Part4摘要数据"""
        summary = {}
        
        if df_ops is not None and not df_ops.empty:
            summary['total_tickets'] = len(df_ops)
            
            # 解析工时
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
            
            if '总工时' in df_ops.columns:
                summary['total_hours'] = df_ops['总工时'].apply(parse_hours).sum()
        
        return summary


# 导入pandas
import pandas as pd


# 主入口函数
def main():
    """主入口"""
    import argparse
    
    parser = argparse.ArgumentParser(description='生成客户经营分析报告')
    parser.add_argument('client_id', help='客户ID')
    parser.add_argument('--output', '-o', help='输出目录', default=None)
    parser.add_argument('--skip-llm', '-s', action='store_true', help='跳过LLM分析')
    
    args = parser.parse_args()
    
    # 创建报告生成器
    generator = ReportGenerator(output_dir=args.output, skip_llm=args.skip_llm)
    
    # 生成报告
    report_path, error = generator.generate_report(args.client_id)
    
    if error:
        print(f"生成失败: {error}")
        sys.exit(1)
    else:
        print(f"报告生成成功: {report_path}")


if __name__ == "__main__":
    main()
