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

# 导入其他模块
try:
    from data_loader import get_data_loader
    from part1_basic_profile import BasicProfileAnalyzer
    from part2_subscription import SubscriptionAnalyzer
    from llm_analyzer import LLMAnalyzer
    # 后续会导入 part3_implementation, part4_operations 等模块
except ImportError as e:
    logging.warning(f"部分模块导入失败: {e}")

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
    
    def generate_report(self, client_id, data):
        """
        生成客户经营分析报告
        
        Args:
            client_id: 客户ID
            data: 加载的数据字典
            
        Returns:
            tuple: (report_path, error_message)
        """
        logger.info(f"开始生成客户报告: {client_id}")
        start_time = datetime.now()
        
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
            
            # 保存处理日志
            self._save_processing_log(client_id, report_path, start_time, data)
            
            # 自动组织文件到client_data文件夹
            organized = self._organize_report_files(client_id, report_filename)
            if organized:
                logger.info(f"报告文件已自动组织到client_data文件夹")
            
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
        content += f"**分析工具**: 商务专家Skill v1.0.0\n"
        content += f"**跳过LLM分析**: {'是' if self.skip_llm else '否'}\n\n"
        
        # 检查数据完整性
        missing_parts = []
        if "part1" not in data:
            missing_parts.append("客户基础档案")
        if "part2" not in data:
            missing_parts.append("订阅数据")
        if "part4" not in data:
            missing_parts.append("收款数据")
        if "part5_ops" not in data:
            missing_parts.append("运维工单")
        
        if missing_parts:
            content += f"**数据缺失警告**: 以下部分数据缺失: {', '.join(missing_parts)}\n\n"
        
        # 使用详细报告生成器
        content += self._generate_detailed_report(client_id, data)
        
        # 报告尾部
        content += "\n---\n"
        content += "*报告生成工具: 商务专家Skill v1.0.0*\n"
        content += f"*生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*\n"
        
        return content
    
    def _generate_detailed_report(self, client_id, data):
        """生成详细版报告（使用新的分析模块）"""
        content = ""
        
        try:
            # Part 1: 客户基础档案
            if "part1" in data and not data["part1"].empty:
                analyzer1 = BasicProfileAnalyzer()
                part1_content = analyzer1.analyze(data["part1"])
                content += part1_content
            else:
                content += "## 1. 客户基础档案\n\n暂无数据\n\n"
            
            # Part 2: 订阅续约与续费
            if "part2" in data and not data["part2"].empty and "part4" in data and not data["part4"].empty:
                analyzer2 = SubscriptionAnalyzer()
                part2_content = analyzer2.analyze(data["part2"], data["part4"])
                content += part2_content
                
                # 集成LLM智能分析
                if not self.skip_llm:
                    try:
                        llm_analyzer = LLMAnalyzer()
                        # 准备数据摘要
                        from part2_subscription import SubscriptionAnalyzer as SA
                        sub_analyzer = SA()
                        subscription_summary = sub_analyzer._prepare_data_summary(data["part2"], data["part4"])
                        
                        # 调用LLM分析
                        llm_analysis = llm_analyzer.analyze_subscription(
                            {"summary": subscription_summary},
                            {"summary": subscription_summary}  # 这里简化处理，实际应该分开
                        )
                        
                        # 替换占位符
                        placeholder = "### 2.4 智能分析\n\n*此部分将由LLM进行智能分析"
                        if placeholder in content:
                            content = content.replace(placeholder, llm_analysis)
                        else:
                            # 如果占位符不存在，直接添加
                            content += llm_analysis
                            
                    except Exception as e:
                        logger.warning(f"LLM分析失败: {str(e)}")
                        # 保留原有的占位符
            else:
                content += "## 2. 订阅续约与续费\n\n"
                if "part2" not in data or data["part2"].empty:
                    content += "订阅数据缺失\n"
                if "part4" not in data or data["part4"].empty:
                    content += "收款数据缺失\n"
                content += "\n"
            
            # Part 3: 实施优化情况（简化版）
            content += "## 3. 实施优化情况\n\n"
            impl_data_exists = False
            
            if "part3_fixed" in data and not data["part3_fixed"].empty:
                df = data["part3_fixed"]
                content += "### 3.1 固定金额实施\n"
                content += f"- 记录数: {len(df)} 条\n"
                if '金额' in df.columns and pd.api.types.is_numeric_dtype(df['金额']):
                    total = df['金额'].sum()
                    avg = df['金额'].mean()
                    content += f"- 总金额: {total:,.0f} 元\n"
                    content += f"- 平均金额: {avg:,.0f} 元\n"
                impl_data_exists = True
            
            if "part3_dayspan" in data and not data["part3_dayspan"].empty:
                df = data["part3_dayspan"]
                content += "### 3.2 人天框架实施\n"
                content += f"- 记录数: {len(df)} 条\n"
                if '人天数' in df.columns and pd.api.types.is_numeric_dtype(df['人天数']):
                    total_days = df['人天数'].sum()
                    content += f"- 总人天数: {total_days:.1f} 天\n"
                if '单价' in df.columns and pd.api.types.is_numeric_dtype(df['单价']):
                    avg_price = df['单价'].mean()
                    content += f"- 平均单价: {avg_price:,.0f} 元/天\n"
                impl_data_exists = True
            
            if not impl_data_exists:
                content += "暂无数据\n"
            content += "\n"
            
            # Part 4: 运维情况（简化版）
            content += "## 4. 运维情况\n\n"
            if "part5_ops" in data and not data["part5_ops"].empty:
                df = data["part5_ops"]
                content += f"- 总工单数: **{len(df)}** 个\n"
                
                # 按月份统计
                if '创建时间' in df.columns:
                    try:
                        df['创建月份'] = pd.to_datetime(df['创建时间']).dt.to_period('M')
                        monthly_counts = df['创建月份'].value_counts().sort_index()
                        
                        content += "\n#### 月度工单分布\n"
                        content += "| 月份 | 工单数 |\n"
                        content += "|------|--------|\n"
                        
                        for month, count in monthly_counts.items():
                            month_str = month.strftime('%Y-%m')
                            content += f"| {month_str} | {count} |\n"
                    except:
                        pass
                
                # 问题分类
                if '问题分类' in df.columns:
                    problem_counts = df['问题分类'].value_counts()
                    content += "\n#### 问题分类统计\n"
                    content += "| 问题分类 | 数量 | 占比 |\n"
                    content += "|----------|------|------|\n"
                    
                    total = len(df)
                    for problem, count in problem_counts.head(10).items():  # 只显示前10个
                        percentage = (count / total * 100)
                        content += f"| {problem} | {count} | {percentage:.1f}% |\n"
                
                # 工单状态
                if '状态' in df.columns:
                    status_counts = df['状态'].value_counts()
                    content += "\n#### 工单状态统计\n"
                    for status, count in status_counts.items():
                        content += f"- {status}: {count} 个\n"
                
                # SLA分析（如果存在）
                if 'SLA状态' in df.columns:
                    sla_counts = df['SLA状态'].value_counts()
                    total_sla = sla_counts.sum()
                    if total_sla > 0:
                        sla_rate = (sla_counts.get('达标', 0) / total_sla * 100) if '达标' in sla_counts else 0
                        content += f"\n- SLA达标率: **{sla_rate:.1f}%**\n"
            else:
                content += "暂无数据\n"
            content += "\n"
            
            # Part 6: 综合经营分析
            content += "## 6. 综合经营分析\n\n"
            
            # 提取关键指标
            key_metrics = self._extract_key_metrics(data)
            
            # 集成LLM智能分析
            if not self.skip_llm:
                try:
                    llm_analyzer = LLMAnalyzer()
                    comprehensive_analysis = llm_analyzer.analyze_comprehensive(key_metrics)
                    content += comprehensive_analysis
                except Exception as e:
                    logger.warning(f"LLM综合分析失败: {str(e)}")
                    # 使用占位符
                    content += "*此部分将由LLM进行综合经营分析，包括：*\n"
                    content += "1) 客户价值分级（A/B/C/D级）\n"
                    content += "2) 经营健康度评估（订阅/收款/运维健康度）\n"
                    content += "3) 机会分析（增购/交叉销售机会）\n"
                    content += "4) 风险预警（流失/收款/服务风险）\n"
                    content += "5) 下一步行动建议（短期/中期/长期）\n\n"
                    
                    content += "*关键指标（供LLM参考）:*\n```\n"
                    for metric, value in key_metrics.items():
                        content += f"{metric}: {value}\n"
                    content += "```\n\n"
            else:
                content += "*LLM分析已跳过*\n\n"
                content += "*关键指标:*\n```\n"
                for metric, value in key_metrics.items():
                    content += f"{metric}: {value}\n"
                content += "```\n\n"
            
        except Exception as e:
            logger.error(f"生成详细报告失败: {str(e)}", exc_info=True)
            content += f"\n报告生成过程中出现错误: {str(e)}\n"
        
        return content
    
    def _extract_key_metrics(self, data):
        """提取关键指标供LLM分析使用"""
        metrics = {}
        
        try:
            # Part 1 指标
            if "part1" in data and not data["part1"].empty:
                df = data["part1"]
                if len(df) > 0:
                    row = df.iloc[0]
                    
                    if '计费ARR' in row and pd.notna(row['计费ARR']):
                        metrics['计费ARR'] = float(row['计费ARR'])
                    
                    if '累计购买金额' in row and pd.notna(row['累计购买金额']):
                        metrics['累计购买金额'] = float(row['累计购买金额'])
                    
                    if '购买次数' in row and pd.notna(row['购买次数']):
                        metrics['购买次数'] = float(row['购买次数'])
                    
                    # 服务阶段
                    metrics['服务阶段'] = row.get('服务阶段', '未知')
                    
                    # 客户状态
                    metrics['客户状态'] = row.get('客户状态', '未知')
            
            # Part 2 指标
            if "part2" in data and not data["part2"].empty:
                df = data["part2"]
                metrics['订阅合同数'] = len(df)
                
                if '金额' in df.columns and pd.api.types.is_numeric_dtype(df['金额']):
                    metrics['总订阅金额'] = float(df['金额'].sum())
                    metrics['平均订阅金额'] = float(df['金额'].mean())
            
            # Part 4 指标（收款）
            if "part4" in data and not data["part4"].empty:
                df = data["part4"]
                metrics['收款记录数'] = len(df)
                
                if '应收金额' in df.columns and pd.api.types.is_numeric_dtype(df['应收金额']):
                    metrics['总应收金额'] = float(df['应收金额'].sum())
                
                if '实收金额' in df.columns and pd.api.types.is_numeric_dtype(df['实收金额']):
                    metrics['总实收金额'] = float(df['实收金额'].sum())
                    
                    # 计算收款率
                    if '总应收金额' in metrics and metrics['总应收金额'] > 0:
                        metrics['收款率'] = float(metrics['总实收金额'] / metrics['总应收金额'] * 100)
            
            # Part 5 指标（运维）
            if "part5_ops" in data and not data["part5_ops"].empty:
                df = data["part5_ops"]
                metrics['总工单数'] = len(df)
                
                if '状态' in df.columns:
                    closed_count = len(df[df['状态'] == '已关闭'])
                    metrics['工单关闭率'] = float(closed_count / len(df) * 100) if len(df) > 0 else 0
            
            # Part 3 指标（实施）
            impl_amount = 0
            if "part3_fixed" in data and not data["part3_fixed"].empty:
                df = data["part3_fixed"]
                if '金额' in df.columns and pd.api.types.is_numeric_dtype(df['金额']):
                    impl_amount += float(df['金额'].sum())
            
            if "part3_dayspan" in data and not data["part3_dayspan"].empty:
                df = data["part3_dayspan"]
                if '金额' in df.columns and pd.api.types.is_numeric_dtype(df['金额']):
                    impl_amount += float(df['金额'].sum())
                elif '人天数' in df.columns and '单价' in df.columns:
                    if pd.api.types.is_numeric_dtype(df['人天数']) and pd.api.types.is_numeric_dtype(df['单价']):
                        impl_amount += float((df['人天数'] * df['单价']).sum())
            
            if impl_amount > 0:
                metrics['总实施金额'] = impl_amount
            
        except Exception as e:
            logger.warning(f"提取关键指标失败: {str(e)}")
        
        return metrics
    
    def _save_processing_log(self, client_id, report_path, start_time, data):
        """保存处理日志"""
        log_filename = f"{client_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        log_path = os.path.join(self.log_dir, log_filename)
        
        elapsed_time = (datetime.now() - start_time).total_seconds()
        
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(f"# 商务专家Skill处理日志\n\n")
            f.write(f"**客户ID**: {client_id}\n")
            f.write(f"**处理时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"**处理状态**: 成功\n")
            f.write(f"**耗时**: {elapsed_time:.2f}秒\n")
            f.write(f"**跳过LLM分析**: {'是' if self.skip_llm else '否'}\n\n")
            
            f.write("## 数据加载统计\n")
            for key, df in data.items():
                if df is not None:
                    f.write(f"- {key}: {len(df)} 行, {len(df.columns)} 列\n")
                else:
                    f.write(f"- {key}: 未加载\n")
            
            f.write(f"\n## 输出文件\n")
            f.write(f"- 报告文件: {report_path}\n")
            f.write(f"- 日志文件: {log_path}\n")
        
        logger.info(f"处理日志已保存: {log_path}")
    
    def _organize_report_files(self, client_id, report_filename):
        """
        自动组织报告文件到client_data文件夹
        
        规则:
        1. 输出文件不要放在skill文件夹里
        2. 要放在user文件夹下client_data里面
        3. 如果里面有该客户的文件夹，就放进去
        4. 如果没有，则建该客户简称的文件夹，放进去
        
        Args:
            client_id: 客户ID
            report_filename: 报告文件名
            
        Returns:
            bool: 是否成功组织文件
        """
        try:
            # 导入文件组织器
            from file_organizer import organize_client_report
            
            logger.info(f"开始自动组织文件到client_data文件夹: {client_id}")
            
            # 组织文件
            success, details = organize_client_report(
                client_name=client_id,
                source_dir=self.output_dir,
                report_filename=report_filename
            )
            
            if success:
                logger.info(f"文件组织成功:")
                logger.info(f"  目标目录: {details['target_directory']}")
                logger.info(f"  移动文件数: {len(details['moved_files'])}")
                for file_info in details['moved_files']:
                    logger.info(f"    - {file_info['name']} ({file_info['size']:,} 字节)")
                return True
            else:
                logger.warning(f"文件组织失败: {details.get('error', '未知错误')}")
                return False
                
        except ImportError as e:
            logger.warning(f"无法导入文件组织器模块: {e}")
            logger.info("文件组织功能将跳过，报告文件保留在原始位置")
            return False
        except Exception as e:
            logger.error(f"文件组织过程中发生异常: {str(e)}")
            return False


# 导入pandas用于类型检查
try:
    import pandas as pd
except ImportError:
    pd = None
    logger.warning("pandas模块未安装，部分功能可能受限")


if __name__ == "__main__":
    # 测试代码
    import sys
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    
    from data_loader import DataLoader
    
    # 测试生成报告
    loader = DataLoader()
    data, error = loader.load_client_data("测试客户")
    
    if data:
        generator = ReportGenerator()
        report_path, gen_error = generator.generate_report("测试客户", data)
        
        if report_path:
            print(f"报告生成成功: {report_path}")
            with open(report_path, 'r', encoding='utf-8') as f:
                print(f.read()[:500])  # 打印前500字符
        else:
            print(f"报告生成失败: {gen_error}")
    else:
        print(f"数据加载失败: {error}")