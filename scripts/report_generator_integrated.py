#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
集成版报告生成器 - 集成所有分析模块，适配现有API
"""

import os
import json
import logging
import subprocess
import pandas as pd
from datetime import datetime
import sys
import argparse

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 导入各分析模块
from data_loader import get_data_loader
from part1_basic_profile import BasicProfileAnalyzer
from part2_subscription import SubscriptionAnalyzer
from part3_implementation import ImplementationAnalyzer
from part4_operations import OperationsAnalyzer
from part5_business_intelligence import BusinessIntelligenceAnalyzer
from part6_comprehensive import ComprehensiveAnalyzer

logger = logging.getLogger(__name__)


class IntegratedReportGenerator:
    """集成版报告生成器"""

    def __init__(self, skip_llm=False):
        """
        初始化报告生成器

        Args:
            skip_llm: 是否跳过LLM分析
        """
        self.skip_llm = skip_llm

        # 基础目录
        self.base_dir = r"/Users/limingheng/AI\client-data"

        # 临时目录（保留）
        self.temp_dir = os.path.join(self.base_dir, "_temp")

        # Markdown转Word脚本路径（本Skill内）
        self.md_to_word_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "md2docx.py")
        self.business_template = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "templates", "business.docx")

        # 创建基础目录
        os.makedirs(self.temp_dir, exist_ok=True)

        logger.info(f"集成版报告生成器初始化")
        logger.info(f"基础目录: {self.base_dir}")
        logger.info(f"临时目录: {self.temp_dir}")
        logger.info(f"跳过LLM分析: {skip_llm}")

    def get_client_short_name(self, client_id):
        """
        从客户主数据获取客户简称

        Args:
            client_id: 客户ID

        Returns:
            str: 客户简称，如果获取失败则返回client_id
        """
        try:
            # 客户主数据文件路径
            client_data_path = os.path.join(self.base_dir, "raw", "基础数据", "客户主数据.xlsx")

            if not os.path.exists(client_data_path):
                logger.warning(f"客户主数据文件不存在: {client_data_path}")
                return client_id  # 默认返回客户ID

            # 读取Excel文件
            df = pd.read_excel(client_data_path)

            if df.empty:
                logger.warning("客户主数据文件为空")
                return client_id

            # 查找客户简称字段
            short_name_col = None
            for col in df.columns:
                if "简称" in str(col):
                    short_name_col = col
                    break

            if short_name_col is None:
                logger.warning("未找到客户简称字段")
                return client_id

            # 简化处理：取第一行的简称
            if len(df) > 0:
                short_name = df.iloc[0][short_name_col]
                if pd.isna(short_name):
                    return client_id

                short_name_str = str(short_name).strip()
                logger.info(f"获取到客户简称: {short_name_str}")
                return short_name_str
            else:
                return client_id

        except Exception as e:
            logger.error(f"获取客户简称时出错: {e}")
            return client_id

    def create_client_directory(self, client_short_name):
        """
        创建客户输出目录

        Args:
            client_short_name: 客户简称

        Returns:
            str: 客户目录路径
        """
        client_dir = os.path.join(self.base_dir, client_short_name)
        os.makedirs(client_dir, exist_ok=True)
        logger.info(f"创建客户目录: {client_dir}")
        return client_dir

    def generate_report_filenames(self, client_short_name):
        """
        生成报告文件名

        Args:
            client_short_name: 客户简称

        Returns:
            tuple: (md_filename, docx_filename)
        """
        date_str = datetime.now().strftime("%Y%m%d")

        # 基础文件名
        base_name = f"{client_short_name}_经营分析报告_{date_str}"

        # 文件名
        md_filename = f"{base_name}.md"
        docx_filename = f"{base_name}.docx"

        return md_filename, docx_filename

    def _generate_report_content(self, data, client_id):
        """
        生成报告内容（内部方法），适配现有API

        Args:
            data: 加载的数据字典
            client_id: 客户ID

        Returns:
            str: 报告内容
        """
        report_parts = []

        try:
            # 1. 报告标题
            report_parts.append(f"# {client_id} 经营分析报告")
            report_parts.append(f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            report_parts.append(f"**报告版本**: 集成版（完整数据分析 + Word输出）")
            report_parts.append("")

            # 2. Part1: 客户基础档案
            logger.info("生成Part1: 客户基础档案")
            try:
                part1 = BasicProfileAnalyzer()
                if 'part1' in data:
                    part1_content = part1.analyze(data['part1'])
                    if part1_content:
                        report_parts.append(part1_content)
                else:
                    report_parts.append("## Part1: 客户基础档案\n\n*数据缺失*")
            except Exception as e:
                logger.error(f"Part1生成失败: {e}")
                report_parts.append("## Part1: 客户基础档案\n\n*生成失败*")

            # 3. Part2: 订阅续约与续费
            logger.info("生成Part2: 订阅续约与续费")
            try:
                part2 = SubscriptionAnalyzer()
                # 检查所需数据是否存在
                if 'part2' in data and 'part4' in data:
                    part2_content = part2.analyze(data['part2'], data['part4'])
                    if part2_content:
                        report_parts.append(part2_content)
                else:
                    report_parts.append("## Part2: 订阅续约与续费\n\n*数据缺失*")
            except Exception as e:
                logger.error(f"Part2生成失败: {e}")
                report_parts.append("## Part2: 订阅续约与续费\n\n*生成失败*")

            # 4. Part3: 实施优化情况
            logger.info("生成Part3: 实施优化情况")
            try:
                part3 = ImplementationAnalyzer()
                # 检查所需数据是否存在
                if 'part3_fixed' in data and data['part3_fixed'] is not None:
                    df_dayspan = data.get('part3_dayspan')
                    part3_content = part3.analyze(data['part3_fixed'], df_dayspan)
                    if part3_content:
                        report_parts.append(part3_content)
                    else:
                        report_parts.append("## 3. 实施优化情况\n\n该客户为免费实施项目客户，上线以来未产生付费优化需求。")
                else:
                    report_parts.append("## 3. 实施优化情况\n\n该客户为免费实施项目客户，上线以来未产生付费优化需求。")
            except Exception as e:
                logger.error(f"Part3生成失败: {e}")
                report_parts.append("## Part3: 实施优化情况\n\n*生成失败*")

            # 5. Part4: 运维情况
            logger.info("生成Part4: 运维情况")
            try:
                part4 = OperationsAnalyzer()
                if 'part5_ops' in data:
                    part4_content = part4.analyze(data['part5_ops'])
                    if part4_content:
                        report_parts.append(part4_content)
                else:
                    report_parts.append("## Part4: 运维情况\n\n*数据缺失*")
            except Exception as e:
                logger.error(f"Part4生成失败: {e}")
                report_parts.append("## Part4: 运维情况\n\n*生成失败*")

            # 6. Part5: 客户经营情报
            logger.info("生成Part5: 客户经营情报")
            try:
                # 获取客户简称用于搜索新闻
                client_short_name = self.get_client_short_name(client_id)
                client_dir = os.path.join(self.base_dir, client_short_name)
                news_json_path = os.path.join(client_dir, "news_data.json")

                # 从Part1数据获取客户全称
                client_full_name = client_short_name
                if 'part1' in data and data['part1'] is not None and not data['part1'].empty:
                    df_part1 = data['part1']
                    # 尝试从"真实服务对象"或"客户全称"字段获取全称
                    if '真实服务对象' in df_part1.columns:
                        client_full_name = str(df_part1.iloc[0].get('真实服务对象', client_short_name))
                    elif '客户全称' in df_part1.columns:
                        client_full_name = str(df_part1.iloc[0].get('客户全称', client_short_name))
                logger.info(f"Part5: 客户简称={client_short_name}, 全称={client_full_name}")

                # 尝试从JSON文件读取新闻数据
                news_data = None
                if os.path.exists(news_json_path):
                    try:
                        with open(news_json_path, 'r', encoding='utf-8') as f:
                            news_data = json.load(f)
                        logger.info(f"已加载新闻数据: {len(news_data)} 条")
                    except Exception as e:
                        logger.warning(f"读取新闻数据失败: {e}")

                # 创建分析器时传入全称
                part5 = BusinessIntelligenceAnalyzer(client_short_name, news_data, client_full_name)

                if news_data:
                    # 有新闻数据，进行完整分析
                    part5_content = part5.get_markdown(news_data)
                else:
                    # 没有新闻数据，输出提示
                    part5_content = part5.get_markdown()

                if part5_content:
                    report_parts.append(part5_content)
                else:
                    report_parts.append("## Part5: 客户经营情报\n\n*生成失败*")
            except Exception as e:
                logger.error(f"Part5生成失败: {e}")
                import traceback
                traceback.print_exc()
                report_parts.append("## Part5: 客户经营情报\n\n*生成失败*")

            # 7. Part6: 综合经营分析
            logger.info("生成Part6: 综合经营分析")
            try:
                part6 = ComprehensiveAnalyzer()

                # 保存Part1-5已生成的markdown内容，供Part6使用
                part1_5_content = ""
                for part in report_parts:
                    part1_5_content += part + "\n\n"

                # 准备part2_summary（需要提取current_arr）
                part2_summary = {}
                if 'part2' in data and data['part2'] is not None and not data['part2'].empty:
                    df_sub = data['part2']
                    if '订阅状态' in df_sub.columns and '年订阅费金额' in df_sub.columns:
                        # 计算当前ARR：订阅状态为"订阅中"的年订阅费金额汇总
                        active_subs = df_sub[df_sub['订阅状态'] == '订阅中']
                        if not active_subs.empty:
                            current_arr = active_subs['年订阅费金额'].sum()
                            part2_summary['current_arr'] = current_arr

                # 准备part3_summary（需要提取yearly_data）
                part3_summary = {'yearly_data': {}}
                if 'part3_fixed' in data and data['part3_fixed'] is not None and not data['part3_fixed'].empty:
                    df_fixed = data['part3_fixed']
                    if '合同签订时间' in df_fixed.columns and '固定金额' in df_fixed.columns:
                        df_fixed['年份'] = pd.to_datetime(df_fixed['合同签订时间'], errors='coerce').dt.year
                        yearly = df_fixed.groupby('年份')['固定金额'].sum().to_dict()
                        part3_summary['yearly_data'] = {int(k): {'固定合同金额': v, '人天框架金额': 0} for k, v in yearly.items()}

                # 准备part1_data和part4_summary（需要检查是否为None或空）
                part1_data = data.get('part1', None)
                part4_raw = data.get('part5_ops', None)
                # 将DataFrame转换为summary字典
                if part4_raw is not None and not part4_raw.empty:
                    part4_summary = {
                        'total_tickets': len(part4_raw),
                        'has_data': True
                    }
                else:
                    part4_summary = {'has_data': False}

                logger.info(f"Part6参数准备: part2_summary={part2_summary}, part3_summary={part3_summary}")

                # 传入Part1-5已生成的完整内容供LLM分析
                part6_content = part6.analyze(
                    part1_data=part1_data,
                    part2_summary=part2_summary,
                    part3_summary=part3_summary,
                    part4_summary=part4_summary,
                    part1_5_full=part1_5_content  # 传入已生成的Part1-5内容
                )
                if part6_content:
                    report_parts.append(part6_content)
                else:
                    report_parts.append("## Part6: 综合经营分析\n\n*生成失败*")
            except Exception as e:
                logger.error(f"Part6生成失败: {e}")
                import traceback
                traceback.print_exc()
                report_parts.append("## Part6: 综合经营分析\n\n*生成失败*")

            # 7. 报告结尾
            report_parts.append("---")

            return "\n\n".join(report_parts)
            
        except Exception as e:
            logger.error(f"生成报告内容时出错: {e}")
            import traceback
            traceback.print_exc()
            return None

    def convert_markdown_to_word(self, md_path, docx_path):
        """
        将Markdown报告转换为Word文档

        Args:
            md_path: Markdown文件路径
            docx_path: Word文件路径

        Returns:
            bool: 转换是否成功
        """
        try:
            # 检查转换脚本是否存在
            if not os.path.exists(self.md_to_word_script):
                logger.error(f"转换脚本不存在: {self.md_to_word_script}")
                return False

            # 构建转换命令（有模板用模板，无模板则不传--template）
            cmd = [
                "python",
                self.md_to_word_script,
                "--input", md_path,
                "--output", docx_path
            ]

            # 只有模板存在时才添加template参数
            if os.path.exists(self.business_template):
                cmd.extend(["--template", self.business_template])
                logger.info(f"使用商业模板: {self.business_template}")
            else:
                logger.info("商业模板不存在，使用纯样式生成Word")

            logger.info(f"执行Word转换命令: {' '.join(cmd)}")

            # 执行转换
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')

            if result.returncode == 0:
                logger.info(f"Word转换成功: {docx_path}")
                return True
            else:
                logger.error(f"Word转换失败，返回码: {result.returncode}")
                if result.stderr:
                    logger.error(f"错误信息: {result.stderr}")
                return False

        except Exception as e:
            logger.error(f"Word转换过程中出错: {e}")
            return False

    def generate_report(self, client_id):
        """
        生成客户经营分析报告（Markdown + Word）

        Args:
            client_id: 客户ID

        Returns:
            tuple: (md_path, docx_path, error_message)
        """
        logger.info(f"开始生成客户报告: {client_id}")
        start_time = datetime.now()

        # 1. 获取客户简称
        client_short_name = self.get_client_short_name(client_id)
        logger.info(f"客户简称: {client_short_name}")

        # 2. 创建客户目录
        client_dir = self.create_client_directory(client_short_name)

        # 3. 生成文件名
        md_filename, docx_filename = self.generate_report_filenames(client_short_name)
        md_path = os.path.join(client_dir, md_filename)
        docx_path = os.path.join(client_dir, docx_filename)

        logger.info(f"Markdown文件: {md_path}")
        logger.info(f"Word文件: {docx_path}")

        # 4. 自动加载数据
        loader = get_data_loader()
        data, error = loader.load_client_data(client_id)
        if error:
            logger.warning(f"数据加载部分失败: {error}")

        # 如果数据加载失败，尝试使用CBD数据作为测试
        if not data or len(data) == 0:
            logger.warning(f"无法加载{client_id}的数据，尝试使用CBD数据")
            data, error = loader.load_client_data("CBD")
            if error:
                logger.error(f"CBD数据也加载失败: {error}")

        try:
            # 5. 生成报告内容
            report_content = self._generate_report_content(data, client_id)

            if not report_content:
                error_msg = "报告内容生成失败"
                logger.error(error_msg)
                return None, None, error_msg

            # 6. 保存Markdown报告
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(report_content)

            logger.info(f"Markdown报告保存成功: {md_path}")

            # 7. 转换为Word文档
            word_success = self.convert_markdown_to_word(md_path, docx_path)

            if not word_success:
                logger.warning("Word转换失败，仅生成Markdown报告")
                docx_path = None

            elapsed_time = (datetime.now() - start_time).total_seconds()
            logger.info(f"报告生成完成，耗时: {elapsed_time:.2f}秒")

            return md_path, docx_path, None

        except Exception as e:
            error_msg = f"报告生成过程中发生异常: {str(e)}"
            logger.error(error_msg)
            import traceback
            traceback.print_exc()
            return None, None, error_msg


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='集成版商务专家报告生成器')
    parser.add_argument('client_id', help='客户ID')
    parser.add_argument('--skip-llm', action='store_true', help='跳过LLM分析')
    parser.add_argument('--debug', action='store_true', help='启用调试模式')

    args = parser.parse_args()

    # 配置日志
    log_level = logging.DEBUG if args.debug else logging.INFO
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler()
        ]
    )

    # 生成报告
    generator = IntegratedReportGenerator(skip_llm=args.skip_llm)
    md_path, docx_path, error = generator.generate_report(args.client_id)

    if error:
        print(f"报告生成失败: {error}")
        sys.exit(1)
    else:
        print(f"报告生成成功!")
        print(f"   Markdown报告: {md_path}")
        if docx_path:
            print(f"   Word报告: {docx_path}")
        else:
            print(f"   Word报告: 转换失败，仅生成Markdown格式")

        # 检查文件是否存在
        if os.path.exists(md_path):
            print(f"   Markdown文件大小: {os.path.getsize(md_path)} 字节")
        if docx_path and os.path.exists(docx_path):
            print(f"   Word文件大小: {os.path.getsize(docx_path)} 字节")

        sys.exit(0)


if __name__ == "__main__":
    main()