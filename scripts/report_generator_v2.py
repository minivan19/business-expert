#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
报告生成器 v2.0 - 集成新输出路径和Word转换功能
负责协调各个分析模块，生成完整的经营分析报告（Markdown + Word）
"""

import os
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
from part6_comprehensive import ComprehensiveAnalyzer

logger = logging.getLogger(__name__)


class ReportGeneratorV2:
    """报告生成器类 v2.0"""
    
    def __init__(self, skip_llm=False):
        """
        初始化报告生成器 v2.0
        
        Args:
            skip_llm: 是否跳过LLM分析
        """
        self.skip_llm = skip_llm
        
        # 基础目录
        self.base_dir = r"C:\Users\mingh\client-data"
        
        # 临时目录（保留）
        self.temp_dir = os.path.join(self.base_dir, "_temp")
        
        # Markdown转Word技能路径
        self.md_to_word_skill_dir = r"C:\Users\mingh\.openclaw\workspace\skills\markdown-to-word-skill"
        self.converter_script = os.path.join(self.md_to_word_skill_dir, "scripts", "md2docx.py")
        self.business_template = os.path.join(self.md_to_word_skill_dir, "templates", "business.docx")
        
        # 创建基础目录
        os.makedirs(self.temp_dir, exist_ok=True)
        
        logger.info(f"报告生成器v2.0初始化")
        logger.info(f"基础目录: {self.base_dir}")
        logger.info(f"临时目录: {self.temp_dir}")
        logger.info(f"跳过LLM分析: {self.skip_llm}")
    
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
            
            # 查找客户简称字段（假设列名包含"简称"）
            short_name_col = None
            for col in df.columns:
                if "简称" in str(col):
                    short_name_col = col
                    break
            
            if short_name_col is None:
                logger.warning("未找到客户简称字段")
                return client_id
            
            # 查找指定客户的行
            # 这里简化处理：假设第一列是客户ID或名称，第一行是我们要的客户
            # 实际应用中需要根据业务逻辑精确匹配
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
            tuple: (md_filename, docx_filename, md_path, docx_path)
        """
        date_str = datetime.now().strftime("%Y%m%d")
        
        # 基础文件名
        base_name = f"{client_short_name}_经营分析报告_{date_str}"
        
        # 文件名
        md_filename = f"{base_name}.md"
        docx_filename = f"{base_name}.docx"
        
        return md_filename, docx_filename
    
    def _generate_report_content(self, client_id, data):
        """
        生成报告内容（内部方法）
        
        Args:
            client_id: 客户ID
            data: 加载的数据字典
            
        Returns:
            str: 报告内容
        """
        report_parts = []
        
        try:
            # 1. 报告标题
            report_parts.append(f"# {client_id} 经营分析报告")
            report_parts.append(f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            report_parts.append(f"**报告版本**: v2.0")
            report_parts.append("")
            
            # 2. Part1: 客户基础档案
            logger.info("生成Part1: 客户基础档案")
            part1 = BasicProfileAnalyzer()
            part1_content = part1.analyze(data, skip_llm=self.skip_llm)
            if part1_content:
                report_parts.append(part1_content)
            
            # 3. Part2: 订阅续约与续费
            logger.info("生成Part2: 订阅续约与续费")
            part2 = SubscriptionAnalyzer()
            part2_content = part2.analyze(data, skip_llm=self.skip_llm)
            if part2_content:
                report_parts.append(part2_content)
            
            # 4. Part3: 实施优化情况
            logger.info("生成Part3: 实施优化情况")
            part3 = ImplementationAnalyzer()
            part3_content = part3.analyze(data, skip_llm=self.skip_llm)
            if part3_content:
                report_parts.append(part3_content)
            
            # 5. Part4: 运维情况
            logger.info("生成Part4: 运维情况")
            part4 = OperationsAnalyzer()
            part4_content = part4.analyze(data, skip_llm=self.skip_llm)
            if part4_content:
                report_parts.append(part4_content)
            
            # 6. Part6: 综合经营分析
            logger.info("生成Part6: 综合经营分析")
            part6 = ComprehensiveAnalyzer()
            part6_content = part6.analyze(data, skip_llm=self.skip_llm)
            if part6_content:
                report_parts.append(part6_content)
            
            # 7. 报告结尾
            report_parts.append("---")
            report_parts.append("## 报告说明")
            report_parts.append("1. 本报告由商务专家系统自动生成")
            report_parts.append("2. 数据来源: 客户业务系统")
            report_parts.append("3. 分析时间: " + datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            report_parts.append("4. 报告版本: v2.0 (集成Word输出)")
            
            return "\n\n".join(report_parts)
            
        except Exception as e:
            logger.error(f"生成报告内容时出错: {e}")
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
            # 检查转换脚本和模板是否存在
            if not os.path.exists(self.converter_script):
                logger.error(f"转换脚本不存在: {self.converter_script}")
                return False
            
            if not os.path.exists(self.business_template):
                logger.error(f"商业模板不存在: {self.business_template}")
                return False
            
            # 构建转换命令
            cmd = [
                "python",
                self.converter_script,
                "--input", md_path,
                "--output", docx_path,
                "--template", self.business_template
            ]
            
            logger.info(f"执行Word转换命令: {' '.join(cmd)}")
            
            # 执行转换
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
            
            if result.returncode == 0:
                logger.info(f"Word转换成功: {docx_path}")
                if result.stdout:
                    logger.info(f"转换输出: {result.stdout}")
                return True
            else:
                logger.error(f"Word转换失败，返回码: {result.returncode}")
                if result.stderr:
                    logger.error(f"错误信息: {result.stderr}")
                return False
                
        except Exception as e:
            logger.error(f"Word转换过程中出错: {e}")
            return False
    
    def generate_report(self, client_id, data=None):
        """
        生成客户经营分析报告（Markdown + Word）
        
        Args:
            client_id: 客户ID
            data: 加载的数据字典（如果为None则自动加载）
            
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
        if data is None:
            loader = get_data_loader()
            data, error = loader.load_client_data(client_id)
            if error:
                logger.warning(f"数据加载部分失败: {error}")
        
        try:
            # 5. 生成报告内容
            report_content = self._generate_report_content(client_id, data)
            
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
                # 仍然返回成功，因为Markdown报告已生成
                docx_path = None
            
            elapsed_time = (datetime.now() - start_time).total_seconds()
            logger.info(f"报告生成完成，耗时: {elapsed_time:.2f}秒")
            
            return md_path, docx_path, None
            
        except Exception as e:
            error_msg = f"报告生成过程中发生异常: {str(e)}"
            logger.error(error_msg)
            return None, None, error_msg


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='商务专家报告生成器 v2.0')
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
            logging.StreamHandler(),
            logging.FileHandler(f'report_generator_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
        ]
    )
    
    # 生成报告
    generator = ReportGeneratorV2(skip_llm=args.skip_llm)
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
        sys.exit(0)


if __name__ == "__main__":
    main()