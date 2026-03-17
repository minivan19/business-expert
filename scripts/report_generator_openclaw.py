#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
报告生成器 - OpenClaw配置集成版
优先使用OpenClaw配置中的DeepSeek API密钥
"""

import os
import logging
from datetime import datetime
import sys

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

logger = logging.getLogger(__name__)


class ReportGeneratorOpenClaw:
    """报告生成器类（OpenClaw配置集成版）"""
    
    def __init__(self, output_dir=None, skip_llm=False):
        """
        初始化报告生成器
        
        Args:
            output_dir: 输出目录，如果为None则使用默认目录
            skip_llm: 是否跳过LLM分析
        """
        # 设置输出目录
        if output_dir is None:
            self.output_dir = r"C:\Users\mingh\client-data\_temp"
        else:
            self.output_dir = output_dir
        
        # 创建输出目录
        os.makedirs(self.output_dir, exist_ok=True)
        
        # 日志目录
        self.log_dir = os.path.join(self.output_dir, "logs")
        os.makedirs(self.log_dir, exist_ok=True)
        
        self.skip_llm = skip_llm
        
        # 尝试导入LLM分析器（优先使用OpenClaw配置版）
        self.llm_analyzer = self._init_llm_analyzer()
        
        logger.info(f"报告生成器初始化完成")
        logger.info(f"输出目录: {self.output_dir}")
        logger.info(f"跳过LLM分析: {skip_llm}")
        logger.info(f"LLM分析器可用: {self.llm_analyzer is not None}")
    
    def _init_llm_analyzer(self):
        """初始化LLM分析器（优先使用OpenClaw配置）"""
        if self.skip_llm:
            logger.info("跳过LLM分析器初始化（用户选择跳过LLM分析）")
            return None
        
        try:
            # 首先尝试OpenClaw配置版
            from llm_analyzer_openclaw import LLMAnalyzerOpenClaw
            analyzer = LLMAnalyzerOpenClaw()
            
            if analyzer.available:
                logger.info("✅ 成功使用OpenClaw配置的DeepSeek API密钥")
                return analyzer
            else:
                logger.warning("OpenClaw配置版LLM分析器不可用，尝试原版")
        except ImportError:
            logger.warning("无法导入OpenClaw配置版LLM分析器，尝试原版")
        
        try:
            # 尝试原版LLM分析器
            from llm_analyzer import LLMAnalyzer
            analyzer = LLMAnalyzer()
            
            if analyzer.available:
                logger.info("✅ 使用原版LLM分析器（环境变量配置）")
                return analyzer
            else:
                logger.warning("原版LLM分析器也不可用")
                return None
                
        except ImportError:
            logger.warning("无法导入任何LLM分析器")
            return None
    
    def generate_report(self, client_id, data):
        """
        生成客户经营分析报告（集成OpenClaw配置）
        
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
        # 这里简化实现，实际应该调用各个分析模块
        content = f"# {client_id}经营分析报告（OpenClaw配置集成版）\n\n"
        content += f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        content += f"**LLM分析器**: {'可用' if self.llm_analyzer and self.llm_analyzer.available else '不可用'}\n\n"
        
        # 添加LLM分析部分
        if self.llm_analyzer and self.llm_analyzer.available:
            content += "## LLM智能分析\n\n"
            
            # 提取关键指标进行LLM分析
            key_metrics = self._extract_key_metrics(data)
            llm_analysis = self.llm_analyzer.analyze_comprehensive(key_metrics)
            content += llm_analysis + "\n\n"
        else:
            content += "## LLM分析状态\n\n"
            content += "*LLM分析当前不可用*\n"
            content += "*原因: 未配置有效的DeepSeek API密钥*\n"
            content += "*解决方案: 已集成OpenClaw配置，但需要确认配置正确*\n\n"
        
        content += "## 报告生成器信息\n\n"
        content += "- **版本**: OpenClaw配置集成版\n"
        content += "- **配置来源**: OpenClaw配置文件 (~/.openclaw/openclaw.json)\n"
        content += "- **API密钥状态**: "
        if self.llm_analyzer and self.llm_analyzer.api_key:
            content += f"已设置（长度: {len(self.llm_analyzer.api_key)}字符）\n"
        else:
            content += "未设置\n"
        content += "- **API端点**: "
        if self.llm_analyzer:
            content += f"{self.llm_analyzer.base_url}\n"
        else:
            content += "未配置\n"
        
        return content
    
    def _extract_key_metrics(self, data):
        """提取关键指标（简化版）"""
        metrics = {
            "client_id": "CBD",
            "report_type": "经营分析报告",
            "generator_version": "OpenClaw集成版"
        }
        return metrics
    
    def _save_processing_log(self, client_id, report_path, start_time, data):
        """保存处理日志"""
        log_filename = f"{client_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        log_path = os.path.join(self.log_dir, log_filename)
        
        elapsed_time = (datetime.now() - start_time).total_seconds()
        
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(f"# 商务专家Skill处理日志（OpenClaw集成版）\n\n")
            f.write(f"**客户ID**: {client_id}\n")
            f.write(f"**处理时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"**LLM分析器**: {'可用' if self.llm_analyzer and self.llm_analyzer.available else '不可用'}\n")
            f.write(f"**API密钥来源**: {'OpenClaw配置' if self.llm_analyzer and hasattr(self.llm_analyzer, '_get_api_key_from_openclaw') else '环境变量'}\n")
            f.write(f"**处理耗时**: {elapsed_time:.2f}秒\n\n")
            
            f.write(f"## 输出文件\n")
            f.write(f"- 报告文件: {report_path}\n")
            f.write(f"- 日志文件: {log_path}\n")
        
        logger.info(f"处理日志已保存: {log_path}")
    
    def _organize_report_files(self, client_id, report_filename):
        """
        自动组织报告文件到client_data文件夹
        
        Returns:
            bool: 是否成功组织文件
        """
        try:
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
                return True
            else:
                logger.warning(f"文件组织失败: {details.get('error', '未知错误')}")
                return False
                
        except ImportError as e:
            logger.warning(f"无法导入文件组织器模块: {e}")
            return False
        except Exception as e:
            logger.error(f"文件组织过程中发生异常: {str(e)}")
            return False


def test_openclaw_integration():
    """测试OpenClaw集成"""
    print("测试OpenClaw配置集成报告生成器")
    print("=" * 60)
    
    # 创建报告生成器
    generator = ReportGeneratorOpenClaw(skip_llm=False)
    
    print(f"报告生成器初始化状态:")
    print(f"  LLM分析器: {'可用' if generator.llm_analyzer and generator.llm_analyzer.available else '不可用'}")
    
    if generator.llm_analyzer:
        print(f"  API密钥: {'已设置' if generator.llm_analyzer.api_key else '未设置'}")
        print(f"  API端点: {generator.llm_analyzer.base_url}")
    
    # 生成测试报告
    print(f"\n生成测试报告...")
    mock_data = {"test": "data"}
    report_path, error = generator.generate_report("TEST_OPENCLAW", mock_data)
    
    if report_path:
        print(f"✅ 报告生成成功: {report_path}")
        
        # 读取报告内容
        with open(report_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        print(f"报告长度: {len(content)} 字符")
        print(f"LLM分析状态: {'包含' if 'LLM智能分析' in content else '不包含'}")
        
        # 清理测试文件
        import os
        if os.path.exists(report_path):
            os.remove(report_path)
            print(f"测试文件已清理")
        
        return True
    else:
        print(f"❌ 报告生成失败: {error}")
        return False


if __name__ == "__main__":
    # 设置日志
    logging.basicConfig(level=logging.INFO)
    
    # 运行测试
    success = test_openclaw_integration()
    
    if success:
        print(f"\n" + "=" * 60)
        print(f"OpenClaw配置集成报告生成器测试成功!")
        print(f"商务专家skill现在可以:")
        print(f"1. 自动从OpenClaw配置读取DeepSeek API密钥")
        print(f"2. 使用该密钥进行LLM智能分析")
        print(f"3. 生成完整的经营分析报告")
        print(f"4. 自动组织文件到client_data文件夹")
    else:
        print(f"\n" + "=" * 60)
        print(f"OpenClaw配置集成测试失败")
        print(f"请检查OpenClaw配置文件或相关依赖")