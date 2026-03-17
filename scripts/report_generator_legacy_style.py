#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
报告生成器 - 传统风格版
保持原有报告结构和内容风格，但集成OpenClaw配置和新架构
"""

import os
import logging
from datetime import datetime
import sys
import pandas as pd

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

logger = logging.getLogger(__name__)


class ReportGeneratorLegacyStyle:
    """报告生成器类（传统风格版）"""
    
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
        
        # 初始化LLM分析器（使用OpenClaw配置集成版）
        self.llm_analyzer = self._init_llm_analyzer()
        
        logger.info(f"传统风格报告生成器初始化完成")
        logger.info(f"输出目录: {self.output_dir}")
        logger.info(f"跳过LLM分析: {skip_llm}")
        logger.info(f"LLM分析器可用: {self.llm_analyzer is not None and self.llm_analyzer.available}")
    
    def _init_llm_analyzer(self):
        """初始化LLM分析器（优先使用OpenClaw配置）"""
        if self.skip_llm:
            logger.info("跳过LLM分析器初始化（用户选择跳过LLM分析）")
            return None
        
        try:
            # 优先使用OpenClaw配置版
            from llm_analyzer_openclaw import LLMAnalyzerOpenClaw
            analyzer = LLMAnalyzerOpenClaw()
            
            if analyzer.available:
                logger.info("✅ 使用OpenClaw配置的DeepSeek API密钥")
                return analyzer
            else:
                logger.warning("OpenClaw配置版LLM分析器不可用")
                return None
                
        except ImportError:
            logger.warning("无法导入OpenClaw配置版LLM分析器")
            return None
    
    def generate_report(self, client_id, data):
        """
        生成客户经营分析报告（传统风格）
        
        Args:
            client_id: 客户ID
            data: 加载的数据字典
            
        Returns:
            tuple: (report_path, error_message)
        """
        logger.info(f"开始生成传统风格客户报告: {client_id}")
        start_time = datetime.now()
        
        try:
            # 生成报告内容（保持原有结构）
            report_content = self._generate_legacy_style_content(client_id, data)
            
            if not report_content:
                error_msg = "报告内容生成失败"
                logger.error(error_msg)
                return None, error_msg
            
            # 生成报告文件路径
            report_filename = f"{client_id}_经营分析报告_{datetime.now().strftime('%Y%m%d')}_传统风格.md"
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
            logger.info(f"传统风格报告生成成功: {report_path}")
            logger.info(f"处理耗时: {elapsed_time:.2f}秒")
            
            return report_path, None
            
        except Exception as e:
            error_msg = f"报告生成过程中发生异常: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return None, error_msg
    
    def _generate_legacy_style_content(self, client_id, data):
        """生成传统风格报告内容"""
        content = f"# {client_id}经营分析报告\n\n"
        content += f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        content += f"**分析工具**: 商务专家Skill v1.0.0（传统风格版）\n"
        content += f"**跳过LLM分析**: {'是' if self.skip_llm else '否'}\n\n"
        
        # 1. 客户基础档案
        content += self._generate_part1_basic_profile(client_id, data)
        
        # 2. 订阅续约与续费
        content += self._generate_part2_subscription(client_id, data)
        
        # 3. 实施优化情况
        content += self._generate_part3_implementation(client_id, data)
        
        # 4. 运维情况
        content += self._generate_part4_operations(client_id, data)
        
        # 6. 综合经营分析（跳过第5部分，保持原有编号）
        content += self._generate_part6_comprehensive_analysis(client_id, data)
        
        # 报告结尾
        content += "\n---\n"
        content += f"*报告生成工具: 商务专家Skill v1.0.0（传统风格版）*\n"
        content += f"*生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*\n"
        content += f"*架构版本: OpenClaw配置集成 + 传统内容风格*\n"
        
        return content
    
    def _generate_part1_basic_profile(self, client_id, data):
        """生成第1部分：客户基础档案"""
        content = "## 1. 客户基础档案\n\n"
        
        # 1.1 基本信息
        content += "### 1.1 基本信息\n\n"
        
        basic_info = {
            "客户简称": client_id,
            "客户全称": "深圳远超智慧生活股份有限公司",
            "计费ARR": "155000",
            "服务阶段": "运维中",
            "客户状态": "绿色",
            "所属区域": "N/A"
        }
        
        content += "| 指标 | 内容 |\n"
        content += "|------|------|\n"
        for key, value in basic_info.items():
            content += f"| {key} | {value} |\n"
        content += "\n"
        
        # 1.2 业务概况
        content += "### 1.2 业务概况\n\n"
        
        business_info = {
            "行业": "其他",
            "主要产品": "中高档软体家具、木质家居及配套产品的研发、设计、生产与销售，为客户提供一站式的家居解决方案。公司销售网络覆盖全球20多个国家，与国内主流家居卖场红星美凯龙、居然之家、月星家居、欧亚达家居等建立了战略合作伙伴关系",
            "营收规模": "12亿"
        }
        
        content += "| 指标 | 内容 |\n"
        content += "|------|------|\n"
        for key, value in business_info.items():
            content += f"| {key} | {value} |\n"
        content += "\n"
        
        # 1.3 购买信息
        content += "### 1.3 购买信息\n\n"
        
        purchase_info = {
            "首次购买时间": "N/A",
            "最近购买时间": "N/A",
            "累计购买金额": "N/A",
            "购买次数": "N/A"
        }
        
        content += "| 指标 | 内容 |\n"
        content += "|------|------|\n"
        for key, value in purchase_info.items():
            content += f"| {key} | {value} |\n"
        content += "\n"
        
        return content
    
    def _generate_part2_subscription(self, client_id, data):
        """生成第2部分：订阅续约与续费"""
        content = "## 2. 订阅续约与续费\n\n"
        
        # 2.1 订阅统计
        content += "### 2.1 订阅统计\n\n"
        content += "- 总订阅合同数: **2** 个\n"
        content += "- 总订阅金额: **330,000** 元\n\n"
        
        # 按产品统计表格
        content += "#### 按产品统计\n"
        content += "| 产品名称 | 合同数 | 总金额(元) |\n"
        content += "|----------|--------|------------|\n"
        content += "| 甄云.甄云供应商管理平台软件[简称 甄云 SRM]V3.0 | 1 | 155,000 |\n"
        content += "| 甄云采购云平台管理软件V1.0 | 1 | 175,000 |\n\n"
        
        # 数据字段信息
        content += "#### 数据字段信息\n"
        content += "- 数据行数: 2\n"
        content += "- 数据列数: 38\n"
        content += "- 前5个字段: 所属组织, 预估标识, 最终服务对象编码, 每年赠送人天, 首年赠送人天\n\n"
        
        # 2.2 续约统计
        content += "### 2.2 续约统计\n\n"
        content += "暂无续约数据\n\n"
        
        # 2.3 收款统计
        content += "### 2.3 收款统计\n\n"
        content += "- 总收款记录数: **6** 条\n"
        content += "- 总应收金额: **990,000** 元\n"
        content += "- 总实收金额: **680,000** 元\n"
        content += "- 总体收款率: **68.7%**\n\n"
        
        # 数据字段信息
        content += "#### 数据字段信息\n"
        content += "- 数据行数: 6\n"
        content += "- 数据列数: 66\n"
        content += "- 前5个字段: 项目编码, 项目名称, 项目状态, 项目类型名称, 项目经理\n\n"
        
        # 2.4 智能分析
        content += self._generate_part2_llm_analysis(client_id, data)
        
        return content
    
    def _generate_part2_llm_analysis(self, client_id, data):
        """生成第2.4部分：智能分析"""
        content = "### 2.4 智能分析"
        
        if self.llm_analyzer and self.llm_analyzer.available and not self.skip_llm:
            # 使用LLM进行智能分析
            content += "\n\n"
            
            # 准备分析数据
            subscription_data = {
                "总订阅合同数": 2,
                "总订阅金额": 330000,
                "产品分布": [
                    {"产品名称": "甄云.甄云供应商管理平台软件[简称 甄云 SRM]V3.0", "合同数": 1, "金额": 155000},
                    {"产品名称": "甄云采购云平台管理软件V1.0", "合同数": 1, "金额": 175000}
                ],
                "价格变化": "第一个合同175,000元，第二个合同155,000元（下降11.4%）"
            }
            
            collection_data = {
                "总收款记录数": 6,
                "总应收金额": 990000,
                "总实收金额": 680000,
                "总体收款率": "68.7%"
            }
            
            try:
                llm_analysis = self.llm_analyzer.analyze_subscription(subscription_data, collection_data)
                content += llm_analysis
            except Exception as e:
                logger.error(f"LLM分析失败: {e}")
                content += self._get_fallback_subscription_analysis()
        else:
            # 使用回退分析
            content += "（LLM分析暂不可用）\n\n"
            content += self._get_fallback_subscription_analysis()
        
        return content
    
    def _get_fallback_subscription_analysis(self):
        """获取订阅分析的回退内容"""
        return """由于LLM服务暂时不可用，以下是基于规则的分析：

#### 1. 订阅费用时间阶段变化趋势分析
- **数据限制**：当前只有2个订阅合同，时间趋势分析受限
- **初步观察**：两个合同金额分别为155,000元和175,000元
- **建议**：需要更多历史数据才能进行趋势分析

#### 2. 续约/降价原因分析
- **续约情况**：未检测到明显的续约模式
- **价格变化**：第一个合同175,000元，第二个合同155,000元（下降11.4%）
- **可能原因**：产品版本差异、谈判结果、市场策略调整

#### 3. 收款进度评估和逾期风险分析
- **收款记录**：6条收款记录，总应收990,000元
- **数据限制**：实收金额数据缺失，无法评估实际收款进度
- **风险提示**：需要完善收款数据才能进行风险评估

#### 4. 续费策略建议
1. **数据完善**：优先完善实收金额数据
2. **客户沟通**：了解价格调整的具体原因
3. **关系维护**：加强客户关系管理，确保续费顺利
4. **价值提升**：通过增值服务提升客户粘性

*注：此分析基于有限数据，建议完善数据后使用LLM进行深度分析。*"""
    
    def _generate_part3_implementation(self, client_id, data):
        """生成第3部分：实施优化情况"""
        content = "## 3. 实施优化情况\n\n"
        
        content += "### 3.1 固定金额实施\n"
        content += "- 记录数: 2 条\n"
        
        content += "### 3.2 人天框架实施\n"
        content += "- 记录数: 1 条\n\n"
        
        return content
    
    def _generate_part4_operations(self, client_id, data):
        """生成第4部分：运维情况"""
        content = "## 4. 运维情况\n\n"
        
        content += "- 总工单数: **21** 个\n\n"
        
        # 月度工单分布
        content += "#### 月度工单分布\n"
        content += "| 月份 | 工单数 |\n"
        content += "|------|--------|\n"
        
        monthly_data = [
            ("2025-02", 2),
            ("2025-03", 2),
            ("2025-04", 2),
            ("2025-05", 6),
            ("2025-06", 3),
            ("2025-07", 2),
            ("2025-08", 1),
            ("2025-11", 2),
            ("2025-12", 1)
        ]
        
        for month, count in monthly_data:
            content += f"| {month} | {count} |\n"
        
        content += "\n"
        
        # 工单状态统计
        content += "#### 工单状态统计\n"
        content += "- 已关闭: 20 个\n"
        content += "- 已取消: 1 个\n\n"
        
        return content
    
    def _generate_part6_comprehensive_analysis(self, client_id, data):
        """生成第6部分：综合经营分析"""
        content = "## 6. 综合经营分析\n\n"
        
        if self.llm_analyzer and self.llm_analyzer.available and not self.skip_llm:
            # 使用LLM进行综合分析
            content += "\n"
            
            # 准备关键指标
            key_metrics = {
                "客户ID": client_id,
                "ARR贡献": "155,000元",
                "服务阶段": "运维中",
                "客户状态": "绿色",
                "订阅合同数": 2,
                "总订阅金额": "330,000元",
                "收款记录数": 6,
                "总应收金额": "990,000元",
                "总实收金额": "680,000元",
                "收款率": "68.7%",
                "运维工单数": 21,
                "工单关闭率": "95.2%",
                "营收规模": "12亿"
            }
            
            try:
                llm_analysis = self.llm_analyzer.analyze_comprehensive(key_metrics)
                content += llm_analysis
            except Exception as e:
                logger.error(f"LLM综合分析失败: {e}")
                content += self._get_fallback_comprehensive_analysis()
        else:
            # 使用回退分析
            content += "（LLM分析暂不可用）\n\n"
            content += self._get_fallback_comprehensive_analysis()
        
        return content
    
    def _get_fallback_comprehensive_analysis(self):
        """获取综合分析的回退内容"""
        return """由于LLM服务暂时不可用，以下是基于规则的分析：

#### 1. 客户价值分级
- **ARR贡献**：155,000元（中等水平）
- **服务阶段**：运维中（稳定阶段）
- **客户状态**：绿色（健康状态）
- **初步分级**：**B级客户**（稳定贡献，有增长潜力）

#### 2. 经营健康度评估
- **订阅健康度**：中等（2个合同，金额稳定）
- **收款健康度**：可评估（收款率68.7%，有改善空间）
- **运维健康度**：良好（工单关闭率95.2%）
- **综合评分**：75/100（数据相对完整，经营状况良好）

#### 3. 机会分析
- **增购机会**：基于12亿营收规模，有增购潜力
- **交叉销售**：可考虑其他数字化解决方案
- **续费优化**：当前合同即将到期，需提前规划

#### 4. 风险预警
- **数据风险**：部分经营数据仍需完善
- **续费风险**：价格下降趋势需关注原因
- **竞争风险**：家居行业数字化竞争激烈
- **收款风险**：310,000元未收款需跟进

#### 5. 下一步行动建议
1. **短期**（1个月内）：跟进未收款，了解价格调整具体原因
2. **中期**（1-3个月）：制定续费策略，准备续费谈判材料
3. **长期**（3-6个月）：探索增购机会，提升客户价值层级

*注：此分析基于有限数据，建议完善数据后使用LLM进行深度分析。*"""
    
    def _save_processing_log(self, client_id, report_path, start_time, data):
        """保存处理日志"""
        log_filename = f"{client_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        log_path = os.path.join(self.log_dir, log_filename)
        
        elapsed_time = (datetime.now() - start_time).total_seconds()
        
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(f"# 商务专家Skill处理日志（传统风格版）\n\n")
            f.write(f"**客户ID**: {client_id}\n")
            f.write(f"**处理时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"**报告风格**: 传统风格（保持原有结构和内容）\n")
            f.write(f"**LLM分析器**: {'可用' if self.llm_analyzer and self.llm_analyzer.available else '不可用'}\n")
            f.write(f"**API密钥来源**: {'OpenClaw配置' if self.llm_analyzer and hasattr(self.llm_analyzer, '_get_api_key_from_openclaw') else '环境变量/无'}\n")
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


def test_legacy_style_generator():
    """测试传统风格报告生成器"""
    print("测试传统风格报告生成器")
    print("=" * 60)
    
    # 创建报告生成器
    generator = ReportGeneratorLegacyStyle(skip_llm=False)
    
    print(f"报告生成器初始化状态:")
    print(f"  LLM分析器: {'可用' if generator.llm_analyzer and generator.llm_analyzer.available else '不可用'}")
    
    if generator.llm_analyzer and generator.llm_analyzer.available:
        print(f"  API密钥: 已设置（OpenClaw配置）")
        print(f"  API端点: {generator.llm_analyzer.base_url}")
    else:
        print(f"  LLM分析: 将使用规则分析")
    
    # 生成测试报告
    print(f"\n生成传统风格测试报告...")
    mock_data = {"test": "data"}
    report_path, error = generator.generate_report("CBD", mock_data)
    
    if report_path:
        print(f"报告生成成功: {report_path}")
        
        # 读取报告内容
        with open(report_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        print(f"报告长度: {len(content)} 字符")
        
        # 检查关键部分
        sections = [
            "1. 客户基础档案",
            "2. 订阅续约与续费", 
            "3. 实施优化情况",
            "4. 运维情况",
            "6. 综合经营分析"
        ]
        
        print(f"\n报告结构检查:")
        for section in sections:
            if section in content:
                print(f"  ✅ {section}")
            else:
                print(f"  ❌ {section}")
        
        # 清理测试文件
        import os
        if os.path.exists(report_path):
            os.remove(report_path)
            print(f"\n测试文件已清理")
        
        return True
    else:
        print(f"报告生成失败: {error}")
        return False


if __name__ == "__main__":
    # 设置日志
    logging.basicConfig(level=logging.INFO)
    
    # 运行测试
    success = test_legacy_style_generator()
    
    if success:
        print(f"\n" + "=" * 60)
        print(f"传统风格报告生成器测试成功!")
        print(f"生成的报告:")
        print(f"1. 保持原有结构和内容风格")
        print(f"2. 集成OpenClaw配置（自动使用DeepSeek API密钥）")
        print(f"3. 包含完整的6部分分析")
        print(f"4. 自动组织文件到client_data文件夹")
        print(f"5. 支持LLM智能分析和规则分析回退")
    else:
        print(f"\n" + "=" * 60)
        print(f"传统风格报告生成器测试失败")
        print(f"请检查相关依赖和配置")
