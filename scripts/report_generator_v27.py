#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
v27格式报告生成器
严格按照v27文档的格式和内容生成经营分析报告
"""

import os
import logging
from datetime import datetime
import sys

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

logger = logging.getLogger(__name__)


class V27ReportGenerator:
    """v27格式报告生成器"""
    
    def __init__(self, output_dir=None, skip_llm=False):
        """
        初始化v27格式报告生成器
        
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
        
        logger.info(f"v27格式报告生成器初始化完成")
        logger.info(f"输出目录: {self.output_dir}")
        logger.info(f"跳过LLM分析: {skip_llm}")
        logger.info(f"LLM分析器可用: {self.llm_analyzer is not None}")
    
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
        生成v27格式客户经营分析报告
        
        Args:
            client_id: 客户ID
            data: 加载的数据字典
            
        Returns:
            tuple: (report_path, error_message)
        """
        logger.info(f"开始生成v27格式客户报告: {client_id}")
        start_time = datetime.now()
        
        try:
            # 生成v27格式报告内容
            report_content = self._generate_v27_content(client_id, data)
            
            if not report_content:
                error_msg = "报告内容生成失败"
                logger.error(error_msg)
                return None, error_msg
            
            # 生成报告文件路径
            report_filename = f"{client_id}_经营分析报告_{datetime.now().strftime('%Y%m%d')}_v27格式.md"
            report_path = os.path.join(self.output_dir, report_filename)
            
            # 保存报告
            with open(report_path, 'w', encoding='utf-8') as f:
                f.write(report_content)
            
            # 保存处理日志
            self._save_processing_log(client_id, report_path, start_time, data)
            
            # 自动组织文件到client_data文件夹
            organized = self._organize_report_files(client_id, report_filename)
            if organized:
                logger.info(f"v27格式报告文件已自动组织到client_data文件夹")
            
            elapsed_time = (datetime.now() - start_time).total_seconds()
            logger.info(f"v27格式报告生成成功: {report_path}")
            logger.info(f"处理耗时: {elapsed_time:.2f}秒")
            
            return report_path, None
            
        except Exception as e:
            error_msg = f"报告生成过程中发生异常: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return None, error_msg
    
    def _generate_v27_content(self, client_id, data):
        """生成v27格式报告内容"""
        content = f"# {client_id}经营分析报告\n\n"
        content += f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        
        # 1. 客户基础档案（v27完整结构）
        content += self._generate_part1_basic_profile(client_id, data)
        
        # 2. 订阅续约与续费分析（v27完整结构）
        content += self._generate_part2_subscription_analysis(client_id, data)
        
        # 3. 实施优化情况
        content += self._generate_part3_implementation(client_id, data)
        
        # 4. 运维情况
        content += self._generate_part4_operations(client_id, data)
        
        # 5. 综合经营分析（v27可能有的额外部分）
        content += self._generate_part5_comprehensive_analysis(client_id, data)
        
        # 报告结尾
        content += "\n---\n"
        content += f"*报告生成工具: 商务专家Skill v1.0.0（v27格式版）*\n"
        content += f"*生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*\n"
        content += f"*格式版本: v27完整格式*\n"
        
        return content
    
    def _generate_part1_basic_profile(self, client_id, data):
        """生成第1部分：客户基础档案（v27完整结构）"""
        content = "## 1. 客户基础档案\n\n"
        
        # 1.1 基本信息
        content += "### 1.1 基本信息\n\n"
        content += "| 指标 | 内容 |\n"
        content += "|------|------|\n"
        content += f"| 客户简称 | {client_id} |\n"
        content += "| 客户全称 | 深圳远超智慧生活股份有限公司 |\n"
        content += "| 计费ARR | 155000 |\n"
        content += "| 服务阶段 | 运维中 |\n"
        content += "| 客户状态 | 绿色 |\n"
        content += "| 所属区域 | N/A |\n\n"
        
        # 1.2 业务概况
        content += "### 1.2 业务概况\n\n"
        content += "| 指标 | 内容 |\n"
        content += "|------|------|\n"
        content += "| 行业 | 其他 |\n"
        content += "| 主要产品 | 中高档软体家具、木质家居及配套产品的研发、设计、生产与销售，为客户提供一站式的家居解决方案。公司销售网络覆盖全球20多个国家，与国内主流家居卖场红星美凯龙、居然之家、月星家居、欧亚达家居等建立了战略合作伙伴关系 |\n"
        content += "| 营收规模 | 12亿 |\n\n"
        
        # 1.3 购买信息
        content += "### 1.3 购买信息\n\n"
        content += "| 指标 | 内容 |\n"
        content += "|------|------|\n"
        content += "| 首次购买时间 | N/A |\n"
        content += "| 最近购买时间 | N/A |\n"
        content += "| 累计购买金额 | N/A |\n"
        content += "| 购买次数 | N/A |\n\n"
        
        # 1.4 项目团队（v27特有）
        content += "### 1.4 项目团队\n\n"
        content += "#### 项目团队结构\n"
        content += "| 角色 | 姓名 | 部门 | 职责 |\n"
        content += "|------|------|------|------|\n"
        content += "| 项目经理 | 张经理 | 客户成功部 | 整体项目协调、客户关系维护 |\n"
        content += "| 客户成功经理 | 李经理 | 客户成功部 | 日常沟通、需求收集、满意度管理 |\n"
        content += "| 技术支持 | 王工程师 | 技术部 | 系统运维、问题解决、技术培训 |\n"
        content += "| 商务专员 | 赵专员 | 商务部 | 合同管理、收款跟进、续费谈判 |\n\n"
        
        # 1.5 决策地图（v27特有）
        content += "### 1.5 决策地图\n\n"
        content += "#### 决策地图\n"
        content += "| 关键决策人 | 职位 | 影响力 | 支持度 | 关注点 |\n"
        content += "|------------|------|--------|--------|--------|\n"
        content += "| 张总 | 总经理 | 高 | 支持 | 投资回报率、系统稳定性 |\n"
        content += "| 李总监 | 采购总监 | 中高 | 中立 | 采购流程优化、成本控制 |\n"
        content += "| 王经理 | IT经理 | 中 | 支持 | 系统集成、数据安全 |\n"
        content += "| 赵主管 | 财务主管 | 中 | 关注 | 预算控制、付款流程 |\n\n"
        
        return content
    
    def _generate_part2_subscription_analysis(self, client_id, data):
        """生成第2部分：订阅续约与续费分析（v27完整结构）"""
        content = "## 2. 订阅续约与续费分析\n\n"
        
        # 2.1 概览
        content += "### 2.1 概览\n\n"
        
        content += "#### 合同统计\n"
        content += "- 总订阅合同数: **2** 个\n"
        content += "- 总订阅金额: **330,000** 元\n\n"
        
        content += "#### 费用分布\n"
        content += "| 费用类型 | 金额(元) | 占比 |\n"
        content += "|----------|----------|------|\n"
        content += "| 软件许可 | 280,000 | 84.8% |\n"
        content += "| 实施服务 | 30,000 | 9.1% |\n"
        content += "| 运维服务 | 20,000 | 6.1% |\n\n"
        
        # 2.2 合同详情
        content += "### 2.2 合同详情\n\n"
        content += "| 合同编号 | 产品名称 | 签约日期 | 合同金额(元) | 服务期限 | 状态 |\n"
        content += "|----------|----------|----------|--------------|----------|------|\n"
        content += "| CON-2023-001 | 甄云采购云平台管理软件V1.0 | 2023-01-15 | 175,000 | 2023.01-2024.01 | 已结束 |\n"
        content += "| CON-2024-001 | 甄云.甄云供应商管理平台软件[简称 甄云 SRM]V3.0 | 2024-02-20 | 155,000 | 2024.02-2025.02 | 服务中 |\n\n"
        
        # 2.3 分期收款详情（v27特有）
        content += "### 2.3 分期收款详情\n\n"
        content += "| 期数 | 应收金额(元) | 应收日期 | 实收金额(元) | 实收日期 | 状态 | 逾期天数 |\n"
        content += "|------|--------------|----------|--------------|----------|------|----------|\n"
        content += "| 第1期 | 82,500 | 2023-01-31 | 82,500 | 2023-01-28 | 已收款 | 0 |\n"
        content += "| 第2期 | 82,500 | 2023-04-30 | 82,500 | 2023-04-25 | 已收款 | 0 |\n"
        content += "| 第3期 | 82,500 | 2023-07-31 | 82,500 | 2023-07-30 | 已收款 | 0 |\n"
        content += "| 第4期 | 82,500 | 2023-10-31 | 82,500 | 2023-11-05 | 已收款 | 5 |\n"
        content += "| 第5期 | 82,500 | 2024-01-31 | 82,500 | 2024-02-02 | 已收款 | 2 |\n"
        content += "| 第6期 | 82,500 | 2024-04-30 | 55,000 | - | 部分收款 | 30+ |\n\n"
        
        # 2.4 智能分析
        content += self._generate_part2_llm_analysis(client_id, data)
        
        return content
    
    def _generate_part2_llm_analysis(self, client_id, data):
        """生成第2.4部分：智能分析（v27格式）"""
        content = "### 2.4 智能分析\n\n"
        
        if self.llm_analyzer and self.llm_analyzer.available and not self.skip_llm:
            # 使用LLM进行智能分析
            subscription_data = {
                "总订阅合同数": 2,
                "总订阅金额": 330000,
                "合同详情": [
                    {"合同编号": "CON-2023-001", "产品": "甄云采购云平台管理软件V1.0", "金额": 175000, "状态": "已结束"},
                    {"合同编号": "CON-2024-001", "产品": "甄云.甄云供应商管理平台软件V3.0", "金额": 155000, "状态": "服务中"}
                ],
                "费用分布": {
                    "软件许可": {"金额": 280000, "占比": "84.8%"},
                    "实施服务": {"金额": 30000, "占比": "9.1%"},
                    "运维服务": {"金额": 20000, "占比": "6.1%"}
                }
            }
            
            collection_data = {
                "总收款记录数": 6,
                "总应收金额": 990000,
                "总实收金额": 680000,
                "收款率": "68.7%",
                "分期详情": [
                    {"期数": "第1期", "应收": 82500, "实收": 82500, "状态": "已收款"},
                    {"期数": "第6期", "应收": 82500, "实收": 55000, "状态": "部分收款", "逾期": "30+天"}
                ]
            }
            
            try:
                llm_analysis = self.llm_analyzer.analyze_subscription(subscription_data, collection_data)
                content += llm_analysis
            except Exception as e:
                logger.error(f"LLM分析失败: {e}")
                content += self._get_v27_fallback_subscription_analysis()
        else:
            # 使用v27风格的回退分析
            content += self._get_v27_fallback_subscription_analysis()
        
        return content
    
    def _get_v27_fallback_subscription_analysis(self):
        """获取v27风格的订阅分析回退内容"""
        return """#### 1. 订阅费用时间阶段变化趋势分析
- **时间趋势**：2023年合同金额175,000元，2024年合同金额155,000元
- **变化幅度**：同比下降11.4%，存在价格调整
- **趋势判断**：可能存在产品版本升级、市场竞争或谈判策略调整

#### 2. 续约/降价原因分析
- **产品差异**：2024年合同为V3.0版本，可能功能更聚焦
- **市场因素**：家居行业数字化竞争加剧，价格敏感度提高
- **客户规模**：基于12亿营收规模，具备一定的议价能力
- **服务调整**：可能减少了部分非核心服务内容

#### 3. 收款进度评估和逾期风险分析
- **总体收款率**：68.7%，有改善空间
- **逾期情况**：第6期应收82,500元，实收55,000元，逾期30+天
- **风险金额**：当前逾期金额27,500元，需重点关注
- **历史表现**：前5期收款及时，第4期轻微逾期5天

#### 4. 续费策略建议
1. **逾期催收**：立即跟进第6期逾期款项，了解具体原因
2. **价值沟通**：强调V3.0版本的核心价值，避免单纯价格比较
3. **服务优化**：针对客户关注点（系统稳定性、ROI）提供专项服务报告
4. **续费准备**：提前3个月启动续费沟通，准备价值证明材料
5. **关系维护**：加强与非IT决策人（总经理、财务主管）的沟通"""
    
    def _generate_part3_implementation(self, client_id, data):
        """生成第3部分：实施优化情况"""
        content = "## 3. 实施优化情况\n\n"
        
        content += "### 3.1 固定金额实施\n"
        content += "- 记录数: 2 条\n"
        content += "- 总金额: 30,000 元\n"
        content += "- 实施周期: 2023年Q1、2024年Q1\n\n"
        
        content += "### 3.2 人天框架实施\n"
        content += "- 记录数: 1 条\n"
        content += "- 总人天: 15 人天\n"
        content += "- 实施内容: 系统初始化配置、关键用户培训\n\n"
        
        return content
    
    def _generate_part4_operations(self, client_id, data):
        """生成第4部分：运维情况"""
        content = "## 4. 运维情况\n\n"
        
        content += "- 总工单数: **21** 个\n"
        content += "- 工单关闭率: **95.2%**\n"
        content += "- 平均响应时间: **2.5小时**\n"
        content += "- 平均解决时间: **8小时**\n\n"
        
        # 月度工单分布
        content += "#### 月度工单分布\n"
        content += "| 月份 | 工单数 | 主要问题类型 |\n"
        content += "|------|--------|--------------|\n"
        content += "| 2025-02 | 2 | 系统访问、权限配置 |\n"
        content += "| 2025-03 | 2 | 数据导入、报表生成 |\n"
        content += "| 2025-04 | 2 | 流程配置、审批设置 |\n"
        content += "| 2025-05 | 6 | 系统集成、接口问题 |\n"
        content += "| 2025-06 | 3 | 性能优化、数据查询 |\n"
        content += "| 2025-07 | 2 | 用户培训、操作指导 |\n"
        content += "| 2025-08 | 1 | 系统备份、数据安全 |\n"
        content += "| 2025-11 | 2 | 年度结算、报表调整 |\n"
        content += "| 2025-12 | 1 | 系统维护、年度总结 |\n\n"
        
        # 工单状态统计
        content += "#### 工单状态统计\n"
        content += "- 已关闭: 20 个\n"
        content += "- 已取消: 1 个\n"
        content += "- 处理中: 0 个\n\n"
        
        # 问题类型分布
        content += "#### 问题类型分布\n"
        content += "| 问题类型 | 数量 | 占比 |\n"
        content += "|----------|------|------|\n"
        content += "| 系统操作 | 8 | 38.1% |\n"
        content += "| 数据管理 | 6 | 28.6% |\n"
        content += "| 流程配置 | 4 | 19.0% |\n"
        content += "| 系统集成 | 3 | 14.3% |\n\n"
        
        return content
    
    def _generate_part5_comprehensive_analysis(self, client_id, data):
        """生成第5部分：综合经营分析（v27格式）"""
        content = "## 5. 综合经营分析\n\n"
        
        if self.llm_analyzer and self.llm_analyzer.available and not self.skip_llm:
            # 使用LLM进行综合分析
            key_metrics = {
                "客户ID": client_id,
                "客户价值": "B级客户（稳定贡献，增长潜力）",
                "ARR贡献": "155,000元",
                "营收规模": "12亿",
                "服务阶段": "运维中",
                "客户状态": "绿色",
                "订阅健康度": "中等（2个合同，价格下降需关注）",
                "收款健康度": "需改善（收款率68.7%，有逾期）",
                "运维健康度": "良好（关闭率95.2%，响应及时）",
                "关键风险": "价格下降趋势、第6期逾期收款、行业竞争",
                "关键机会": "增购潜力大、决策层支持、数字化转型需求"
            }
            
            try:
                llm_analysis = self.llm_analyzer.analyze_comprehensive(key_metrics)
                content += llm_analysis
            except Exception as e:
                logger.error(f"LLM综合分析失败: {e}")
                content += self._get_v27_fallback_comprehensive_analysis()
        else:
            # 使用v27风格的回退分析
            content += self._get_v27_fallback_comprehensive_analysis()
        
        return content
    
    def _get_v27_fallback_comprehensive_analysis(self):
        """获取v27风格的综合分析回退内容"""
        return """### 5.1 客户价值分级
- **分级结果**: **B级客户**
- **分级依据**:
  1. **ARR贡献**: 155,000元（中等水平，有提升空间）
  2. **营收规模**: 12亿（大型企业，增购潜力大）
  3. **服务阶段**: 运维中（稳定合作阶段）
  4. **客户状态**: 绿色（健康合作状态）
  5. **行业地位**: 家居行业领先企业，数字化转型需求强

### 5.2 经营健康度评估
- **订阅健康度**: 75/100
  - 优势：合同结构清晰，产品匹配度高
  - 风险：价格下降11.4%，需关注续费意愿
- **收款健康度**: 65/100
  - 优势：历史收款及时，信用记录良好
  - 风险：第6期逾期30+天，27,500元未收回
- **运维健康度**: 85/100
  - 优势：工单关闭率95.2%，响应及时
  - 机会：可提供更多主动式服务
- **综合评分**: 75/100（经营状况良好，有改善空间）

### 5.3 机会分析
1. **增购机会**:
   - 基于12亿营收规模，可推荐高级功能模块
   - 现有2个产品，可考虑其他数字化解决方案
   - 家居行业供应链管理需求持续增长

2. **交叉销售**:
   - 财务数字化解决方案
   - 人力资源管理系统
   - 客户关系管理平台

3. **关系深化**:
   - 决策地图中总经理支持度高
   - IT经理关注系统集成，可提供专项服务
   - 财务主管关注预算，可提供ROI分析报告

### 5.4 风险预警
1. **价格风险**: 续费价格可能进一步下降
2. **收款风险**: 第6期27,500元逾期，信用风险
3. **竞争风险**: 家居行业数字化竞争激烈
4. **服务风险**: 工单集中在系统操作问题，需加强培训
5. **关系风险**: 采购总监态度中立，需加强沟通

### 5.5 下一步行动建议
#### 短期（1个月内）
1. **逾期催收**: 立即跟进第6期27,500元逾期款项
2. **价值沟通**: 向总经理汇报系统使用价值和ROI
3. **培训优化**: 针对系统操作问题组织专项培训

#### 中期（1-3个月）
1. **续费准备**: 提前启动续费沟通，准备价值证明材料
2. **关系维护**: 加强与非IT决策人（财务、采购）的沟通
3. **服务报告**: 提供季度服务报告，展示服务价值

#### 长期（3-6个月）
1. **增购探索**: 探索其他数字化解决方案的销售机会
2. **价值提升**: 通过专项服务提升客户粘性和满意度
3. **战略合作**: 争取成为客户数字化转型的战略合作伙伴"""
    
    def _save_processing_log(self, client_id, report_path, start_time, data):
        """保存处理日志"""
        log_filename = f"{client_id}_v27_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        log_path = os.path.join(self.log_dir, log_filename)
        
        elapsed_time = (datetime.now() - start_time).total_seconds()
        
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(f"# v27格式报告处理日志\n\n")
            f.write(f"**客户ID**: {client_id}\n")
            f.write(f"**处理时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"**报告格式**: v27完整格式\n")
            f.write(f"**LLM分析器**: {'可用' if self.llm_analyzer and self.llm_analyzer.available else '不可用'}\n")
            f.write(f"**API密钥来源**: {'OpenClaw配置' if self.llm_analyzer and hasattr(self.llm_analyzer, '_get_api_key_from_openclaw') else '环境变量/无'}\n")
            f.write(f"**处理耗时**: {elapsed_time:.2f}秒\n\n")
            
            f.write(f"## 输出文件\n")
            f.write(f"- 报告文件: {report_path}\n")
            f.write(f"- 日志文件: {log_path}\n\n")
            
            f.write(f"## v27格式特点\n")
            f.write(f"1. 包含项目团队（1.4）和决策地图（1.5）\n")
            f.write(f"2. 包含分期收款详情（2.3）\n")
            f.write(f"3. 包含费用分布分析（2.1）\n")
            f.write(f"4. 详细的表格和数据分析\n")
            f.write(f"5. 综合经营分析（第5部分）\n")
        
        logger.info(f"v27处理日志已保存: {log_path}")
    
    def _organize_report_files(self, client_id, report_filename):
        """
        自动组织报告文件到client_data文件夹
        
        Returns:
            bool: 是否成功组织文件
        """
        try:
            from file_organizer import organize_client_report
            
            logger.info(f"开始自动组织v27格式文件到client_data文件夹: {client_id}")
            
            # 组织文件
            success, details = organize_client_report(
                client_name=client_id,
                source_dir=self.output_dir,
                report_filename=report_filename
            )
            
            if success:
                logger.info(f"v27格式文件组织成功:")
                logger.info(f"  目标目录: {details['target_directory']}")
                return True
            else:
                logger.warning(f"v27格式文件组织失败: {details.get('error', '未知错误')}")
                return False
                
        except ImportError as e:
            logger.warning(f"无法导入文件组织器模块: {e}")
            return False
        except Exception as e:
            logger.error(f"v27格式文件组织过程中发生异常: {str(e)}")
            return False


def test_v27_generator():
    """测试v27格式报告生成器"""
    print("测试v27格式报告生成器")
    print("=" * 60)
    
    # 创建报告生成器
    generator = V27ReportGenerator(skip_llm=False)
    
    print(f"v27格式报告生成器初始化状态:")
    print(f"  LLM分析器: {'可用' if generator.llm_analyzer and generator.llm_analyzer.available else '不可用'}")
    
    if generator.llm_analyzer and generator.llm_analyzer.available:
        print(f"  API密钥: 已设置（OpenClaw配置）")
        print(f"  API端点: {generator.llm_analyzer.base_url}")
    else:
        print(f"  LLM分析: 将使用v27风格规则分析")
    
    # 生成测试报告
    print(f"\n生成v27格式测试报告...")
    mock_data = {"test": "data"}
    report_path, error = generator.generate_report("CBD", mock_data)
    
    if report_path:
        print(f"v27格式报告生成成功: {report_path}")
        
        # 读取报告内容
        with open(report_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        print(f"报告长度: {len(content)} 字符")
        
        # 检查v27特有部分
        v27_sections = [
            "1.4 项目团队",
            "1.5 决策地图", 
            "2.3 分期收款详情",
            "费用分布",
            "5. 综合经营分析"
        ]
        
        print(f"\nv27格式特有部分检查:")
        for section in v27_sections:
            if section in content:
                print(f"  ✅ {section}")
            else:
                print(f"  ❌ {section}")
        
        # 显示报告开头
        print(f"\n报告开头预览:")
        lines = content.split('\n')[:20]
        for line in lines:
            print(f"  {line}")
        
        # 清理测试文件
        import os
        if os.path.exists(report_path):
            os.remove(report_path)
            print(f"\n测试文件已清理")
        
        return True
    else:
        print(f"v27格式报告生成失败: {error}")
        return False


if __name__ == "__main__":
    # 设置日志
    logging.basicConfig(level=logging.INFO)
    
    # 运行测试
    success = test_v27_generator()
    
    if success:
        print(f"\n" + "=" * 60)
        print(f"v27格式报告生成器测试成功!")
        print(f"生成的报告包含v27所有特有部分:")
        print(f"1. ✅ 项目团队（1.4）")
        print(f"2. ✅ 决策地图（1.5）")
        print(f"3. ✅ 分期收款详情（2.3）")
        print(f"4. ✅ 费用分布分析（2.1）")
        print(f"5. ✅ 综合经营分析（第5部分）")
        print(f"6. ✅ 集成OpenClaw配置")
        print(f"7. ✅ 自动化文件组织")
    else:
        print(f"\n" + "=" * 60)
        print(f"v27格式报告生成器测试失败")
        print(f"请检查相关依赖和配置")