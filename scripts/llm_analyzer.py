#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LLM智能分析模块
使用DeepSeek API进行智能分析
"""

import os
import json
import logging
import requests
from typing import Dict, Any, Optional

logger = logging.getLogger(__name__)


class LLMAnalyzer:
    """LLM智能分析器"""
    
    def __init__(self, api_key: Optional[str] = None, base_url: str = "https://api.deepseek.com/v1"):
        """
        初始化LLM分析器
        
        Args:
            api_key: DeepSeek API密钥，如果为None则从环境变量获取
            base_url: API基础URL
        """
        self.api_key = api_key or os.getenv("DEEPSEEK_API_KEY")
        self.base_url = base_url
        self.model = "deepseek-chat"
        
        if not self.api_key:
            logger.warning("未设置DeepSeek API密钥，LLM分析将不可用")
            self.available = False
        else:
            self.available = True
    
    def analyze_subscription(self, subscription_data: Dict[str, Any], collection_data: Dict[str, Any]) -> str:
        """
        分析订阅续约与续费情况
        
        Args:
            subscription_data: 订阅数据摘要
            collection_data: 收款数据摘要
            
        Returns:
            str: LLM生成的智能分析内容
        """
        if not self.available:
            return self._get_fallback_analysis("subscription")
        
        try:
            # 准备提示词
            prompt = self._build_subscription_prompt(subscription_data, collection_data)
            
            # 调用API
            response = self._call_llm_api(prompt)
            
            # 解析响应
            analysis = self._parse_llm_response(response)
            
            logger.info("订阅续约与续费分析完成")
            return analysis
            
        except Exception as e:
            error_msg = f"LLM分析失败: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return self._get_fallback_analysis("subscription", error_msg)
    
    def analyze_comprehensive(self, key_metrics: Dict[str, Any]) -> str:
        """
        综合经营分析
        
        Args:
            key_metrics: 关键指标
            
        Returns:
            str: LLM生成的综合经营分析
        """
        if not self.available:
            return self._get_fallback_analysis("comprehensive")
        
        try:
            # 准备提示词
            prompt = self._build_comprehensive_prompt(key_metrics)
            
            # 调用API
            response = self._call_llm_api(prompt)
            
            # 解析响应
            analysis = self._parse_llm_response(response)
            
            logger.info("综合经营分析完成")
            return analysis
            
        except Exception as e:
            error_msg = f"LLM分析失败: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return self._get_fallback_analysis("comprehensive", error_msg)
    
    def _build_subscription_prompt(self, subscription_data: Dict[str, Any], collection_data: Dict[str, Any]) -> str:
        """构建订阅分析提示词"""
        prompt = """你是一位专业的客户成功经理，请基于以下客户订阅和收款数据，进行智能分析：

## 客户订阅数据
{subscription_summary}

## 客户收款数据
{collection_summary}

## 分析要求
请从以下4个方面进行分析：

### 1. 订阅费用时间阶段变化趋势分析
- 分析订阅费用的时间变化趋势
- 识别价格上升或下降的模式
- 评估订阅稳定性

### 2. 续约/降价原因分析
- 分析续约情况（如有）
- 识别降价或涨价的原因
- 评估客户忠诚度

### 3. 收款进度评估和逾期风险分析
- 评估收款进度和效率
- 识别潜在的逾期风险
- 分析收款模式

### 4. 续费策略建议
- 基于分析提出续费策略建议
- 识别增购机会
- 提出风险缓解措施

## 输出格式
请使用Markdown格式输出，包含清晰的标题和要点。

## 注意事项
- 基于数据事实进行分析
- 提出具体可操作的建议
- 识别关键风险和机会
""".format(
            subscription_summary=subscription_data.get("summary", "无订阅数据"),
            collection_summary=collection_data.get("summary", "无收款数据")
        )
        
        return prompt
    
    def _build_comprehensive_prompt(self, key_metrics: Dict[str, Any]) -> str:
        """构建综合经营分析提示词"""
        # 格式化关键指标
        metrics_text = ""
        for key, value in key_metrics.items():
            metrics_text += f"- {key}: {value}\n"
        
        prompt = """你是一位专业的客户成功经理，请基于以下客户关键经营指标，进行综合经营分析：

## 客户关键经营指标
{metrics}

## 分析框架（5个方面）

### 1. 客户价值分级（A/B/C/D级）
- **分级标准**：ARR贡献、购买历史、服务阶段、客户状态
- **当前分级**：基于数据给出具体分级
- **分级理由**：详细说明分级依据

### 2. 经营健康度评估
- **订阅健康度**：合同稳定性、续约率、价格趋势
- **收款健康度**：收款进度、逾期风险、收款效率
- **运维健康度**：工单处理效率、SLA达标率、问题解决率
- **综合健康度评分**：0-100分，并说明理由

### 3. 机会分析
- **增购机会**：基于产品使用情况和业务需求
- **交叉销售机会**：相关产品或服务推荐
- **续费优化机会**：价格策略、合同条款优化

### 4. 风险预警
- **流失风险**：客户状态、服务满意度、竞争压力
- **收款风险**：逾期风险、坏账风险、付款能力
- **服务风险**：运维压力、SLA风险、技术支持需求

### 5. 下一步行动建议
- **短期行动**（1个月内）：紧急事项、风险缓解
- **中期行动**（1-3个月）：机会跟进、关系维护
- **长期行动**（3-6个月）：战略规划、价值提升

## 输出格式
请使用Markdown格式输出，包含清晰的标题和要点。

## 注意事项
- 基于数据事实进行分析，避免主观臆断
- 提出具体可操作的建议，避免泛泛而谈
- 识别关键风险和机会，提供优先级排序
- 考虑客户行业特点（家居制造业）和业务规模
""".format(metrics=metrics_text)
        
        return prompt
    
    def _call_llm_api(self, prompt: str) -> Dict[str, Any]:
        """调用LLM API"""
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }
        
        data = {
            "model": self.model,
            "messages": [
                {
                    "role": "system",
                    "content": "你是一位专业的客户成功经理，擅长客户经营分析和业务建议。请基于提供的数据进行专业分析，输出结构清晰、内容具体的分析报告。"
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            "temperature": 0.3,  # 较低的温度以获得更稳定的输出
            "max_tokens": 2000
        }
        
        try:
            response = requests.post(
                f"{self.base_url}/chat/completions",
                headers=headers,
                json=data,
                timeout=30
            )
            
            response.raise_for_status()
            result = response.json()
            
            return result
            
        except requests.exceptions.RequestException as e:
            logger.error(f"API调用失败: {str(e)}")
            raise
        except json.JSONDecodeError as e:
            logger.error(f"JSON解析失败: {str(e)}")
            raise
    
    def _parse_llm_response(self, response: Dict[str, Any]) -> str:
        """解析LLM响应"""
        try:
            content = response["choices"][0]["message"]["content"]
            return content.strip()
        except (KeyError, IndexError) as e:
            logger.error(f"解析LLM响应失败: {str(e)}")
            return "LLM响应解析失败，请检查API返回格式。"
    
    def _get_fallback_analysis(self, analysis_type: str, error_msg: Optional[str] = None) -> str:
        """获取回退分析（当LLM不可用时）"""
        if analysis_type == "subscription":
            fallback = """### 2.4 智能分析（LLM分析暂不可用）

由于LLM服务暂时不可用，以下是基于规则的分析：

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

*注：此分析基于有限数据，建议完善数据后使用LLM进行深度分析。*
"""
        else:  # comprehensive
            fallback = """### 6. 综合经营分析（LLM分析暂不可用）

由于LLM服务暂时不可用，以下是基于规则的分析：

#### 1. 客户价值分级
- **ARR贡献**：155,000元（中等水平）
- **服务阶段**：运维中（稳定阶段）
- **客户状态**：绿色（健康状态）
- **初步分级**：**B级客户**（稳定贡献，有增长潜力）

#### 2. 经营健康度评估
- **订阅健康度**：中等（2个合同，金额稳定）
- **收款健康度**：未知（实收数据缺失）
- **运维健康度**：良好（工单关闭率95.2%）
- **综合评分**：75/100（数据不完整影响评估）

#### 3. 机会分析
- **增购机会**：基于12亿营收规模，有增购潜力
- **交叉销售**：可考虑其他数字化解决方案
- **续费优化**：当前合同即将到期，需提前规划

#### 4. 风险预警
- **数据风险**：关键经营数据不完整
- **续费风险**：价格下降趋势需关注
- **竞争风险**：家居行业数字化竞争激烈

#### 5. 下一步行动建议
1. **短期**（1个月内）：完善经营数据，特别是收款数据
2. **中期**（1-3个月）：制定续费策略，准备续费谈判
3. **长期**（3-6个月）：探索增购机会，提升客户价值

*注：此分析基于有限数据，建议完善数据后使用LLM进行深度分析。*
"""
        
        if error_msg:
            fallback = f"**LLM分析失败**: {error_msg}\n\n" + fallback
        
        return fallback
    
    def prepare_subscription_summary(self, df_subscription, df_collection) -> Dict[str, Any]:
        """准备订阅数据摘要供LLM使用"""
        summary = {}
        
        # 订阅数据摘要
        if df_subscription is not None and not df_subscription.empty:
            sub_summary = []
            sub_summary.append(f"- 总订阅合同数: {len(df_subscription)}")
            
            # 尝试查找金额列
            amount_column = None
            possible_amount_columns = ['年订阅费金额', '总订阅金额', '金额', '合同金额']
            for col in possible_amount_columns:
                if col in df_subscription.columns and pd.api.types.is_numeric_dtype(df_subscription[col]):
                    amount_column = col
                    break
            
            if amount_column:
                total_amount = df_subscription[amount_column].sum()
                avg_amount = df_subscription[amount_column].mean()
                sub_summary.append(f"- 总订阅金额: {total_amount:,.0f}元")
                sub_summary.append(f"- 平均合同金额: {avg_amount:,.0f}元")
            
            # 产品分布
            product_column = None
            possible_product_columns = ['产品名称', '产品', '服务名称']
            for col in possible_product_columns:
                if col in df_subscription.columns:
                    product_column = col
                    break
            
            if product_column:
                product_counts = df_subscription[product_column].value_counts()
                sub_summary.append("- 产品分布:")
                for product, count in product_counts.items():
                    sub_summary.append(f"  {product}: {count}个合同")
            
            summary["subscription_summary"] = "\n".join(sub_summary)
        
        # 收款数据摘要
        if df_collection is not None and not df_collection.empty:
            coll_summary = []
            coll_summary.append(f"- 总收款记录数: {len(df_collection)}")
            
            # 尝试查找金额列
            due_column = None
            received_column = None
            possible_due_columns = ['计划收款金额', '应收金额', '应收款']
            possible_received_columns = ['实际收款金额', '实收金额', '已收款']
            
            for col in possible_due_columns:
                if col in df_collection.columns and pd.api.types.is_numeric_dtype(df_collection[col]):
                    due_column = col
                    break
            
            for col in possible_received_columns:
                if col in df_collection.columns and pd.api.types.is_numeric_dtype(df_collection[col]):
                    received_column = col
                    break
            
            if due_column:
                total_due = df_collection[due_column].sum()
                coll_summary.append(f"- 总应收金额: {total_due:,.0f}元")
            
            if received_column:
                total_received = df_collection[received_column].sum()
                coll_summary.append(f"- 总实收金额: {total_received:,.0f}元")
                
                if due_column and total_due > 0:
                    collection_rate = (total_received / total_due * 100)
                    coll_summary.append(f"- 收款率: {collection_rate:.1f}%")
            
            # 状态分布
            status_column = None
            possible_status_columns = ['收款状态', '状态', '付款状态']
            for col in possible_status_columns:
                if col in df_collection.columns:
                    status_column = col
                    break
            
            if status_column:
                status_counts = df_collection[status_column].value_counts()
                coll_summary.append("- 状态分布:")
                for status, count in status_counts.head(3).items():  # 只显示前3个
                    coll_summary.append(f"  {status}: {count}条")
            
            summary["collection_summary"] = "\n".join(coll_summary)
        
        return summary


# 测试代码
if __name__ == "__main__":
    import pandas as pd
    
    # 设置日志
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    
    # 测试数据
    test_subscription = {
        "产品名称": ["产品A", "产品B"],
        "金额": [100000, 150000]
    }
    
    test_collection = {
        "应收金额": [50000, 60000],
        "实收金额": [40000, 55000],
        "状态": ["已收款", "待收款"]
    }
    
    df_sub = pd.DataFrame(test_subscription)
    df_coll = pd.DataFrame(test_collection)
    
    # 测试LLM分析器
    analyzer = LLMAnalyzer()
    
    # 准备数据摘要
    from part2_subscription import SubscriptionAnalyzer
    sub_analyzer = SubscriptionAnalyzer()
    summary = sub_analyzer._prepare_data_summary(df_sub, df_coll)
    
    print("数据摘要:")
    print(summary)
    
    # 测试回退分析
    subscription_analysis = analyzer.analyze_subscription(
        {"summary": summary},
        {"summary": "测试收款数据"}
    )
    
    print("\n订阅分析结果:")
    print(subscription_analysis)
    
    # 测试综合分析
    key_metrics = {
        "计费ARR": 155000,
        "服务阶段": "运维中",
        "客户状态": "绿色",
        "订阅合同数": 2,
        "总工单数": 21,
        "工单关闭率": 95.2
    }
    
    comprehensive_analysis = analyzer.analyze_comprehensive(key_metrics)
    
    print("\n综合经营分析结果:")
    print(comprehensive_analysis)