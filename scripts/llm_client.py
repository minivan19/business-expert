#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LLM客户端 - 统一管理DeepSeek API调用
"""

import os
import json
import logging
import requests
from typing import Optional, Dict, Any
from datetime import datetime

logger = logging.getLogger(__name__)

# 获取当前日期
def get_current_date():
    return datetime.now().strftime('%Y-%m-%d')

# DeepSeek API配置
DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "sk-340ed7819c2346508c0a46a80df85999")
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
DEEPSEEK_MODEL = "deepseek-chat"


class LLMClient:
    """DeepSeek LLM客户端"""
    
    def __init__(self, api_key: Optional[str] = None, model: str = DEEPSEEK_MODEL):
        """
        初始化LLM客户端
        
        Args:
            api_key: DeepSeek API Key
            model: 模型名称
        """
        self.api_key = api_key or DEEPSEEK_API_KEY
        self.model = model
        self.api_url = DEEPSEEK_API_URL
        self.max_tokens = 4096
        self.temperature = 0.7
        
        logger.info(f"LLM客户端初始化完成")
        logger.info(f"API Key: {self.api_key[:10]}...")
        logger.info(f"Model: {self.model}")
    
    def call(self, prompt: str, system_prompt: Optional[str] = None, 
             temperature: Optional[float] = None, max_tokens: Optional[int] = None) -> Optional[str]:
        """
        调用LLM生成回复
        
        Args:
            prompt: 用户提示
            system_prompt: 系统提示（可选）
            temperature: 温度参数（可选）
            max_tokens: 最大token数（可选）
            
        Returns:
            str: LLM回复内容，如果失败返回None
        """
        try:
            # 构建请求头
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}"
            }
            
            # 构建消息
            messages = []
            
            if system_prompt:
                messages.append({
                    "role": "system",
                    "content": system_prompt
                })
            
            messages.append({
                "role": "user",
                "content": prompt
            })
            
            # 构建请求体
            payload = {
                "model": self.model,
                "messages": messages,
                "temperature": temperature or self.temperature,
                "max_tokens": max_tokens or self.max_tokens
            }
            
            logger.info(f"开始调用LLM API...")
            logger.info(f"Prompt长度: {len(prompt)} 字符")
            
            # 发送请求
            response = requests.post(
                self.api_url,
                headers=headers,
                json=payload,
                timeout=120  # 120秒超时（完整数据需要更长时间）
            )
            
            # 检查响应
            if response.status_code == 200:
                result = response.json()
                
                if "choices" in result and len(result["choices"]) > 0:
                    content = result["choices"][0]["message"]["content"]
                    logger.info(f"LLM调用成功，回复长度: {len(content)} 字符")
                    return content
                else:
                    logger.error(f"LLM返回格式异常: {result}")
                    return None
            else:
                logger.error(f"LLM API调用失败，状态码: {response.status_code}")
                logger.error(f"错误信息: {response.text}")
                return None
                
        except requests.exceptions.Timeout:
            logger.error("LLM API调用超时")
            return None
        except requests.exceptions.RequestException as e:
            logger.error(f"LLM API请求异常: {e}")
            return None
        except Exception as e:
            logger.error(f"LLM调用未知异常: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def analyze_subscription(self, data_summary: str) -> str:
        """
        分析订阅续约情况
        
        Args:
            data_summary: 数据摘要
            
        Returns:
            str: 分析结果
        """
        system_prompt = """你是一位资深的客户成功经理，擅长分析客户的订阅续约情况。
请基于提供的数据摘要，分析客户的续约风险和机会。
分析要点：
1. 订阅状态和金额趋势
2. 续约风险评估（降价风险、流失风险）
3. 续费收款情况（坏账风险）
4. 给出具体的建议

请用专业的商业语言给出分析结果，结构清晰，建议具体可行。"""

        prompt = f"""请分析以下客户的订阅续约情况：

## 数据摘要
{data_summary}

请给出详细的分析和建议。"""

        return self.call(prompt, system_prompt, max_tokens=600) or "LLM分析失败"
    
    def analyze_subscription_full(self, subscription_data: str, collection_data: str) -> str:
        """
        分析订阅续约情况（完整数据）
        
        Args:
            subscription_data: 订阅数据完整内容
            collection_data: 收款数据完整内容
            
        Returns:
            str: 分析结果
        """
        system_prompt = """你是一位资深的客户成功经理，擅长分析客户的订阅续约情况。
请基于提供的本章节统计数据分析客户的续约风险和机会。
分析要点：
1. 订阅状态和金额趋势
2. 续约风险评估（降价风险、流失风险）
3. 续费收款情况（坏账风险、逾期情况）
4. 给出具体的建议

**重要：输出必须限制在300字以内，语言精炼，直击重点。
**不要输出任何标题（如"分析"、"结论"等），直接输出正文内容。**"""

        prompt = f"""请分析以下客户的订阅续约情况：

## 订阅统计（2.1-2.3本章节数据）
{subscription_data}

请基于以上本章节统计数据给出分析，**输出限制在300字以内**。"""

        return self.call(prompt, system_prompt, max_tokens=600) or "LLM分析失败"
    
    def analyze_subscription_from_content(self, content: str) -> str:
        """
        基于已生成的章节内容进行订阅续约分析
        
        Args:
            content: 已生成的2.1-2.3章节markdown内容
            
        Returns:
            str: 分析结果
        """
        system_prompt = """你是一位资深的客户成功经理，擅长分析客户的订阅续约情况。
请基于已经生成好的2.1-2.3章节内容（包含续约概览、订阅明细、收款明细）进行智能分析。
分析要点：
1. 订阅状态和金额趋势
2. 续约风险评估（降价风险、流失风险）
3. 续费收款情况（坏账风险、逾期情况）
4. 给出具体的建议

**重要要求：**
1. 必须基于提供的章节内容分析，不要自行补充数据
2. 如果章节中没有某项数据，明确说明"数据缺失"而非假设
3. 输出必须限制在300字以内，语言精炼，直击重点
4. 不要输出任何标题（如"分析"、"结论"等），直接输出正文内容"""

        prompt = f"""【当前日期】{get_current_date()}

请基于以下已生成的2.1-2.3章节内容进行智能分析：

【重要】请根据【当前日期】{get_current_date()}判断款项是否逾期：
- 如果"考核收款日期"早于当前日期，且"已收款金额"为0，则该款项已逾期
- 例如：考核收款日期2025-09-30，当前日期2026-03-19，则该款项已逾期超过5个月

## 2.1-2.3 章节内容
{content}

请基于以上章节内容给出分析，**输出限制在300字以内**。"""

        return self.call(prompt, system_prompt, max_tokens=600) or "LLM分析失败"
    
    def analyze_implementation_full(self, fixed_data: str, dayspan_data: str) -> str:
        """
        分析实施优化情况（统计数据）
        
        Args:
            fixed_data: 本章节统计数据
            
        Returns:
            str: 分析结果
        """
        system_prompt = """你是一位资深的实施顾问，擅长分析客户的实施优化情况。
请基于提供的本章节统计数据分析客户的实施进度和优化机会。
分析要点：
1. 实施合同执行情况
2. 人天框架使用情况
3. 优化空间和建议
4. 给出具体的建议

**重要：输出必须限制在300字以内，语言精炼，直击重点。
**不要输出任何标题（如"分析"、"结论"等），直接输出正文内容。**"""

        prompt = f"""请分析以下客户的实施优化情况：

## 实施统计（3.1-3.3本章节数据）
{fixed_data}

请基于以上本章节统计数据给出分析，**输出限制在300字以内**。"""

        return self.call(prompt, system_prompt, max_tokens=600) or "LLM分析失败"
    
    def analyze_implementation_from_content(self, content: str) -> str:
        """
        基于已生成的章节内容进行实施优化分析
        
        Args:
            content: 已生成的3.1-3.3章节markdown内容
            
        Returns:
            str: 分析结果
        """
        system_prompt = """你是一位资深的实施顾问，擅长分析客户的实施优化情况。
请基于已经生成好的3.1-3.3章节内容（包含实施概览、固定合同明细、人天框架明细）进行智能分析。
分析要点：
1. 实施合同执行情况
2. 人天框架使用情况
3. 优化空间和建议
4. 给出具体的建议

**重要要求：**
1. 必须基于提供的章节内容分析，不要自行补充数据
2. 如果章节中没有某项数据，明确说明"数据缺失"而非假设
3. 输出必须限制在300字以内，语言精炼，直击重点
4. 不要输出任何标题（如"分析"、"结论"等），直接输出正文内容"""

        prompt = f"""【当前日期】{get_current_date()}

请基于以下已生成的3.1-3.3章节内容进行智能分析：

## 3.1-3.3 章节内容
{content}

请基于以上章节内容给出分析，**输出限制在300字以内**。"""

        return self.call(prompt, system_prompt, max_tokens=600) or "LLM分析失败"
    
    def analyze_operations_full(self, operations_data: str) -> str:
        """
        分析运维情况（完整数据）
        
        Args:
            operations_data: 运维工单统计数据
            
        Returns:
            str: 分析结果
        """
        system_prompt = """你是一位资深的运维经理，擅长分析客户的运维情况。
请基于提供的本章节统计数据，分析客户的系统稳定性和运维需求。
分析要点：
1. 系统稳定性评估
2. 运维工单分布和趋势
3. 潜在风险预警
4. 给出具体的建议

**重要：输出必须限制在300字以内，语言精炼，直击重点。
**不要输出任何标题（如"分析"、"结论"等），直接输出正文内容。**"""

        prompt = f"""请分析以下客户的运维情况：

## 运维统计（4.1-4.3本章节数据）
{operations_data}

请基于以上本章节统计数据给出分析，**输出限制在300字以内**。"""

        return self.call(prompt, system_prompt, max_tokens=600) or "LLM分析失败"
    
    def analyze_operations_from_content(self, content: str) -> str:
        """
        基于已生成的章节内容进行运维分析
        
        Args:
            content: 已生成的4.1-4.3章节markdown内容
            
        Returns:
            str: 分析结果
        """
        system_prompt = """你是一位资深的运维经理，擅长分析客户的运维情况。
请基于已经生成好的4.1-4.3章节内容（包含运维概览、模块分布、类型分布）进行智能分析。
分析要点：
1. 系统稳定性评估
2. 运维工单分布和趋势
3. 潜在风险预警
4. 给出具体的建议

**重要要求：**
1. 必须基于提供的章节内容分析，不要自行补充数据
2. 如果章节中没有某项数据，明确说明"数据缺失"而非假设
3. 输出必须限制在300字以内，语言精炼，直击重点
4. 不要输出任何标题（如"分析"、"结论"等），直接输出正文内容"""

        prompt = f"""【当前日期】{get_current_date()}

请基于以下已生成的4.1-4.3章节内容进行智能分析：

## 4.1-4.3 章节内容
{content}

请基于以上章节内容给出分析，**输出限制在300字以内**。"""

        return self.call(prompt, system_prompt, max_tokens=600) or "LLM分析失败"
    
    def analyze_implementation(self, data_summary: str) -> str:
        """
        分析实施优化情况
        
        Args:
            data_summary: 数据摘要
            
        Returns:
            str: 分析结果
        """
        system_prompt = """你是一位资深的实施顾问，擅长分析客户的实施优化情况。
请基于提供的数据摘要，分析客户的实施进度和优化机会。
分析要点：
1. 实施合同执行情况
2. 人天框架使用情况
3. 优化空间和建议
4. 给出具体的建议

请用专业的商业语言给出分析结果，结构清晰，建议具体可行。"""

        prompt = f"""请分析以下客户的实施优化情况：

## 数据摘要
{data_summary}

请给出详细的分析和建议。"""

        return self.call(prompt, system_prompt) or "LLM分析失败"
    
    def analyze_operations(self, data_summary: str) -> str:
        """
        分析运维情况
        
        Args:
            data_summary: 数据摘要
            
        Returns:
            str: 分析结果
        """
        system_prompt = """你是一位资深的运维经理，擅长分析客户的运维情况。
请基于提供的数据摘要，分析客户的系统稳定性和运维需求。
分析要点：
1. 系统稳定性评估
2. 运维工单分布和趋势
3. 潜在风险预警
4. 给出具体的建议

请用专业的商业语言给出分析结果，结构清晰，建议具体可行。"""

        prompt = f"""请分析以下客户的运维情况：

## 数据摘要
{data_summary}

请给出详细的分析和建议。"""

        return self.call(prompt, system_prompt) or "LLM分析失败"
    
    def analyze_comprehensive(self, data_summary: str) -> str:
        """
        综合经营分析
        
        Args:
            data_summary: 数据摘要
            
        Returns:
            str: 分析结果
        """
        system_prompt = """你是一位资深的商务专家，擅长分析客户的综合经营情况。
请基于提供的数据摘要，给出客户的全方位分析。
分析要点：
1. 客户价值评估
2. 经营健康度评估
3. 机会分析
4. 风险预警
5. 综合建议

请用专业的商业语言给出分析结果，结构清晰，建议具体可行。"""

        prompt = f"""请分析以下客户的综合经营情况：

## 数据摘要
{data_summary}

请给出详细的分析和建议。"""

        return self.call(prompt, system_prompt) or "LLM分析失败"
    
    def analyze_comprehensive_full(self, part1_data: str, part2_data: str, part3_data: str, part4_data: str, part5_data: str) -> str:
        """
        综合经营分析（完整数据，Part1-5）
        
        Args:
            part1_data: Part1客户基础档案完整数据
            part2_data: Part2订阅续约完整数据
            part3_data: Part3实施优化完整数据
            part4_data: Part4运维完整数据
            part5_data: Part5客户经营情报分析结果
            
        Returns:
            str: 分析结果（800字内）
        """
        system_prompt = """你是一位资深的商务专家，擅长分析客户的综合经营情况。
请基于提供的Part1-5完整数据，给出客户的全方位综合经营分析。

分析要求：
1. 基于Part1客户档案，了解客户基本情况
2. 结合Part2订阅续约数据，分析客户续约意愿和付款能力
3. 结合Part3实施优化数据，分析客户业务发展需求
4. 结合Part4运维数据，分析客户系统使用状况和满意度
5. 结合Part5经营情报，分析客户经营变化对SRM业务的影响
6. 给出综合评估：客户价值分级、经营健康度、机会分析、风险预警
7. 给出具体的行动建议（短期、中期、长期）

请用专业的商业语言给出分析结果，结构清晰，建议具体可行。
**重要：输出必须控制在800字以内，语言精炼，直击重点。
**不要输出任何标题（如"一、"、"（一）"、"分析如下"等），直接输出正文内容。**"""

        prompt = f"""请分析以下客户的综合经营情况：

## Part1: 客户基础档案
{part1_data}

## Part2: 订阅续约与续费
{part2_data}

## Part3: 实施优化情况
{part3_data}

## Part4: 运维情况
{part4_data}

## Part5: 客户经营情报
{part5_data}

请基于以上Part1-5的完整数据，给出综合经营分析。**输出必须控制在800字以内**。"""

        return self.call(prompt, system_prompt, max_tokens=1500) or "LLM分析失败"
    
    def analyze_comprehensive_from_content(self, content: str) -> str:
        """
        基于已生成的Part1-5章节内容进行综合经营分析
        
        Args:
            content: 已生成的Part1-5完整markdown内容
            
        Returns:
            str: 分析结果（800字内）
        """
        system_prompt = """你是一位资深的商务专家，擅长分析客户的综合经营情况。
请基于已经生成好的Part1-5章节内容（包含客户基础档案、订阅续约、实施优化、运维情况、客户经营情报），给出客户的全方位综合经营分析。

分析要求：
1. 基于Part1客户档案，了解客户基本情况
2. 结合Part2订阅续约数据，分析客户续约意愿和付款能力
3. 结合Part3实施优化数据，分析客户业务发展需求
4. 结合Part4运维数据，分析客户系统使用状况和满意度
5. 结合Part5经营情报，分析客户经营变化对SRM业务的影响
6. 给出综合评估：客户价值分级、经营健康度、机会分析、风险预警
7. 给出具体的行动建议（短期、中期、长期）

**重要要求：**
1. 必须基于提供的Part1-5章节内容分析，不要自行补充数据
2. 如果章节中缺少某部分数据，明确说明"数据缺失"而非假设
3. 输出必须控制在800字以内，语言精炼，直击重点
4. 不要输出任何标题（如"一、"、"（一）"、"分析如下"等），直接输出正文内容"""

        prompt = f"""【当前日期】{get_current_date()}

请基于以下已生成的Part1-5章节内容进行综合经营分析：

## Part1-5 章节内容
{content}

请基于以上章节内容给出综合经营分析，**输出必须控制在800字以内**。"""

        return self.call(prompt, system_prompt, max_tokens=1500) or "LLM分析失败"


# 全局LLM客户端实例
_llm_client: Optional[LLMClient] = None


def get_llm_client() -> LLMClient:
    """获取全局LLM客户端实例"""
    global _llm_client
    if _llm_client is None:
        _llm_client = LLMClient()
    return _llm_client


def test_llm_connection() -> bool:
    """
    测试LLM连接
    
    Returns:
        bool: 连接是否成功
    """
    client = get_llm_client()
    result = client.call("你好，请回复'测试成功'")
    return result is not None and "测试成功" in result


if __name__ == "__main__":
    # 测试LLM连接
    print("=" * 60)
    print("LLM连接测试")
    print("=" * 60)
    
    if test_llm_connection():
        print("LLM连接测试成功!")
    else:
        print("LLM连接测试失败")
    
    # 测试分析功能
    print("\n测试订阅分析功能:")
    client = get_llm_client()
    test_summary = """
订阅数据摘要:
- 总订阅记录数: 2
- 状态分布:
  - 订阅中: 2条
- 年订阅费总金额: 300,000元

收款数据摘要:
- 总收款记录数: 4
- 计划收款金额: 300,000元
- 已收款金额: 250,000元
- 未收款金额: 50,000元
"""
    result = client.analyze_subscription(test_summary)
    print(f"分析结果:\n{result}")
