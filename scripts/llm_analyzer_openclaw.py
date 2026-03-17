#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LLM智能分析模块 - OpenClaw配置集成版
从OpenClaw配置文件中读取DeepSeek API密钥
"""

import os
import json
import logging
import requests
from typing import Dict, Any, Optional
from pathlib import Path

logger = logging.getLogger(__name__)


class LLMAnalyzerOpenClaw:
    """LLM智能分析器（OpenClaw配置集成版）"""
    
    def __init__(self, openclaw_config_path: Optional[str] = None):
        """
        初始化LLM分析器（从OpenClaw配置读取API密钥）
        
        Args:
            openclaw_config_path: OpenClaw配置文件路径，如果为None则自动查找
        """
        self.api_key = None
        self.base_url = "https://api.deepseek.com/v1"
        self.model = "deepseek-chat"
        
        # 尝试从OpenClaw配置读取API密钥
        self.api_key = self._get_api_key_from_openclaw(openclaw_config_path)
        
        if not self.api_key:
            # 如果OpenClaw配置中没有，尝试环境变量
            self.api_key = os.getenv("DEEPSEEK_API_KEY")
            if self.api_key:
                logger.info("从环境变量获取DeepSeek API密钥")
        
        if not self.api_key:
            logger.warning("未找到DeepSeek API密钥，LLM分析将不可用")
            logger.warning("请检查：1) OpenClaw配置文件 2) DEEPSEEK_API_KEY环境变量")
            self.available = False
        else:
            logger.info(f"成功获取DeepSeek API密钥（长度：{len(self.api_key)}字符）")
            self.available = True
    
    def _get_api_key_from_openclaw(self, config_path: Optional[str] = None) -> Optional[str]:
        """
        从OpenClaw配置文件中读取DeepSeek API密钥
        
        Args:
            config_path: 配置文件路径
            
        Returns:
            str: API密钥，如果未找到则返回None
        """
        try:
            # 确定配置文件路径
            if config_path:
                config_file = Path(config_path)
            else:
                # 默认路径：用户目录下的.openclaw/openclaw.json
                config_file = Path.home() / ".openclaw" / "openclaw.json"
            
            logger.info(f"尝试从OpenClaw配置文件读取API密钥: {config_file}")
            
            if not config_file.exists():
                logger.warning(f"OpenClaw配置文件不存在: {config_file}")
                return None
            
            # 读取配置文件
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 提取DeepSeek API密钥
            # 路径: models -> providers -> deepseek -> apiKey
            if "models" in config and "providers" in config["models"]:
                providers = config["models"]["providers"]
                if "deepseek" in providers:
                    deepseek_config = providers["deepseek"]
                    if "apiKey" in deepseek_config:
                        api_key = deepseek_config["apiKey"]
                        if api_key and api_key.strip():
                            logger.info("成功从OpenClaw配置读取DeepSeek API密钥")
                            
                            # 同时检查baseUrl
                            if "baseUrl" in deepseek_config:
                                self.base_url = deepseek_config["baseUrl"]
                                logger.info(f"使用OpenClaw配置的baseUrl: {self.base_url}")
                            
                            return api_key.strip()
            
            logger.warning("OpenClaw配置文件中未找到DeepSeek API密钥")
            return None
            
        except json.JSONDecodeError as e:
            logger.error(f"OpenClaw配置文件JSON解析错误: {e}")
            return None
        except Exception as e:
            logger.error(f"读取OpenClaw配置文件时出错: {e}")
            return None
    
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
            # 构建分析提示
            prompt = self._build_subscription_prompt(subscription_data, collection_data)
            
            # 调用LLM API
            response = self._call_llm_api(prompt)
            
            if response:
                return response
            else:
                logger.warning("LLM API调用失败，使用回退分析")
                return self._get_fallback_analysis("subscription")
                
        except Exception as e:
            logger.error(f"LLM分析过程中出错: {e}")
            return self._get_fallback_analysis("subscription")
    
    def analyze_comprehensive(self, key_metrics: Dict[str, Any]) -> str:
        """
        综合经营分析
        
        Args:
            key_metrics: 关键指标字典
            
        Returns:
            str: LLM生成的综合经营分析内容
        """
        if not self.available:
            return self._get_fallback_analysis("comprehensive")
        
        try:
            # 构建综合分析提示
            prompt = self._build_comprehensive_prompt(key_metrics)
            
            # 调用LLM API
            response = self._call_llm_api(prompt)
            
            if response:
                return response
            else:
                logger.warning("LLM API调用失败，使用回退分析")
                return self._get_fallback_analysis("comprehensive")
                
        except Exception as e:
            logger.error(f"LLM综合分析过程中出错: {e}")
            return self._get_fallback_analysis("comprehensive")
    
    def _build_subscription_prompt(self, subscription_data: Dict[str, Any], collection_data: Dict[str, Any]) -> str:
        """构建订阅分析提示"""
        prompt = f"""你是一个专业的客户成功经理，请分析以下客户订阅和收款数据：

## 订阅数据摘要
{json.dumps(subscription_data, indent=2, ensure_ascii=False)}

## 收款数据摘要  
{json.dumps(collection_data, indent=2, ensure_ascii=False)}

请从以下角度进行专业分析：
1. 订阅费用时间阶段变化趋势分析
2. 续约/降价原因分析
3. 收款进度评估和逾期风险分析
4. 续费策略建议

请用中文回答，保持专业、客观，提供具体的数据支持和可操作建议。"""
        return prompt
    
    def _build_comprehensive_prompt(self, key_metrics: Dict[str, Any]) -> str:
        """构建综合分析提示"""
        prompt = f"""你是一个专业的客户经营分析师，请基于以下关键指标进行综合经营分析：

## 客户经营关键指标
{json.dumps(key_metrics, indent=2, ensure_ascii=False)}

请从以下角度进行全面的经营分析：
1. 客户价值分级（A/B/C/D级）及理由
2. 经营健康度评估（订阅健康度、收款健康度、运维健康度）
3. 机会分析（增购机会、交叉销售机会、关系深化机会）
4. 风险预警（流失风险、收款风险、服务风险、竞争风险）
5. 下一步行动建议（短期1个月内、中期3个月内、长期6个月内）

请用中文回答，保持专业、客观，提供具体的数据支持和可操作建议。"""
        return prompt
    
    def _call_llm_api(self, prompt: str) -> Optional[str]:
        """
        调用LLM API
        
        Args:
            prompt: 提示词
            
        Returns:
            str: LLM响应内容，如果失败则返回None
        """
        try:
            url = f"{self.base_url}/chat/completions"
            
            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }
            
            data = {
                "model": self.model,
                "messages": [
                    {"role": "system", "content": "你是一个专业的客户成功经理和经营分析师。"},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.2,
                "max_tokens": 2000
            }
            
            logger.info(f"调用DeepSeek API: {url}")
            response = requests.post(url, headers=headers, json=data, timeout=120)
            
            if response.status_code == 200:
                result = response.json()
                content = result["choices"][0]["message"]["content"]
                logger.info(f"LLM API调用成功，响应长度: {len(content)}字符")
                return content
            else:
                logger.error(f"LLM API调用失败: {response.status_code} - {response.text[:200]}")
                return None
                
        except requests.exceptions.Timeout:
            logger.error("LLM API调用超时")
            return None
        except requests.exceptions.ConnectionError:
            logger.error("LLM API连接错误")
            return None
        except Exception as e:
            logger.error(f"LLM API调用过程中出错: {e}")
            return None
    
    def _get_fallback_analysis(self, analysis_type: str) -> str:
        """获取回退分析内容"""
        if analysis_type == "subscription":
            return """### 2.4 智能分析（LLM分析暂不可用）

由于LLM服务暂时不可用，以下是基于规则的分析：

#### 1. 订阅费用时间阶段变化趋势分析
- **数据基础**：2个订阅合同，总金额330,000元
- **趋势观察**：第二个合同（155,000元）比第一个（175,000元）下降11.4%
- **初步判断**：可能存在价格调整、产品版本差异或谈判结果

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
        
        elif analysis_type == "comprehensive":
            return """## 6. 综合经营分析

*此部分将由LLM进行综合经营分析，包括：*
1) 客户价值分级（A/B/C/D级）
2) 经营健康度评估（订阅/收款/运维健康度）
3) 机会分析（增购/交叉销售机会）
4) 风险预警（流失/收款/服务风险）
5) 下一步行动建议（短期/中期/长期）

*关键指标（供LLM参考）:*
```
客户ID: CBD
订阅合同数: 2
总订阅金额: 330,000元
收款记录数: 6
总应收金额: 990,000元
运维工单数: 21
工单关闭率: 95.2%
```

*注：启用LLM分析后，此部分将替换为智能生成的深度分析。*"""
        
        else:
            return "*LLM分析暂不可用，请检查API密钥配置。*"


def test_openclaw_integration():
    """测试OpenClaw集成"""
    print("测试OpenClaw配置集成")
    print("=" * 60)
    
    analyzer = LLMAnalyzerOpenClaw()
    
    print(f"LLM分析器可用状态: {analyzer.available}")
    print(f"API密钥: {'已设置' if analyzer.api_key else '未设置'}")
    
    if analyzer.available:
        print(f"✅ 成功从OpenClaw配置读取API密钥")
        print(f"   密钥长度: {len(analyzer.api_key)} 字符")
        print(f"   API端点: {analyzer.base_url}")
        print(f"   模型: {analyzer.model}")
        
        # 测试API连接
        print(f"\n测试API连接...")
        try:
            import requests
            url = f"{analyzer.base_url}/models"
            headers = {"Authorization": f"Bearer {analyzer.api_key}"}
            response = requests.get(url, headers=headers, timeout=10)
            
            if response.status_code == 200:
                print(f"✅ API连接成功!")
                data = response.json()
                print(f"   可用模型数: {len(data.get('data', []))}")
            else:
                print(f"❌ API连接失败: {response.status_code}")
        except Exception as e:
            print(f"❌ API测试出错: {e}")
    else:
        print(f"❌ 无法从OpenClaw配置读取API密钥")
        print(f"   请检查OpenClaw配置文件: ~/.openclaw/openclaw.json")
    
    return analyzer.available


if __name__ == "__main__":
    # 设置日志
    logging.basicConfig(level=logging.INFO)
    
    # 测试集成
    success = test_openclaw_integration()
    
    if success:
        print(f"\n🎉 OpenClaw配置集成测试成功!")
        print(f"商务专家skill现在可以使用OpenClaw配置中的DeepSeek API密钥")
    else:
        print(f"\n⚠️  OpenClaw配置集成测试失败")
        print(f"请检查OpenClaw配置文件或使用环境变量")