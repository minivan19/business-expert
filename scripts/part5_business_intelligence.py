#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Part5: 客户经营情报分析
分析客户近一年的经营新闻，评估对SRM应用的机遇和挑战
"""

import os
import sys
import json
import logging
from typing import Optional, Dict, Any, List

logger = logging.getLogger(__name__)


# 客户简称→全称映射
MAPPING_FILE = r"C:\Users\mingh\client-data\raw\_mapping.json"


def get_full_company_name(client_short_name: str) -> str:
    """
    获取客户全称
    
    Args:
        client_short_name: 客户简称
        
    Returns:
        str: 客户全称，如果没有则返回简称
    """
    try:
        if os.path.exists(MAPPING_FILE):
            with open(MAPPING_FILE, 'r', encoding='utf-8') as f:
                mapping = json.load(f)
            
            full_name = mapping.get(client_short_name)
            if full_name:
                logger.info(f"找到客户全称: {full_name}")
                return full_name
    except Exception as e:
        logger.warning(f"读取客户映射失败: {e}")
    
    # 如果没有映射，返回简称
    logger.info(f"未找到全称，使用简称: {client_short_name}")
    return client_short_name


def search_news_tavily(client_name: str, max_results: int = 10) -> List[Dict[str, Any]]:
    """
    使用Tavily搜索客户经营新闻
    
    Args:
        client_name: 客户名称
        max_results: 最大结果数
        
    Returns:
        List[Dict]: 新闻列表
    """
    try:
        # 构建查询
        query = f"{client_name} 经营 2025 2026"
        
        # 调用tavily_search脚本
        script_dir = os.path.dirname(os.path.abspath(__file__))
        tavily_script = os.path.join(script_dir, "..", "..", "openclaw-tavily-search", "scripts", "tavily_search.py")
        
        if not os.path.exists(tavily_script):
            logger.warning(f"Tavily脚本不存在: {tavily_script}")
            return []
        
        # 使用subprocess调用
        import subprocess
        
        cmd = [
            sys.executable,
            tavily_script,
            "--query", query,
            "--max-results", str(max_results),
            "--format", "raw"
        ]
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30
        )
        
        if result.returncode != 0:
            logger.error(f"Tavily搜索失败: {result.stderr}")
            return []
        
        # 解析JSON结果
        try:
            data = json.loads(result.stdout)
            results = data.get("results", [])
            
            # 转换为统一格式
            news_list = []
            for r in results:
                news_list.append({
                    "title": r.get("title", ""),
                    "snippet": r.get("content", "")[:200],
                    "url": r.get("url", "")
                })
            
            logger.info(f"Tavily搜索成功，找到 {len(news_list)} 条新闻")
            return news_list
            
        except json.JSONDecodeError as e:
            logger.error(f"解析Tavily结果失败: {e}")
            return []
            
    except Exception as e:
        logger.error(f"Tavily搜索异常: {e}")
        import traceback
        traceback.print_exc()
        return []


class BusinessIntelligenceAnalyzer:
    """客户经营情报分析器"""
    
    def __init__(self, client_name: str, news_data: Optional[List[Dict[str, Any]]] = None, client_full_name: Optional[str] = None):
        """
        初始化分析器
        
        Args:
            client_name: 客户简称
            news_data: 可选的新闻数据列表（如果外部已搜索）
            client_full_name: 客户全称（可选，如果不传则从mapping读取）
        """
        self.client_short_name = client_name
        
        # 如果提供了全称直接使用，否则从mapping读取
        if client_full_name:
            self.client_full_name = client_full_name
        else:
            # 尝试从mapping读取全称
            self.client_full_name = get_full_company_name(client_name)
        
        self.news_data = news_data or []
        self.analysis_result: str = ""
        
        logger.info(f"初始化客户经营情报分析器: {client_name} (全称: {self.client_full_name})")
    
    def search_news(self, max_results: int = 10) -> List[Dict[str, Any]]:
        """
        搜索客户经营新闻（优先使用Tavily）
        
        Args:
            max_results: 最大搜索结果数
            
        Returns:
            List[Dict]: 新闻列表
        """
        # 如果已经有数据，直接返回
        if self.news_data:
            return self.news_data
        
        # 尝试使用Tavily搜索（使用全称）
        self.news_data = search_news_tavily(self.client_full_name, max_results)
        return self.news_data
    
    def set_news_data(self, news_data: List[Dict[str, Any]]):
        """
        设置新闻数据（从外部传入）
        
        Args:
            news_data: 新闻列表
        """
        self.news_data = news_data
        logger.info(f"设置新闻数据: {len(news_data)} 条")
    
    def format_news_for_llm(self) -> str:
        """
        格式化新闻数据供LLM分析
        
        Returns:
            str: 格式化后的新闻摘要
        """
        if not self.news_data:
            return f"未找到{self.client_full_name}（简称：{self.client_short_name}）的近期经营新闻。请使用Tavily搜索工具获取该客户的经营动态。"
        
        formatted = []
        for i, news in enumerate(self.news_data[:10], 1):
            title = news.get("title", "")
            snippet = news.get("snippet", "")
            url = news.get("url", "")
            
            formatted.append(f"{i}. {title}")
            if snippet:
                formatted.append(f"   {snippet}")
            if url:
                formatted.append(f"   来源: {url}")
            formatted.append("")
        
        return "\n".join(formatted)
    
    def analyze(self, news_data: Optional[List[Dict[str, Any]]] = None) -> str:
        """
        执行分析
        
        Args:
            news_data: 可选的新闻数据（优先使用）
            
        Returns:
            str: 分析结果（300字内）
        """
        # 如果传入了新闻数据，使用传入的数据
        if news_data:
            self.news_data = news_data
        elif not self.news_data:
            # 没有数据时自动搜索
            self.search_news()
        
        # 构建prompt
        news_summary = self.format_news_for_llm()
        
        system_prompt = """你是一位资深的商务分析师，擅长从公开信息中分析客户的经营状况和趋势。
请基于搜索到的客户经营新闻，进行分析并给出对SRM业务的机遇和挑战评估。

分析要求：
1. 总结客户近一年最重要的经营动态（融资、上市、收购、裁员、业务扩张/收缩、政策影响等）
2. 分析客户面临的机遇（发展机会、市场机遇、业务增长点）
3. 分析客户面临的挑战（竞争压力、成本压力、政策风险、经营困难）
4. 评估对SRM业务的机遇：
   - 如果客户发展良好 → 增购机会（人员增加、业务扩大、采购需求增长）
   - 如果客户经营萎缩 → 降价/弃用风险（预算压缩、人员减少）
   - 如果客户战略调整 → 新模块/新业务机会
5. 浓缩在300字以内，语言精炼，直接给出结论

重要：只基于提供的新闻信息分析，不要臆测。
**不要输出任何标题（如"分析"、"结论"等），直接输出正文内容。**"""

        prompt = f"""请分析以下客户的经营情报：

## 1. 客户名称
{self.client_full_name}（简称：{self.client_short_name}）

## 1. 搜索到的新闻
{news_summary}

请基于以上新闻，分析客户的经营状况，评估对SRM业务的机遇和挑战。300字以内。"""

        # 调用LLM
        try:
            from llm_client import get_llm_client
            client = get_llm_client()
            
            logger.info("开始LLM分析...")
            result = client.call(prompt, system_prompt, temperature=0.7, max_tokens=600)
            
            if result:
                self.analysis_result = result  # 不截断
                logger.info(f"分析完成，结果长度: {len(self.analysis_result)} 字")
                return self.analysis_result
            else:
                self.analysis_result = "LLM分析失败"
                return self.analysis_result
                
        except Exception as e:
            logger.error(f"LLM分析失败: {e}")
            import traceback
            traceback.print_exc()
            self.analysis_result = f"分析失败: {str(e)}"
            return self.analysis_result
    
    def get_markdown(self, news_data: Optional[List[Dict[str, Any]]] = None) -> str:
        """
        生成Markdown格式的分析报告
        
        Args:
            news_data: 可选的新闻数据
            
        Returns:
            str: Markdown报告
        """
        # 如果有新闻数据，先分析
        if news_data:
            self.news_data = news_data
            self.analyze()
        elif not self.analysis_result:
            # 没有数据时尝试分析（会调用tavily搜索）
            self.analyze()
        
        # 清理LLM输出
        cleaned = ""
        if self.analysis_result:
            import re
            cleaned = self.analysis_result
            cleaned = re.sub(r'^\d+\.\d+\s+', '', cleaned, flags=re.MULTILINE)
            cleaned = cleaned.replace('**', '')
            cleaned = re.sub(r'^\s*[-*]\s+', '- ', cleaned, flags=re.MULTILINE)
            cleaned = cleaned.strip()
        
        markdown = f"""## 5. 客户经营情报

{cleaned}

---
*数据来源：公开新闻搜索（Tavily）*
"""
        return markdown


def analyze_client_business_intelligence(client_name: str, news_data: Optional[List[Dict[str, Any]]] = None) -> str:
    """
    分析客户经营情报的便捷函数
    
    Args:
        client_name: 客户名称
        news_data: 可选的新闻数据
        
    Returns:
        str: Markdown格式的分析报告
    """
    analyzer = BusinessIntelligenceAnalyzer(client_name, news_data)
    return analyzer.get_markdown(news_data)


if __name__ == "__main__":
    # 测试
    print("=" * 60)
    print("客户经营情报分析测试")
    print("=" * 60)
    
    # 测试（自动搜索）
    print("\n分析 CBD 客户（自动搜索新闻）...")
    result = analyze_client_business_intelligence("CBD")
    print(result)
