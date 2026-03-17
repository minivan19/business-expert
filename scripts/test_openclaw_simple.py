#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简单测试OpenClaw配置集成
"""

import sys
sys.path.append(r'C:\Users\mingh\.openclaw\workspace\skills\business-expert\scripts')

from llm_analyzer_openclaw import LLMAnalyzerOpenClaw

def main():
    print("测试OpenClaw配置集成")
    print("=" * 60)
    
    analyzer = LLMAnalyzerOpenClaw()
    
    print(f"LLM分析器可用状态: {analyzer.available}")
    print(f"API密钥设置: {'是' if analyzer.api_key else '否'}")
    
    if analyzer.available:
        print("成功! 可以从OpenClaw配置使用DeepSeek API密钥")
        print(f"密钥长度: {len(analyzer.api_key)} 字符")
        print(f"API端点: {analyzer.base_url}")
        print(f"模型: {analyzer.model}")
        
        # 测试API连接
        print("\n测试API连接...")
        try:
            import requests
            url = f"{analyzer.base_url}/models"
            headers = {"Authorization": f"Bearer {analyzer.api_key}"}
            response = requests.get(url, headers=headers, timeout=10)
            
            if response.status_code == 200:
                print("API连接成功!")
                data = response.json()
                print(f"可用模型数: {len(data.get('data', []))}")
                return True
            else:
                print(f"API连接失败: {response.status_code}")
                return False
        except Exception as e:
            print(f"API测试出错: {e}")
            return False
    else:
        print("失败: 无法从OpenClaw配置读取API密钥")
        return False

if __name__ == "__main__":
    success = main()
    
    if success:
        print("\n" + "=" * 60)
        print("OpenClaw配置集成测试成功!")
        print("商务专家skill现在可以使用OpenClaw配置中的DeepSeek API密钥")
    else:
        print("\n" + "=" * 60)
        print("OpenClaw配置集成测试失败")
        print("请检查OpenClaw配置文件或使用环境变量")