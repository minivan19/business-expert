#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试LLM连接和API密钥
"""

import os
import sys

def test_llm_connection():
    """测试LLM连接"""
    print("测试LLM连接状态")
    print("=" * 60)
    
    # 检查环境变量
    api_key = os.getenv("DEEPSEEK_API_KEY")
    print(f"1. 检查DEEPSEEK_API_KEY环境变量:")
    if api_key:
        print(f"   ✅ 找到API密钥")
        print(f"     密钥长度: {len(api_key)} 字符")
        print(f"     密钥前缀: {api_key[:10]}...")
    else:
        print(f"   ❌ 未设置DEEPSEEK_API_KEY环境变量")
        print(f"     请设置环境变量: set DEEPSEEK_API_KEY=your_key")
        return False
    
    # 测试导入LLM分析器
    print(f"\n2. 测试导入LLM分析器:")
    try:
        from llm_analyzer import LLMAnalyzer
        print(f"   ✅ 成功导入LLM分析器")
        
        # 创建实例
        analyzer = LLMAnalyzer()
        print(f"   ✅ LLM分析器实例化成功")
        print(f"     可用状态: {analyzer.available}")
        print(f"     API密钥: {'已设置' if analyzer.api_key else '未设置'}")
        
        if analyzer.available:
            print(f"   ✅ LLM分析功能已启用")
            return True
        else:
            print(f"   ❌ LLM分析功能未启用")
            return False
            
    except ImportError as e:
        print(f"   ❌ 导入LLM分析器失败: {e}")
        return False
    except Exception as e:
        print(f"   ❌ 测试过程中出错: {e}")
        return False

def test_api_connection():
    """测试API连接"""
    print(f"\n3. 测试DeepSeek API连接:")
    
    api_key = os.getenv("DEEPSEEK_API_KEY")
    if not api_key:
        print("   ⚠️  跳过API连接测试（无API密钥）")
        return False
    
    try:
        import requests
        
        url = "https://api.deepseek.com/v1/models"
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        
        print(f"   正在连接DeepSeek API...")
        response = requests.get(url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            print(f"   ✅ API连接成功!")
            data = response.json()
            print(f"      可用模型: {len(data.get('data', []))} 个")
            for model in data.get('data', [])[:3]:  # 显示前3个模型
                print(f"      - {model.get('id', '未知')}")
            return True
        else:
            print(f"   ❌ API连接失败")
            print(f"      状态码: {response.status_code}")
            print(f"      响应: {response.text[:100]}")
            return False
            
    except requests.exceptions.Timeout:
        print(f"   ❌ API连接超时")
        return False
    except requests.exceptions.ConnectionError:
        print(f"   ❌ 网络连接错误")
        return False
    except Exception as e:
        print(f"   ❌ API测试出错: {e}")
        return False

def main():
    """主函数"""
    print("商务专家skill - LLM连接测试")
    print("=" * 60)
    
    # 添加当前目录到Python路径
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    
    # 运行测试
    llm_import_ok = test_llm_connection()
    
    if llm_import_ok:
        # 如果LLM分析器可用，测试API连接
        api_ok = test_api_connection()
        
        if api_ok:
            print(f"\n" + "=" * 60)
            print(f"🎉 所有测试通过!")
            print(f"LLM分析功能已完全启用")
            print(f"可以生成完整的智能分析报告")
        else:
            print(f"\n" + "=" * 60)
            print(f"⚠️  LLM分析器可用，但API连接失败")
            print(f"请检查:")
            print(f"1. API密钥是否正确")
            print(f"2. 网络连接是否正常")
            print(f"3. DeepSeek服务状态")
    else:
        print(f"\n" + "=" * 60)
        print(f"❌ LLM分析功能未启用")
        print(f"请执行以下步骤:")
        print(f"1. 获取DeepSeek API密钥")
        print(f"2. 设置环境变量: set DEEPSEEK_API_KEY=your_key")
        print(f"3. 重新运行此测试")
    
    return 0

if __name__ == "__main__":
    sys.exit(main())