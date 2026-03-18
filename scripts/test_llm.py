#!/usr/bin/env python3
"""测试 DeepSeek LLM 调用"""

import requests
import json

API_KEY = "sk-340ed7819c2346508c0a46a80df85999"
URL = "https://api.deepseek.com/v1/chat/completions"

def test_deepseek():
    print("=" * 50)
    print("DeepSeek API 测试")
    print("=" * 50)
    
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {API_KEY}"
    }
    
    data = {
        "model": "deepseek-chat",
        "messages": [
            {"role": "user", "content": "你好，请用一句话介绍自己"}
        ],
        "temperature": 0.3,
        "max_tokens": 100
    }
    
    try:
        print("\n正在连接 api.deepseek.com ...")
        response = requests.post(URL, headers=headers, json=data, timeout=30)
        
        print(f"HTTP 状态码: {response.status_code}")
        print(f"响应内容: {response.text[:500]}")
        
        if response.status_code == 200:
            result = response.json()
            content = result["choices"][0]["message"]["content"]
            print("\n✅ DeepSeek 调用成功!")
            print(f"回复: {content}")
            return True
        else:
            print(f"\n❌ 调用失败: {response.status_code}")
            return False
            
    except requests.exceptions.Timeout:
        print("\n❌ 连接超时 - 网络无法访问 api.deepseek.com")
        return False
    except requests.exceptions.ConnectionError as e:
        print(f"\n❌ 连接错误: {e}")
        return False
    except Exception as e:
        print(f"\n❌ 未知错误: {e}")
        return False

if __name__ == "__main__":
    import sys
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    test_deepseek()
