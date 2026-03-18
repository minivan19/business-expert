#!/usr/bin/env python3
"""测试商务专家 Skill 的 LLM 模块"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from llm_analyzer import LLMAnalyzer

# 测试 LLMAnalyzer
analyzer = LLMAnalyzer(api_key='sk-340ed7819c2346508c0a46a80df85999')
print('LLMAnalyzer 初始化:', '可用' if analyzer.available else '不可用')

# 测试综合分析
key_metrics = {
    '计费ARR': 155000,
    '服务阶段': '运维中',
    '客户状态': '绿色',
    '订阅合同数': 2,
    '总工单数': 21,
    '工单关闭率': 95.2
}

print('\n正在调用 DeepSeek LLM...')
result = analyzer.analyze_comprehensive(key_metrics)
print('\n=== 综合经营分析结果 ===')
print(result)
