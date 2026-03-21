---
name: business-expert
description: >
  商务专家 - 360°客户档案与经营分析。基于客户档案数据生成经营分析报告，采用混合模式（脚本处理结构化数据+LLM智能分析）。

  **当以下情况时使用此Skill**：
  (1) 需要生成客户经营分析报告
  (2) 需要分析客户订阅、实施、运维数据
  (3) 需要评估客户价值和经营健康度
  (4) 用户提到"客户分析"、"经营报告"、"商务分析"
dependency:
  python:
    - pandas>=2.0.0
    - openpyxl>=3.0.0
    - requests>=2.28.0
---

# 商务专家 Skill

## 🚀 快速开始

### 生成客户经营分析报告
```bash
cd skills/business-expert/scripts

# 生成报告
python report_generator_integrated.py 客户名称

# 示例
python report_generator_integrated.py 高义
python report_generator_integrated.py CBD

# 跳过LLM分析（只生成数据）
python report_generator_integrated.py CBD --skip-llm

# 调试模式
python report_generator_integrated.py CBD --debug
```

### 环境变量
```bash
export DEEPSEEK_API_KEY="sk-your-api-key"
```

## 🎯 核心功能

### 混合模式分析
- **脚本处理**：结构化数据（Excel表格）由Python脚本直接处理
- **LLM智能分析**：需要推理的部分调用大模型

### 6层数据框架
1. **客户基础档案** - 基本信息、业务概况、购买信息
2. **订阅续约与续费** - 订阅明细、续约分析、收款情况
3. **实施优化情况** - 实施费用、优化趋势、模式分析
4. **运维情况** - 工单统计、问题分类、SLA分析
5. **客户经营情报** - 网络搜索经营动态
6. **综合经营分析** - 价值分级、健康度评估、机会风险分析

## 📁 详细文档

- [数据框架](references/data_framework.md) - 6层数据框架详细说明
- [确定性分析手册](references/deterministic_analysis_manual.md) - 标准化分析流程和输出格式
- [Prompt约束](references/prompt_constraints.md) - LLM使用规范和模板

## 🔧 脚本说明

| 脚本 | 说明 |
|------|------|
| `report_generator_integrated.py` | **主入口**，生成完整报告 |
| `data_loader.py` | 数据加载模块 |
| `llm_client.py` | LLM客户端封装 |
| `md2docx.py` | Markdown转Word |
| `part1_basic_profile.py` | Part1客户基础档案 |
| `part2_subscription.py` | Part2订阅续约分析 |
| `part3_implementation.py` | Part3实施优化分析 |
| `part4_operations.py` | Part4运维情况分析 |
| `part5_business_intelligence.py` | Part5客户经营情报 |
| `part6_comprehensive.py` | Part6综合经营分析 |

## ⚙️ 配置说明

报告保存到 `/Users/limingheng/AI\client-data\{客户名称}\` 目录下。

**Word转换**：支持两种模式 — 有模板时使用模板样式，无模板时使用纯样式生成。

## 🛠️ 依赖

```bash
pip install pandas openpyxl requests python-docx
```
