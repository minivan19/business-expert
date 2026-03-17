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
# 进入skill目录
cd skills/business-expert

# 生成单个客户报告
python scripts/cli.py --client=客户编号

# 生成所有客户报告
python scripts/cli.py --client=all

# 指定输出目录
python scripts/cli.py --client=客户编号 --output=自定义输出目录

# 详细输出模式
python scripts/cli.py --client=客户编号 --verbose
```

### 环境变量配置
```bash
# 设置API密钥
export DEEPSEEK_API_KEY="sk-your-api-key"

# 设置数据路径（可选）
export CLIENT_DATA_PATH="C:\\your\\path\\to\\客户档案"
export OUTPUT_DIR="C:\\your\\output\\path"
```

## 🎯 核心功能

### 混合模式分析
- **脚本处理**：结构化数据（Excel表格）由Python脚本直接处理，确保准确性
- **LLM智能分析**：需要推理和分析的部分调用大模型，提供深度洞察
- **优势**：结合脚本的准确性和LLM的智能性，降低成本减少幻觉

### 6层数据框架
完整覆盖客户经营的各个方面：
1. **客户基础档案** - 基本信息、业务概况、购买信息
2. **订阅续约与续费** - 订阅明细、续约分析、收款情况
3. **实施优化情况** - 实施费用、优化趋势、模式分析
4. **运维情况** - 工单统计、问题分类、SLA分析
5. **客户经营信息** - 网络搜索信息（待开发）
6. **综合经营分析** - 价值分级、健康度评估、机会风险分析

## 📁 详细文档

### 数据框架与架构
- [数据框架](references/data_framework.md) - 6层数据框架详细说明
- [架构设计](references/architecture.md) - 混合模式架构和工作流程
- [输出格式](references/output_formats.md) - 报告格式规范和示例
- [Prompt约束](references/prompt_constraints.md) - LLM使用规范和模板

### 参考手册
- [确定性分析手册](references/deterministic_analysis_manual.md) - 标准化分析流程和输出格式

## 🔧 脚本说明

### 主要脚本
- `scripts/cli.py` - 命令行接口（主入口）
- `scripts/report_generator.py` - 报告生成器（协调各模块）
- `scripts/data_loader.py` - 数据加载和预处理
- `scripts/llm_analyzer.py` - LLM分析器

### 数据模块
- `scripts/part1_basic_profile.py` - Part 1: 客户基础档案
- `scripts/part2_subscription.py` - Part 2: 订阅续约与续费
- `scripts/part3_implementation.py` - Part 3: 实施优化情况
- `scripts/part4_operations.py` - Part 4: 运维情况

## ⚙️ 配置说明

### 配置文件
Skill配置在 `openclaw.yaml` 中，支持以下配置：

1. **数据路径配置**：
   - `client_data_path`: 客户数据根目录
   - `output_dir`: 报告输出目录

2. **LLM配置**：
   - `api_key_env`: API密钥环境变量名
   - `model`: 使用的模型
   - `temperature`: 温度参数（控制随机性）

3. **功能开关**：
   - 可以启用/禁用各个分析模块

### 环境变量优先级
1. 命令行参数（最高优先级）
2. 环境变量
3. 配置文件默认值

## 📊 输出示例

### 报告结构
```
{客户}_经营分析报告.md
├── 1. 客户基础档案
├── 2. 订阅续约与续费
├── 3. 实施优化情况
├── 4. 运维情况
└── 6. 综合经营分析
```

### 关键指标
- **客户价值分级**：A/B/C/D级
- **经营健康度**：订阅/收款/运维健康度
- **机会分析**：增购/交叉销售机会
- **风险预警**：流失/收款风险

## 🛠️ 故障排除

### 常见问题
1. **依赖安装失败**：确保Python版本>=3.8，使用 `pip install -r requirements.txt`
2. **API调用失败**：检查API密钥和环境变量设置
3. **数据文件找不到**：检查 `client_data_path` 配置
4. **权限问题**：确保有读写输出目录的权限

### 日志查看
```bash
# 启用详细日志
python scripts/cli.py --client=客户编号 --verbose

# 查看错误日志
检查 output_dir/logs/ 目录
```

## 🔄 更新与维护

### 版本信息
- **当前版本**: 1.0.0
- **创建时间**: 2026-03-16
- **最后更新**: 2026-03-17

### 更新记录
- 2026-03-17: 按照Skill creator规范进行规范化改进
- 2026-03-16: 初始版本创建

---

**提示**: 详细的技术文档和参考手册请查看 `references/` 目录下的对应文件。