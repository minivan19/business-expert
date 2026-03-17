# 混合模式架构设计

## 设计理念

商务专家skill采用**混合模式**架构，结合脚本处理结构化数据和大模型智能分析的优势：

### 核心优势
1. **准确性**：结构化数据由脚本直接处理，避免LLM幻觉
2. **智能性**：需要推理和分析的部分调用LLM，提供深度洞察
3. **效率**：降低LLM使用成本，提高处理速度
4. **一致性**：确保每次生成结果格式一致

## 系统架构

### 整体架构图
```
┌─────────────────────────────────────────────────────┐
│                   数据加载层                          │
│  ┌─────────┐ ┌─────────┐ ┌─────────┐ ┌─────────┐   │
│  │客户主数据│ │订阅明细 │ │实施数据 │ │运维工单 │   │
│  └─────────┘ └─────────┘ └─────────┘ └─────────┘   │
└─────────────────────────────────────────────────────┘
                          │
         ┌────────────────┼────────────────┐
         ▼                ▼                ▼
    ┌─────────┐    ┌─────────┐    ┌─────────┐
    │ 数据清洗  │    │ 数据清洗  │    │ 数据清洗  │
    │  与预处理 │    │  与预处理 │    │  与预处理 │
    └─────────┘    └─────────┘    └─────────┘
         │                │                │
         ▼                ▼                ▼
    ┌──────────────────────────────────────┐
    │          分析引擎层                   │
    │  ┌─────┐ ┌─────┐ ┌─────┐ ┌─────┐    │
    │  │Part1│ │Part2│ │Part3│ │Part4│    │
    │  │脚本 │ │混合 │ │混合 │ │脚本 │    │
    │  └─────┘ └─────┘ └─────┘ └─────┘    │
    └──────────────────────────────────────┘
                          │
                          ▼
                  ┌──────────────┐
                  │  LLM分析器   │
                  │  (Part 2.4)  │
                  │  (Part 3.4)  │
                  │  (Part 5)    │
                  └──────────────┘
                          │
                          ▼
                  ┌──────────────┐
                  │  报告生成器  │
                  │  整合所有部分 │
                  │  格式标准化  │
                  └──────────────┘
                          │
                          ▼
                  ┌──────────────┐
                  │  输出层      │
                  │  Markdown报告│
                  │  日志文件    │
                  └──────────────┘
```

## 模块详细设计

### 1. 数据加载层 (Data Loader)

#### 功能
- 读取Excel文件
- 处理不同编码格式
- 验证数据完整性
- 统一数据格式

#### 实现
```python
class DataLoader:
    def __init__(self, client_data_path):
        self.client_data_path = client_data_path
    
    def load_basic_data(self, client_id):
        """加载客户主数据"""
        path = f"{self.client_data_path}/{client_id}/基础数据/客户主数据.xlsx"
        return pd.read_excel(path)
    
    def load_subscription_data(self, client_id):
        """加载订阅数据"""
        path = f"{self.client_data_path}/{client_id}/订阅合同行/订阅明细.xlsx"
        return pd.read_excel(path)
    
    # ... 其他数据加载方法
```

### 2. 数据清洗与预处理层 (Data Preprocessor)

#### 功能
- 处理空值和异常值
- 统一日期格式
- 标准化金额单位（元）
- 数据去重和验证

#### 关键处理
```python
def preprocess_data(df):
    """数据预处理"""
    # 处理空值
    df = df.fillna('')
    
    # 统一日期格式
    date_columns = ['购买时间', '合同开始时间', '合同结束时间']
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # 标准化金额（转为数值，单位：元）
    amount_columns = ['金额', '费用', 'ARR']
    for col in amount_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    return df
```

### 3. 分析引擎层 (Analysis Engine)

#### Part 1: 客户基础档案（纯脚本）
- 直接从139个字段提取关键信息
- 生成标准化表格
- 无需LLM参与

#### Part 2: 订阅续约与续费（混合模式）
- **2.1-2.3（脚本）**: 统计表格生成
- **2.4（LLM）**: 续费分析

#### Part 3: 实施优化情况（混合模式）
- **3.1-3.3（脚本）**: 统计表格生成
- **3.4（LLM）**: 优化分析

#### Part 4: 运维情况（纯脚本）
- 工单统计和分析
- 问题分类和SLA分析
- 无需LLM参与

### 4. LLM分析器 (LLM Analyzer)

#### 设计原则
1. **最小化调用**：只在需要智能分析时调用LLM
2. **结构化Prompt**：使用标准化Prompt模板
3. **结果验证**：验证LLM输出格式和内容
4. **错误处理**：处理API调用失败和超时

#### Prompt设计
```python
def create_analysis_prompt(data_summary, analysis_type):
    """创建分析Prompt"""
    prompt_template = """
你是一个客户成功经理助手。根据以下数据，进行智能分析。

## 重要约束
- 禁止使用**加粗**等Markdown格式
- 禁止使用##标题标记
- 禁止使用"好的"、"作为您的"等客套话
- 直接输出分析内容，不要任何开场白

## 数据摘要
{data}

请分析以下内容（直接输出，不要带标题）：
{analysis_requirements}
"""
    return prompt_template.format(data=data_summary, analysis_requirements=analysis_type)
```

### 5. 报告生成器 (Report Generator)

#### 功能
- 整合所有分析结果
- 生成标准化Markdown报告
- 添加时间戳和元数据
- 保存到指定目录

#### 报告结构
```markdown
# {客户}经营分析报告

**生成时间**: {timestamp}
**分析周期**: {period}

## 1. 客户基础档案
{part1_content}

## 2. 订阅续约与续费
{part2_content}

## 3. 实施优化情况
{part3_content}

## 4. 运维情况
{part4_content}

## 5. 综合经营分析
{part6_content}
```

## 工作流程

### 1. 初始化阶段
```python
# 1. 读取配置
config = load_config("openclaw.yaml")

# 2. 初始化数据加载器
loader = DataLoader(config.client_data_path)

# 3. 初始化LLM分析器
llm = LLMAnalyzer(config.api_key, config.model)
```

### 2. 数据处理阶段
```python
# 1. 加载数据
basic_data = loader.load_basic_data(client_id)
subscription_data = loader.load_subscription_data(client_id)
# ... 加载其他数据

# 2. 数据预处理
basic_data = preprocess_data(basic_data)
subscription_data = preprocess_data(subscription_data)
# ... 预处理其他数据
```

### 3. 分析生成阶段
```python
# 1. Part 1: 客户基础档案（脚本）
part1 = generate_basic_profile(basic_data)

# 2. Part 2: 订阅续约与续费（混合）
part2_1_3 = generate_subscription_stats(subscription_data)  # 脚本
part2_4 = llm.analyze_subscription(subscription_data)       # LLM

# 3. Part 3: 实施优化情况（混合）
part3_1_3_4 = generate_implementation_stats(impl_data)      # 脚本
part3_2 = llm.analyze_implementation(impl_data)             # LLM

# 4. Part 4: 运维情况（脚本）
part4 = generate_operations_stats(ops_data)

# 5. Part 6: 综合经营分析（LLM）
part6 = llm.analyze_comprehensive(all_data)
```

### 4. 报告生成阶段
```python
# 1. 整合所有部分
report_content = assemble_report(
    part1, part2_1_3, part2_4, part3_1_3_4, 
    part3_2, part4, part6
)

# 2. 生成Markdown报告
report_path = generate_markdown_report(client_id, report_content)

# 3. 保存日志
save_processing_log(client_id, report_path)
```

## 错误处理机制

### 1. 数据加载错误
- 文件不存在 → 记录错误，跳过该客户
- 格式错误 → 尝试其他编码，记录警告
- 权限错误 → 提示用户检查权限

### 2. LLM调用错误
- API密钥错误 → 提示用户检查配置
- 网络超时 → 重试机制（最多3次）
- 响应格式错误 → 尝试解析，记录警告

### 3. 处理逻辑错误
- 数据验证失败 → 记录详细错误信息
- 内存不足 → 分批处理大数据
- 磁盘空间不足 → 提示用户清理空间

## 性能优化

### 1. 缓存机制
- 缓存已处理的数据摘要
- 缓存LLM分析结果（相同输入）
- 定期清理过期缓存

### 2. 并行处理
- 多个客户可以并行处理
- 数据加载和预处理可以并行
- LLM调用可以批量处理

### 3. 增量更新
- 只处理新增或修改的数据
- 记录上次处理时间戳
- 支持增量报告生成

## 扩展性设计

### 1. 插件式架构
- 新的分析模块可以轻松添加
- 支持自定义数据处理逻辑
- 支持不同的LLM提供商

### 2. 配置驱动
- 所有路径和参数可配置
- 支持环境变量覆盖
- 支持多环境配置

### 3. 日志和监控
- 详细的处理日志
- 性能指标监控
- 错误统计和报警