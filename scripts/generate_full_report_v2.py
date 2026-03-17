#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
商务专家 - 完整报告生成脚本 v2
采用混合模式：脚本处理结构化数据 + LLM处理智能分析
"""

import os
import sys
import re
import requests
import pandas as pd
from datetime import datetime
from glob import glob

# ============== 配置 ==============
CLIENT_DATA_RAW = r"C:\Users\mingh\client-data\raw\客户档案"
OUTPUT_DIR = r"C:\Users\mingh\client-data\_temp"

# API配置
API_KEY = os.environ.get("DEEPSEEK_API_KEY", "sk-340ed7819c2346508c0a46a80df85999")
API_URL = "https://api.deepseek.com/v1/chat/completions"
MODEL = "deepseek-chat"

import re

def format_llm_list(text):
    """为LLM输出添加格式：将1.改成1）"""
    lines = text.split('\n')
    result = []
    for line in lines:
        import re
        new_line = re.sub(r'^(\d+)\.\s', r'\1） ', line)
        result.append(new_line)
    return '\n'.join(result)

def call_llm(prompt, system_prompt):
    """调用LLM"""
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    data = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.2
    }
    try:
        response = requests.post(API_URL, headers=headers, json=data, timeout=120)
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"LLM调用失败: {str(e)}"

def load_client_data(client_name):
    """加载客户所有数据"""
    client_dir = os.path.join(CLIENT_DATA_RAW, client_name)
    if not os.path.exists(client_dir):
        return None, f"客户目录不存在: {client_dir}"
    
    data = {}
    
    # Part 1: 客户主数据
    basic_path = os.path.join(client_dir, "基础数据", "客户主数据.xlsx")
    if os.path.exists(basic_path):
        data["part1"] = pd.read_excel(basic_path)
        data["part1"].columns = data["part1"].columns.str.strip()
        # 创建行业字段映射（处理列名中的不可见字符）
        # 行业字段在第18列（索引17之后，因为有隐藏字符）
        # 直接使用iloc获取
        if len(data["part1"]) > 0:
            data["part1"].loc[:, "行业_clean"] = data["part1"].iloc[:, 18]
    
    # Part 2: 订阅数据
    sub_path = os.path.join(client_dir, "订阅合同行", "订阅明细.xlsx")
    if os.path.exists(sub_path):
        data["part2"] = pd.read_excel(sub_path)
    
    # Part 2: 收款数据
    coll_path = os.path.join(client_dir, "订阅合同收款情况", "订阅合同收款情况.xlsx")
    if os.path.exists(coll_path):
        data["part4"] = pd.read_excel(coll_path)
    
    # Part 3: 实施数据
    impl_dir = os.path.join(client_dir, "实施合同行")
    if os.path.exists(impl_dir):
        fixed_path = os.path.join(impl_dir, "固定金额台账.xlsx")
        dayspan_path = os.path.join(impl_dir, "人天框架台账.xlsx")
        if os.path.exists(fixed_path):
            data["part3_fixed"] = pd.read_excel(fixed_path)
        if os.path.exists(dayspan_path):
            data["part3_dayspan"] = pd.read_excel(dayspan_path)
    
    # Part 4: 运维工单
    ops_dir = os.path.join(client_dir, "运维工单")
    if os.path.exists(ops_dir):
        ops_files = glob(os.path.join(ops_dir, "*.xlsx"))
        if ops_files:
            dfs = [pd.read_excel(f) for f in ops_files]
            data["part5_ops"] = pd.concat(dfs, ignore_index=True)
    
    return data, None

# ============== Part 1: 客户基础档案 ==============
def generate_part1(df):
    """生成客户基础档案"""
    content = "## 1. 客户基础档案\n\n"
    
    if df is None or len(df) == 0:
        return content + "暂无数据\n"
    
    d = df.iloc[0]
    
    # 1.1-1.5 各自独立处理
    content += generate_part1_1_basic(d)
    content += generate_part1_2_business(d)
    content += generate_part1_3_purchase(d)
    content += generate_part1_4_contact(d)
    content += generate_part1_5_decision(d)
    
    return content

def generate_part1_1_basic(d):
    """1.1 基本信息"""
    return """### 1.1 基本信息
| 指标 | 内容 |
|------|------|
| 客户简称 | {客户简称} |
| 客户全称 | {真实服务对象} |
| 计费ARR | {计费ARR} 元 |
| 服务阶段 | {服务阶段} |
| 客户状态 | {客户状态} |
| 所属区域 | {客户所属区域} |

""".format(
        客户简称=d.get('客户简称', 'N/A'),
        真实服务对象=d.get('真实服务对象', 'N/A'),
        计费ARR=d.get('计费ARR', 'N/A'),
        服务阶段=d.get('服务阶段', 'N/A'),
        客户状态=d.get('客户状态', 'N/A'),
        客户所属区域=d.get('客户所属区域', 'N/A')
    )

def generate_part1_2_business(d):
    """1.2 业务概况"""
    return """### 1.2 业务概况
| 指标 | 内容 |
|------|------|
| 行业 | {行业} |
| 主要产品 | {主要产品} |
| 营收规模 | {营收规模} |

""".format(
        行业=d.get('行业_clean', d.get('行业', 'N/A')),
        主要产品=d.get('主要产品', 'N/A'),
        营收规模=d.get('营收规模', 'N/A')
    )

def generate_part1_3_purchase(d):
    """1.3 购买信息"""
    return """### 1.3 购买信息
| 指标 | 内容 |
|------|------|
| 购买模块 | {购买模块} |
| 部署方式 | {部署方式} |
| 产品版本 | {产品版本} |

""".format(
        购买模块=d.get('购买模块', 'N/A'),
        部署方式=d.get('部署方式', 'N/A'),
        产品版本=d.get('产品版本', 'N/A')
    )

def generate_part1_4_contact(d):
    """1.4 服务团队"""
    return """### 1.4 服务团队
| 角色 | 姓名 |
|------|------|
| 客户成功经理 | {客户成功经理} |
| 销售主责 | {销售主责} |
| 项目经理 | {项目经理} |
| 运维主责 | {运维主责} |

""".format(
        客户成功经理=d.get('客户成功经理', 'N/A'),
        销售主责=d.get('销售主责', 'N/A'),
        项目经理=d.get('项目经理', 'N/A'),
        运维主责=d.get('运维主责', 'N/A')
    )

def generate_part1_5_decision(d):
    """1.5 决策地图"""
    return """### 1.5 决策地图
| 角色 | 信息 |
|------|------|
| IT总 | {IT高层} |
| IT经理 | {IT中层} |
| 采购总 | {采购高层} |
| 采购经理 | {采购中层} |
| 对接人 | {客户对接人} |
| 决策链 | {决策链} |

""".format(
        IT高层=d.get('IT高层（姓名-职位）', 'N/A'),
        IT中层=d.get('IT中层（姓名-职位）', 'N/A'),
        采购高层=d.get('采购高层（姓名-职位）', 'N/A'),
        采购中层=d.get('采购中层（姓名-职位）', 'N/A'),
        客户对接人=d.get('客户对接人', 'N/A'),
        决策链=d.get('决策链', 'N/A')
    )

# ============== Part 2: 订阅续约与续费 ==============
def generate_part2_static(sub_df, coll_df):
    """生成订阅续约与续费情况（脚本处理结构化数据）"""
    content = "## 2. 订阅续约与续费情况\n\n"
    
    # 2.1 概览 - 续约
    if sub_df is not None and len(sub_df) > 0:
        # 当前ARR：只看"订阅中"状态的合同行
        current_arr = 0
        if '订阅状态' in sub_df.columns and '年订阅费金额' in sub_df.columns:
            active_subs = sub_df[sub_df['订阅状态'].str.contains('订阅中', na=False)]
            current_arr = active_subs['年订阅费金额'].sum() if len(active_subs) > 0 else 0
        
        # 累计订阅费：区分计算
        # 已结束：ARR × 订阅年限
        # 订阅中：ARR × 已过年限(向上取整)
        total_subscription = 0
        today = pd.Timestamp.now()
        
        if '订阅状态' in sub_df.columns and '年订阅费金额' in sub_df.columns and '订阅有效期从' in sub_df.columns and '订阅有效期至' in sub_df.columns:
            for _, row in sub_df.iterrows():
                arr = row.get('年订阅费金额', 0)
                if pd.isna(arr) or arr == 0:
                    continue
                
                status = row.get('订阅状态', '')
                start_date = pd.to_datetime(row.get('订阅有效期从', None), errors='coerce')
                end_date = pd.to_datetime(row.get('订阅有效期至', None), errors='coerce')
                
                if pd.isna(start_date):
                    continue
                
                # 清理状态中的不可见字符
                status_str = str(status) if pd.notna(status) else ''
                status_clean = ''.join(c for c in status_str if c.isprintable() or c.isspace())
                
                if '结束' in status_clean and pd.notna(end_date):
                    # 已结束：用 (有效期至 - 有效期从) 的天数 ÷ 365 → 向上取整
                    total_days = (end_date - start_date).days
                    years = -(-total_days // 365)  # 向上取整
                    if years < 1:
                        years = 1
                    total_subscription += arr * years
                elif '订阅中' in status_clean:
                    # 订阅中：用 (今天 - 有效期从) 的天数 → 向上取整
                    days = (today - start_date).days
                    years = -(-days // 365)  # 向上取整
                    if years < 1:
                        years = 1  # 至少算1年
                    total_subscription += arr * years
        
        # 起始订阅日期：最早的有ARR（不为0）的订阅有效期从
        start_date = None
        valid_subs = sub_df[(sub_df['年订阅费金额'] > 0) & (sub_df['年订阅费金额'].notna())] if '年订阅费金额' in sub_df.columns else pd.DataFrame()
        if len(valid_subs) > 0 and '订阅有效期从' in valid_subs.columns:
            dates = pd.to_datetime(valid_subs['订阅有效期从'], errors='coerce')
            dates = dates.dropna()
            if len(dates) > 0:
                start_date = dates.min().strftime('%Y-%m-%d')
        
        # 在手订阅有效截至日期：最晚的"订阅中"状态的订阅有效期至
        end_date = None
        if '订阅状态' in sub_df.columns and '订阅有效期至' in sub_df.columns:
            active_subs = sub_df[sub_df['订阅状态'].str.contains('订阅中', na=False)]
            if len(active_subs) > 0:
                dates = pd.to_datetime(active_subs['订阅有效期至'], errors='coerce')
                dates = dates.dropna()
                if len(dates) > 0:
                    end_date = dates.max().strftime('%Y-%m-%d')
        
        content += f"""### 2.1 概览

#### 续约概览
| 指标 | 内容 |
|------|------|
| 当前ARR | {current_arr} 元 |
| 累计订阅费 | {int(total_subscription)} 元 |
| 起始订阅日期 | {start_date if start_date else 'N/A'} |
| 在手订阅有效截至日期 | {end_date if end_date else 'N/A'} |

"""
    else:
        content += "### 2.1 概览\n\n暂无订阅数据\n\n"
    
    # 2.1 概览 - 续费（收款数据）
    if coll_df is not None and len(coll_df) > 0:
        plan_total = coll_df['计划收款金额'].sum() if '计划收款金额' in coll_df.columns else 0
        received = coll_df['已收款金额'].sum() if '已收款金额' in coll_df.columns else 0
        unreceived = coll_df['未收款金额'].sum() if '未收款金额' in coll_df.columns else 0
        progress = (received / plan_total * 100) if plan_total > 0 else 0
        
        content += """#### 续费概览
| 指标 | 内容 |
|------|------|
| 计划收款总额 | {plan} 元 |
| 已收款金额 | {received} 元 |
| 未收款金额 | {unreceived} 元 |
| 整体收款进度 | {progress}% |

""".format(
            plan=plan_total,
            received=received,
            unreceived=unreceived,
            progress=round(progress, 2)
        )
    
    # 2.2 订阅明细
    content += "### 2.2 订阅明细\n\n"
    if sub_df is not None and len(sub_df) > 0:
        cols = ['合同名称', '产品名称', '订阅类别', '年订阅费金额', '订阅有效期从', '订阅有效期至', '订阅状态']
        existing_cols = [c for c in cols if c in sub_df.columns]
        if existing_cols:
            content += "| " + " | ".join(existing_cols) + " |\n"
            content += "| " + " | ".join(["---"] * len(existing_cols)) + " |\n"
            for _, row in sub_df.iterrows():
                content += "| " + " | ".join([str(row.get(c, '')) for c in existing_cols]) + " |\n"
    content += "\n"
    
    # 2.3 续费收款明细
    content += "### 2.3 续费收款明细\n\n"
    if coll_df is not None and len(coll_df) > 0:
        cols = ['项目名称', '款项名称', '期数', '计划收款金额', '已收款金额', '未收款金额', '考核收款日期']
        existing_cols = [c for c in cols if c in coll_df.columns]
        if existing_cols:
            # 添加状态列
            display_cols = existing_cols + ['状态']
            content += "| " + " | ".join(display_cols) + " |\n"
            content += "| " + " | ".join(["---"] * len(display_cols)) + " |\n"
            
            today = pd.Timestamp.now()
            for _, row in coll_df.iterrows():
                # 判断状态
                received_amt = row.get('已收款金额', 0)
                check_date = row.get('考核收款日期', None)
                status = "已收款"
                if pd.notna(received_amt) and received_amt > 0:
                    status = "已收款"
                elif pd.notna(check_date):
                    days_diff = (today - pd.to_datetime(check_date)).days
                    if days_diff < 0:
                        status = "未到期"
                    elif days_diff <= 30:
                        status = "预警"
                    else:
                        status = "严重逾期"
                else:
                    status = "未知"
                
                content += "| " + " | ".join([str(row.get(c, '')) for c in existing_cols]) + " | " + status + " |\n"
    
        return content

def generate_part2_analysis(sub_df, coll_df):
    """生成Part 2.4智能分析（LLM）"""
    if sub_df is None or len(sub_df) == 0:
        return ""
    
    # 准备订阅数据摘要
    sub_summary = ""
    if sub_df is not None:
        sub_summary = f"订阅明细共{len(sub_df)}条\n"
        if '合同名称' in sub_df.columns and '年订阅费金额' in sub_df.columns:
            for _, r in sub_df.iterrows():
                sub_summary += f"- {r.get('合同名称', '')}: {r.get('年订阅费金额', 0)}元\n"
    
    # 准备收款数据摘要
    coll_summary = ""
    if coll_df is not None:
        coll_summary = f"收款明细共{len(coll_df)}条\n"
        if '项目名称' in coll_df.columns and '计划收款金额' in coll_df.columns:
            for _, r in coll_df.iterrows():
                coll_summary += f"- {r.get('项目名称', '')}: 计划{r.get('计划收款金额', 0)}元, 已收{r.get('已收款金额', 0)}元\n"
    
    system_prompt = """你是一个客户成功经理助手。根据订阅和收款数据，进行智能分析。

## 重要约束（必须遵守）
- 禁止使用**加粗**等Markdown格式
- 禁止使用##标题标记
- 禁止使用"好的"、"作为您的"、"我已"、"这是"等客套话
- 直接输出分析内容，用纯文本
- 编号用"1. "格式，不要用"-"列表
- 禁止使用1.1.、2.2.等嵌套编号"""

    prompt = f"""请分析以下订阅和收款数据：

订阅数据：
{sub_summary}

收款数据：
{coll_summary}

请直接输出分析内容（不要带### 2.4标题），包括：

1. 续约分析：订阅费用的时间阶段变化、续约/降价原因分析
2. 续费分析：收款进度评估、逾期风险分析、预测建议"""

    result = call_llm(prompt, system_prompt)
    result = format_llm_list(result)
    return f"\n### 2.4 智能分析\n\n{result}\n"

# ============== Part 3: 实施优化 ==============
def generate_part3_overview(fixed_df, dayspan_df):
    """生成3.1实施概览"""
    content = "### 3.1 实施概览\n\n"
    
    yearly_data = {}
    
    # 处理固定金额合同
    if fixed_df is not None and len(fixed_df) > 0:
        if '合同签订时间' in fixed_df.columns and '固定金额' in fixed_df.columns:
            for _, row in fixed_df.iterrows():
                try:
                    year = pd.to_datetime(row['合同签订时间']).year
                    amt = row.get('固定金额', 0) or 0
                    if year not in yearly_data:
                        yearly_data[year] = {'fixed': 0, 'dayspan': 0}
                    yearly_data[year]['fixed'] += amt
                except:
                    pass
    
    # 处理人天框架
    if dayspan_df is not None and len(dayspan_df) > 0:
        if '创建日期' in dayspan_df.columns:
            for _, row in dayspan_df.iterrows():
                try:
                    year = pd.to_datetime(row['创建日期']).year
                    amt = row.get('应收金额', 0) or 0
                    if year not in yearly_data:
                        yearly_data[year] = {'fixed': 0, 'dayspan': 0}
                    yearly_data[year]['dayspan'] += amt
                except:
                    pass
    
    if yearly_data:
        content += "| 年份 | 固定金额合同(元) | 人天框架(元) | 实施优化总金额(元) |\n"
        content += "|------|----------------|-------------|-----------------|\n"
        total_fixed = 0
        total_dayspan = 0
        for year in sorted(yearly_data.keys()):
            fixed = yearly_data[year]['fixed']
            dayspan = yearly_data[year]['dayspan']
            total_fixed += fixed
            total_dayspan += dayspan
            content += f"| {year} | {fixed} | {dayspan} | {fixed + dayspan} |\n"
        content += f"| 汇总 | {total_fixed} | {total_dayspan} | {total_fixed + total_dayspan} |\n"
    
    return content + "\n"

def generate_part3_fixed_table(fixed_df):
    """生成3.3固定金额合同明细"""
    content = "### 3.2 固定金额合同明细\n\n"
    if fixed_df is not None and len(fixed_df) > 0:
        cols = ['合同名称', '产品名称', '固定金额', '合同签订时间']
        existing_cols = [c for c in cols if c in fixed_df.columns]
        if existing_cols:
            content += "| " + " | ".join(existing_cols) + " |\n"
            content += "| " + " | ".join(["---"] * len(existing_cols)) + " |\n"
            for _, row in fixed_df.iterrows():
                content += "| " + " | ".join([str(row.get(c, '')) for c in existing_cols]) + " |\n"
    return content + "\n"

def generate_part3_dayspan_table(dayspan_df):
    """生成3.4人天框架合同明细"""
    content = "### 3.3 人天框架合同明细\n\n"
    if dayspan_df is not None and len(dayspan_df) > 0:
        cols = ['合同名称', '产品名称', '总人天', '应收金额']
        existing_cols = [c for c in cols if c in dayspan_df.columns]
        if existing_cols:
            content += "| " + " | ".join(existing_cols) + " |\n"
            content += "| " + " | ".join(["---"] * len(existing_cols)) + " |\n"
            for _, row in dayspan_df.iterrows():
                content += "| " + " | ".join([str(row.get(c, '')) for c in existing_cols]) + " |\n"
    return content

def generate_part3_analysis(fixed_df, dayspan_df):
    """生成3.2智能分析（LLM）"""
    content = "### 3.4 智能分析\n\n"
    
    # 准备数据摘要
    summary = ""
    if fixed_df is not None and len(fixed_df) > 0:
        summary += "固定金额合同：\n"
        for _, r in fixed_df.iterrows():
            summary += f"- {r.get('合同名称', '')}: {r.get('固定金额', 0)}元\n"
    
    if dayspan_df is not None and len(dayspan_df) > 0:
        summary += "\n人天框架合同：\n"
        for _, r in dayspan_df.iterrows():
            summary += f"- {r.get('合同名称', '')}: {r.get('应收金额', 0)}元\n"
    
    system_prompt = """你是一个客户成功经理助手。根据实施合同数据，进行智能分析。

## 重要约束（必须遵守）
- 禁止使用**加粗**等Markdown格式
- 禁止使用##标题标记
- 禁止使用客套话
- 直接输出分析内容，用纯文本
- 编号用"1. "格式
- 禁止使用1.1.、2.2.等嵌套编号"""

    prompt = f"""请分析以下实施合同数据：

{summary}

请直接输出分析内容（不要带### 3.2标题），包括：
1. 实施优化费用的年度变化趋势
2. 实施模式分析（固定金额 vs 人天框架）
3. 预测和建议"""

    result = call_llm(prompt, system_prompt)
    result = format_llm_list(result)
    return content + result + "\n"

# ============== Part 4: 运维情况 ==============
def generate_part4(ops_df):
    """生成运维情况（从运维工单数据）"""
    content = "## 4. 运维情况\n\n"
    
    if ops_df is None or len(ops_df) == 0:
        content += "暂无运维工单数据\n"
        return content
    
    # 提单时间在第13列，工时在第14列（索引）
    time_col = ops_df.columns[12]  # 提单时间
    hours_col = ops_df.columns[13]  # 工时
    
    # 按年份统计
    ops_df['年份'] = pd.to_datetime(ops_df[time_col], errors='coerce').dt.year
    
    # 解析工时（格式如"1小时15分"）
    def parse_hours(s):
        if pd.isna(s):
            return 0
        s = str(s)
        total = 0
        # 匹配小时
        import re
        hour_match = re.search(r'(\d+)小时', s)
        if hour_match:
            total += int(hour_match.group(1))
        # 匹配分钟
        min_match = re.search(r'(\d+)分', s)
        if min_match:
            total += int(min_match.group(1)) / 60
        return total
    
    ops_df['工时_数值'] = ops_df[hours_col].apply(parse_hours)
    
    # 按年份汇总
    yearly = ops_df.groupby('年份').agg({
        '工时_数值': 'sum'
    }).reset_index()
    yearly['工单数量'] = ops_df.groupby('年份').size().values
    
    content += "### 4.1 运维工单概览\n\n"
    content += "| 年份 | 工单数量 | 工时汇总(小时) |\n"
    content += "|------|---------|---------------|\n"
    
    total_count = 0
    total_hours = 0
    for _, row in yearly.iterrows():
        if pd.notna(row['年份']):
            content += f"| {int(row['年份'])} | {int(row['工单数量'])} | {round(row['工时_数值'], 2)} |\n"
            total_count += row['工单数量']
            total_hours += row['工时_数值']
    
    content += f"| 汇总 | {int(total_count)} | {round(total_hours, 2)} |\n"
    
    return content

# ============== Part 6: 综合分析 ==============
def generate_part6(data):
    """生成综合经营分析（LLM）"""
    # 汇总所有数据摘要
    summary = "## 客户经营分析数据摘要\n\n"
    
    # Part 1
    if 'part1' in data and data['part1'] is not None:
        d = data['part1'].iloc[0]
        summary += f"客户: {d.get('客户简称', 'N/A')}, ARR: {d.get('计费ARR', 'N/A')}元\n"
    
    # Part 2
    if 'part2' in data and data['part2'] is not None:
        sub = data['part2']
        summary += f"订阅: {len(sub)}个合同\n"
    
    # Part 4 (收款)
    if 'part4' in data and data['part4'] is not None:
        coll = data['part4']
        plan = coll['计划收款金额'].sum() if '计划收款金额' in coll.columns else 0
        received = coll['已收款金额'].sum() if '已收款金额' in coll.columns else 0
        summary += f"收款: 已收{received}元/计划{plan}元 ({received/plan*100:.1f}%)\n"
    
    # Part 5 (运维)
    if 'part5_ops' in data and data['part5_ops'] is not None:
        ops = data['part5_ops']
        summary += f"运维: {len(ops)}个工单\n"
    
    system_prompt = """你是一个客户成功经理助手。根据客户经营数据，生成综合经营分析报告。

## 重要约束（必须遵守）
- 禁止使用**加粗**等Markdown格式
- 禁止使用##标题标记
- 禁止使用"好的"、"作为您的"、"我已"、"这些都是"等客套话
- 直接输出内容，用纯文本
- 编号用"1. "格式，不要用"-"列表
- 禁止使用1.1.、2.2.等嵌套编号"""

    prompt = f"""{summary}

请直接输出分析内容（不要带## 6和### 6.x标题），包括：

6.1 客户价值分级
价值等级（A/B/C）
评估依据

6.2 经营健康度评估
订阅健康度：
收款健康度：
运维健康度：

6.3 机会分析
增购机会：
续约机会：

6.4 风险预警
流失风险：
收款风险：

6.5 下一步行动建议
短期（1-3个月）：
中期（3-6个月）："""
    
    result = call_llm(prompt, system_prompt)
    result = format_llm_list(result)
    return f"\n## 6. 综合经营分析\n\n{result}\n"

# ============== 主函数 ==============
def generate_report(client_name):
    """生成完整报告"""
    print(f"\n{'='*60}")
    print(f"生成客户经营分析报告: {client_name}")
    print("=" * 60)
    
    # 加载数据
    data, error = load_client_data(client_name)
    if error:
        print(f"错误: {error}")
        return None
    
    report = []
    report.append(f"# {client_name} 经营分析报告")
    report.append(f"\n生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    # Part 1: 客户基础档案（脚本处理）
    print("生成 Part 1: 客户基础档案...")
    if "part1" in data:
        report.append(generate_part1(data["part1"]))
    
    # Part 2: 订阅续约与续费
    print("生成 Part 2: 订阅续约与续费...")
    # 合并2.1-2.4，统一输出顺序（函数内已包含标题）
    
    # 2.1-2.3 脚本处理
    part2_content = generate_part2_static(
        data.get('part2'),
        data.get('part4')
    )
    
    # 2.4 LLM分析
    part2_content += generate_part2_analysis(
        data.get('part2'),
        data.get('part4')
    )
    
    report.append(part2_content)
    
    # Part 3: 实施优化
    print("生成 Part 3: 实施优化...")
    # 正确顺序：3.1概览 → 3.2固定金额 → 3.3人天 → 3.4智能分析
    part3_content = "## 3. 实施优化情况\n\n"
    
    # 3.1 实施概览
    part3_content += generate_part3_overview(data.get('part3_fixed'), data.get('part3_dayspan'))
    
    # 3.2 固定金额合同明细
    part3_content += generate_part3_fixed_table(data.get('part3_fixed'))
    
    # 3.3 人天框架合同明细
    part3_content += generate_part3_dayspan_table(data.get('part3_dayspan'))
    
    # 3.4 智能分析（LLM）
    part3_content += generate_part3_analysis(data.get('part3_fixed'), data.get('part3_dayspan'))
    
    report.append(part3_content)
    
    # Part 4: 运维情况
    print("生成 Part 4: 运维情况...")
    if "part5_ops" in data:
        report.append(generate_part4(data["part5_ops"]))
    else:
        report.append("## 4. 运维情况\n\n暂无运维工单数据\n")
    
    # Part 5: 客户经营信息（待补充）
    report.append("\n## 5. 客户经营信息\n\n请手动搜索补充...\n")
    
    # Part 6: 综合分析（LLM）
    print("生成 Part 6: 综合经营分析...")
    report.append(generate_part6(data))
    
    # 保存
    report_text = "\n".join(report)
    output_file = os.path.join(OUTPUT_DIR, f"{client_name}_经营分析报告.md")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(report_text)
    
    print(f"\n报告已保存: {output_file}")
    return report_text

def main():
    print("=" * 60)
    print("商务专家 - 完整报告生成脚本 v2")
    print("=" * 60)
    
    if not os.path.exists(CLIENT_DATA_RAW):
        print(f"错误: 客户档案目录不存在: {CLIENT_DATA_RAW}")
        return
    
    clients = sorted([d for d in os.listdir(CLIENT_DATA_RAW) 
                     if os.path.isdir(os.path.join(CLIENT_DATA_RAW, d))])
    
    print(f"\n找到 {len(clients)} 个客户")
    for i, c in enumerate(clients[:10]):
        print(f"  {i+1}. {c}")
    if len(clients) > 10:
        print(f"  ... 共{len(clients)}个")
    
    print(f"\n输入客户编号（1-{len(clients)}），或输入 'all' 生成所有客户：")
    choice = input("> ").strip()
    
    if choice.lower() == "all":
        selected = clients
    else:
        try:
            idx = int(choice) - 1
            selected = [clients[idx]] if 0 <= idx < len(clients) else []
        except:
            selected = []
    
    if not selected:
        print("无效选择")
        return
    
    for client in selected:
        generate_report(client)
    
    print("\n完成!")

if __name__ == "__main__":
    main()
