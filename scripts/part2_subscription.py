#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Part 2: 订阅续约与续费分析模块
分析客户的订阅情况、续约趋势和收款进度
"""

import pandas as pd
import numpy as np
import logging
from datetime import datetime

# 导入LLM客户端
try:
    from llm_client import get_llm_client
    LLM_AVAILABLE = True
except ImportError:
    LLM_AVAILABLE = False
    logger.warning("LLM客户端不可用，智能分析将使用默认内容")

logger = logging.getLogger(__name__)


class SubscriptionAnalyzer:
    """订阅续约与续费分析器"""

    def __init__(self):
        # 订阅明细字段映射
        self.sub_fields = {
            '签约日期': '合同签约日期',
            '合同行号': '合同行号',
            '订阅类别': '订阅类别',
            '年订阅费金额': '年订阅费金额',
            '订阅有效期从': '订阅有效期从',
            '订阅有效期至': '订阅有效期至',
            '订阅状态': '订阅状态',
        }

        # 收款明细字段映射
        self.coll_fields = {
            '项目编号': '项目编码',
            '期数': '期数',
            '计划收款金额': '计划收款金额',
            '考核收款日期': '考核收款日期',
            '已收款金额': '已收款金额',
            '未收款金额': '未收款金额',
        }

    def analyze(self, df_subscription, df_collection):
        """
        分析订阅续约与续费情况

        Args:
            df_subscription: 订阅数据DataFrame
            df_collection: 收款数据DataFrame

        Returns:
            str: Markdown格式的分析报告
        """
        logger.info("开始分析订阅续约与续费情况")

        content = "## 2. 订阅续约与续费\n\n"

        try:
            # 2.1 概览
            overview_content = self._generate_overview(df_subscription, df_collection)
            content += overview_content

            # 2.2 订阅明细
            details_content = self._generate_subscription_details(df_subscription)
            content += details_content

            # 2.3 续费收款明细
            collection_content = self._generate_collection_details(df_collection)
            content += collection_content

            # 保存已生成的分析内容，供2.4使用
            self._generated_analysis = overview_content + "\n" + details_content + "\n" + collection_content

            # 2.4 智能分析（基于已生成的分析内容）
            content += self._generate_intelligent_analysis(df_subscription, df_collection)

            logger.info("订阅续约与续费分析完成")
            return content

        except Exception as e:
            error_msg = f"订阅续约与续费分析失败: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return content + f"分析失败: {error_msg}\n"

    def _generate_overview(self, df_subscription, df_collection):
        """生成2.1概览"""
        content = "### 2.1 概览\n\n"

        # 2.1.1 续约概览
        content += "#### 2.1.1 续约概览\n\n"

        if df_subscription is None or df_subscription.empty:
            content += "| 当前ARR | 起始订阅日期 | 在手订阅有效截至日期 |\n"
            content += "|---------|-------------|---------------------|\n"
            content += "| - | - | - |\n\n"
        else:
            # 当前ARR：订阅状态="订阅中"，年订阅费金额汇总
            active_subs = df_subscription[df_subscription['订阅状态'] == '订阅中']
            current_arr = active_subs['年订阅费金额'].sum() if not active_subs.empty else 0

            # 起始订阅日期：年订阅费金额≠0，订阅有效期从取最早
            valid_subs = df_subscription[df_subscription['年订阅费金额'].notna() & (df_subscription['年订阅费金额'] > 0)]
            if not valid_subs.empty:
                start_dates = pd.to_datetime(valid_subs['订阅有效期从'], errors='coerce')
                start_date_min = start_dates.min()
                start_date_str = start_date_min.strftime('%Y-%m-%d') if pd.notna(start_date_min) else '-'

                # 在手订阅有效截至日期：年订阅费金额≠0，订阅有效期至取最晚
                end_dates = pd.to_datetime(valid_subs['订阅有效期至'], errors='coerce')
                end_date_max = end_dates.max()
                end_date_str = end_date_max.strftime('%Y-%m-%d') if pd.notna(end_date_max) else '-'
            else:
                start_date_str = '-'
                end_date_str = '-'

            content += "| 当前ARR | 起始订阅日期 | 在手订阅有效截至日期 |\n"
            content += "|---------|-------------|---------------------|\n"
            content += f"| {current_arr:,.0f}元 | {start_date_str} | {end_date_str} |\n\n"

        # 2.1.2 续费概览
        content += "#### 2.1.2 续费概览\n\n"

        if df_collection is None or df_collection.empty:
            content += "| 计划收款总额 | 已收款金额 | 未收款金额 |\n"
            content += "|-------------|-----------|-----------|\n"
            content += "| - | - | - |\n\n"
        else:
            total_planned = df_collection['计划收款金额'].sum()
            total_received = df_collection['已收款金额'].sum()
            total_unreceived = df_collection['未收款金额'].sum()

            content += "| 计划收款总额 | 已收款金额 | 未收款金额 |\n"
            content += "|-------------|-----------|-----------|\n"
            content += f"| {total_planned:,.0f}元 | {total_received:,.0f}元 | {total_unreceived:,.0f}元 |\n\n"

        return content

    def _generate_subscription_details(self, df):
        """生成2.2订阅明细"""
        content = "### 2.2 订阅明细\n\n"

        if df is None or df.empty:
            content += "暂无订阅明细数据\n\n"
            return content

        # 筛选需要的字段
        available_fields = {}
        for field_name, col_name in self.sub_fields.items():
            if col_name in df.columns:
                available_fields[field_name] = col_name

        if not available_fields:
            content += "订阅明细数据格式不正确\n\n"
            return content

        # 选择字段并排序
        df_select = df[list(available_fields.values())].copy()
        df_select.columns = list(available_fields.keys())

        # 按订阅有效期从 从新到旧排序
        if '订阅有效期从' in df_select.columns:
            df_select['订阅有效期从'] = pd.to_datetime(df_select['订阅有效期从'], errors='coerce')
            df_select = df_select.sort_values('订阅有效期从', ascending=False)

        # 格式化日期 - 先转换为datetime再格式化
        for col in ['签约日期', '订阅有效期从', '订阅有效期至']:
            if col in df_select.columns:
                # 先转换为datetime
                df_select[col] = pd.to_datetime(df_select[col], errors='coerce')
                # 再格式化为字符串
                df_select[col] = df_select[col].apply(
                    lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else '-'
                )

        # 格式化金额
        if '年订阅费金额' in df_select.columns:
            df_select['年订阅费金额'] = df_select['年订阅费金额'].apply(
                lambda x: f"{x:,.0f}元" if pd.notna(x) else '-'
            )

        # 生成表格
        content += "| " + " | ".join(df_select.columns) + " |\n"
        content += "|" + "|".join([" --- " for _ in df_select.columns]) + "|\n"

        for _, row in df_select.iterrows():
            content += "| " + " | ".join(str(v) for v in row.values) + " |\n"

        content += "\n"
        return content

    def _generate_collection_details(self, df):
        """生成2.3续费收款明细"""
        content = "### 2.3 续费收款明细\n\n"

        if df is None or df.empty:
            content += "暂无收款明细数据\n\n"
            return content

        # 筛选需要的字段
        available_fields = {}
        for field_name, col_name in self.coll_fields.items():
            if col_name in df.columns:
                available_fields[field_name] = col_name

        if not available_fields:
            content += "收款明细数据格式不正确\n\n"
            return content

        # 选择字段并排序
        df_select = df[list(available_fields.values())].copy()
        df_select.columns = list(available_fields.keys())

        # 按考核收款日期 从新到旧排序
        if '考核收款日期' in df_select.columns:
            df_select['考核收款日期'] = pd.to_datetime(df_select['考核收款日期'], errors='coerce')
            df_select = df_select.sort_values('考核收款日期', ascending=False)

        # 格式化日期
        if '考核收款日期' in df_select.columns:
            df_select['考核收款日期'] = df_select['考核收款日期'].apply(
                lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else '-'
            )

        # 格式化金额
        for col in ['计划收款金额', '已收款金额', '未收款金额']:
            if col in df_select.columns:
                df_select[col] = df_select[col].apply(
                    lambda x: f"{x:,.0f}元" if pd.notna(x) else '-'
                )

        # 生成表格
        content += "| " + " | ".join(df_select.columns) + " |\n"
        content += "|" + "|".join([" --- " for _ in df_select.columns]) + "|\n"

        for _, row in df_select.iterrows():
            content += "| " + " | ".join(str(v) for v in row.values) + " |\n"

        content += "\n"
        return content

    def _generate_intelligent_analysis(self, df_subscription, df_collection):
        """生成2.4智能分析（基于已生成的分析内容）"""
        content = "### 2.4 智能分析\n\n"

        # 使用已生成的分析内容供LLM分析
        generated_content = getattr(self, '_generated_analysis', '') or ''

        # 调用LLM进行智能分析（使用已生成的分析内容）
        if LLM_AVAILABLE:
            try:
                logger.info("开始调用LLM进行订阅续约智能分析（基于2.1-2.3分析内容）...")
                llm_client = get_llm_client()
                llm_result = llm_client.analyze_subscription_from_content(generated_content)

                if llm_result:
                    # 清理LLM输出中的标题和序号
                    cleaned = self._clean_llm_output(llm_result)
                    content += f"{cleaned}\n\n"
                else:
                    content += "- LLM调用失败\n\n"

            except Exception as e:
                logger.error(f"LLM调用失败: {e}")
                content += f"- LLM调用异常: {str(e)}\n\n"
        else:
            content += "- LLM不可用\n\n"

        return content

    def _clean_llm_output(self, text):
        """清理LLM输出，去除标题序号和多余符号"""
        import re
        # 去除以"2.4"、"3.4"、"4.4"、"5."、"6."等开头的标题
        text = re.sub(r'^\d+\.\d+\s+', '', text, flags=re.MULTILINE)
        # 去除"**"加粗标记
        text = text.replace('**', '')
        # 去除多余的*号（列表项保留但去掉开头的*）
        text = re.sub(r'^\s*[-*]\s+', '- ', text, flags=re.MULTILINE)
        return text.strip()

    def _prepare_full_data(self, df_subscription, df_collection):
        """准备本章节统计好的数据供LLM分析使用"""
        lines = []

        # 2.1 订阅概览统计
        lines.append("【2.1 订阅概览】")
        if df_subscription is not None and not df_subscription.empty:
            total = len(df_subscription)
            lines.append(f"- 总订阅记录数: {total}")

            if '订阅状态' in df_subscription.columns:
                status_counts = df_subscription['订阅状态'].value_counts()
                for status, count in status_counts.items():
                    lines.append(f"  - {status}: {count}条")

            if '年订阅费金额' in df_subscription.columns:
                total_amount = df_subscription['年订阅费金额'].sum()
                lines.append(f"- 年订阅费总金额: {total_amount:,.0f}元")
        else:
            lines.append("- 无订阅数据")

        # 2.2 订阅明细统计（关键数据）
        lines.append("\n【2.2 订阅明细（关键数据）】")
        if df_subscription is not None and not df_subscription.empty:
            key_cols = ['合同编号', '产品名称', '订阅状态', '年订阅费金额']
            available = [c for c in key_cols if c in df_subscription.columns]
            if available:
                df_show = df_subscription[available].head(10)
                lines.append("| " + " | ".join(available) + " |")
                lines.append("|---" * len(available) + "|")
                for _, row in df_show.iterrows():
                    vals = []
                    for c in available:
                        v = row.get(c, '')
                        if pd.isna(v): vals.append('-')
                        elif isinstance(v, (int,float)): vals.append(f"{v:,.0f}")
                        else: vals.append(str(v)[:20])
                    lines.append("| " + " | ".join(vals) + " |")
            if len(df_subscription) > 10:
                lines.append(f"... (共{total}条)")
        else:
            lines.append("- 无订阅明细数据")

        # 2.3 续费收款明细统计
        lines.append("\n【2.3 续费收款明细统计】")
        if df_collection is not None and not df_collection.empty:
            total_planned = df_collection['计划收款金额'].sum() if '计划收款金额' in df_collection.columns else 0
            total_received = df_collection['已收款金额'].sum() if '已收款金额' in df_collection.columns else 0
            total_unpaid = df_collection['未收款金额'].sum() if '未收款金额' in df_collection.columns else 0

            lines.append(f"- 计划收款总额: {total_planned:,.0f}元")
            lines.append(f"- 已收款金额: {total_received:,.0f}元")
            lines.append(f"- 未收款金额: {total_unpaid:,.0f}元")

            if '项目状态' in df_collection.columns:
                status_counts = df_collection['项目状态'].value_counts()
                lines.append("- 收款状态分布:")
                for status, count in status_counts.items():
                    lines.append(f"  - {status}: {count}条")
        else:
            lines.append("- 无收款数据")

        return "\n".join(lines)

    def _prepare_data_summary(self, df_subscription, df_collection):
        """准备数据摘要供LLM分析使用"""
        summary = []

        # 订阅数据摘要
        if df_subscription is not None and not df_subscription.empty:
            summary.append("订阅数据摘要:")
            summary.append(f"- 总订阅记录数: {len(df_subscription)}")

            if '订阅状态' in df_subscription.columns:
                status_counts = df_subscription['订阅状态'].value_counts()
                summary.append("- 状态分布:")
                for status, count in status_counts.items():
                    summary.append(f"  {status}: {count}条")

            if '年订阅费金额' in df_subscription.columns:
                total = df_subscription['年订阅费金额'].sum()
                summary.append(f"- 年订阅费总金额: {total:,.0f}元")

        # 收款数据摘要
        if df_collection is not None and not df_collection.empty:
            summary.append("\n收款数据摘要:")
            summary.append(f"- 总收款记录数: {len(df_collection)}")

            for col in ['计划收款金额', '已收款金额', '未收款金额']:
                if col in df_collection.columns:
                    total = df_collection[col].sum()
                    summary.append(f"- {col}: {total:,.0f}元")

            if '项目状态' in df_collection.columns:
                status_counts = df_collection['项目状态'].value_counts()
                summary.append("- 项目状态分布:")
                for status, count in status_counts.items():
                    summary.append(f"  {status}: {count}条")

        return "\n".join(summary)


# 测试代码
if __name__ == "__main__":
    import os

    # 测试加载数据
    client_path = r"C:\Users\mingh\client-data\raw\客户档案\CBD"
    sub_path = os.path.join(client_path, "订阅合同行", "订阅明细.xlsx")
    coll_path = os.path.join(client_path, "订阅合同收款情况", "订阅合同收款情况.xlsx")

    df_sub = pd.read_excel(sub_path) if os.path.exists(sub_path) else None
    df_coll = pd.read_excel(coll_path) if os.path.exists(coll_path) else None

    # 分析
    analyzer = SubscriptionAnalyzer()
    result = analyzer.analyze(df_sub, df_coll)

    print(result)
