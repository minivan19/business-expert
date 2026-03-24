#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""虎牙2025年运维报告生成器"""
import sys, os, logging
from pathlib import Path
from datetime import datetime
import pandas as pd
import openpyxl
import subprocess

# Setup
sys.path.insert(0, os.path.dirname(__file__))
logging.basicConfig(level=logging.INFO, format='%(message)s', stream=sys.stdout)
logger = logging.getLogger(__name__)

CLIENT = '虎牙'
CLIENT_DIR = Path('/Users/limingheng/AI/client-data/客户档案') / CLIENT
OUTPUT_DIR = Path('/Users/limingheng/AI/client-data/客户报告') / CLIENT

def load_2025_work_orders():
    ops_dir = CLIENT_DIR / '运维工单'
    all_rows = []
    for f in sorted(ops_dir.glob('*.xlsx')):
        if '2025' not in f.name:
            continue
        logger.info(f"加载: {f.name}")
        try:
            wb = openpyxl.load_workbook(f, data_only=True)
            for ws in wb.worksheets:
                rows = list(ws.values)
                if not rows:
                    continue
                header = [str(c) if c is not None else '' for c in rows[0]]
                for row in rows[1:]:
                    all_rows.append(dict(zip(header, row)))
        except Exception as e:
            logger.error(f"  加载失败: {e}")
    df = pd.DataFrame(all_rows)
    logger.info(f"2025年工单总数: {len(df)}")
    return df

def main():
    df = load_2025_work_orders()
    if df.empty:
        logger.error("没有2025年运维工单数据")
        return

    from part4_operations import OperationsAnalyzer
    analyzer = OperationsAnalyzer()
    content = analyzer.analyze(df)

    today = datetime.now().strftime('%Y%m%d')
    title = f"# {CLIENT} 2025年运维情况报告\n\n_报告日期: {datetime.now().strftime('%Y-%m-%d')}_\n\n"
    content = title + content

    md_path = OUTPUT_DIR / f'{CLIENT}_2025年运维报告_{today}.md'
    docx_path = OUTPUT_DIR / f'{CLIENT}_2025年运维报告_{today}.docx'

    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(content)
    logger.info(f"Markdown: {md_path}")

    # Convert to docx
    result = subprocess.run([
        sys.executable,
        os.path.join(os.path.dirname(__file__), 'md2docx.py'),
        '--input', str(md_path),
        '--output', str(docx_path)
    ], capture_output=True, text=True)
    if result.returncode == 0:
        logger.info(f"Word: {docx_path}")
    else:
        logger.error(f"Word转换失败: {result.stderr}")
        print(result.stdout)

if __name__ == '__main__':
    main()
