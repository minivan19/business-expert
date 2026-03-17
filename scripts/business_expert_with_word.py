#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
商务专家skill完整流程：生成报告 + 转换为Word文档
"""

import os
import sys
import argparse
import logging
from pathlib import Path

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def generate_business_report(client_name, output_dir, skip_llm=False):
    """生成商务分析报告"""
    try:
        # 导入商务专家skill模块
        sys.path.insert(0, str(Path(__file__).parent))
        
        from cli import main as business_main
        from report_generator import ReportGenerator
        from data_loader import get_data_loader
        
        logger.info(f"开始生成{client_name}客户商务分析报告...")
        
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        
        # 设置参数
        class Args:
            def __init__(self):
                self.client = client_name
                self.output = output_dir
                self.data_path = None
                self.skip_llm = skip_llm
                self.verbose = True
        
        args = Args()
        
        # 生成报告
        business_main(args)
        
        # 查找生成的Markdown文件
        md_files = list(Path(output_dir).glob(f"{client_name}_经营分析报告_*.md"))
        if not md_files:
            raise FileNotFoundError(f"未找到{client_name}的Markdown报告")
        
        latest_md = max(md_files, key=lambda x: x.stat().st_mtime)
        logger.info(f"Markdown报告生成成功: {latest_md}")
        
        return str(latest_md)
        
    except Exception as e:
        logger.error(f"生成商务分析报告失败: {str(e)}")
        raise

def convert_to_word(md_path, use_template=None):
    """将Markdown报告转换为Word文档"""
    try:
        # 导入markdown-to-word-skill模块
        skill_path = Path(r"C:\Users\mingh\.openclaw\workspace\skills\markdown-to-word-skill\scripts")
        sys.path.insert(0, str(skill_path))
        
        from md2docx import MarkdownToDocxConverter
        
        # 确定输出路径
        md_file = Path(md_path)
        docx_path = md_file.with_suffix('.docx')
        
        logger.info(f"开始转换Markdown为Word: {md_path} -> {docx_path}")
        
        # 读取Markdown内容
        with open(md_path, 'r', encoding='utf-8') as f:
            markdown_content = f.read()
        
        # 选择模板
        template_path = None
        if use_template:
            template_dir = Path(r"C:\Users\mingh\.openclaw\workspace\skills\markdown-to-word-skill\templates")
            if use_template == "business":
                template_path = template_dir / "business.docx"
            elif use_template == "academic":
                template_path = template_dir / "academic.docx"
            elif use_template == "technical":
                template_path = template_dir / "technical.docx"
            
            if template_path and template_path.exists():
                logger.info(f"使用模板: {template_path}")
            else:
                logger.warning(f"模板不存在: {template_path}")
                template_path = None
        
        # 转换文档
        converter = MarkdownToDocxConverter(
            template_path=str(template_path) if template_path else None,
            debug=True
        )
        converter.convert(markdown_content)
        converter.doc.save(str(docx_path))
        
        logger.info(f"Word文档生成成功: {docx_path}")
        logger.info(f"文件大小: {docx_path.stat().st_size:,} 字节")
        
        return str(docx_path)
        
    except Exception as e:
        logger.error(f"转换为Word文档失败: {str(e)}")
        raise

def main():
    parser = argparse.ArgumentParser(description='商务专家skill完整流程')
    parser.add_argument('--client', required=True, help='客户名称')
    parser.add_argument('--output', default='./reports/', help='输出目录')
    parser.add_argument('--skip-llm', action='store_true', help='跳过LLM分析')
    parser.add_argument('--template', choices=['business', 'academic', 'technical'], 
                       default='business', help='Word模板类型')
    parser.add_argument('--skip-report', action='store_true', help='跳过报告生成，只转换现有文件')
    
    args = parser.parse_args()
    
    print("=" * 60)
    print("商务专家skill完整流程")
    print("=" * 60)
    print(f"客户: {args.client}")
    print(f"输出目录: {args.output}")
    print(f"跳过LLM: {args.skip_llm}")
    print(f"Word模板: {args.template}")
    print("=" * 60)
    
    try:
        # 步骤1: 生成商务分析报告
        md_path = None
        if not args.skip_report:
            md_path = generate_business_report(
                client_name=args.client,
                output_dir=args.output,
                skip_llm=args.skip_llm
            )
        else:
            # 查找现有的Markdown文件
            md_files = list(Path(args.output).glob(f"{args.client}_经营分析报告_*.md"))
            if not md_files:
                raise FileNotFoundError(f"未找到{args.client}的现有Markdown报告")
            latest_md = max(md_files, key=lambda x: x.stat().st_mtime)
            md_path = str(latest_md)
            logger.info(f"使用现有Markdown报告: {md_path}")
        
        # 步骤2: 转换为Word文档
        docx_path = convert_to_word(md_path, use_template=args.template)
        
        print("\n" + "=" * 60)
        print("✅ 流程完成!")
        print("=" * 60)
        print(f"Markdown报告: {md_path}")
        print(f"Word文档: {docx_path}")
        print(f"文件大小: {Path(docx_path).stat().st_size:,} 字节")
        print("=" * 60)
        
        return 0
        
    except Exception as e:
        logger.error(f"流程执行失败: {str(e)}")
        return 1

if __name__ == "__main__":
    sys.exit(main())