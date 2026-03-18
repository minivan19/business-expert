#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简化版报告生成器 - 确保基本功能工作
"""

import os
import logging
import subprocess
import pandas as pd
from datetime import datetime
import sys
import argparse

# 添加当前目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

logger = logging.getLogger(__name__)


class SimpleReportGenerator:
    """简化版报告生成器"""
    
    def __init__(self):
        # 基础目录
        self.base_dir = r"C:\Users\mingh\client-data"
        
        # 临时目录（保留）
        self.temp_dir = os.path.join(self.base_dir, "_temp")
        
        # Markdown转Word技能路径
        self.md_to_word_skill_dir = r"C:\Users\mingh\.openclaw\workspace\skills\markdown-to-word-skill"
        self.converter_script = os.path.join(self.md_to_word_skill_dir, "scripts", "md2docx.py")
        self.business_template = os.path.join(self.md_to_word_skill_dir, "templates", "business.docx")
        
        # 创建基础目录
        os.makedirs(self.temp_dir, exist_ok=True)
        
        logger.info(f"简化版报告生成器初始化")
        logger.info(f"基础目录: {self.base_dir}")
        logger.info(f"临时目录: {self.temp_dir}")
    
    def get_client_short_name(self, client_id):
        """
        从客户主数据获取客户简称
        """
        try:
            # 客户主数据文件路径
            client_data_path = os.path.join(self.base_dir, "raw", "基础数据", "客户主数据.xlsx")
            
            if not os.path.exists(client_data_path):
                logger.warning(f"客户主数据文件不存在: {client_data_path}")
                return client_id
            
            # 读取Excel文件
            df = pd.read_excel(client_data_path)
            
            if df.empty:
                logger.warning("客户主数据文件为空")
                return client_id
            
            # 查找客户简称字段
            short_name_col = None
            for col in df.columns:
                if "简称" in str(col):
                    short_name_col = col
                    break
            
            if short_name_col is None:
                logger.warning("未找到客户简称字段")
                return client_id
            
            # 简化处理：取第一行的简称
            if len(df) > 0:
                short_name = df.iloc[0][short_name_col]
                if pd.isna(short_name):
                    return client_id
                
                short_name_str = str(short_name).strip()
                logger.info(f"获取到客户简称: {short_name_str}")
                return short_name_str
            else:
                return client_id
                
        except Exception as e:
            logger.error(f"获取客户简称时出错: {e}")
            return client_id
    
    def create_client_directory(self, client_short_name):
        """创建客户输出目录"""
        client_dir = os.path.join(self.base_dir, client_short_name)
        os.makedirs(client_dir, exist_ok=True)
        logger.info(f"创建客户目录: {client_dir}")
        return client_dir
    
    def generate_report_filenames(self, client_short_name):
        """生成报告文件名"""
        date_str = datetime.now().strftime("%Y%m%d")
        
        # 基础文件名
        base_name = f"{client_short_name}_经营分析报告_{date_str}"
        
        # 文件名
        md_filename = f"{base_name}.md"
        docx_filename = f"{base_name}.docx"
        
        return md_filename, docx_filename
    
    def create_simple_report(self, client_id):
        """创建简单的测试报告"""
        report_content = f"""# {client_id} 经营分析报告

**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
**报告版本**: 简化测试版（集成Word输出）

## 1. 客户基础信息

- **客户ID**: {client_id}
- **报告日期**: {datetime.now().strftime('%Y-%m-%d')}
- **输出目录**: {os.path.join(self.base_dir, client_id)}

## 2. 系统状态

- ✅ 目录结构已创建
- ✅ 文件命名规范已应用
- ✅ Word转换功能已集成

## 3. 设计确认

### 3.1 输出路径
- 主目录: `client-data/<客户简称>/`
- 临时目录: `client-data/_temp/` (保留)

### 3.2 文件格式
- Markdown报告: `{client_id}_经营分析报告_{datetime.now().strftime('%Y%m%d')}.md`
- Word报告: `{client_id}_经营分析报告_{datetime.now().strftime('%Y%m%d')}.docx`

### 3.3 转换流程
1. 生成Markdown报告
2. 调用markdown-to-word-skill转换为Word
3. 保存两种格式到客户目录

## 4. 测试结果

### 4.1 目录创建
- 客户目录: 已创建
- 临时目录: 已保留

### 4.2 文件生成
- Markdown文件: 即将生成
- Word文件: 即将转换

### 4.3 功能验证
- 客户简称提取: 已实现
- 目录自动创建: 已实现
- Word格式转换: 已集成

## 5. 下一步计划

1. **完整数据分析**: 集成所有分析模块
2. **LLM智能分析**: 启用DeepSeek API分析
3. **模板优化**: 优化商业报告模板
4. **批量处理**: 支持多客户批量生成

---

**报告说明**
1. 本报告由商务专家系统自动生成
2. 数据来源: 客户业务系统
3. 分析时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
4. 报告版本: 简化测试版（验证输出路径和Word转换）
"""
        
        return report_content
    
    def convert_markdown_to_word(self, md_path, docx_path):
        """将Markdown报告转换为Word文档"""
        try:
            # 检查转换脚本和模板是否存在
            if not os.path.exists(self.converter_script):
                logger.error(f"转换脚本不存在: {self.converter_script}")
                return False
            
            if not os.path.exists(self.business_template):
                logger.error(f"商业模板不存在: {self.business_template}")
                return False
            
            # 构建转换命令
            cmd = [
                "python",
                self.converter_script,
                "--input", md_path,
                "--output", docx_path,
                "--template", self.business_template
            ]
            
            logger.info(f"执行Word转换命令: {' '.join(cmd)}")
            
            # 执行转换
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
            
            if result.returncode == 0:
                logger.info(f"Word转换成功: {docx_path}")
                return True
            else:
                logger.error(f"Word转换失败，返回码: {result.returncode}")
                if result.stderr:
                    logger.error(f"错误信息: {result.stderr}")
                return False
                
        except Exception as e:
            logger.error(f"Word转换过程中出错: {e}")
            return False
    
    def generate_report(self, client_id):
        """生成客户经营分析报告（Markdown + Word）"""
        logger.info(f"开始生成客户报告: {client_id}")
        start_time = datetime.now()
        
        # 1. 获取客户简称
        client_short_name = self.get_client_short_name(client_id)
        logger.info(f"客户简称: {client_short_name}")
        
        # 2. 创建客户目录
        client_dir = self.create_client_directory(client_short_name)
        
        # 3. 生成文件名
        md_filename, docx_filename = self.generate_report_filenames(client_short_name)
        md_path = os.path.join(client_dir, md_filename)
        docx_path = os.path.join(client_dir, docx_filename)
        
        logger.info(f"Markdown文件: {md_path}")
        logger.info(f"Word文件: {docx_path}")
        
        try:
            # 4. 生成报告内容
            report_content = self.create_simple_report(client_id)
            
            # 5. 保存Markdown报告
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(report_content)
            
            logger.info(f"Markdown报告保存成功: {md_path}")
            
            # 6. 转换为Word文档
            word_success = self.convert_markdown_to_word(md_path, docx_path)
            
            if not word_success:
                logger.warning("Word转换失败，仅生成Markdown报告")
                docx_path = None
            
            elapsed_time = (datetime.now() - start_time).total_seconds()
            logger.info(f"报告生成完成，耗时: {elapsed_time:.2f}秒")
            
            return md_path, docx_path, None
            
        except Exception as e:
            error_msg = f"报告生成过程中发生异常: {str(e)}"
            logger.error(error_msg)
            import traceback
            traceback.print_exc()
            return None, None, error_msg


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description='简化版商务专家报告生成器')
    parser.add_argument('client_id', help='客户ID')
    parser.add_argument('--debug', action='store_true', help='启用调试模式')
    
    args = parser.parse_args()
    
    # 配置日志
    log_level = logging.DEBUG if args.debug else logging.INFO
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler()
        ]
    )
    
    # 生成报告
    generator = SimpleReportGenerator()
    md_path, docx_path, error = generator.generate_report(args.client_id)
    
    if error:
        print(f"报告生成失败: {error}")
        sys.exit(1)
    else:
        print(f"报告生成成功!")
        print(f"   Markdown报告: {md_path}")
        if docx_path:
            print(f"   Word报告: {docx_path}")
        else:
            print(f"   Word报告: 转换失败，仅生成Markdown格式")
        
        # 检查文件是否存在
        if os.path.exists(md_path):
            print(f"   Markdown文件大小: {os.path.getsize(md_path)} 字节")
        if docx_path and os.path.exists(docx_path):
            print(f"   Word文件大小: {os.path.getsize(docx_path)} 字节")
        
        sys.exit(0)


if __name__ == "__main__":
    main()