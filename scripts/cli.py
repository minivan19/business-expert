#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
商务专家Skill - 命令行接口
"""

import argparse
import sys
import os
import logging
from datetime import datetime
import io

# 设置标准输出的编码为UTF-8，解决Windows控制台编码问题
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# 添加当前目录到Python路径，以便导入其他模块
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from data_loader import get_data_loader
from report_generator import ReportGenerator

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('business_expert.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)


def parse_arguments():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(
        description='商务专家 - 客户经营分析报告生成工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 生成单个客户报告
  python cli.py --client=步步高
  
  # 生成所有客户报告
  python cli.py --client=all
  
  # 指定输出目录
  python cli.py --client=华为 --output=./reports/
  
  # 详细输出模式
  python cli.py --client=测试客户 --verbose
  
  # 指定数据路径
  python cli.py --client=客户1 --data-path=/path/to/客户档案
        """
    )
    
    parser.add_argument(
        '--client', '-c',
        required=True,
        help='客户编号或"all"（生成所有客户报告）'
    )
    
    parser.add_argument(
        '--output', '-o',
        default=None,
        help='输出目录（默认使用配置中的output_dir）'
    )
    
    parser.add_argument(
        '--data-path', '-d',
        default=None,
        help='客户数据根目录（默认使用配置中的client_data_path）'
    )
    
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='详细输出模式'
    )
    
    parser.add_argument(
        '--skip-llm',
        action='store_true',
        help='跳过LLM分析，只生成脚本处理的部分'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version='商务专家Skill v1.0.0'
    )
    
    return parser.parse_args()


def setup_logging(verbose):
    """设置日志级别"""
    if verbose:
        logging.getLogger().setLevel(logging.DEBUG)
        logger.debug("启用详细日志模式")
    else:
        logging.getLogger().setLevel(logging.INFO)


def generate_single_client_report(client_id, data_path, output_dir, skip_llm):
    """生成单个客户报告"""
    logger.info(f"开始生成客户报告: {client_id}")
    start_time = datetime.now()
    
    try:
        # 初始化数据加载器
        loader = get_data_loader(data_path)
        
        # 加载数据
        data, error = loader.load_client_data(client_id)
        if error:
            logger.error(f"数据加载失败: {error}")
            return False, error
        
        if not data:
            error_msg = "未加载到有效数据"
            logger.error(error_msg)
            return False, error_msg
        
        logger.info(f"成功加载数据，模块数: {len(data)}")
        
        # 初始化报告生成器
        generator = ReportGenerator(output_dir=output_dir, skip_llm=skip_llm)
        
        # 生成报告
        report_path, gen_error = generator.generate_report(client_id, data)
        if gen_error:
            logger.error(f"报告生成失败: {gen_error}")
            return False, gen_error
        
        elapsed_time = (datetime.now() - start_time).total_seconds()
        logger.info(f"报告生成成功: {report_path}")
        logger.info(f"处理耗时: {elapsed_time:.2f}秒")
        
        return True, report_path
        
    except Exception as e:
        error_msg = f"生成报告过程中发生异常: {str(e)}"
        logger.error(error_msg, exc_info=True)
        return False, error_msg


def generate_all_clients_reports(data_path, output_dir, skip_llm):
    """生成所有客户报告"""
    logger.info("开始生成所有客户报告")
    start_time = datetime.now()
    
    try:
        # 初始化数据加载器
        loader = get_data_loader(data_path)
        
        # 获取所有客户列表
        clients = loader.list_all_clients()
        if not clients:
            error_msg = "未找到任何有效客户"
            logger.error(error_msg)
            return False, error_msg
        
        logger.info(f"找到 {len(clients)} 个客户: {', '.join(clients)}")
        
        success_count = 0
        failed_clients = []
        
        # 为每个客户生成报告
        for i, client_id in enumerate(clients, 1):
            logger.info(f"处理客户 {i}/{len(clients)}: {client_id}")
            
            success, result = generate_single_client_report(
                client_id, data_path, output_dir, skip_llm
            )
            
            if success:
                success_count += 1
                logger.info(f"客户 {client_id} 处理成功: {result}")
            else:
                failed_clients.append((client_id, result))
                logger.error(f"客户 {client_id} 处理失败: {result}")
        
        # 生成汇总报告
        if success_count > 0:
            summary_path = os.path.join(output_dir, f"all_clients_summary_{datetime.now().strftime('%Y%m%d')}.md")
            with open(summary_path, 'w', encoding='utf-8') as f:
                f.write(f"# 所有客户报告生成汇总\n\n")
                f.write(f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"**总客户数**: {len(clients)}\n")
                f.write(f"**成功数**: {success_count}\n")
                f.write(f"**失败数**: {len(failed_clients)}\n\n")
                
                if failed_clients:
                    f.write("## 失败客户列表\n")
                    for client_id, error in failed_clients:
                        f.write(f"- **{client_id}**: {error}\n")
            
            logger.info(f"汇总报告已生成: {summary_path}")
        
        elapsed_time = (datetime.now() - start_time).total_seconds()
        logger.info(f"所有客户报告生成完成")
        logger.info(f"成功: {success_count}/{len(clients)}")
        logger.info(f"总耗时: {elapsed_time:.2f}秒")
        
        if failed_clients:
            return False, f"{success_count}个成功，{len(failed_clients)}个失败"
        else:
            return True, f"所有{success_count}个客户处理成功"
        
    except Exception as e:
        error_msg = f"批量生成报告过程中发生异常: {str(e)}"
        logger.error(error_msg, exc_info=True)
        return False, error_msg


def main():
    """主函数"""
    # 解析命令行参数
    args = parse_arguments()
    
    # 设置日志级别
    setup_logging(args.verbose)
    
    logger.info("=" * 60)
    logger.info("商务专家Skill - 客户经营分析报告生成")
    logger.info(f"开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"客户: {args.client}")
    logger.info(f"输出目录: {args.output or '默认'}")
    logger.info(f"数据路径: {args.data_path or '默认'}")
    logger.info(f"跳过LLM: {args.skip_llm}")
    logger.info("=" * 60)
    
    try:
        # 确定输出目录
        if args.output:
            output_dir = args.output
            # 创建输出目录（如果不存在）
            os.makedirs(output_dir, exist_ok=True)
        else:
            # 使用默认输出目录
            output_dir = r"C:\Users\mingh\client-data\_temp"
            os.makedirs(output_dir, exist_ok=True)
        
        # 确定数据路径
        data_path = args.data_path
        
        # 生成报告
        if args.client.lower() == 'all':
            success, result = generate_all_clients_reports(
                data_path, output_dir, args.skip_llm
            )
        else:
            success, result = generate_single_client_report(
                args.client, data_path, output_dir, args.skip_llm
            )
        
        # 输出结果
        if success:
            logger.info("=" * 60)
            logger.info("[SUCCESS] 报告生成成功!")
            logger.info(f"结果: {result}")
            logger.info("=" * 60)
            print(f"\n[SUCCESS] 报告生成成功: {result}")
            return 0
        else:
            logger.error("=" * 60)
            logger.error("[FAILED] 报告生成失败!")
            logger.error(f"错误: {result}")
            logger.error("=" * 60)
            print(f"\n[FAILED] 报告生成失败: {result}")
            return 1
            
    except KeyboardInterrupt:
        logger.warning("用户中断操作")
        print("\n⚠️  操作被用户中断")
        return 130
    except Exception as e:
        logger.error(f"程序执行失败: {str(e)}", exc_info=True)
        print(f"\n❌ 程序执行失败: {str(e)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())