#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件组织器 - 自动将生成的报告文件组织到client_data文件夹中

规则:
1. 输出文件不要放在skill文件夹里
2. 要放在user文件夹下client_data里面
3. 如果里面有该客户的文件夹，就放进去
4. 如果没有，则建该客户简称的文件夹，放进去
"""

import os
import shutil
import logging
from pathlib import Path
from datetime import datetime

logger = logging.getLogger(__name__)


class FileOrganizer:
    """文件组织器类"""
    
    def __init__(self, client_name, source_dir=None):
        """
        初始化文件组织器
        
        Args:
            client_name: 客户简称
            source_dir: 源文件目录（如果不指定，则使用默认的报告目录）
        """
        self.client_name = client_name
        
        # 源文件目录（报告生成的位置）
        if source_dir is None:
            # 默认的报告输出目录
            self.source_dir = Path(r"C:\Users\mingh\client-data\_temp")
        else:
            self.source_dir = Path(source_dir)
        
        # 目标目录（client_data文件夹）
        self.client_data_root = Path(r"C:\Users\mingh\client_data")
        self.target_dir = self.client_data_root / client_name
        
        logger.info(f"文件组织器初始化:")
        logger.info(f"  客户: {client_name}")
        logger.info(f"  源目录: {self.source_dir}")
        logger.info(f"  目标目录: {self.target_dir}")
    
    def organize_files(self, report_filename=None):
        """
        组织文件到client_data文件夹
        
        Args:
            report_filename: 特定的报告文件名（如果不指定，则移动所有相关文件）
            
        Returns:
            tuple: (成功与否, 移动的文件列表, 错误信息)
        """
        try:
            # 确保源目录存在
            if not self.source_dir.exists():
                error_msg = f"源目录不存在: {self.source_dir}"
                logger.error(error_msg)
                return False, [], error_msg
            
            # 创建目标目录（如果不存在）
            os.makedirs(self.target_dir, exist_ok=True)
            logger.info(f"目标目录已创建/确认: {self.target_dir}")
            
            # 确定要移动的文件
            files_to_move = []
            if report_filename:
                # 移动特定文件
                source_file = self.source_dir / report_filename
                if source_file.exists():
                    files_to_move.append(source_file)
                else:
                    error_msg = f"源文件不存在: {source_file}"
                    logger.error(error_msg)
                    return False, [], error_msg
            else:
                # 移动所有相关文件（按客户名称和日期过滤）
                today_str = datetime.now().strftime("%Y%m%d")
                
                # 查找所有相关文件
                for ext in ['.md', '.docx', '.log', '.txt', '.json', '.csv']:
                    pattern = f"*{self.client_name}*{ext}"
                    for file in self.source_dir.glob(pattern):
                        if file.is_file():
                            files_to_move.append(file)
                
                # 如果没有找到文件，尝试更宽松的匹配
                if not files_to_move:
                    for file in self.source_dir.glob("*"):
                        if file.is_file() and self.client_name.lower() in file.name.lower():
                            files_to_move.append(file)
            
            if not files_to_move:
                error_msg = f"在源目录中未找到{self.client_name}的相关文件"
                logger.warning(error_msg)
                return False, [], error_msg
            
            logger.info(f"找到 {len(files_to_move)} 个相关文件:")
            for file in files_to_move:
                logger.info(f"  - {file.name}")
            
            # 移动文件
            moved_files = []
            for source_file in files_to_move:
                target_file = self.target_dir / source_file.name
                
                try:
                    # 复制文件（保留元数据）
                    shutil.copy2(source_file, target_file)
                    moved_files.append({
                        'source': str(source_file),
                        'target': str(target_file),
                        'size': source_file.stat().st_size,
                        'name': source_file.name
                    })
                    logger.info(f"  已移动: {source_file.name} -> {target_file}")
                    
                except Exception as e:
                    logger.error(f"  移动文件失败 {source_file.name}: {e}")
                    # 继续处理其他文件
            
            if not moved_files:
                error_msg = "所有文件移动都失败了"
                logger.error(error_msg)
                return False, [], error_msg
            
            # 记录目标目录内容
            self._log_target_directory_contents()
            
            success_msg = f"成功移动 {len(moved_files)} 个文件到 {self.target_dir}"
            logger.info(success_msg)
            
            return True, moved_files, success_msg
            
        except Exception as e:
            error_msg = f"文件组织过程中发生异常: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return False, [], error_msg
    
    def _log_target_directory_contents(self):
        """记录目标目录内容"""
        try:
            if self.target_dir.exists():
                all_items = list(self.target_dir.glob("*"))
                if all_items:
                    logger.info(f"目标目录内容 ({self.target_dir}):")
                    for item in sorted(all_items):
                        if item.is_file():
                            size = item.stat().st_size
                            logger.info(f"  📄 {item.name} ({size:,} 字节)")
                        else:
                            logger.info(f"  📁 {item.name}/")
                else:
                    logger.info(f"目标目录为空: {self.target_dir}")
            else:
                logger.warning(f"目标目录不存在: {self.target_dir}")
        except Exception as e:
            logger.error(f"记录目标目录内容失败: {e}")
    
    def get_target_path(self, filename=None):
        """
        获取目标文件路径
        
        Args:
            filename: 文件名（如果不指定，则返回目标目录路径）
            
        Returns:
            str: 目标文件路径或目录路径
        """
        if filename:
            return str(self.target_dir / filename)
        else:
            return str(self.target_dir)
    
    def cleanup_source_files(self, keep_original=False):
        """
        清理源文件（可选）
        
        Args:
            keep_original: 是否保留源文件
            
        Returns:
            tuple: (成功与否, 清理的文件数, 错误信息)
        """
        if keep_original:
            logger.info("保留源文件，不进行清理")
            return True, 0, "保留源文件"
        
        try:
            # 这里可以实现清理逻辑
            # 注意：生产环境中可能需要更谨慎的清理策略
            logger.info("源文件清理功能暂未实现（安全考虑）")
            return True, 0, "源文件保留"
            
        except Exception as e:
            error_msg = f"清理源文件失败: {str(e)}"
            logger.error(error_msg)
            return False, 0, error_msg


def organize_client_report(client_name, source_dir=None, report_filename=None):
    """
    组织客户报告文件的便捷函数
    
    Args:
        client_name: 客户简称
        source_dir: 源文件目录
        report_filename: 特定的报告文件名
        
    Returns:
        tuple: (成功与否, 详细信息)
    """
    organizer = FileOrganizer(client_name, source_dir)
    success, moved_files, message = organizer.organize_files(report_filename)
    
    if success:
        target_path = organizer.get_target_path()
        details = {
            'client': client_name,
            'target_directory': target_path,
            'moved_files': moved_files,
            'message': message
        }
        return True, details
    else:
        return False, {'error': message}


if __name__ == "__main__":
    # 测试文件组织器
    import sys
    
    # 设置日志
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    if len(sys.argv) > 1:
        client = sys.argv[1]
    else:
        client = "CBD"  # 默认测试客户
    
    print(f"测试文件组织器 - 客户: {client}")
    print("=" * 60)
    
    success, details = organize_client_report(client)
    
    if success:
        print(f"✅ 文件组织成功!")
        print(f"   目标目录: {details['target_directory']}")
        print(f"   移动文件数: {len(details['moved_files'])}")
        for file_info in details['moved_files']:
            print(f"     - {file_info['name']} ({file_info['size']:,} 字节)")
    else:
        print(f"❌ 文件组织失败: {details['error']}")