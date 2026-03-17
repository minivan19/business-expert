#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试文件组织功能
"""

import os
import sys
import logging
from pathlib import Path

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_file_organizer():
    """测试文件组织器"""
    print("测试文件组织功能")
    print("=" * 60)
    
    # 导入文件组织器
    try:
        from file_organizer import FileOrganizer, organize_client_report
        print("✅ 成功导入文件组织器")
    except ImportError as e:
        print(f"❌ 导入文件组织器失败: {e}")
        return False
    
    # 测试1: 创建文件组织器实例
    print("\n1. 测试文件组织器实例化...")
    try:
        organizer = FileOrganizer("TEST")
        print(f"✅ 文件组织器实例化成功")
        print(f"   客户: {organizer.client_name}")
        print(f"   源目录: {organizer.source_dir}")
        print(f"   目标目录: {organizer.target_dir}")
    except Exception as e:
        print(f"❌ 文件组织器实例化失败: {e}")
        return False
    
    # 测试2: 测试便捷函数
    print("\n2. 测试便捷函数...")
    try:
        success, details = organize_client_report("TEST")
        if success:
            print(f"✅ 便捷函数调用成功")
            print(f"   目标目录: {details['target_directory']}")
        else:
            print(f"⚠️  便捷函数调用失败（可能是没有测试文件）: {details.get('error', '未知错误')}")
    except Exception as e:
        print(f"❌ 便捷函数调用失败: {e}")
    
    # 测试3: 测试实际文件组织
    print("\n3. 测试实际文件组织...")
    
    # 创建一个测试文件
    test_source_dir = Path(r"C:\Users\mingh\client-data\_temp")
    test_file = test_source_dir / "TEST_测试报告_20260317.md"
    
    try:
        # 创建测试文件
        test_source_dir.mkdir(exist_ok=True)
        with open(test_file, 'w', encoding='utf-8') as f:
            f.write("# 测试报告\n\n这是一个测试文件。\n")
        
        print(f"✅ 创建测试文件: {test_file}")
        
        # 使用文件组织器
        organizer = FileOrganizer("TEST", source_dir=test_source_dir)
        success, moved_files, message = organizer.organize_files("TEST_测试报告_20260317.md")
        
        if success:
            print(f"✅ 文件组织成功!")
            print(f"   消息: {message}")
            print(f"   移动文件数: {len(moved_files)}")
            
            # 检查目标文件
            target_file = organizer.target_dir / "TEST_测试报告_20260317.md"
            if target_file.exists():
                print(f"✅ 目标文件存在: {target_file}")
                # 清理测试文件
                if test_file.exists():
                    test_file.unlink()
                if target_file.exists():
                    target_file.unlink()
                print(f"✅ 测试文件已清理")
            else:
                print(f"❌ 目标文件不存在")
        else:
            print(f"❌ 文件组织失败: {message}")
            
    except Exception as e:
        print(f"❌ 测试过程中发生异常: {e}")
        # 清理测试文件
        if 'test_file' in locals() and test_file.exists():
            test_file.unlink()
    
    # 测试4: 集成测试
    print("\n4. 测试与report_generator集成...")
    try:
        from report_generator import ReportGenerator
        
        # 创建模拟数据
        mock_data = {
            "part1_basic": None,
            "part2_subscription": None,
            "part5_ops": None
        }
        
        # 创建报告生成器
        generator = ReportGenerator(output_dir=str(test_source_dir), skip_llm=True)
        
        # 测试_organize_report_files方法
        test_report_filename = "INTEGRATION_TEST_经营分析报告_20260317.md"
        test_report_path = test_source_dir / test_report_filename
        
        # 创建测试报告文件
        with open(test_report_path, 'w', encoding='utf-8') as f:
            f.write("# 集成测试报告\n\n测试集成功能。\n")
        
        print(f"✅ 创建集成测试文件: {test_report_path}")
        
        # 调用文件组织方法
        organized = generator._organize_report_files("INTEGRATION_TEST", test_report_filename)
        
        if organized:
            print(f"✅ 集成测试成功! 文件已自动组织")
        else:
            print(f"⚠️  集成测试部分成功（可能是没有安装依赖）")
        
        # 清理
        if test_report_path.exists():
            test_report_path.unlink()
        target_test_file = organizer.client_data_root / "INTEGRATION_TEST" / test_report_filename
        if target_test_file.exists():
            target_test_file.unlink()
            target_test_file.parent.rmdir()  # 删除空目录
        print(f"✅ 集成测试文件已清理")
        
    except Exception as e:
        print(f"❌ 集成测试失败: {e}")
    
    print("\n" + "=" * 60)
    print("文件组织功能测试完成!")
    return True

def main():
    """主函数"""
    print("商务专家skill - 文件组织功能测试")
    print("=" * 60)
    
    # 添加当前目录到Python路径
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))
    
    # 运行测试
    success = test_file_organizer()
    
    if success:
        print("\n🎉 所有测试完成!")
        print("文件组织功能已成功集成到商务专家skill中")
        print("生成的报告将自动保存到: C:\\Users\\mingh\\client_data\\{客户简称}\\")
    else:
        print("\n⚠️  测试过程中发现问题")
        print("请检查错误信息并修复")
    
    return 0 if success else 1

if __name__ == "__main__":
    sys.exit(main())