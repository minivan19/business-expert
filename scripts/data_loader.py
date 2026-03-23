#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据加载模块 - 负责加载和预处理所有客户数据
"""

import os
import pandas as pd
from glob import glob
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class DataLoader:
    """数据加载器类"""
    
    def __init__(self, client_data_path=None):
        """
        初始化数据加载器
        
        Args:
            client_data_path: 客户数据根目录，如果为None则使用默认路径
        """
        if client_data_path is None:
            # 默认路径，可以从环境变量或配置文件中读取
            self.client_data_path = r"/Users/limingheng/AI/client-data/raw/客户档案"
        else:
            self.client_data_path = client_data_path
        
        logger.info(f"数据加载器初始化，数据路径: {self.client_data_path}")
    
    def load_client_data(self, client_name):
        """
        加载指定客户的所有数据
        
        Args:
            client_name: 客户名称
            
        Returns:
            tuple: (data_dict, error_message)
                   data_dict: 包含各部分数据的字典
                   error_message: 错误信息，如果成功则为None
        """
        client_dir = os.path.join(self.client_data_path, client_name)
        
        # 检查客户目录是否存在
        if not os.path.exists(client_dir):
            error_msg = f"客户目录不存在: {client_dir}"
            logger.error(error_msg)
            return None, error_msg
        
        logger.info(f"开始加载客户数据: {client_name}")
        
        data = {}
        errors = []
        
        try:
            # Part 1: 客户主数据
            basic_data, basic_error = self._load_basic_data(client_dir)
            if basic_data is not None:
                data["part1"] = basic_data
            if basic_error:
                errors.append(basic_error)
            
            # Part 2: 订阅数据
            sub_data, sub_error = self._load_subscription_data(client_dir)
            if sub_data is not None:
                data["part2"] = sub_data
            if sub_error:
                errors.append(sub_error)
            
            # Part 2: 收款数据
            coll_data, coll_error = self._load_collection_data(client_dir)
            if coll_data is not None:
                data["part4"] = coll_data
            if coll_error:
                errors.append(coll_error)
            
            # Part 3: 实施数据
            impl_data, impl_error = self._load_implementation_data(client_dir)
            if impl_data:
                data.update(impl_data)
            if impl_error:
                errors.append(impl_error)
            
            # Part 4: 运维工单
            ops_data, ops_error = self._load_operations_data(client_dir)
            if ops_data is not None:
                data["part5_ops"] = ops_data
            if ops_error:
                errors.append(ops_error)
            
            # 数据预处理
            data = self._preprocess_data(data)
            
            logger.info(f"客户数据加载完成: {client_name}, 加载模块数: {len(data)}")
            
            # 如果有错误但仍有数据，返回数据和警告
            if errors and data:
                error_msg = "; ".join(errors)
                return data, f"部分数据加载失败: {error_msg}"
            elif not data:
                return None, "未加载到任何有效数据"
            else:
                return data, None
                
        except Exception as e:
            error_msg = f"数据加载过程中发生异常: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return None, error_msg
    
    def _load_basic_data(self, client_dir):
        """加载客户主数据"""
        basic_path = os.path.join(client_dir, "基础数据", "客户主数据.xlsx")
        
        if not os.path.exists(basic_path):
            logger.warning(f"客户主数据文件不存在: {basic_path}")
            return None, "客户主数据文件不存在"
        
        try:
            df = pd.read_excel(basic_path)
            logger.info(f"成功加载客户主数据，行数: {len(df)}，列数: {len(df.columns)}")
            
            # 清理列名
            df.columns = df.columns.str.strip()
            
            # 处理行业字段（第18列，索引17）
            if len(df) > 0 and len(df.columns) > 18:
                df.loc[:, "行业_clean"] = df.iloc[:, 18]
                logger.info(f"行业字段处理完成: {df.iloc[0, 18] if len(df) > 0 else '无数据'}")
            
            return df, None
            
        except Exception as e:
            error_msg = f"加载客户主数据失败: {str(e)}"
            logger.error(error_msg)
            return None, error_msg
    
    def _load_subscription_data(self, client_dir):
        """加载订阅数据"""
        sub_path = os.path.join(client_dir, "订阅合同行", "订阅明细.xlsx")
        
        if not os.path.exists(sub_path):
            logger.warning(f"订阅数据文件不存在: {sub_path}")
            return None, "订阅数据文件不存在"
        
        try:
            df = pd.read_excel(sub_path)
            logger.info(f"成功加载订阅数据，行数: {len(df)}")
            return df, None
            
        except Exception as e:
            error_msg = f"加载订阅数据失败: {str(e)}"
            logger.error(error_msg)
            return None, error_msg
    
    def _load_collection_data(self, client_dir):
        """加载收款数据"""
        coll_path = os.path.join(client_dir, "订阅合同收款情况", "订阅合同收款情况.xlsx")
        
        if not os.path.exists(coll_path):
            logger.warning(f"收款数据文件不存在: {coll_path}")
            return None, "收款数据文件不存在"
        
        try:
            df = pd.read_excel(coll_path)
            logger.info(f"成功加载收款数据，行数: {len(df)}")
            return df, None
            
        except Exception as e:
            error_msg = f"加载收款数据失败: {str(e)}"
            logger.error(error_msg)
            return None, error_msg
    
    def _load_implementation_data(self, client_dir):
        """加载实施数据"""
        impl_dir = os.path.join(client_dir, "实施合同行")
        
        if not os.path.exists(impl_dir):
            logger.warning(f"实施数据目录不存在: {impl_dir}")
            return {}, "实施数据目录不存在"
        
        data = {}
        errors = []
        
        # 固定金额台账
        fixed_path = os.path.join(impl_dir, "固定金额台账.xlsx")
        if os.path.exists(fixed_path):
            try:
                df_fixed = pd.read_excel(fixed_path)
                data["part3_fixed"] = df_fixed
                logger.info(f"成功加载固定金额台账，行数: {len(df_fixed)}")
            except Exception as e:
                error_msg = f"加载固定金额台账失败: {str(e)}"
                logger.error(error_msg)
                errors.append(error_msg)
        else:
            logger.warning(f"固定金额台账文件不存在: {fixed_path}")
            errors.append("固定金额台账文件不存在")
        
        # 人天框架台账
        dayspan_path = os.path.join(impl_dir, "人天框架台账.xlsx")
        if os.path.exists(dayspan_path):
            try:
                df_dayspan = pd.read_excel(dayspan_path)
                data["part3_dayspan"] = df_dayspan
                logger.info(f"成功加载人天框架台账，行数: {len(df_dayspan)}")
            except Exception as e:
                error_msg = f"加载人天框架台账失败: {str(e)}"
                logger.error(error_msg)
                errors.append(error_msg)
        else:
            logger.warning(f"人天框架台账文件不存在: {dayspan_path}")
            errors.append("人天框架台账文件不存在")
        
        error_msg = "; ".join(errors) if errors else None
        return data, error_msg
    
    def _load_operations_data(self, client_dir):
        """加载运维工单数据"""
        ops_dir = os.path.join(client_dir, "运维工单")
        
        if not os.path.exists(ops_dir):
            logger.warning(f"运维工单目录不存在: {ops_dir}")
            return None, "运维工单目录不存在"
        
        ops_files = glob(os.path.join(ops_dir, "*.xlsx"))
        if not ops_files:
            logger.warning(f"运维工单目录中没有Excel文件: {ops_dir}")
            return None, "运维工单目录中没有Excel文件"
        
        try:
            dfs = []
            for file_path in ops_files:
                df = pd.read_excel(file_path)
                dfs.append(df)
                logger.info(f"加载运维工单文件: {os.path.basename(file_path)}，行数: {len(df)}")
            
            if dfs:
                combined_df = pd.concat(dfs, ignore_index=True)
                logger.info(f"合并运维工单数据，总行数: {len(combined_df)}")
                return combined_df, None
            else:
                return None, "未加载到有效的运维工单数据"
                
        except Exception as e:
            error_msg = f"加载运维工单数据失败: {str(e)}"
            logger.error(error_msg)
            return None, error_msg
    
    def _preprocess_data(self, data):
        """数据预处理"""
        processed_data = {}
        
        for key, df in data.items():
            if df is not None and not df.empty:
                try:
                    # 处理空值
                    df = df.fillna('')
                    
                    # 统一日期格式
                    date_columns = [col for col in df.columns if '时间' in col or '日期' in col or 'date' in col.lower()]
                    for col in date_columns:
                        if col in df.columns:
                            df[col] = pd.to_datetime(df[col], errors='coerce')
                    
                    # 标准化金额（转为数值，单位：元）
                    amount_columns = [col for col in df.columns if '金额' in col or '费用' in col or '价格' in col or 'arr' in col.lower()]
                    for col in amount_columns:
                        if col in df.columns:
                            # 尝试转换为数值
                            df[col] = pd.to_numeric(df[col], errors='coerce')
                            # 处理可能的万元单位
                            if df[col].max() < 100 and df[col].min() >= 0:
                                # 可能是万元单位，转换为元
                                df[col] = df[col] * 10000
                    
                    processed_data[key] = df
                    logger.debug(f"数据预处理完成: {key}，行数: {len(df)}")
                    
                except Exception as e:
                    logger.warning(f"数据预处理失败 {key}: {str(e)}")
                    processed_data[key] = df  # 保留原始数据
        
        return processed_data
    
    def list_all_clients(self):
        """列出所有可用的客户"""
        if not os.path.exists(self.client_data_path):
            logger.error(f"客户数据根目录不存在: {self.client_data_path}")
            return []
        
        try:
            clients = []
            for item in os.listdir(self.client_data_path):
                item_path = os.path.join(self.client_data_path, item)
                if os.path.isdir(item_path):
                    # 检查是否有基础数据文件
                    basic_path = os.path.join(item_path, "基础数据", "客户主数据.xlsx")
                    if os.path.exists(basic_path):
                        clients.append(item)
            
            logger.info(f"找到 {len(clients)} 个有效客户")
            return sorted(clients)
            
        except Exception as e:
            logger.error(f"列出客户失败: {str(e)}")
            return []


# 单例实例
_loader_instance = None

def get_data_loader(client_data_path=None):
    """获取数据加载器实例（单例模式）"""
    global _loader_instance
    if _loader_instance is None:
        _loader_instance = DataLoader(client_data_path)
    return _loader_instance


if __name__ == "__main__":
    # 测试代码
    loader = DataLoader()
    
    # 测试列出所有客户
    clients = loader.list_all_clients()
    print(f"找到 {len(clients)} 个客户: {clients}")
    
    if clients:
        # 测试加载第一个客户的数据
        test_client = clients[0]
        print(f"\n测试加载客户: {test_client}")
        
        data, error = loader.load_client_data(test_client)
        if data:
            print(f"成功加载数据，模块数: {len(data)}")
            for key, df in data.items():
                print(f"  {key}: {len(df)} 行, {len(df.columns)} 列")
        else:
            print(f"加载失败: {error}")