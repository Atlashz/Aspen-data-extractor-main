#!/usr/bin/env python3
"""
读取aspen_flowsheet.xlsx文件，分析设备和流股的链接信息

Author: 数据分析工具
Date: 2025-07-27
"""

import pandas as pd
import logging
from pathlib import Path

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def read_flowsheet_excel(file_path: str):
    """
    读取flowsheet Excel文件并分析结构
    
    Args:
        file_path: Excel文件路径
    """
    try:
        # 检查文件是否存在
        if not Path(file_path).exists():
            logger.error(f"文件不存在: {file_path}")
            return
        
        logger.info(f"📖 正在读取文件: {file_path}")
        
        # 读取Excel文件的所有工作表
        excel_file = pd.ExcelFile(file_path)
        logger.info(f"发现 {len(excel_file.sheet_names)} 个工作表: {excel_file.sheet_names}")
        
        # 分析每个工作表
        all_data = {}
        for sheet_name in excel_file.sheet_names:
            logger.info(f"\n🔍 分析工作表: {sheet_name}")
            
            try:
                # 读取工作表
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                all_data[sheet_name] = df
                
                logger.info(f"  📊 数据维度: {df.shape} (行x列)")
                
                if not df.empty:
                    logger.info(f"  📋 列名: {list(df.columns)}")
                    
                    # 显示前几行数据
                    logger.info("  📝 前5行数据:")
                    print(df.head().to_string())
                    
                    # 分析数据类型
                    logger.info("  🔢 数据类型:")
                    for col in df.columns:
                        non_null_count = df[col].count()
                        total_count = len(df)
                        logger.info(f"    {col}: {df[col].dtype} ({non_null_count}/{total_count} 非空)")
                    
                    # 如果有设备和流股相关的列，进行特殊分析
                    analyze_equipment_stream_connections(df, sheet_name)
                    
            except Exception as e:
                logger.error(f"  ❌ 读取工作表 {sheet_name} 时出错: {str(e)}")
        
        return all_data
        
    except Exception as e:
        logger.error(f"❌ 读取Excel文件时出错: {str(e)}")
        return None

def analyze_equipment_stream_connections(df, sheet_name):
    """
    分析设备和流股连接信息
    
    Args:
        df: DataFrame
        sheet_name: 工作表名称
    """
    logger.info(f"\n🔗 分析 {sheet_name} 中的设备流股连接:")
    
    # 查找可能包含设备信息的列
    equipment_cols = [col for col in df.columns if any(keyword in col.lower() 
                     for keyword in ['equipment', 'block', 'unit', '设备', '块'])]
    
    # 查找可能包含流股信息的列
    stream_cols = [col for col in df.columns if any(keyword in col.lower() 
                  for keyword in ['stream', 'flow', 'inlet', 'outlet', '流股', '进料', '出料', 'feed', 'product'])]
    
    # 查找可能包含连接信息的列
    connection_cols = [col for col in df.columns if any(keyword in col.lower() 
                      for keyword in ['from', 'to', 'source', 'destination', '来源', '目标', 'connect'])]
    
    if equipment_cols:
        logger.info(f"  🏭 设备相关列: {equipment_cols}")
        
    if stream_cols:
        logger.info(f"  🌊 流股相关列: {stream_cols}")
        
    if connection_cols:
        logger.info(f"  🔗 连接相关列: {connection_cols}")
    
    # 分析唯一值
    for col in df.columns:
        if df[col].dtype == 'object':  # 文本列
            unique_values = df[col].dropna().unique()
            if len(unique_values) <= 20:  # 只显示不超过20个唯一值
                logger.info(f"  📋 '{col}' 的唯一值: {list(unique_values)}")
            else:
                logger.info(f"  📋 '{col}' 有 {len(unique_values)} 个唯一值")
                logger.info(f"      前10个值: {list(unique_values[:10])}")
    
    # 尝试识别连接模式
    identify_connection_patterns(df, sheet_name)

def identify_connection_patterns(df, sheet_name):
    """
    识别连接模式和关系
    
    Args:
        df: DataFrame
        sheet_name: 工作表名称
    """
    logger.info(f"\n📊 识别 {sheet_name} 中的连接模式:")
    
    # 查找包含箭头或连接符的列
    for col in df.columns:
        if df[col].dtype == 'object':
            sample_values = df[col].dropna().head(10).tolist()
            
            # 检查是否包含常见的连接符号
            connection_indicators = ['→', '->', '-->', '==>', '|', '流向', '到', 'to', 'from']
            
            for value in sample_values:
                if isinstance(value, str):
                    for indicator in connection_indicators:
                        if indicator in value:
                            logger.info(f"  🎯 发现连接模式在列 '{col}': {value}")
                            break
    
    # 如果有多列，尝试分析关系
    if len(df.columns) >= 2:
        logger.info("  🔍 尝试分析列之间的关系...")
        
        # 查找可能的源-目标对
        for i, col1 in enumerate(df.columns):
            for j, col2 in enumerate(df.columns):
                if i != j and df[col1].dtype == 'object' and df[col2].dtype == 'object':
                    # 检查是否有相同的值（可能表示连接关系）
                    common_values = set(df[col1].dropna()) & set(df[col2].dropna())
                    if common_values and len(common_values) > 1:
                        logger.info(f"  🔗 '{col1}' 和 '{col2}' 有共同值: {list(common_values)[:5]}...")

def main():
    """主函数"""
    file_path = "aspen_flowsheet.xlsx"
    
    logger.info("=" * 60)
    logger.info("🚀 开始分析 Aspen Flowsheet 文件")
    logger.info("=" * 60)
    
    # 读取并分析文件
    data = read_flowsheet_excel(file_path)
    
    if data:
        logger.info("\n" + "=" * 60)
        logger.info("📋 文件分析完成!")
        logger.info("=" * 60)
        
        # 提供数据访问总结
        logger.info("\n📈 数据总结:")
        for sheet_name, df in data.items():
            logger.info(f"  工作表 '{sheet_name}': {df.shape[0]} 行, {df.shape[1]} 列")
    else:
        logger.error("❌ 文件分析失败")

if __name__ == "__main__":
    main()
