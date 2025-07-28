#!/usr/bin/env python3
"""
专门解析aspen_flowsheet.xlsx文件中的设备和流股连接信息

根据之前的分析，文件结构如下：
- 第2行：'Material'
- 第3行：'Stream Name' + 流股名称
- 第4行：'Description'
- 第5行：'From' + 源设备
- 第6行：'To' + 目标设备

Author: 流股连接分析工具
Date: 2025-07-27
"""

import pandas as pd
import logging
from pathlib import Path
from typing import Dict, List, Tuple

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class FlowsheetConnectionAnalyzer:
    """分析flowsheet中的设备和流股连接关系"""
    
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.df = None
        self.stream_connections = {}
        self.equipment_connections = {}
        
    def load_data(self):
        """加载Excel数据"""
        try:
            self.df = pd.read_excel(self.file_path, sheet_name='Aspen Data Tables')
            logger.info(f"✅ 成功加载数据: {self.df.shape}")
            return True
        except Exception as e:
            logger.error(f"❌ 加载数据失败: {str(e)}")
            return False
    
    def parse_stream_connections(self):
        """解析流股连接信息"""
        if self.df is None:
            logger.error("❌ 数据未加载")
            return
        
        logger.info("🔍 解析流股连接信息...")
        
        # 找到关键行
        stream_name_row = None
        from_row = None
        to_row = None
        
        for idx, row in self.df.iterrows():
            first_col_value = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
            
            if first_col_value == "Stream Name":
                stream_name_row = idx
                logger.info(f"  📋 找到流股名称行: {idx}")
            elif first_col_value == "From":
                from_row = idx
                logger.info(f"  📍 找到源设备行: {idx}")
            elif first_col_value == "To":
                to_row = idx
                logger.info(f"  📍 找到目标设备行: {idx}")
        
        if not all([stream_name_row is not None, from_row is not None, to_row is not None]):
            logger.error("❌ 找不到必要的行信息")
            return
        
        # 提取流股连接信息
        stream_names = []
        from_equipment = []
        to_equipment = []
        
        # 从列3开始提取数据（跳过前3列标题列）
        for col_idx in range(3, len(self.df.columns)):
            stream_name = str(self.df.iloc[stream_name_row, col_idx]) if pd.notna(self.df.iloc[stream_name_row, col_idx]) else None
            from_eq = str(self.df.iloc[from_row, col_idx]) if pd.notna(self.df.iloc[from_row, col_idx]) else None
            to_eq = str(self.df.iloc[to_row, col_idx]) if pd.notna(self.df.iloc[to_row, col_idx]) else None
            
            # 过滤掉无效数据
            if stream_name and stream_name != 'nan' and stream_name != '':
                stream_names.append(stream_name)
                from_equipment.append(from_eq if from_eq and from_eq != 'nan' else None)
                to_equipment.append(to_eq if to_eq and to_eq != 'nan' else None)
        
        # 构建连接字典
        for i, stream_name in enumerate(stream_names):
            self.stream_connections[stream_name] = {
                'from': from_equipment[i],
                'to': to_equipment[i]
            }
        
        logger.info(f"✅ 解析了 {len(self.stream_connections)} 个流股连接")
        
        return self.stream_connections
    
    def build_equipment_connections(self):
        """构建设备连接关系"""
        if not self.stream_connections:
            logger.error("❌ 流股连接信息未解析")
            return
        
        logger.info("🔧 构建设备连接关系...")
        
        # 初始化设备连接字典
        all_equipment = set()
        for stream_info in self.stream_connections.values():
            if stream_info['from']:
                all_equipment.add(stream_info['from'])
            if stream_info['to']:
                all_equipment.add(stream_info['to'])
        
        # 为每个设备创建连接信息
        for equipment in all_equipment:
            self.equipment_connections[equipment] = {
                'inlet_streams': [],
                'outlet_streams': []
            }
        
        # 填充连接信息
        for stream_name, stream_info in self.stream_connections.items():
            from_eq = stream_info['from']
            to_eq = stream_info['to']
            
            # 对于源设备，这是出料流股
            if from_eq and from_eq in self.equipment_connections:
                self.equipment_connections[from_eq]['outlet_streams'].append(stream_name)
            
            # 对于目标设备，这是进料流股
            if to_eq and to_eq in self.equipment_connections:
                self.equipment_connections[to_eq]['inlet_streams'].append(stream_name)
        
        logger.info(f"✅ 构建了 {len(self.equipment_connections)} 个设备的连接关系")
        
        return self.equipment_connections
    
    def print_analysis_results(self):
        """打印分析结果"""
        print("\n" + "="*80)
        print("🌊 流股连接分析结果")
        print("="*80)
        
        if self.stream_connections:
            print(f"\n📊 发现 {len(self.stream_connections)} 个流股:")
            for stream_name, connection in self.stream_connections.items():
                from_info = f"从 {connection['from']}" if connection['from'] else "未知源"
                to_info = f"到 {connection['to']}" if connection['to'] else "未知目标"
                print(f"  🌊 {stream_name:12s}: {from_info:15s} → {to_info}")
        
        print("\n" + "="*80)
        print("🏭 设备连接分析结果")
        print("="*80)
        
        if self.equipment_connections:
            print(f"\n📊 发现 {len(self.equipment_connections)} 个设备:")
            for equipment, connections in self.equipment_connections.items():
                inlet_count = len(connections['inlet_streams'])
                outlet_count = len(connections['outlet_streams'])
                print(f"\n🏭 {equipment}:")
                print(f"  📥 进料流股 ({inlet_count}): {', '.join(connections['inlet_streams']) if connections['inlet_streams'] else '无'}")
                print(f"  📤 出料流股 ({outlet_count}): {', '.join(connections['outlet_streams']) if connections['outlet_streams'] else '无'}")
    
    def get_equipment_stream_summary(self):
        """获取设备流股连接摘要"""
        summary = {}
        
        if self.equipment_connections:
            for equipment, connections in self.equipment_connections.items():
                summary[equipment] = {
                    'inlet_count': len(connections['inlet_streams']),
                    'outlet_count': len(connections['outlet_streams']),
                    'inlet_streams': connections['inlet_streams'],
                    'outlet_streams': connections['outlet_streams']
                }
        
        return summary
    
    def export_connections_to_json(self, output_file: str = "flowsheet_connections.json"):
        """导出连接信息到JSON文件"""
        import json
        
        export_data = {
            'timestamp': pd.Timestamp.now().isoformat(),
            'source_file': self.file_path,
            'stream_connections': self.stream_connections,
            'equipment_connections': self.equipment_connections,
            'summary': {
                'total_streams': len(self.stream_connections),
                'total_equipment': len(self.equipment_connections),
                'equipment_summary': self.get_equipment_stream_summary()
            }
        }
        
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, indent=2, ensure_ascii=False)
            logger.info(f"✅ 连接信息已导出到: {output_file}")
        except Exception as e:
            logger.error(f"❌ 导出失败: {str(e)}")

def main():
    """主函数"""
    file_path = "aspen_flowsheet.xlsx"
    
    logger.info("=" * 80)
    logger.info("🚀 开始分析 Aspen Flowsheet 连接信息")
    logger.info("=" * 80)
    
    # 创建分析器
    analyzer = FlowsheetConnectionAnalyzer(file_path)
    
    # 加载数据
    if not analyzer.load_data():
        return
    
    # 解析流股连接
    stream_connections = analyzer.parse_stream_connections()
    if not stream_connections:
        logger.error("❌ 流股连接解析失败")
        return
    
    # 构建设备连接
    equipment_connections = analyzer.build_equipment_connections()
    if not equipment_connections:
        logger.error("❌ 设备连接构建失败")
        return
    
    # 打印结果
    analyzer.print_analysis_results()
    
    # 导出结果
    analyzer.export_connections_to_json()
    
    logger.info("\n" + "=" * 80)
    logger.info("✅ 流股连接分析完成!")
    logger.info("=" * 80)

if __name__ == "__main__":
    main()
