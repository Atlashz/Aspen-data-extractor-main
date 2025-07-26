#!/usr/bin/env python3
"""
设备模型功能加载器 - 严格按照 Equipment_Model_Functions.xlsx 进行设备匹配
"""

import pandas as pd
import os
from pathlib import Path
from typing import Dict, Optional, Any
import logging

logger = logging.getLogger(__name__)

class EquipmentModelMatcher:
    """
    设备模型匹配器 - 严格按照 Equipment_Model_Functions.xlsx 文件进行设备匹配
    确保所有数据来自于 ASPEN 读取和 HEX 表格
    """
    
    def __init__(self, excel_file_path: str = None):
        """
        初始化设备模型匹配器
        
        Args:
            excel_file_path: Equipment_Model_Functions.xlsx 文件路径
        """
        self.equipment_mapping = {}
        self.model_to_type = {}
        self.model_to_function = {}
        
        # 默认文件路径
        if excel_file_path is None:
            current_dir = Path(__file__).parent
            excel_file_path = current_dir / "Equipment_Model_Functions.xlsx"
        
        self.excel_file_path = excel_file_path
        self._load_equipment_mapping()
    
    def _load_equipment_mapping(self):
        """从 Excel 文件加载设备映射关系"""
        try:
            if not os.path.exists(self.excel_file_path):
                logger.error(f"Equipment model file not found: {self.excel_file_path}")
                return
            
            # 读取 Excel 文件
            df = pd.read_excel(self.excel_file_path, sheet_name='Sheet1')
            logger.info(f"✅ 加载设备模型文件: {self.excel_file_path}")
            logger.info(f"   总设备数: {len(df)}")
            
            # 构建映射字典
            for _, row in df.iterrows():
                model_name = str(row['Model Name']).strip()
                module_type = str(row['Module Type']).strip()
                function = str(row['Function']).strip()
                
                # 设备名称映射
                self.equipment_mapping[model_name] = {
                    'module_type': module_type,
                    'function': function,
                    'equipment_type': self._map_function_to_equipment_type(function)
                }
                
                # 构建类型映射
                self.model_to_type[model_name] = module_type
                self.model_to_function[model_name] = function
            
            logger.info("✅ 设备映射加载完成")
            logger.info(f"   映射设备: {list(self.equipment_mapping.keys())}")
            
        except Exception as e:
            logger.error(f"Failed to load equipment mapping: {e}")
    
    def _map_function_to_equipment_type(self, function: str) -> str:
        """将功能映射到标准设备类型"""
        function_mapping = {
            'Heater': 'Heat Exchanger',
            'Flash Column': 'Separator', 
            'Compressor': 'Compressor',
            'Valve': 'Valve',
            'Distillation Tower': 'Distillation Column',
            'Reactor': 'Reactor',
            'Mixer': 'Mixer',
            'Split Device': 'Splitter',
            'Untility': 'Utility'
        }
        
        return function_mapping.get(function, f"Unknown ({function})")
    
    def get_equipment_info(self, model_name: str) -> Optional[Dict[str, Any]]:
        """
        根据模型名称获取设备信息
        
        Args:
            model_name: Aspen 中的设备模型名称
            
        Returns:
            设备信息字典或 None
        """
        model_name = str(model_name).strip()
        return self.equipment_mapping.get(model_name)
    
    def get_equipment_type(self, model_name: str) -> str:
        """
        获取设备类型
        
        Args:
            model_name: Aspen 中的设备模型名称
            
        Returns:
            设备类型字符串
        """
        equipment_info = self.get_equipment_info(model_name)
        if equipment_info:
            return equipment_info['equipment_type']
        else:
            return f"Unknown ({model_name})"
    
    def get_module_type(self, model_name: str) -> str:
        """
        获取 Aspen 模块类型
        
        Args:
            model_name: Aspen 中的设备模型名称
            
        Returns:
            Aspen 模块类型
        """
        return self.model_to_type.get(model_name, "Unknown")
    
    def get_function(self, model_name: str) -> str:
        """
        获取设备功能
        
        Args:
            model_name: Aspen 中的设备模型名称
            
        Returns:
            设备功能
        """
        return self.model_to_function.get(model_name, "Unknown")
    
    def is_known_equipment(self, model_name: str) -> bool:
        """
        检查是否为已知设备
        
        Args:
            model_name: Aspen 中的设备模型名称
            
        Returns:
            是否为已知设备
        """
        return str(model_name).strip() in self.equipment_mapping
    
    def get_all_equipment_models(self) -> Dict[str, Dict[str, Any]]:
        """获取所有设备模型信息"""
        return self.equipment_mapping.copy()
    
    def get_equipment_count_by_type(self) -> Dict[str, int]:
        """按设备类型统计数量"""
        type_counts = {}
        for equipment_info in self.equipment_mapping.values():
            eq_type = equipment_info['equipment_type']
            type_counts[eq_type] = type_counts.get(eq_type, 0) + 1
        return type_counts
    
    def print_equipment_summary(self):
        """打印设备映射摘要"""
        logger.info("\n" + "="*60)
        logger.info("EQUIPMENT MODEL MAPPING SUMMARY")
        logger.info("="*60)
        
        logger.info(f"Total Equipment Models: {len(self.equipment_mapping)}")
        
        # 按类型统计
        type_counts = self.get_equipment_count_by_type()
        logger.info("\nEquipment by Type:")
        for eq_type, count in sorted(type_counts.items()):
            logger.info(f"  {eq_type}: {count}")
        
        logger.info("\nDetailed Equipment Mapping:")
        for model_name, info in self.equipment_mapping.items():
            logger.info(f"  {model_name}: {info['module_type']} → {info['equipment_type']} ({info['function']})")
        
        logger.info("="*60)


# 测试代码
if __name__ == "__main__":
    import logging
    logging.basicConfig(level=logging.INFO)
    
    # 创建设备匹配器
    matcher = EquipmentModelMatcher()
    
    # 打印摘要
    matcher.print_equipment_summary()
    
    # 测试几个设备
    test_models = ['COOL2', 'MC1', 'B1', 'C-301', 'UNKNOWN']
    
    print("\n🧪 设备匹配测试:")
    for model in test_models:
        info = matcher.get_equipment_info(model)
        if info:
            print(f"   {model}: {info['equipment_type']} ({info['function']})")
        else:
            print(f"   {model}: 未知设备")
