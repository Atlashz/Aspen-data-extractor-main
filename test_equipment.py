#!/usr/bin/env python3
import logging
from aspen_data_extractor import AspenDataExtractor

# 设置日志级别
logging.basicConfig(level=logging.INFO, format='%(levelname)s:%(name)s:%(message)s')

print("Creating AspenDataExtractor instance...")
extractor = AspenDataExtractor()

print("Testing extract_all_equipment method...")
try:
    # 不实际连接Aspen，只测试方法调用
    equipment = extractor.extract_all_equipment()
    print(f"Equipment extraction returned: {type(equipment)}")
    print(f"Equipment count: {len(equipment) if equipment else 0}")
    
    if equipment:
        print("First few equipment items:")
        for i, (name, data) in enumerate(equipment.items()):
            if i >= 3:  # 只显示前3个
                break
            print(f"  {name}: {data.get('type', 'Unknown')}")
    
    print(f"Equipment connections storage type: {type(extractor.equipment_connections)}")
    print(f"Equipment connections count: {len(extractor.equipment_connections)}")
    
except Exception as e:
    print(f"Error during equipment extraction: {e}")
    import traceback
    traceback.print_exc()
