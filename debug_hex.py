#!/usr/bin/env python3
import logging
from aspen_data_extractor import AspenDataExtractor

# 设置日志级别
logging.basicConfig(level=logging.DEBUG, format='%(levelname)s:%(name)s:%(message)s')

print("Creating AspenDataExtractor instance...")
extractor = AspenDataExtractor()

hex_file = r"C:\Users\61723\Downloads\Aspen-data-extractor-main\Aspen-data-extractor-main\BFG-CO2H-HEX.xlsx"

print("Loading hex data...")
success = extractor.load_hex_data(hex_file)

if success:
    print("✅ Hex data loaded successfully!")
    
    # 获取hex数据用于检查
    hex_data = extractor.get_hex_data_for_tea()
    
    print(f"Heat exchangers count: {hex_data.get('hex_count', 0)}")
    
    # 检查前几个热交换器的数据
    heat_exchangers = hex_data.get('heat_exchangers', [])
    print("\nFirst 3 heat exchangers data:")
    for i, hex_info in enumerate(heat_exchangers[:3]):
        print(f"\n{i+1}. {hex_info.get('name', 'Unknown')}:")
        print(f"   hot_stream_name: '{hex_info.get('hot_stream_name', 'None')}' (type: {type(hex_info.get('hot_stream_name'))})")
        print(f"   cold_stream_name: '{hex_info.get('cold_stream_name', 'None')}' (type: {type(hex_info.get('cold_stream_name'))})")
        print(f"   duty: {hex_info.get('duty', 0)} kW")
        print(f"   area: {hex_info.get('area', 0)} m²")
else:
    print("❌ Failed to load hex data")
