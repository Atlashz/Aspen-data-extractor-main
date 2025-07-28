#!/usr/bin/env python3
import logging
from aspen_data_extractor import AspenDataExtractor

# 设置日志级别
logging.basicConfig(level=logging.INFO, format='%(levelname)s:%(name)s:%(message)s')

print("Creating AspenDataExtractor instance...")
extractor = AspenDataExtractor()

# 设置文件路径
aspen_file = r"C:\Users\61723\Downloads\Aspen-data-extractor-main\Aspen-data-extractor-main\aspen_files\BFG-CO2H-MEOH V2 (purge burning).apw"
hex_file = r"C:\Users\61723\Downloads\Aspen-data-extractor-main\Aspen-data-extractor-main\BFG-CO2H-HEX.xlsx"

print("Running complete data extraction and storage...")
result = extractor.extract_and_store_all_data(aspen_file, hex_file)

if result['success']:
    print("\n✅ Data extraction completed successfully!")
    
    print("\n🔗 Testing stream connection functionality:")
    
    # 测试获取所有设备的流股连接
    all_connections = extractor.get_equipment_stream_connections()
    print(f"Total equipment with connections: {len(all_connections)}")
    
    # 显示每个设备的流股连接信息
    print("\nEquipment Stream Connections:")
    print("=" * 60)
    for eq_name, connections in all_connections.items():
        inlet_count = len(connections.get('inlet_streams', []))
        outlet_count = len(connections.get('outlet_streams', []))
        inlet_streams = ', '.join(connections.get('inlet_streams', []))
        outlet_streams = ', '.join(connections.get('outlet_streams', []))
        
        print(f"{eq_name}:")
        print(f"  进料流股 ({inlet_count}): {inlet_streams if inlet_streams else '无'}")
        print(f"  出料流股 ({outlet_count}): {outlet_streams if outlet_streams else '无'}")
        print()
    
    # 测试获取特定设备的连接信息
    print("Testing specific equipment lookup:")
    test_equipment = ['B1', 'COOL2', 'MIX3']
    for eq_name in test_equipment:
        connections = extractor.get_equipment_stream_connections(eq_name)
        if connections:
            print(f"{eq_name}: {len(connections.get('inlet_streams', []))} in, {len(connections.get('outlet_streams', []))} out")
        else:
            print(f"{eq_name}: No connection data found")
    
else:
    print(f"❌ Data extraction failed: {result.get('errors', [])}")
