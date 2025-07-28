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
try:
    result = extractor.extract_and_store_all_data(aspen_file, hex_file)
    
    if result['success']:
        print("✅ Data extraction completed successfully!")
        print(f"Data counts: {result.get('data_counts', {})}")
        
        # 检查数据库中的数据
        import sqlite3
        conn = sqlite3.connect('aspen_data.db')
        cursor = conn.cursor()
        
        # 检查heat exchanger数据
        cursor.execute('SELECT name, hot_stream_name, cold_stream_name FROM heat_exchangers LIMIT 3')
        hex_records = cursor.fetchall()
        print("\nHeat exchanger records:")
        for record in hex_records:
            print(f'  {record[0]}: hot={record[1]}, cold={record[2]}')
        
        # 检查equipment数据
        cursor.execute('SELECT name, equipment_type, inlet_streams, outlet_streams FROM aspen_equipment LIMIT 3')
        eq_records = cursor.fetchall()
        print("\nEquipment records:")
        for record in eq_records:
            print(f'  {record[0]} ({record[1]}): inlet={record[2]}, outlet={record[3]}')
        
        conn.close()
    else:
        print(f"❌ Data extraction failed: {result.get('errors', [])}")
        
except Exception as e:
    print(f"❌ Error during extraction: {e}")
    import traceback
    traceback.print_exc()
