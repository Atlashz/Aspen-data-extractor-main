#!/usr/bin/env python3
import logging
from aspen_data_extractor import AspenDataExtractor
import os

# 设置日志级别
logging.basicConfig(level=logging.INFO, format='%(levelname)s:%(name)s:%(message)s')

print("Creating AspenDataExtractor instance...")
extractor = AspenDataExtractor()

# 设置文件路径
aspen_file = r"C:\Users\61723\Downloads\Aspen-data-extractor-main\Aspen-data-extractor-main\aspen_files\BFG-CO2H-MEOH V2 (purge burning).apw"
hex_file = r"C:\Users\61723\Downloads\Aspen-data-extractor-main\Aspen-data-extractor-main\BFG-CO2H-HEX.xlsx"

print(f"Aspen file exists: {os.path.exists(aspen_file)}")
print(f"HEX file exists: {os.path.exists(hex_file)}")

print("Running complete data extraction and storage...")
try:
    result = extractor.extract_and_store_all_data(aspen_file, hex_file)
    print("✅ Data extraction and storage completed!")
    print(f"Success: {result.get('success', False)}")
    print(f"Session ID: {result.get('session_id', 'None')}")
    print(f"Data counts: {result.get('data_counts', {})}")
    if result.get('errors'):
        print(f"Errors: {result.get('errors')}")
except Exception as e:
    print(f"❌ Error during extraction: {e}")
    import traceback
    traceback.print_exc()
