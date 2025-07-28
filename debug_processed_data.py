#!/usr/bin/env python3
"""Debug processed data structure"""

import sys
import traceback

try:
    from aspen_data_extractor import HeatExchangerDataLoader
    
    loader = HeatExchangerDataLoader('BFG-CO2H-HEX.xlsx')
    loader.load_data()
    processed = loader._process_hex_data()
    
    print("🔍 Processed data structure:")
    print(f"Keys: {list(processed.keys())}")
    print(f"hex_count: {processed['hex_count']}")
    print(f"equipment_list length: {len(processed['equipment_list'])}")
    
    if processed['equipment_list']:
        print("\n📋 First equipment item structure:")
        first_item = processed['equipment_list'][0]
        print(f"Keys: {list(first_item.keys())}")
        print(f"Name: {first_item.get('name')}")
        print(f"Hot stream: {first_item.get('hot_stream_name')}")
        print(f"Cold stream: {first_item.get('cold_stream_name')}")
        print(f"Inlet streams: {first_item.get('inlet_streams')}")
        print(f"Outlet streams: {first_item.get('outlet_streams')}")
        print(f"Hot inlet temp: {first_item.get('hot_stream_inlet_temp')}")
        print(f"Hot outlet temp: {first_item.get('hot_stream_outlet_temp')}")
        print(f"Cold inlet temp: {first_item.get('cold_stream_inlet_temp')}")
        print(f"Cold outlet temp: {first_item.get('cold_stream_outlet_temp')}")
    
except Exception as e:
    print(f"❌ Error: {e}")
    traceback.print_exc()
