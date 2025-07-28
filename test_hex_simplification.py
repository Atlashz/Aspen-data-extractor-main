#!/usr/bin/env python3
"""Test script for simplified heat exchanger processing"""

import sys
import traceback

try:
    # Test basic import
    from aspen_data_extractor import HeatExchangerDataLoader
    print("✅ Import successful")
    
    # Test with Excel file
    loader = HeatExchangerDataLoader('BFG-CO2H-HEX.xlsx')
    print("✅ Loader created successfully")
    
    # Load the data
    df = loader.load_data()
    if df is not None:
        print(f"✅ Data loaded successfully: {df.shape}")
    else:
        print("❌ No data loaded")
        
    # Test heat exchanger processing
    processed = loader._process_hex_data()
    print(f"✅ Heat exchanger processing successful")
    print(f"   Found {processed['hex_count']} heat exchangers")
    print(f"   Total duty: {processed['total_heat_duty']:.1f} kW")
    
    # Show first few results
    if processed['equipment_list']:
        print("\n📋 Sample results:")
        for i, hex_data in enumerate(processed['equipment_list'][:3]):
            print(f"   {i+1}. {hex_data['name']}")
            print(f"      Hot stream: {hex_data.get('hot_stream_name', 'N/A')}")
            print(f"      Cold stream: {hex_data.get('cold_stream_name', 'N/A')}")
            print(f"      Inlet streams: {hex_data.get('inlet_streams', [])}")
            print(f"      Outlet streams: {hex_data.get('outlet_streams', [])}")
            print(f"      Duty: {hex_data.get('duty', 0):.1f} kW")
            print()
    
    print("🎉 All tests passed!")
    
except Exception as e:
    print(f"❌ Error: {e}")
    print("\n📍 Traceback:")
    traceback.print_exc()
    sys.exit(1)
