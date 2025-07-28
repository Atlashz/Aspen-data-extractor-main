#!/usr/bin/env python3
"""Debug script for heat exchanger processing"""

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
        print(f"Columns: {list(df.columns)}")
    else:
        print("❌ No data loaded")
        sys.exit(1)
        
    # Try to process line by line to find the error
    print("\n🔍 Debug: Starting _process_hex_data...")
    
    # Access the method and try to debug step by step
    import logging
    logging.basicConfig(level=logging.DEBUG)
    
    try:
        processed = loader._process_hex_data()
        print(f"✅ Processing successful - {processed['hex_count']} heat exchangers found")
    except Exception as e:
        print(f"❌ Error in _process_hex_data: {e}")
        print("Full traceback:")
        traceback.print_exc()
    
except Exception as e:
    print(f"❌ Error: {e}")
    print("\n📍 Traceback:")
    traceback.print_exc()
    sys.exit(1)
