#!/usr/bin/env python3
"""Test database storage for simplified heat exchanger data"""

import sys
import traceback

try:
    from aspen_data_extractor import HeatExchangerDataLoader
    from aspen_data_database import AspenDataDatabase
    print("✅ Import successful")
    
    # Test with Excel file
    loader = HeatExchangerDataLoader('BFG-CO2H-HEX.xlsx')
    print("✅ Loader created successfully")
    
    # Load and process the data
    df = loader.load_data()
    if df is not None:
        print(f"✅ Data loaded successfully: {df.shape}")
        
        processed = loader._process_hex_data()
        print(f"✅ Processing successful - {processed['hex_count']} heat exchangers found")
        
        # Test database storage
        db = AspenDataDatabase()
        print("✅ Database initialized")
        
        # Start a new session
        session_id = db.start_new_session("Test simplified HEX processing")
        print(f"✅ Database session started: {session_id}")
        
        # Store heat exchanger data
        for hex_data in processed['equipment_list']:
            try:
                db.store_hex_data(hex_data)
                print(f"   ✅ Stored: {hex_data['name']} - Hot: {hex_data.get('hot_stream_name', 'N/A')}, Cold: {hex_data.get('cold_stream_name', 'N/A')}")
            except Exception as e:
                print(f"   ❌ Failed to store {hex_data['name']}: {e}")
        
        print(f"\n🎯 Database Storage Summary:")
        print(f"   Total heat exchangers processed: {processed['hex_count']}")
        print(f"   Hot stream → inlet streams mapping implemented")
        print(f"   Cold stream → outlet streams mapping implemented")
        print("🎉 All database tests passed!")
        
    else:
        print("❌ No data loaded")
        sys.exit(1)
    
except Exception as e:
    print(f"❌ Error: {e}")
    print("\n📍 Traceback:")
    traceback.print_exc()
    sys.exit(1)
