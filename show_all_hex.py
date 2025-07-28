#!/usr/bin/env python3
"""Display all heat exchanger data with stream mapping"""

import sqlite3
import json

try:
    conn = sqlite3.connect('aspen_data.db')
    cursor = conn.cursor()

    cursor.execute('SELECT name, inlet_streams, outlet_streams, duty_kw FROM heat_exchangers ORDER BY name')
    results = cursor.fetchall()

    print('🔥 All Heat Exchangers with Stream Mapping:')
    print('='*80)
    for row in results:
        name, inlet_json, outlet_json, duty = row
        inlet_streams = json.loads(inlet_json) if inlet_json else []
        outlet_streams = json.loads(outlet_json) if outlet_json else []
        
        inlet_name = inlet_streams[0] if inlet_streams else "None"
        outlet_name = outlet_streams[0] if outlet_streams else "None"
        
        print(f'{name:8} | {duty:6.1f}kW | Inlet: {inlet_name:30} | Outlet: {outlet_name:30}')

    print('='*80)
    print(f'✅ Total: {len(results)} heat exchangers stored successfully')
    conn.close()
    
except Exception as e:
    print(f"Error: {e}")
