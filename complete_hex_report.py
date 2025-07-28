#!/usr/bin/env python3
"""Complete heat exchanger data with temperatures"""

import sqlite3
import json

try:
    conn = sqlite3.connect('aspen_data.db')
    cursor = conn.cursor()

    cursor.execute('''
        SELECT name, inlet_streams, outlet_streams, duty_kw, area_m2,
               hot_inlet_temp, hot_outlet_temp, cold_inlet_temp, cold_outlet_temp 
        FROM heat_exchangers 
        ORDER BY name
    ''')
    results = cursor.fetchall()

    print('🔥 Complete Heat Exchanger Data with Temperatures:')
    print('='*100)
    print(f'{"Name":8} | {"Duty":6} | {"Area":8} | {"Hot In":7} | {"Hot Out":8} | {"Cold In":8} | {"Cold Out":9} | {"Inlet Stream":25} | {"Outlet Stream":25}')
    print('-'*100)
    
    for row in results:
        name, inlet_json, outlet_json, duty, area, h_in, h_out, c_in, c_out = row
        
        inlet_streams = json.loads(inlet_json) if inlet_json else []
        outlet_streams = json.loads(outlet_json) if outlet_json else []
        
        inlet_name = inlet_streams[0] if inlet_streams else "None"
        outlet_name = outlet_streams[0] if outlet_streams else "None"
        
        # Format temperature data
        h_in_str = f'{h_in:7.1f}' if h_in is not None else '    N/A'
        h_out_str = f'{h_out:8.1f}' if h_out is not None else '     N/A'  
        c_in_str = f'{c_in:8.1f}' if c_in is not None else '     N/A'
        c_out_str = f'{c_out:9.1f}' if c_out is not None else '      N/A'
        
        print(f'{name:8} | {duty:4.0f}kW | {area:6.1f}m² | {h_in_str} | {h_out_str} | {c_in_str} | {c_out_str} | {inlet_name[:25]:25} | {outlet_name[:25]:25}')

    print('='*100)
    print(f'✅ Total: {len(results)} heat exchangers with complete temperature data')
    
    # Calculate temperature differences
    print('\n📊 Temperature Analysis:')
    cursor.execute('''
        SELECT name, 
               (hot_inlet_temp - hot_outlet_temp) as hot_temp_drop,
               (cold_outlet_temp - cold_inlet_temp) as cold_temp_rise
        FROM heat_exchangers 
        WHERE hot_inlet_temp IS NOT NULL AND hot_outlet_temp IS NOT NULL 
        AND cold_inlet_temp IS NOT NULL AND cold_outlet_temp IS NOT NULL
    ''')
    temp_analysis = cursor.fetchall()
    
    print(f'{"Name":8} | {"Hot ΔT":8} | {"Cold ΔT":9}')
    print('-'*30)
    for name, hot_drop, cold_rise in temp_analysis:
        print(f'{name:8} | {hot_drop:6.1f}°C | {cold_rise:7.1f}°C')
    
    conn.close()
    
except Exception as e:
    print(f"Error: {e}")
