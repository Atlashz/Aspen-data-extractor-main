#!/usr/bin/env python3
"""Check database content"""

import sqlite3

try:
    conn = sqlite3.connect('aspen_data.db')
    cursor = conn.cursor()

    # 检查heat_exchangers表
    cursor.execute('SELECT COUNT(*) FROM heat_exchangers')
    hex_count = cursor.fetchone()[0]
    print(f'Heat exchangers count: {hex_count}')

    if hex_count > 0:
        cursor.execute('SELECT name, inlet_streams, outlet_streams, duty_kw, area_m2 FROM heat_exchangers LIMIT 5')
        results = cursor.fetchall()
        print('Sample heat exchanger data:')
        for row in results:
            print(f'  {row[0]}: Inlet={row[1]}, Outlet={row[2]}, Duty={row[3]:.1f}kW, Area={row[4]:.1f}m²')

    # 检查所有表
    cursor.execute('SELECT name FROM sqlite_master WHERE type="table"')
    tables = cursor.fetchall()
    print(f'Available tables: {[t[0] for t in tables]}')

    # 检查每个表的记录数
    for table in tables:
        cursor.execute(f'SELECT COUNT(*) FROM {table[0]}')
        count = cursor.fetchone()[0]
        print(f'{table[0]}: {count} records')

    conn.close()
    
except Exception as e:
    print(f"Error: {e}")
