#!/usr/bin/env python3
import sqlite3
import os

# 检查数据库文件
if os.path.exists('aspen_data.db'):
    conn = sqlite3.connect('aspen_data.db')
    cursor = conn.cursor()
    
    # 检查表是否存在
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = cursor.fetchall()
    print('Tables in database:', [t[0] for t in tables])
    
    # 检查设备记录
    try:
        cursor.execute('SELECT COUNT(*) FROM aspen_equipment')
        count = cursor.fetchone()[0]
        print(f'Equipment records: {count}')
        
        if count > 0:
            cursor.execute('SELECT name, equipment_type FROM aspen_equipment LIMIT 5')
            records = cursor.fetchall()
            print('Sample records:', records)
    except Exception as e:
        print(f'Error querying equipment: {e}')
    
    conn.close()
else:
    print('Database file not found')
