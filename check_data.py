#!/usr/bin/env python3
import sqlite3
import os

print('Current directory:', os.getcwd())
conn = sqlite3.connect('aspen_data.db')
cursor = conn.cursor()

# 检查heat_exchangers表的数据
cursor.execute('SELECT COUNT(*) FROM heat_exchangers')
hex_count = cursor.fetchone()[0]
print(f'Heat exchanger records: {hex_count}')

if hex_count > 0:
    cursor.execute('SELECT name, duty_kw, area_m2, hot_stream_name, cold_stream_name FROM heat_exchangers LIMIT 3')
    records = cursor.fetchall()
    print('\nSample heat exchanger records:')
    for record in records:
        print(f'  {record[0]}: duty={record[1]}kW, area={record[2]}m², hot={record[3]}, cold={record[4]}')

# 检查aspen_equipment表是否有inlet/outlet字段
cursor.execute('PRAGMA table_info(aspen_equipment)')
columns = cursor.fetchall()
print('\naspen_equipment table columns:')
for col in columns:
    print(f'  {col[1]} ({col[2]})')

conn.close()
