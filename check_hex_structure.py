#!/usr/bin/env python3
import sqlite3

conn = sqlite3.connect('aspen_data.db')
cursor = conn.cursor()

# 检查heat_exchangers表的完整结构
cursor.execute('PRAGMA table_info(heat_exchangers)')
columns = cursor.fetchall()
print('Heat exchangers table structure:')
for col in columns:
    print(f'  {col[1]} ({col[2]})')

print('\nSample data:')
cursor.execute('SELECT name, duty_kw, area_m2, hot_stream_name, cold_stream_name, hot_stream_inlet_temp, cold_stream_inlet_temp FROM heat_exchangers LIMIT 3')
records = cursor.fetchall()
for record in records:
    print(f'  {record[0]}: duty={record[1]}kW, area={record[2]}m², hot={record[3]}, cold={record[4]}, hot_temp={record[5]}, cold_temp={record[6]}')

conn.close()
