#!/usr/bin/env python3
import sqlite3

conn = sqlite3.connect('aspen_data.db')
cursor = conn.cursor()

# 获取完整的heat_exchangers数据
cursor.execute('SELECT name, hot_stream_name, cold_stream_name, column_i_data, column_l_data FROM heat_exchangers LIMIT 5')
records = cursor.fetchall()

print('Heat exchanger data in database:')
for record in records:
    print(f'  {record[0]}:')
    print(f'    hot_stream_name: {record[1]} (type: {type(record[1])})')
    print(f'    cold_stream_name: {record[2]} (type: {type(record[2])})')
    print(f'    column_i_data: {record[3]} (type: {type(record[3])})')
    print(f'    column_l_data: {record[4]} (type: {type(record[4])})')
    print()

conn.close()
