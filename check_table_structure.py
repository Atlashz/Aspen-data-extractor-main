#!/usr/bin/env python3
import sqlite3
import os

# 检查数据库文件
if os.path.exists('aspen_data.db'):
    conn = sqlite3.connect('aspen_data.db')
    cursor = conn.cursor()
    
    # 获取heat_exchangers表结构
    cursor.execute("PRAGMA table_info(heat_exchangers)")
    columns = cursor.fetchall()
    
    print("heat_exchangers table structure:")
    for col in columns:
        print(f"  {col[1]} ({col[2]})")
    
    conn.close()
else:
    print('Database file not found')
