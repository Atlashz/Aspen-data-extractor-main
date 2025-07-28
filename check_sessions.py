#!/usr/bin/env python3
import sqlite3
import os

# 检查数据库文件
if os.path.exists('aspen_data.db'):
    conn = sqlite3.connect('aspen_data.db')
    cursor = conn.cursor()
    
    # 检查提取会话
    try:
        cursor.execute('SELECT COUNT(*) FROM extraction_sessions')
        session_count = cursor.fetchone()[0]
        print(f'Total extraction sessions: {session_count}')
        
        if session_count > 0:
            cursor.execute('SELECT session_id, extraction_time FROM extraction_sessions ORDER BY extraction_time DESC LIMIT 5')
            sessions = cursor.fetchall()
            print('Recent sessions:')
            for session in sessions:
                print(f'  {session[0]}: {session[1]}')
        
        # 检查流股记录  
        cursor.execute('SELECT COUNT(*) FROM aspen_streams')
        stream_count = cursor.fetchone()[0]
        print(f'Stream records: {stream_count}')
        
        # 检查换热器记录
        cursor.execute('SELECT COUNT(*) FROM heat_exchangers')  
        hex_count = cursor.fetchone()[0]
        print(f'Heat exchanger records: {hex_count}')
        
    except Exception as e:
        print(f'Error querying database: {e}')
    
    conn.close()
else:
    print('Database file not found')
