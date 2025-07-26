#!/usr/bin/env python3
"""
重新提取HEX数据并更新数据库
"""

import os
import sys
import sqlite3
import pandas as pd
import json
from datetime import datetime

def convert_kj_to_kw(kj_per_hour):
    """将kJ/h转换为kW"""
    return kj_per_hour / 3600

def extract_hex_data_from_excel():
    """从Excel文件提取HEX数据"""
    excel_file = "BFG-CO2H-HEX.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"❌ Excel文件不存在: {excel_file}")
        return None
    
    try:
        print(f"📊 从 {excel_file} 读取数据...")
        df = pd.read_excel(excel_file)
        
        print(f"✅ 成功读取 {len(df)} 行数据")
        print(f"📋 列名: {df.columns.tolist()}")
        
        hex_data = []
        
        for idx, row in df.iterrows():
            # 提取换热器名称
            name = str(row['Heat Exchanger Name']) if pd.notna(row['Heat Exchanger Name']) else f"HEX-{idx:03d}"
            
            # 提取Load并转换为kW
            load_kj_h = row['Load Kj/h'] if pd.notna(row['Load Kj/h']) else 0.0
            duty_kw = convert_kj_to_kw(load_kj_h)
            
            # 提取面积
            area_m2 = row['Area m2'] if pd.notna(row['Area m2']) else 0.0
            
            # 提取温度数据
            temperatures = {}
            temp_cols = ['Hot T in ( C )', 'Hot T out ( C )', 'Cold T in', 'Cold T out']
            for col in temp_cols:
                if col in row and pd.notna(row[col]):
                    temperatures[col] = float(row[col])
            
            # 提取压力数据（如果有的话）
            pressures = {}
            # 目前Excel中没有压力数据，留空
            
            hex_info = {
                'name': name,
                'duty_kw': duty_kw,
                'area_m2': area_m2,
                'temperatures': temperatures,
                'pressures': pressures,
                'load_kj_h': load_kj_h,  # 原始数据
                'hot_stream': str(row['hot stream']) if pd.notna(row['hot stream']) else '',
                'cold_stream': str(row['Cold stream']) if pd.notna(row['Cold stream']) else ''
            }
            
            hex_data.append(hex_info)
            print(f"  📦 {name}: {duty_kw:.1f} kW, {area_m2:.1f} m²")
        
        total_duty = sum(h['duty_kw'] for h in hex_data)
        total_area = sum(h['area_m2'] for h in hex_data)
        
        print(f"\n📊 汇总统计:")
        print(f"  • 换热器总数: {len(hex_data)}")
        print(f"  • 总热负荷: {total_duty:,.1f} kW")
        print(f"  • 总面积: {total_area:,.1f} m²")
        
        return hex_data
        
    except Exception as e:
        print(f"❌ 提取HEX数据失败: {e}")
        return None

def update_hex_database(hex_data):
    """更新数据库中的HEX数据"""
    if not hex_data:
        return False
    
    try:
        conn = sqlite3.connect('aspen_data.db')
        cursor = conn.cursor()
        
        # 获取当前session_id
        cursor.execute("SELECT session_id FROM extraction_sessions ORDER BY session_id DESC LIMIT 1")
        result = cursor.fetchone()
        session_id = result[0] if result else f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        print(f"🔄 更新数据库中的HEX数据 (session: {session_id})...")
        
        # 删除旧的HEX数据
        cursor.execute("DELETE FROM heat_exchangers WHERE session_id = ?", (session_id,))
        deleted_count = cursor.rowcount
        print(f"  🗑️ 删除了 {deleted_count} 条旧记录")
        
        # 插入新的HEX数据
        insert_count = 0
        for hex_info in hex_data:
            cursor.execute("""
                INSERT INTO heat_exchangers 
                (session_id, name, duty_kw, area_m2, temperatures, pressures, source, extraction_time)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                session_id,
                hex_info['name'],
                hex_info['duty_kw'],
                hex_info['area_m2'],
                json.dumps(hex_info['temperatures']),
                json.dumps(hex_info['pressures']),
                'excel_corrected',
                datetime.now().isoformat()
            ))
            insert_count += 1
        
        conn.commit()
        conn.close()
        
        print(f"  ✅ 成功插入 {insert_count} 条新记录")
        return True
        
    except Exception as e:
        print(f"❌ 更新数据库失败: {e}")
        return False

def verify_hex_data():
    """验证更新后的HEX数据"""
    try:
        conn = sqlite3.connect('aspen_data.db')
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT name, duty_kw, area_m2, temperatures, source 
            FROM heat_exchangers 
            ORDER BY name
        """)
        
        results = cursor.fetchall()
        
        print(f"\n🔍 验证更新后的HEX数据:")
        print(f"📊 数据库中的换热器 ({len(results)} 个):")
        
        total_duty = 0
        total_area = 0
        
        for name, duty_kw, area_m2, temps_json, source in results:
            total_duty += duty_kw
            total_area += area_m2
            
            try:
                temps = json.loads(temps_json)
                temp_info = f"T范围: {min(temps.values()):.1f}-{max(temps.values()):.1f}°C" if temps else "无温度数据"
            except:
                temp_info = "温度数据格式错误"
            
            print(f"  📦 {name}: {duty_kw:.1f} kW, {area_m2:.1f} m², {temp_info}")
        
        print(f"\n📈 汇总:")
        print(f"  • 总热负荷: {total_duty:,.1f} kW")
        print(f"  • 总面积: {total_area:,.1f} m²")
        
        conn.close()
        return True
        
    except Exception as e:
        print(f"❌ 验证失败: {e}")
        return False

if __name__ == "__main__":
    print("🔧 HEX数据修复工具")
    print("=" * 50)
    
    # 1. 从Excel提取数据
    hex_data = extract_hex_data_from_excel()
    if not hex_data:
        sys.exit(1)
    
    # 2. 更新数据库
    if update_hex_database(hex_data):
        print("✅ HEX数据更新成功")
    else:
        print("❌ HEX数据更新失败")
        sys.exit(1)
    
    # 3. 验证结果
    verify_hex_data()
    
    print("\n🎉 HEX数据修复完成!")
