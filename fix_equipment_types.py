#!/usr/bin/env python3
"""
修复数据库中的设备类型信息
"""

import sqlite3
import os
from pathlib import Path

def connect_to_database():
    """连接到数据库"""
    db_path = "aspen_data.db"
    if not os.path.exists(db_path):
        print(f"❌ 数据库文件 {db_path} 不存在")
        return None
    
    conn = sqlite3.connect(db_path)
    return conn

def update_equipment_types():
    """根据设备名称和Excel映射更新设备类型"""
    
    # 从Excel文件读取设备映射
    equipment_mapping = {
        'COOL2': {'module_type': 'Heater', 'function': 'Heater', 'category': 'Heat Exchanger'},
        'HT8': {'module_type': 'Heater', 'function': 'Heater', 'category': 'Heat Exchanger'},
        'HT9': {'module_type': 'Heater', 'function': 'Heater', 'category': 'Heat Exchanger'},
        'S2': {'module_type': 'Flash2', 'function': 'Flash Column', 'category': 'Separator'},
        'S3': {'module_type': 'Flash2', 'function': 'Flash Column', 'category': 'Separator'},
        'MC1': {'module_type': 'MCompr', 'function': 'Compressor', 'category': 'Compressor'},
        'V3': {'module_type': 'Valve', 'function': 'Valve', 'category': 'Valve'},
        'C-301': {'module_type': 'RadFrac', 'function': 'Distillation Tower', 'category': 'Distillation Column'},
        'DI': {'module_type': 'RadFrac', 'function': 'Distillation Tower', 'category': 'Distillation Column'},
        'B1': {'module_type': 'RStoic', 'function': 'Reactor', 'category': 'Reactor'},
        'MEOH': {'module_type': 'RPlug', 'function': 'Reactor', 'category': 'Reactor'},
        'B11': {'module_type': 'Mixer', 'function': 'Mixer', 'category': 'Mixer'},
        'MIX3': {'module_type': 'Mixer', 'function': 'Mixer', 'category': 'Mixer'},
        'MX1': {'module_type': 'Mixer', 'function': 'Mixer', 'category': 'Mixer'},
        'MX2': {'module_type': 'Mixer', 'function': 'Mixer', 'category': 'Mixer'},
        'F1': {'module_type': 'FSplit', 'function': 'Split Device', 'category': 'Splitter'},
        'U-1': {'module_type': 'Utility', 'function': 'Utility', 'category': 'Utility'}
    }
    
    conn = connect_to_database()
    if not conn:
        return
    
    cursor = conn.cursor()
    
    print("🔄 修复设备类型信息...")
    print("=" * 50)
    
    updated_count = 0
    
    for equipment_name, info in equipment_mapping.items():
        try:
            # 更新设备类型和功能
            cursor.execute("""
                UPDATE aspen_equipment 
                SET 
                    aspen_type = ?,
                    equipment_type = ?,
                    function = ?
                WHERE name = ?
            """, (
                info['module_type'],
                info['category'], 
                info['function'],
                equipment_name
            ))
            
            if cursor.rowcount > 0:
                print(f"✅ {equipment_name}: {info['module_type']} -> {info['category']} ({info['function']})")
                updated_count += 1
            else:
                print(f"⚠️  未找到设备: {equipment_name}")
                
        except Exception as e:
            print(f"❌ 更新 {equipment_name} 时出错: {e}")
    
    conn.commit()
    
    print("=" * 50)
    print(f"📊 更新统计:")
    print(f"  • 成功更新: {updated_count} 个设备")
    print(f"  • 映射总数: {len(equipment_mapping)} 个设备")
    
    # 验证更新结果
    print("\n🔍 验证更新结果:")
    cursor.execute("""
        SELECT name, aspen_type, equipment_type, function 
        FROM aspen_equipment 
        ORDER BY name
    """)
    
    results = cursor.fetchall()
    
    for name, aspen_type, eq_type, function in results:
        status = "✅" if aspen_type != "Unknown" else "❌"
        print(f"  {status} {name}: {aspen_type} -> {eq_type} ({function})")
    
    conn.close()
    print("\n✅ 设备类型修复完成!")

if __name__ == "__main__":
    update_equipment_types()
