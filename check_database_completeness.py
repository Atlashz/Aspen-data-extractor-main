"""
检查数据库完整性并恢复缺失的功能
"""
import sqlite3
import os

def check_database_completeness():
    print("🔍 检查数据库完整性")
    print("="*50)
    
    if not os.path.exists('aspen_data.db'):
        print("❌ aspen_data.db 不存在")
        return
    
    conn = sqlite3.connect('aspen_data.db')
    cursor = conn.cursor()
    
    # 检查所有表
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = [t[0] for t in cursor.fetchall()]
    
    print("📋 当前数据库表:")
    for table in tables:
        cursor.execute(f"SELECT COUNT(*) FROM {table}")
        count = cursor.fetchone()[0]
        print(f"  - {table}: {count} 条记录")
    
    # 检查缺失的重要表
    required_tables = {
        'heat_exchangers': 'HEX换热器数据',
        'improved_stream_mappings': '改进的流股映射',
        'stream_mappings': '基础流股映射'
    }
    
    print(f"\n🔍 检查重要功能:")
    missing_tables = []
    
    for table, description in required_tables.items():
        if table in tables:
            cursor.execute(f"SELECT COUNT(*) FROM {table}")
            count = cursor.fetchone()[0]
            if count > 0:
                print(f"  ✅ {description}: {count} 条记录")
            else:
                print(f"  ⚠️ {description}: 表存在但无数据")
                missing_tables.append(table)
        else:
            print(f"  ❌ {description}: 表不存在")
            missing_tables.append(table)
    
    # 检查流股数据是否包含分类信息
    cursor.execute("PRAGMA table_info(aspen_streams)")
    stream_columns = [col[1] for col in cursor.fetchall()]
    
    print(f"\n🌊 流股表列结构:")
    important_columns = ['stream_category', 'stream_sub_category', 'classification_confidence']
    for col in important_columns:
        if col in stream_columns:
            cursor.execute(f"SELECT COUNT(*) FROM aspen_streams WHERE {col} IS NOT NULL")
            count = cursor.fetchone()[0]
            print(f"  ✅ {col}: {count} 条有数据")
        else:
            print(f"  ❌ {col}: 列不存在")
    
    # 检查设备数据是否包含类型信息
    cursor.execute("SELECT COUNT(*) FROM aspen_equipment WHERE equipment_type != 'Unknown'")
    typed_equipment = cursor.fetchone()[0]
    print(f"\n⚙️ 设备类型识别: {typed_equipment}/16 个设备有明确类型")
    
    if typed_equipment < 5:
        print("  ⚠️ 设备类型识别功能可能缺失")
    
    conn.close()
    
    return missing_tables

def check_external_files():
    print(f"\n📁 检查外部数据文件:")
    
    # 检查HEX Excel文件
    hex_file = "BFG-CO2H-HEX.xlsx"
    if os.path.exists(hex_file):
        print(f"  ✅ HEX数据文件: {hex_file}")
    else:
        print(f"  ❌ HEX数据文件缺失: {hex_file}")
    
    # 检查映射文件
    mapping_files = [
        "stream_mappings_20250725_151156.xlsx",
        "equipment_mapping_summary_20250725_143048.xlsx"
    ]
    
    for file in mapping_files:
        if os.path.exists(file):
            print(f"  ✅ 映射文件: {file}")
        else:
            print(f"  ❌ 映射文件缺失: {file}")
    
    # 检查关键脚本文件
    key_scripts = [
        "improved_stream_mapping.py",
        "stream_mapping.py", 
        "query_stream_mappings.py"
    ]
    
    print(f"\n🔧 检查功能脚本:")
    for script in key_scripts:
        if os.path.exists(script):
            print(f"  ✅ {script}")
        else:
            print(f"  ❌ {script}")

if __name__ == "__main__":
    missing_tables = check_database_completeness()
    check_external_files()
    
    print(f"\n📝 总结:")
    if missing_tables:
        print(f"❌ 缺失功能: {', '.join(missing_tables)}")
        print("需要恢复这些功能")
    else:
        print("✅ 所有核心功能完整")
