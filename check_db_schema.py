#!/usr/bin/env python3
"""
数据库表结构检查脚本
检查heat_exchangers表是否包含I-N列字段
"""

import sqlite3
import os

def check_heat_exchangers_schema():
    """检查heat_exchangers表结构"""
    db_path = "aspen_data.db"
    
    if not os.path.exists(db_path):
        print(f"❌ 数据库文件不存在: {db_path}")
        return False
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 检查表是否存在
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='heat_exchangers'")
        if not cursor.fetchone():
            print("❌ heat_exchangers表不存在")
            conn.close()
            return False
        
        print("✅ heat_exchangers表存在")
        
        # 获取表结构
        cursor.execute("PRAGMA table_info(heat_exchangers)")
        columns = cursor.fetchall()
        
        print(f"\n📋 当前heat_exchangers表结构 ({len(columns)} 个字段):")
        print("-" * 60)
        
        existing_columns = []
        for col in columns:
            col_name = col[1]
            col_type = col[2]
            not_null = "NOT NULL" if col[3] else "NULL"
            default_val = col[4] if col[4] else "None"
            print(f"  {col_name:25s} {col_type:15s} {not_null:8s} Default: {default_val}")
            existing_columns.append(col_name)
        
        # 检查I-N列字段
        required_i_to_n_columns = [
            'column_i_data', 'column_i_header',
            'column_j_data', 'column_j_header', 
            'column_k_data', 'column_k_header',
            'column_l_data', 'column_l_header',
            'column_m_data', 'column_m_header',
            'column_n_data', 'column_n_header',
            'columns_i_to_n_raw'
        ]
        
        print(f"\n🔍 I-N列字段检查:")
        print("-" * 40)
        
        existing_i_to_n = []
        missing_i_to_n = []
        
        for col in required_i_to_n_columns:
            if col in existing_columns:
                existing_i_to_n.append(col)
                print(f"  ✅ {col}")
            else:
                missing_i_to_n.append(col)
                print(f"  ❌ {col} (缺失)")
        
        print(f"\n📊 I-N列字段统计:")
        print(f"  存在字段: {len(existing_i_to_n)}/{len(required_i_to_n_columns)}")
        print(f"  缺失字段: {len(missing_i_to_n)}")
        
        if missing_i_to_n:
            print(f"\n⚠️ 需要添加的字段:")
            for col in missing_i_to_n:
                print(f"    {col}")
            
            # 检查记录数
            cursor.execute("SELECT COUNT(*) FROM heat_exchangers")
            record_count = cursor.fetchone()[0]
            print(f"\n📊 当前记录数: {record_count}")
            
            conn.close()
            return False
        else:
            print(f"\n✅ 所有I-N列字段都存在!")
            conn.close()
            return True
            
    except Exception as e:
        print(f"❌ 数据库检查失败: {e}")
        return False

if __name__ == "__main__":
    print("🔍 检查数据库表结构")
    print("=" * 50)
    
    schema_ok = check_heat_exchangers_schema()
    
    if schema_ok:
        print("\n🎉 数据库表结构完整，可以直接运行I-N列数据填充")
    else:
        print("\n🛠️ 需要先修复数据库表结构，然后再填充I-N列数据")