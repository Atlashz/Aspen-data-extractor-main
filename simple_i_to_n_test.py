#!/usr/bin/env python3
"""
Simple I-N Column Test
简单测试I-N列数据读取
"""

import pandas as pd
import sqlite3
import os

def test_excel_i_to_n():
    """测试Excel文件I-N列读取"""
    excel_file = "BFG-CO2H-HEX.xlsx"
    
    print("🔍 测试Excel文件I-N列读取")
    print("=" * 50)
    
    try:
        if not os.path.exists(excel_file):
            print(f"❌ 文件不存在: {excel_file}")
            return False
        
        print(f"✅ 文件存在: {excel_file}")
        
        # 读取Excel
        df = pd.read_excel(excel_file)
        print(f"📊 文件大小: {len(df)} 行 × {len(df.columns)} 列")
        
        # 检查I-N列 (索引8-13)
        i_to_n_indices = list(range(8, 14))  # I-N对应索引8-13
        
        print(f"\n🔍 I-N列检查:")
        for i, idx in enumerate(i_to_n_indices):
            excel_col = chr(73 + i)  # I=73, J=74, etc.
            if idx < len(df.columns):
                header = df.columns[idx]
                non_null = df.iloc[:, idx].notna().sum()
                sample = df.iloc[:, idx].dropna().head(3).tolist()
                print(f"  列{excel_col} (索引{idx}): '{header}'")
                print(f"    有效数据: {non_null}/{len(df)}")
                print(f"    样本: {sample}")
            else:
                print(f"  列{excel_col}: 不存在")
        
        return True
        
    except Exception as e:
        print(f"❌ Excel读取失败: {e}")
        return False

def test_database_i_to_n():
    """测试数据库I-N列状态"""
    db_file = "aspen_data.db"
    
    print(f"\n🔍 测试数据库I-N列状态")  
    print("-" * 50)
    
    try:
        if not os.path.exists(db_file):
            print(f"❌ 数据库不存在: {db_file}")
            return False
        
        print(f"✅ 数据库存在: {db_file}")
        
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        
        # 检查heat_exchangers表是否存在
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='heat_exchangers'")
        if not cursor.fetchone():
            print("❌ heat_exchangers表不存在")
            conn.close()
            return False
        
        print("✅ heat_exchangers表存在")
        
        # 检查总记录数
        cursor.execute("SELECT COUNT(*) FROM heat_exchangers")
        total_count = cursor.fetchone()[0]
        print(f"📊 总记录数: {total_count}")
        
        if total_count == 0:
            print("⚠️ 表中无数据")
            conn.close()
            return total_count > 0
        
        # 检查I-N列数据
        i_to_n_columns = [
            'column_i_data', 'column_j_data', 'column_k_data',
            'column_l_data', 'column_m_data', 'column_n_data'
        ]
        
        print(f"\n🔍 I-N列数据检查:")
        total_i_to_n = 0
        
        for col in i_to_n_columns:
            cursor.execute(f"SELECT COUNT(*) FROM heat_exchangers WHERE {col} IS NOT NULL")
            count = cursor.fetchone()[0]
            excel_col = col.split('_')[1].upper()
            print(f"  列{excel_col}: {count}/{total_count} 有数据")
            total_i_to_n += count
        
        print(f"\n📊 I-N列数据汇总:")
        print(f"  总I-N数据点: {total_i_to_n}")
        print(f"  平均每行: {total_i_to_n/total_count:.1f}" if total_count > 0 else "  平均每行: 0")
        
        # 显示样本数据
        cursor.execute("""
            SELECT name, column_i_data, column_j_data, column_k_data,
                   column_l_data, column_m_data, column_n_data
            FROM heat_exchangers LIMIT 3
        """)
        
        print(f"\n🔬 样本数据:")
        for row in cursor.fetchall():
            name = row[0]
            i_to_n_values = row[1:7]
            print(f"  {name}: I={i_to_n_values[0]}, J={i_to_n_values[1]}, K={i_to_n_values[2]}, L={i_to_n_values[3]}, M={i_to_n_values[4]}, N={i_to_n_values[5]}")
        
        conn.close()
        return total_i_to_n > 0
        
    except Exception as e:
        print(f"❌ 数据库检查失败: {e}")
        return False

def main():
    """主测试函数"""
    print("🚀 I-N列数据简单测试")
    print("=" * 80)
    
    # 测试Excel文件
    excel_ok = test_excel_i_to_n()
    
    # 测试数据库
    db_ok = test_database_i_to_n()
    
    print(f"\n📋 测试结果:")
    print(f"  Excel文件读取: {'✅ 成功' if excel_ok else '❌ 失败'}")
    print(f"  数据库I-N数据: {'✅ 有数据' if db_ok else '❌ 无数据'}")
    
    if excel_ok and not db_ok:
        print(f"\n💡 建议:")
        print(f"  Excel文件中有I-N列数据，但数据库中缺失")
        print(f"  需要运行数据提取或修复脚本")
    elif not excel_ok:
        print(f"\n⚠️ 警告:")
        print(f"  Excel文件读取失败，请检查文件是否存在且格式正确")
    elif db_ok:
        print(f"\n🎉 成功:")
        print(f"  I-N列数据已正常存储在数据库中")

if __name__ == "__main__":
    main()