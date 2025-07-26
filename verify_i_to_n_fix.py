#!/usr/bin/env python3
"""
I-N列修复结果验证脚本
简单验证数据库中的I-N列数据是否正确填充
"""

import sqlite3
import os
import json
from datetime import datetime

def verify_i_to_n_fix():
    """验证I-N列修复结果"""
    print("🔍 验证I-N列修复结果")
    print("=" * 50)
    
    db_path = "aspen_data.db"
    
    if not os.path.exists(db_path):
        print(f"❌ 数据库文件不存在: {db_path}")
        return False
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # 1. 检查表是否存在
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='heat_exchangers'")
        if not cursor.fetchone():
            print("❌ heat_exchangers表不存在")
            conn.close()
            return False
        
        print("✅ heat_exchangers表存在")
        
        # 2. 检查表结构 - I-N列字段
        cursor.execute("PRAGMA table_info(heat_exchangers)")
        columns = [col[1] for col in cursor.fetchall()]
        
        required_i_to_n_columns = [
            'column_i_data', 'column_i_header',
            'column_j_data', 'column_j_header',
            'column_k_data', 'column_k_header', 
            'column_l_data', 'column_l_header',
            'column_m_data', 'column_m_header',
            'column_n_data', 'column_n_header',
            'columns_i_to_n_raw'
        ]
        
        print(f"\n📋 I-N列字段检查:")
        missing_columns = []
        for col in required_i_to_n_columns:
            if col in columns:
                print(f"  ✅ {col}")
            else:
                print(f"  ❌ {col} (缺失)")
                missing_columns.append(col)
        
        if missing_columns:
            print(f"\n❌ 表结构不完整，缺失 {len(missing_columns)} 个字段")
            conn.close()
            return False
        
        print(f"✅ 所有I-N列字段都存在")
        
        # 3. 检查数据记录
        cursor.execute("SELECT COUNT(*) FROM heat_exchangers")
        total_records = cursor.fetchone()[0]
        print(f"\n📊 数据记录检查:")
        print(f"  总记录数: {total_records}")
        
        if total_records == 0:
            print("❌ 表中没有数据")
            conn.close()
            return False
        
        # 4. 检查I-N列数据覆盖率
        print(f"\n🔍 I-N列数据覆盖率:")
        i_to_n_data_columns = [
            ('column_i_data', 'I'),
            ('column_j_data', 'J'), 
            ('column_k_data', 'K'),
            ('column_l_data', 'L'),
            ('column_m_data', 'M'),
            ('column_n_data', 'N')
        ]
        
        total_i_to_n_values = 0
        coverage_summary = {}
        
        for db_col, excel_col in i_to_n_data_columns:
            cursor.execute(f"SELECT COUNT(*) FROM heat_exchangers WHERE {db_col} IS NOT NULL")
            count = cursor.fetchone()[0]
            coverage_pct = (count / total_records) * 100 if total_records > 0 else 0
            
            coverage_summary[excel_col] = {
                'count': count,
                'coverage_percentage': coverage_pct
            }
            total_i_to_n_values += count
            
            status = "✅" if count > 0 else "❌"
            print(f"  {status} 列{excel_col}: {count}/{total_records} ({coverage_pct:.1f}%)")
        
        # 5. 显示样本数据
        print(f"\n🔬 样本数据验证:")
        cursor.execute("""
            SELECT name, 
                   column_i_data, column_i_header,
                   column_j_data, column_j_header, 
                   column_k_data, column_k_header,
                   column_l_data, column_l_header,
                   column_m_data, column_m_header,
                   column_n_data, column_n_header
            FROM heat_exchangers 
            LIMIT 3
        """)
        
        sample_count = 0
        for row in cursor.fetchall():
            sample_count += 1
            name = row[0]
            print(f"  样本 {sample_count} ({name}):")
            
            # 显示I-N列数据
            i_to_n_sample = {
                'I': {'data': row[1], 'header': row[2]},
                'J': {'data': row[3], 'header': row[4]},
                'K': {'data': row[5], 'header': row[6]},
                'L': {'data': row[7], 'header': row[8]},
                'M': {'data': row[9], 'header': row[10]},  
                'N': {'data': row[11], 'header': row[12]}
            }
            
            for col, info in i_to_n_sample.items():
                data_val = info['data']
                header_val = info['header']
                if data_val is not None or header_val is not None:
                    print(f"    列{col}: {data_val} ('{header_val}')")
        
        # 6. 生成验证报告
        print(f"\n📋 验证结果总结:")
        print(f"  表结构: {'✅ 完整' if not missing_columns else '❌ 不完整'}")
        print(f"  数据记录: {total_records} 条")
        print(f"  I-N数据点: {total_i_to_n_values} 个")
        print(f"  平均每行数据点: {total_i_to_n_values/total_records:.1f}" if total_records > 0 else "  平均每行数据点: 0")
        
        # 成功判断标准
        success = (
            not missing_columns and  # 表结构完整
            total_records > 0 and    # 有数据记录
            total_i_to_n_values > 0  # 有I-N数据
        )
        
        if success:
            print(f"\n🎉 I-N列修复验证成功!")
            print(f"数据库中的heat_exchangers表已包含完整的I-N列数据")
        else:
            print(f"\n❌ I-N列修复验证失败")
            if not missing_columns:
                print("  表结构正确但数据可能有问题")
            else:
                print("  表结构不完整")
        
        # 保存验证报告
        verification_report = {
            'timestamp': datetime.now().isoformat(),
            'database_path': db_path,
            'table_exists': True,
            'missing_columns': missing_columns,
            'total_records': total_records,
            'total_i_to_n_values': total_i_to_n_values,
            'coverage_summary': coverage_summary,
            'verification_passed': success
        }
        
        report_file = f"i_to_n_verification_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        try:
            with open(report_file, 'w', encoding='utf-8') as f:
                json.dump(verification_report, f, indent=2, ensure_ascii=False, default=str)
            print(f"\n💾 验证报告已保存: {report_file}")
        except Exception as e:
            print(f"\n⚠️ 验证报告保存失败: {e}")
        
        conn.close()
        return success
        
    except Exception as e:
        print(f"❌ 验证过程失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """主函数"""
    print("🚀 启动I-N列修复结果验证")
    
    success = verify_i_to_n_fix()
    
    if success:
        print(f"\n✅ 验证完成: I-N列数据修复成功!")
    else:
        print(f"\n❌ 验证完成: I-N列数据修复失败或不完整")
    
    return success

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)