#!/usr/bin/env python3
"""
数据库表结构修复脚本

专门修复heat_exchangers表缺少I-N列字段的问题
安全地添加缺失的列，保留现有数据

Author: TEA Analysis Framework
Date: 2025-07-26
Version: 1.0 - Database Schema Fix
"""

import sqlite3
import os
import json
from datetime import datetime
from typing import Dict, List, Any

class DatabaseSchemaFixer:
    """
    数据库表结构修复器
    """
    
    def __init__(self, db_path: str = "aspen_data.db"):
        self.db_path = db_path
        
    def fix_heat_exchangers_schema(self) -> Dict[str, Any]:
        """
        修复heat_exchangers表结构，添加I-N列字段
        """
        print("\n" + "="*80)
        print("🔧 数据库表结构修复工具")
        print("="*80)
        print(f"数据库: {self.db_path}")
        print(f"修复时间: {datetime.now().isoformat()}")
        
        result = {
            'success': False,
            'database_exists': False,
            'table_exists': False,
            'backup_created': False,
            'columns_added': [],
            'existing_records': 0,
            'error': None
        }
        
        try:
            # 1. 检查数据库是否存在
            if not os.path.exists(self.db_path):
                print(f"❌ 数据库文件不存在: {self.db_path}")
                return result
            
            result['database_exists'] = True
            print(f"✅ 数据库文件存在: {self.db_path}")
            
            # 2. 连接数据库
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 3. 检查heat_exchangers表是否存在
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='heat_exchangers'")
            if not cursor.fetchone():
                print("❌ heat_exchangers表不存在，需要创建完整表结构")
                self._create_complete_heat_exchangers_table(cursor)
                result['table_exists'] = True
            else:
                result['table_exists'] = True
                print("✅ heat_exchangers表存在")
            
            # 4. 检查现有记录数
            cursor.execute("SELECT COUNT(*) FROM heat_exchangers")
            result['existing_records'] = cursor.fetchone()[0]
            print(f"📊 现有记录数: {result['existing_records']}")
            
            # 5. 创建备份
            if result['existing_records'] > 0:
                backup_time = datetime.now().strftime('%Y%m%d_%H%M%S')
                backup_table = f"heat_exchangers_backup_{backup_time}"
                cursor.execute(f"""
                    CREATE TABLE {backup_table} AS 
                    SELECT * FROM heat_exchangers
                """)
                result['backup_created'] = True
                print(f"✅ 数据备份创建: {backup_table}")
            
            # 6. 获取当前表结构
            cursor.execute("PRAGMA table_info(heat_exchangers)")
            existing_columns = [col[1] for col in cursor.fetchall()]
            print(f"📋 现有字段数: {len(existing_columns)}")
            
            # 7. 定义需要的I-N列字段
            required_i_to_n_columns = [
                ('column_i_data', 'REAL'),
                ('column_i_header', 'TEXT'),
                ('column_j_data', 'REAL'), 
                ('column_j_header', 'TEXT'),
                ('column_k_data', 'REAL'),
                ('column_k_header', 'TEXT'),
                ('column_l_data', 'REAL'),
                ('column_l_header', 'TEXT'),
                ('column_m_data', 'REAL'),
                ('column_m_header', 'TEXT'),
                ('column_n_data', 'REAL'),
                ('column_n_header', 'TEXT'),
                ('columns_i_to_n_raw', 'TEXT')
            ]
            
            # 8. 添加缺失的列
            print(f"\n🔧 添加缺失的I-N列字段:")
            columns_added = 0
            
            for col_name, col_type in required_i_to_n_columns:
                if col_name not in existing_columns:
                    try:
                        cursor.execute(f"ALTER TABLE heat_exchangers ADD COLUMN {col_name} {col_type}")
                        result['columns_added'].append(col_name)
                        columns_added += 1
                        print(f"   ✅ 添加字段: {col_name} ({col_type})")
                    except Exception as e:
                        print(f"   ❌ 添加字段失败 {col_name}: {e}")
                else:
                    print(f"   ⚪ 字段已存在: {col_name}")
            
            # 9. 提交更改
            conn.commit()
            
            # 10. 验证表结构
            cursor.execute("PRAGMA table_info(heat_exchangers)")
            final_columns = [col[1] for col in cursor.fetchall()]
            final_i_to_n_count = sum(1 for col in final_columns if col.startswith('column_') and ('_data' in col or '_header' in col or 'i_to_n_raw' in col))
            
            print(f"\n📊 表结构修复结果:")
            print(f"   总字段数: {len(existing_columns)} -> {len(final_columns)}")
            print(f"   新增字段: {columns_added}")
            print(f"   I-N相关字段: {final_i_to_n_count}")
            
            if columns_added > 0 or final_i_to_n_count >= 13:
                result['success'] = True
                print(f"✅ 表结构修复成功!")
            else:
                print(f"⚠️ 表结构可能仍有问题")
            
            conn.close()
            
        except Exception as e:
            print(f"❌ 表结构修复失败: {e}")
            result['error'] = str(e)
            import traceback
            traceback.print_exc()
        
        return result
    
    def _create_complete_heat_exchangers_table(self, cursor):
        """
        创建完整的heat_exchangers表结构（包含I-N列）
        """
        print("🏗️ 创建完整的heat_exchangers表结构")
        
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS heat_exchangers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL,
                name TEXT NOT NULL,
                duty_kw REAL DEFAULT 0.0,
                area_m2 REAL DEFAULT 0.0,
                temperatures TEXT,
                pressures TEXT,
                source TEXT DEFAULT 'unknown',
                extraction_time TEXT,
                hot_stream_name TEXT,
                hot_stream_inlet_temp REAL,
                hot_stream_outlet_temp REAL,
                hot_stream_flow_rate REAL,
                hot_stream_composition TEXT,
                cold_stream_name TEXT,
                cold_stream_inlet_temp REAL,
                cold_stream_outlet_temp REAL,
                cold_stream_flow_rate REAL,
                cold_stream_composition TEXT,
                column_i_data REAL,
                column_i_header TEXT,
                column_j_data REAL,
                column_j_header TEXT,
                column_k_data REAL,
                column_k_header TEXT,
                column_l_data REAL,
                column_l_header TEXT,
                column_m_data REAL,
                column_m_header TEXT,
                column_n_data REAL,
                column_n_header TEXT,
                columns_i_to_n_raw TEXT,
                FOREIGN KEY (session_id) REFERENCES extraction_sessions (session_id)
            )
        """)
        
        print("✅ 完整表结构创建成功")


def main():
    """
    主修复函数
    """
    print("🚀 启动数据库表结构修复工具")
    
    fixer = DatabaseSchemaFixer()
    result = fixer.fix_heat_exchangers_schema()
    
    # 保存修复报告
    report_file = f"schema_fix_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    try:
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, indent=2, ensure_ascii=False, default=str)
        print(f"\n💾 修复报告已保存: {report_file}")
    except Exception as e:
        print(f"\n❌ 报告保存失败: {e}")
    
    if result['success']:
        print(f"\n🎉 数据库表结构修复完成!")
        print("现在可以运行I-N列数据填充脚本了")
        return True
    else:
        print(f"\n⚠️ 数据库表结构修复失败，请检查错误信息")
        return False


if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)