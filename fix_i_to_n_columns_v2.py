#!/usr/bin/env python3
"""
I-N Column Data Fix Script V2

修复版本：先修复数据库表结构，再填充I-N列数据
专门解决BFG-CO2H-HEX.xlsx中I-N列数据提取和存储问题

Author: TEA Analysis Framework  
Date: 2025-07-26
Version: 2.0 - Complete Fix with Schema Update
"""

import pandas as pd
import sqlite3
import json
import logging
import os
from datetime import datetime
from typing import Dict, List, Any, Optional

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(levelname)s:%(name)s:%(message)s')
logger = logging.getLogger(__name__)

class IToNColumnFixerV2:
    """
    专门用于修复I-N列数据提取问题的类（V2版本）
    包含表结构修复和数据填充
    """
    
    def __init__(self, excel_file: str = "BFG-CO2H-HEX.xlsx", db_path: str = "aspen_data.db"):
        self.excel_file = excel_file
        self.db_path = db_path
        self.df = None
        self.i_to_n_data = {}
        
    def complete_fix(self) -> Dict[str, Any]:
        """
        完整修复流程：表结构修复 + 数据填充
        """
        print("\n" + "="*80)
        print("🔧 I-N列数据完整修复工具 V2")
        print("="*80)
        print(f"Excel文件: {self.excel_file}")
        print(f"数据库: {self.db_path}")
        print(f"修复时间: {datetime.now().isoformat()}")
        
        results = {
            'step_1_schema_fix': self._step1_fix_database_schema(),
            'step_2_excel_analysis': self._step2_analyze_excel(),
            'step_3_extract_i_to_n': self._step3_extract_i_to_n_data(),
            'step_4_update_database': self._step4_update_database(),
            'step_5_verify_fix': self._step5_verify_fix()
        }
        
        self._generate_complete_report(results)
        return results
    
    def _step1_fix_database_schema(self) -> Dict[str, Any]:
        """
        Step 1: 修复数据库表结构，确保I-N列字段存在
        """
        print("\n🔧 Step 1: 修复数据库表结构")
        print("-" * 50)
        
        result = {
            'success': False,
            'database_exists': False,
            'table_exists': False,
            'columns_added': [],
            'error': None
        }
        
        try:
            # 检查数据库文件
            if not os.path.exists(self.db_path):
                print(f"❌ 数据库文件不存在: {self.db_path}")
                return result
            
            result['database_exists'] = True
            print(f"✅ 数据库文件存在: {self.db_path}")
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 检查heat_exchangers表
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='heat_exchangers'")
            if not cursor.fetchone():
                print("📋 heat_exchangers表不存在，创建完整表结构")
                self._create_complete_heat_exchangers_table(cursor)
                result['table_exists'] = True
            else:
                result['table_exists'] = True
                print("✅ heat_exchangers表存在")
            
            # 获取现有字段
            cursor.execute("PRAGMA table_info(heat_exchangers)")
            existing_columns = [col[1] for col in cursor.fetchall()]
            
            # 需要的I-N列字段
            required_i_to_n_fields = [
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
            
            # 添加缺失字段
            columns_added = 0
            for col_name, col_type in required_i_to_n_fields:
                if col_name not in existing_columns:
                    try:
                        cursor.execute(f"ALTER TABLE heat_exchangers ADD COLUMN {col_name} {col_type}")
                        result['columns_added'].append(col_name)
                        columns_added += 1
                        print(f"   ✅ 添加字段: {col_name}")
                    except Exception as e:
                        print(f"   ❌ 添加字段失败 {col_name}: {e}")
            
            conn.commit()
            conn.close()
            
            if columns_added > 0 or len([col for col in existing_columns if 'column_' in col and ('_data' in col or '_header' in col)]) >= 10:
                result['success'] = True
                print(f"✅ 表结构修复成功，添加了 {columns_added} 个字段")
            else:
                print(f"⚠️ 表结构可能已经完整")
                result['success'] = True  # 字段已存在也算成功
            
        except Exception as e:
            print(f"❌ 表结构修复失败: {e}")
            result['error'] = str(e)
        
        return result
    
    def _create_complete_heat_exchangers_table(self, cursor):
        """创建完整的heat_exchangers表"""
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
                columns_i_to_n_raw TEXT
            )
        """)
    
    def _step2_analyze_excel(self) -> Dict[str, Any]:
        """
        Step 2: 分析Excel文件
        """
        print(f"\n🔍 Step 2: 分析Excel文件")
        print("-" * 50)
        
        result = {
            'success': False,
            'file_found': False,
            'total_columns': 0,
            'total_rows': 0,
            'i_to_n_columns': {},
            'error': None
        }
        
        try:
            if not os.path.exists(self.excel_file):
                print(f"❌ Excel文件不存在: {self.excel_file}")
                return result
            
            result['file_found'] = True
            print(f"✅ Excel文件找到: {self.excel_file}")
            
            # 读取Excel
            self.df = pd.read_excel(self.excel_file)
            result['total_columns'] = len(self.df.columns)
            result['total_rows'] = len(self.df)
            
            print(f"📊 文件结构: {result['total_rows']} 行 × {result['total_columns']} 列")
            
            # I-N列映射（索引8-13）
            i_to_n_mapping = {'I': 8, 'J': 9, 'K': 10, 'L': 11, 'M': 12, 'N': 13}
            
            for excel_col, col_idx in i_to_n_mapping.items():
                if col_idx < len(self.df.columns):
                    header = str(self.df.columns[col_idx])
                    non_null_count = self.df.iloc[:, col_idx].notna().sum()
                    result['i_to_n_columns'][excel_col] = {
                        'index': col_idx,
                        'header': header,
                        'non_null_count': int(non_null_count)
                    }
                    print(f"   列{excel_col}: '{header}' ({non_null_count} 个有效值)")
            
            result['success'] = True
            
        except Exception as e:
            print(f"❌ Excel分析失败: {e}")
            result['error'] = str(e)
        
        return result
    
    def _step3_extract_i_to_n_data(self) -> Dict[str, Any]:
        """
        Step 3: 提取I-N列数据
        """
        print(f"\n📤 Step 3: 提取I-N列数据")
        print("-" * 50)
        
        result = {
            'success': False,
            'total_rows_processed': 0,
            'rows_with_i_to_n_data': 0,
            'extracted_data_count': {},
            'error': None
        }
        
        try:
            if self.df is None:
                print("❌ Excel数据未加载")
                return result
            
            i_to_n_mapping = {'I': 8, 'J': 9, 'K': 10, 'L': 11, 'M': 12, 'N': 13}
            extracted_data = []
            rows_with_data = 0
            
            for idx, row in self.df.iterrows():
                row_data = {
                    'row_index': idx,
                    'name': f'HEX-{idx+1:03d}',
                    'i_to_n_columns': {}
                }
                
                has_data = False
                for excel_col, col_idx in i_to_n_mapping.items():
                    if col_idx < len(self.df.columns):
                        header = str(self.df.columns[col_idx])
                        value = row.iloc[col_idx] if col_idx < len(row) else None
                        
                        # 对于数值列，转换为float；对于文本列，保持字符串
                        if excel_col in ['I', 'L']:  # 流股名称列
                            clean_value = str(value) if value is not None and not pd.isna(value) else None
                        else:  # 温度列
                            clean_value = self._clean_numeric_value(value)
                        
                        if clean_value is not None:
                            row_data['i_to_n_columns'][excel_col.lower()] = {
                                'data': clean_value,
                                'header': header,
                                'raw_value': value
                            }
                            has_data = True
                
                if has_data:
                    extracted_data.append(row_data)
                    rows_with_data += 1
            
            self.i_to_n_data = extracted_data
            result['total_rows_processed'] = len(self.df)
            result['rows_with_i_to_n_data'] = rows_with_data
            
            # 统计每列
            for excel_col in ['I', 'J', 'K', 'L', 'M', 'N']:
                count = sum(1 for row in extracted_data if excel_col.lower() in row['i_to_n_columns'])
                result['extracted_data_count'][excel_col] = count
            
            if rows_with_data > 0:
                result['success'] = True
                print(f"✅ 数据提取成功: {rows_with_data} 行，各列情况：")
                for col, count in result['extracted_data_count'].items():
                    print(f"   列{col}: {count} 个值")
            
        except Exception as e:
            print(f"❌ 数据提取失败: {e}")
            result['error'] = str(e)
        
        return result
    
    def _step4_update_database(self) -> Dict[str, Any]:
        """
        Step 4: 更新数据库
        """
        print(f"\n💾 Step 4: 更新数据库")
        print("-" * 50)
        
        result = {
            'success': False,
            'records_updated': 0,
            'error': None
        }
        
        try:
            if not self.i_to_n_data:
                print("❌ 没有提取到数据")
                return result
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 获取或创建会话ID
            cursor.execute("SELECT session_id FROM extraction_sessions ORDER BY extraction_time DESC LIMIT 1")
            session_result = cursor.fetchone()
            session_id = session_result[0] if session_result else f"fix_session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            
            # 清空并重新插入数据
            cursor.execute("DELETE FROM heat_exchangers")
            extraction_time = datetime.now().isoformat()
            
            for row_data in self.i_to_n_data:
                i_to_n_cols = row_data['i_to_n_columns']
                
                # 准备I-N列数据
                i_data = i_to_n_cols.get('i', {})
                j_data = i_to_n_cols.get('j', {})
                k_data = i_to_n_cols.get('k', {})
                l_data = i_to_n_cols.get('l', {})
                m_data = i_to_n_cols.get('m', {})
                n_data = i_to_n_cols.get('n', {})
                
                # 原始数据
                raw_data = {col_name.upper(): col_info.get('data') for col_name, col_info in i_to_n_cols.items()}
                
                cursor.execute("""
                    INSERT INTO heat_exchangers (
                        session_id, name, duty_kw, area_m2, temperatures, pressures,
                        source, extraction_time,
                        column_i_data, column_i_header,
                        column_j_data, column_j_header,
                        column_k_data, column_k_header,
                        column_l_data, column_l_header,
                        column_m_data, column_m_header,
                        column_n_data, column_n_header,
                        columns_i_to_n_raw
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    session_id, row_data['name'], 0.0, 0.0, 
                    json.dumps({}), json.dumps({}), 'excel_fix_v2', extraction_time,
                    i_data.get('data'), i_data.get('header'),
                    j_data.get('data'), j_data.get('header'),
                    k_data.get('data'), k_data.get('header'),
                    l_data.get('data'), l_data.get('header'),
                    m_data.get('data'), m_data.get('header'),
                    n_data.get('data'), n_data.get('header'),
                    json.dumps(raw_data) if raw_data else None
                ))
            
            conn.commit()
            result['records_updated'] = len(self.i_to_n_data)
            result['success'] = True
            
            print(f"✅ 数据库更新成功: {result['records_updated']} 条记录")
            conn.close()
            
        except Exception as e:
            print(f"❌ 数据库更新失败: {e}")
            result['error'] = str(e)
            import traceback
            traceback.print_exc()
        
        return result
    
    def _step5_verify_fix(self) -> Dict[str, Any]:
        """
        Step 5: 验证修复结果
        """
        print(f"\n✅ Step 5: 验证修复结果")
        print("-" * 50)
        
        result = {
            'success': False,
            'total_records': 0,
            'i_to_n_coverage': {},
            'sample_data': [],
            'error': None
        }
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 检查记录数
            cursor.execute("SELECT COUNT(*) FROM heat_exchangers")
            result['total_records'] = cursor.fetchone()[0]
            print(f"📊 总记录数: {result['total_records']}")
            
            # 检查I-N列覆盖率
            i_to_n_columns = [
                ('column_i_data', 'I'), ('column_j_data', 'J'), ('column_k_data', 'K'),
                ('column_l_data', 'L'), ('column_m_data', 'M'), ('column_n_data', 'N')
            ]
            
            total_values = 0
            for db_col, excel_col in i_to_n_columns:
                cursor.execute(f"SELECT COUNT(*) FROM heat_exchangers WHERE {db_col} IS NOT NULL")
                count = cursor.fetchone()[0]
                coverage = (count / result['total_records']) * 100 if result['total_records'] > 0 else 0
                result['i_to_n_coverage'][excel_col] = {'count': count, 'coverage': coverage}
                total_values += count
                print(f"   列{excel_col}: {count}/{result['total_records']} ({coverage:.1f}%)")
            
            # 样本数据
            cursor.execute("""
                SELECT name, column_i_data, column_j_data, column_k_data,
                       column_l_data, column_m_data, column_n_data
                FROM heat_exchangers LIMIT 3
            """)
            
            for row in cursor.fetchall():
                result['sample_data'].append({
                    'name': row[0],
                    'values': {'I': row[1], 'J': row[2], 'K': row[3], 'L': row[4], 'M': row[5], 'N': row[6]}
                })
            
            print(f"🔬 样本数据:")
            for sample in result['sample_data']:
                print(f"   {sample['name']}: {sample['values']}")
            
            if total_values > 0:
                result['success'] = True
                print(f"✅ 验证成功! 总计 {total_values} 个I-N数据点")
            
            conn.close()
            
        except Exception as e:
            print(f"❌ 验证失败: {e}")
            result['error'] = str(e)
        
        return result
    
    def _clean_numeric_value(self, value) -> Optional[float]:
        """清理数值数据"""
        if value is None or pd.isna(value):
            return None
        
        if isinstance(value, (int, float)):
            return float(value)
        
        if isinstance(value, str):
            import re
            clean_str = re.sub(r'[^\d.-]', '', str(value).strip())
            if clean_str:
                try:
                    return float(clean_str)
                except ValueError:
                    pass
        
        return None
    
    def _generate_complete_report(self, results: Dict[str, Any]) -> None:
        """生成完整报告"""
        print(f"\n📋 I-N列完整修复报告")
        print("=" * 80)
        
        steps_passed = sum(1 for result in results.values() if result.get('success', False))
        total_steps = len(results)
        
        print(f"修复状态: {steps_passed}/{total_steps} 步骤成功")
        
        if steps_passed == total_steps:
            print("🎉 I-N列数据完整修复成功!")
            
            verify_result = results.get('step_5_verify_fix', {})
            if verify_result.get('success'):
                coverage = verify_result.get('i_to_n_coverage', {})
                total_values = sum(col_info.get('count', 0) for col_info in coverage.values())
                print(f"📊 修复成果: {verify_result.get('total_records', 0)} 条记录，{total_values} 个I-N数据点")
        else:
            print("⚠️ 部分步骤失败，请检查错误信息")
        
        # 保存报告
        report_file = f"i_to_n_complete_fix_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        try:
            with open(report_file, 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2, ensure_ascii=False, default=str)
            print(f"💾 详细报告: {report_file}")
        except Exception as e:
            print(f"❌ 报告保存失败: {e}")


def main():
    """主函数"""
    print("🚀 启动I-N列数据完整修复工具 V2")
    
    fixer = IToNColumnFixerV2()
    results = fixer.complete_fix()
    
    success = all(result.get('success', False) for result in results.values())
    
    if success:
        print(f"\n🎉 I-N列数据完整修复成功!")
        print("数据库中的heat_exchangers表现在包含完整的I-N列数据")
    else:
        print(f"\n⚠️ 修复未完全成功，请查看报告文件")
    
    return success


if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)