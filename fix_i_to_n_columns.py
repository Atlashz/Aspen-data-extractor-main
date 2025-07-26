#!/usr/bin/env python3
"""
I-N Column Data Fix Script

专门修复BFG-CO2H-HEX.xlsx中I-N列数据提取和存储问题
强制重新提取并填充数据库中的I-N列数据

Author: TEA Analysis Framework  
Date: 2025-07-26
Version: 1.0 - I-N Column Fix
"""

import pandas as pd
import sqlite3
import json
import logging
from datetime import datetime
from typing import Dict, List, Any, Optional

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(levelname)s:%(name)s:%(message)s')
logger = logging.getLogger(__name__)

class IToNColumnFixer:
    """
    专门用于修复I-N列数据提取问题的类
    """
    
    def __init__(self, excel_file: str = "BFG-CO2H-HEX.xlsx", db_path: str = "aspen_data.db"):
        self.excel_file = excel_file
        self.db_path = db_path
        self.df = None
        self.i_to_n_data = {}
        
    def diagnose_and_fix(self) -> Dict[str, Any]:
        """
        诊断并修复I-N列数据问题
        """
        print("\n" + "="*80)
        print("🔧 I-N列数据修复工具")
        print("="*80)
        print(f"Excel文件: {self.excel_file}")
        print(f"数据库: {self.db_path}")
        print(f"修复时间: {datetime.now().isoformat()}")
        
        results = {
            'step_1_excel_analysis': self._step1_analyze_excel(),
            'step_2_extract_i_to_n': self._step2_extract_i_to_n_data(),
            'step_3_update_database': self._step3_update_database(),
            'step_4_verify_fix': self._step4_verify_fix()
        }
        
        self._generate_fix_report(results)
        return results
    
    def _step1_analyze_excel(self) -> Dict[str, Any]:
        """
        Step 1: 分析Excel文件结构，确定I-N列位置
        """
        print("\n🔍 Step 1: 分析Excel文件结构")
        print("-" * 50)
        
        result = {
            'success': False,
            'file_found': False,
            'total_columns': 0,
            'total_rows': 0,
            'i_to_n_columns': {},
            'sample_data': {},
            'error': None
        }
        
        try:
            import os
            if not os.path.exists(self.excel_file):
                print(f"❌ Excel文件不存在: {self.excel_file}")
                return result
            
            result['file_found'] = True
            print(f"✅ Excel文件找到: {self.excel_file}")
            
            # 读取Excel文件
            self.df = pd.read_excel(self.excel_file)
            result['total_columns'] = len(self.df.columns)
            result['total_rows'] = len(self.df)
            
            print(f"📊 文件结构: {result['total_rows']} 行 × {result['total_columns']} 列")
            
            # 强制定位I-N列 (Excel列I=索引8, J=9, K=10, L=11, M=12, N=13)
            i_to_n_mapping = {
                'I': 8, 'J': 9, 'K': 10, 'L': 11, 'M': 12, 'N': 13
            }
            
            print(f"🔍 强制定位I-N列:")
            for excel_col, col_idx in i_to_n_mapping.items():
                if col_idx < len(self.df.columns):
                    header = str(self.df.columns[col_idx])
                    result['i_to_n_columns'][excel_col] = {
                        'index': col_idx,
                        'header': header,
                        'non_null_count': self.df.iloc[:, col_idx].notna().sum(),
                        'data_type': str(self.df.iloc[:, col_idx].dtype)
                    }
                    
                    print(f"   列{excel_col} (索引{col_idx}): '{header}'")
                    print(f"      有效数据: {result['i_to_n_columns'][excel_col]['non_null_count']}/{result['total_rows']}")
                    print(f"      数据类型: {result['i_to_n_columns'][excel_col]['data_type']}")
                    
                    # 采样前3个非空值
                    sample_values = self.df.iloc[:, col_idx].dropna().head(3).tolist()
                    result['sample_data'][excel_col] = sample_values
                    print(f"      样本数据: {sample_values}")
                else:
                    print(f"   列{excel_col}: 不存在 (文件只有{len(self.df.columns)}列)")
            
            if result['i_to_n_columns']:
                result['success'] = True
                print(f"✅ 成功识别 {len(result['i_to_n_columns'])} 个I-N列")
            else:
                print(f"❌ 未找到任何I-N列")
            
        except Exception as e:
            print(f"❌ Excel分析失败: {e}")
            result['error'] = str(e)
        
        return result
    
    def _step2_extract_i_to_n_data(self) -> Dict[str, Any]:
        """
        Step 2: 提取I-N列数据
        """
        print(f"\n📤 Step 2: 提取I-N列数据")
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
                print("❌ Excel数据未加载，无法提取")
                return result
            
            # 强制映射I-N列
            i_to_n_mapping = {
                'I': 8, 'J': 9, 'K': 10, 'L': 11, 'M': 12, 'N': 13
            }
            
            extracted_data = []
            rows_with_data = 0
            
            print(f"🔄 处理 {len(self.df)} 行数据...")
            
            for idx, row in self.df.iterrows():
                row_data = {
                    'row_index': idx,
                    'name': f'HEX-{idx+1:03d}',  # 默认名称
                    'i_to_n_columns': {}
                }
                
                has_i_to_n_data = False
                
                # 提取每个I-N列的数据
                for excel_col, col_idx in i_to_n_mapping.items():
                    if col_idx < len(self.df.columns):
                        header = str(self.df.columns[col_idx])
                        value = row.iloc[col_idx] if col_idx < len(row) else None
                        
                        # 数据清理和转换
                        clean_value = self._clean_numeric_value(value)
                        
                        if clean_value is not None:
                            row_data['i_to_n_columns'][excel_col.lower()] = {
                                'data': clean_value,
                                'header': header,
                                'raw_value': value
                            }
                            has_i_to_n_data = True
                
                if has_i_to_n_data:
                    extracted_data.append(row_data)
                    rows_with_data += 1
            
            self.i_to_n_data = extracted_data
            result['total_rows_processed'] = len(self.df)
            result['rows_with_i_to_n_data'] = rows_with_data
            
            # 统计每列提取的数据量
            for excel_col in ['I', 'J', 'K', 'L', 'M', 'N']:
                count = sum(1 for row in extracted_data if excel_col.lower() in row['i_to_n_columns'])
                result['extracted_data_count'][excel_col] = count
            
            if rows_with_data > 0:
                result['success'] = True
                print(f"✅ 数据提取成功:")
                print(f"   总行数: {result['total_rows_processed']}")
                print(f"   有I-N数据的行: {rows_with_data}")
                print(f"   各列提取情况:")
                for col, count in result['extracted_data_count'].items():
                    print(f"      列{col}: {count} 个值")
            else:
                print(f"❌ 未提取到任何I-N列数据")
            
        except Exception as e:
            print(f"❌ 数据提取失败: {e}")
            result['error'] = str(e)
            import traceback
            traceback.print_exc()
        
        return result
    
    def _step3_update_database(self) -> Dict[str, Any]:
        """
        Step 3: 更新数据库中的I-N列数据
        """
        print(f"\n💾 Step 3: 更新数据库I-N列数据")
        print("-" * 50)
        
        result = {
            'success': False,
            'records_updated': 0,
            'database_connected': False,
            'backup_created': False,
            'error': None
        }
        
        try:
            if not self.i_to_n_data:
                print("❌ 没有提取到I-N数据，无法更新数据库")
                return result
            
            # 连接数据库
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            result['database_connected'] = True
            print(f"✅ 数据库连接成功: {self.db_path}")
            
            # 创建备份
            backup_time = datetime.now().strftime('%Y%m%d_%H%M%S')
            cursor.execute(f"""
                CREATE TABLE IF NOT EXISTS heat_exchangers_backup_{backup_time} AS 
                SELECT * FROM heat_exchangers
            """)
            result['backup_created'] = True
            print(f"✅ 数据备份创建: heat_exchangers_backup_{backup_time}")
            
            # 清空现有的I-N列数据并重新插入
            cursor.execute("DELETE FROM heat_exchangers")
            print(f"🗑️ 清空原有heat_exchangers数据")
            
            # 获取当前会话ID
            cursor.execute("SELECT session_id FROM extraction_sessions ORDER BY extraction_time DESC LIMIT 1")
            session_result = cursor.fetchone()
            session_id = session_result[0] if session_result else f"fix_session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            
            # 插入带有I-N列数据的记录
            records_inserted = 0
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
                
                # 创建原始数据字典
                raw_i_to_n = {
                    col_name.upper(): col_info.get('data')
                    for col_name, col_info in i_to_n_cols.items()
                    if col_info.get('data') is not None
                }
                
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
                    session_id,
                    row_data['name'],
                    0.0,  # 默认duty
                    0.0,  # 默认area
                    json.dumps({}),  # 默认temperatures
                    json.dumps({}),  # 默认pressures
                    'excel_fix',
                    extraction_time,
                    i_data.get('data'),
                    i_data.get('header'),
                    j_data.get('data'),
                    j_data.get('header'),
                    k_data.get('data'),
                    k_data.get('header'),
                    l_data.get('data'),
                    l_data.get('header'),
                    m_data.get('data'),
                    m_data.get('header'),
                    n_data.get('data'),
                    n_data.get('header'),
                    json.dumps(raw_i_to_n) if raw_i_to_n else None
                ))
                
                records_inserted += 1
            
            conn.commit()
            result['records_updated'] = records_inserted
            result['success'] = True
            
            print(f"✅ 数据库更新成功:")
            print(f"   插入记录数: {records_inserted}")
            print(f"   会话ID: {session_id}")
            
            conn.close()
            
        except Exception as e:
            print(f"❌ 数据库更新失败: {e}")
            result['error'] = str(e)
            import traceback
            traceback.print_exc()
        
        return result
    
    def _step4_verify_fix(self) -> Dict[str, Any]:
        """
        Step 4: 验证修复结果
        """
        print(f"\n✅ Step 4: 验证修复结果")
        print("-" * 50)
        
        result = {
            'success': False,
            'total_records': 0,
            'i_to_n_coverage': {},
            'sample_verification': [],
            'error': None
        }
        
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 检查总记录数
            cursor.execute("SELECT COUNT(*) FROM heat_exchangers")
            result['total_records'] = cursor.fetchone()[0]
            print(f"📊 heat_exchangers表总记录数: {result['total_records']}")
            
            # 检查I-N列覆盖率
            i_to_n_columns = [
                ('column_i_data', 'I'),
                ('column_j_data', 'J'),
                ('column_k_data', 'K'),
                ('column_l_data', 'L'),
                ('column_m_data', 'M'),
                ('column_n_data', 'N')
            ]
            
            print(f"🔍 I-N列数据覆盖率:")
            total_i_to_n_values = 0
            
            for db_col, excel_col in i_to_n_columns:
                cursor.execute(f"SELECT COUNT(*) FROM heat_exchangers WHERE {db_col} IS NOT NULL")
                count = cursor.fetchone()[0]
                coverage_pct = (count / result['total_records']) * 100 if result['total_records'] > 0 else 0
                
                result['i_to_n_coverage'][excel_col] = {
                    'count': count,
                    'coverage_percentage': coverage_pct
                }
                total_i_to_n_values += count
                
                print(f"   列{excel_col}: {count}/{result['total_records']} ({coverage_pct:.1f}%)")
            
            # 抽样验证
            cursor.execute("""
                SELECT name, column_i_data, column_j_data, column_k_data, 
                       column_l_data, column_m_data, column_n_data
                FROM heat_exchangers 
                WHERE column_i_data IS NOT NULL OR column_j_data IS NOT NULL 
                   OR column_k_data IS NOT NULL OR column_l_data IS NOT NULL
                   OR column_m_data IS NOT NULL OR column_n_data IS NOT NULL
                LIMIT 3
            """)
            
            sample_rows = cursor.fetchall()
            for row in sample_rows:
                sample_item = {
                    'name': row[0],
                    'i_to_n_values': {
                        'I': row[1], 'J': row[2], 'K': row[3],
                        'L': row[4], 'M': row[5], 'N': row[6]
                    }
                }
                result['sample_verification'].append(sample_item)
            
            print(f"🔬 样本验证 (前3条记录):")
            for sample in result['sample_verification']:
                print(f"   {sample['name']}: {sample['i_to_n_values']}")
            
            # 判断修复是否成功
            if total_i_to_n_values > 0:
                result['success'] = True
                print(f"✅ 修复验证成功!")
                print(f"   总I-N数据点: {total_i_to_n_values}")
                print(f"   平均每行I-N数据: {total_i_to_n_values/result['total_records']:.1f}")
            else:
                print(f"❌ 修复验证失败 - 仍无I-N列数据")
            
            conn.close()
            
        except Exception as e:
            print(f"❌ 验证失败: {e}")
            result['error'] = str(e)
            import traceback
            traceback.print_exc()
        
        return result
    
    def _clean_numeric_value(self, value) -> Optional[float]:
        """
        清理和转换数值数据
        """
        if value is None or pd.isna(value):
            return None
        
        if isinstance(value, (int, float)):
            return float(value)
        
        if isinstance(value, str):
            # 清理字符串中的非数字字符
            import re
            clean_str = re.sub(r'[^\d.-]', '', str(value).strip())
            if clean_str:
                try:
                    return float(clean_str)
                except ValueError:
                    pass
        
        return None
    
    def _generate_fix_report(self, results: Dict[str, Any]) -> None:
        """
        生成修复报告
        """
        print(f"\n📋 I-N列修复报告")
        print("=" * 80)
        
        # 修复状态概览
        steps_passed = sum(1 for step_result in results.values() if step_result.get('success', False))
        total_steps = len(results)
        
        print(f"修复状态: {steps_passed}/{total_steps} 步骤成功")
        
        if steps_passed == total_steps:
            print("🎉 I-N列数据修复完全成功!")
            
            # 显示关键指标
            extract_result = results.get('step_2_extract_i_to_n', {})
            verify_result = results.get('step_4_verify_fix', {})
            
            if extract_result.get('success') and verify_result.get('success'):
                print(f"\n📊 修复成果:")
                print(f"   Excel行数: {extract_result.get('total_rows_processed', 0)}")
                print(f"   有效数据行: {extract_result.get('rows_with_i_to_n_data', 0)}")
                print(f"   数据库记录: {verify_result.get('total_records', 0)}")
                
                coverage = verify_result.get('i_to_n_coverage', {})
                total_values = sum(col_info.get('count', 0) for col_info in coverage.values())
                print(f"   I-N数据点总数: {total_values}")
        else:
            print("⚠️ I-N列数据修复部分成功，需要进一步检查")
            
            # 显示失败的步骤
            for step_name, step_result in results.items():
                if not step_result.get('success', False):
                    error_msg = step_result.get('error', '未知错误')
                    print(f"   {step_name}: 失败 ({error_msg})")
        
        # 保存详细报告
        report_file = f"i_to_n_fix_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        try:
            with open(report_file, 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2, ensure_ascii=False, default=str)
            print(f"\n💾 详细报告已保存: {report_file}")
        except Exception as e:
            print(f"\n❌ 报告保存失败: {e}")


def main():
    """
    主修复函数
    """
    print("🚀 启动I-N列数据修复工具")
    
    fixer = IToNColumnFixer()
    results = fixer.diagnose_and_fix()
    
    # 最终状态
    success = all(step_result.get('success', False) for step_result in results.values())
    
    if success:
        print(f"\n🎉 I-N列数据修复完成！")
        print("现在@aspen_data.db中的heat_exchangers表应该包含完整的I-N列数据")
    else:
        print(f"\n⚠️ I-N列数据修复未完全成功，请检查报告文件")
    
    return success


if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)