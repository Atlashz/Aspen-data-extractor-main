#!/usr/bin/env python3
"""
流股映射查询工具

用于查询和使用数据库中的流股名称映射关系
"""

import sqlite3
import pandas as pd
from typing import Dict, List, Optional, Tuple
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class StreamMappingQuery:
    """流股映射查询器"""
    
    def __init__(self, db_path: str = "aspen_data.db"):
        self.db_path = db_path
    
    def get_all_mappings(self, table_name: str = "improved_stream_mappings") -> List[Tuple]:
        """获取所有映射关系"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute(f"""
                SELECT database_name, aspen_name, confidence, explanation, mapping_type 
                FROM {table_name}
                ORDER BY confidence DESC, database_name
            """)
            
            mappings = cursor.fetchall()
            conn.close()
            
            return mappings
            
        except Exception as e:
            logger.error(f"查询映射时出错: {e}")
            return []
    
    def get_mapping_by_db_name(self, db_name: str, table_name: str = "improved_stream_mappings") -> Optional[Tuple]:
        """根据数据库名称查询映射"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute(f"""
                SELECT database_name, aspen_name, confidence, explanation 
                FROM {table_name}
                WHERE database_name = ?
            """, (db_name,))
            
            result = cursor.fetchone()
            conn.close()
            
            return result
            
        except Exception as e:
            logger.error(f"查询映射时出错: {e}")
            return None
    
    def get_mapping_by_aspen_name(self, aspen_name: str, table_name: str = "improved_stream_mappings") -> List[Tuple]:
        """根据Aspen名称查询映射（可能有多个数据库名称对应同一个Aspen名称）"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute(f"""
                SELECT database_name, aspen_name, confidence, explanation 
                FROM {table_name}
                WHERE aspen_name = ?
                ORDER BY confidence DESC
            """, (aspen_name,))
            
            results = cursor.fetchall()
            conn.close()
            
            return results
            
        except Exception as e:
            logger.error(f"查询映射时出错: {e}")
            return []
    
    def get_mapping_dict(self, min_confidence: float = 0.0, table_name: str = "improved_stream_mappings") -> Dict[str, str]:
        """获取映射字典 (数据库名 -> Aspen名)"""
        mappings = {}
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute(f"""
                SELECT database_name, aspen_name, confidence 
                FROM {table_name}
                WHERE confidence >= ?
                ORDER BY database_name
            """, (min_confidence,))
            
            for db_name, aspen_name, confidence in cursor.fetchall():
                mappings[db_name] = aspen_name
            
            conn.close()
            
        except Exception as e:
            logger.error(f"获取映射字典时出错: {e}")
        
        return mappings
    
    def get_reverse_mapping_dict(self, min_confidence: float = 0.0, table_name: str = "improved_stream_mappings") -> Dict[str, List[str]]:
        """获取反向映射字典 (Aspen名 -> [数据库名列表])"""
        mappings = {}
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute(f"""
                SELECT database_name, aspen_name, confidence 
                FROM {table_name}
                WHERE confidence >= ?
                ORDER BY aspen_name, confidence DESC
            """, (min_confidence,))
            
            for db_name, aspen_name, confidence in cursor.fetchall():
                if aspen_name not in mappings:
                    mappings[aspen_name] = []
                mappings[aspen_name].append(db_name)
            
            conn.close()
            
        except Exception as e:
            logger.error(f"获取反向映射字典时出错: {e}")
        
        return mappings
    
    def export_to_excel(self, filename: str = None, table_name: str = "improved_stream_mappings") -> bool:
        """导出映射到Excel文件"""
        if filename is None:
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"stream_mappings_{timestamp}.xlsx"
        
        try:
            conn = sqlite3.connect(self.db_path)
            
            # 查询映射数据
            df = pd.read_sql_query(f"""
                SELECT 
                    database_name as '数据库流股名',
                    aspen_name as 'Aspen流股名',
                    confidence as '置信度',
                    explanation as '映射说明',
                    mapping_type as '映射类型',
                    created_at as '创建时间'
                FROM {table_name}
                ORDER BY confidence DESC, database_name
            """, conn)
            
            conn.close()
            
            # 保存到Excel
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='流股映射', index=False)
                
                # 添加统计信息
                stats_data = {
                    '统计项目': ['总映射数', '高置信度(≥0.85)', '中等置信度(0.75-0.84)', '低置信度(<0.75)'],
                    '数量': [
                        len(df),
                        len(df[df['置信度'] >= 0.85]),
                        len(df[(df['置信度'] >= 0.75) & (df['置信度'] < 0.85)]),
                        len(df[df['置信度'] < 0.75])
                    ]
                }
                stats_df = pd.DataFrame(stats_data)
                stats_df.to_excel(writer, sheet_name='统计信息', index=False)
            
            logger.info(f"✅ 映射数据已导出到: {filename}")
            return True
            
        except Exception as e:
            logger.error(f"导出Excel时出错: {e}")
            return False
    
    def print_mapping_summary(self, table_name: str = "improved_stream_mappings"):
        """打印映射摘要"""
        mappings = self.get_all_mappings(table_name)
        
        if not mappings:
            print("❌ 没有找到映射数据")
            return
        
        print("\n" + "="*80)
        print(f"📋 流股映射摘要 (表: {table_name})")
        print("="*80)
        
        # 按置信度分组
        high_conf = [m for m in mappings if m[2] >= 0.85]
        medium_conf = [m for m in mappings if 0.75 <= m[2] < 0.85]
        low_conf = [m for m in mappings if m[2] < 0.75]
        
        print(f"\n📊 映射统计:")
        print(f"  • 总映射数: {len(mappings)}")
        print(f"  • 高置信度 (≥0.85): {len(high_conf)} ({len(high_conf)/len(mappings)*100:.1f}%)")
        print(f"  • 中等置信度 (0.75-0.84): {len(medium_conf)} ({len(medium_conf)/len(mappings)*100:.1f}%)")
        print(f"  • 低置信度 (<0.75): {len(low_conf)} ({len(low_conf)/len(mappings)*100:.1f}%)")
        
        print(f"\n🟢 高置信度映射 ({len(high_conf)} 个):")
        print("-" * 70)
        for db_name, aspen_name, confidence, explanation, _ in high_conf:
            print(f"  {db_name:25} → {aspen_name:15} ({confidence:.2f})")
        
        if medium_conf:
            print(f"\n🟡 中等置信度映射 ({len(medium_conf)} 个):")
            print("-" * 70)
            for db_name, aspen_name, confidence, explanation, _ in medium_conf:
                print(f"  {db_name:25} → {aspen_name:15} ({confidence:.2f})")
        
        if low_conf:
            print(f"\n🔴 低置信度映射 ({len(low_conf)} 个):")
            print("-" * 70)
            for db_name, aspen_name, confidence, explanation, _ in low_conf:
                print(f"  {db_name:25} → {aspen_name:15} ({confidence:.2f})")
    
    def search_mapping(self, keyword: str, table_name: str = "improved_stream_mappings") -> List[Tuple]:
        """搜索包含关键词的映射"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute(f"""
                SELECT database_name, aspen_name, confidence, explanation 
                FROM {table_name}
                WHERE database_name LIKE ? OR aspen_name LIKE ? OR explanation LIKE ?
                ORDER BY confidence DESC
            """, (f'%{keyword}%', f'%{keyword}%', f'%{keyword}%'))
            
            results = cursor.fetchall()
            conn.close()
            
            return results
            
        except Exception as e:
            logger.error(f"搜索映射时出错: {e}")
            return []

def main():
    """主函数 - 演示查询功能"""
    print("🔍 流股映射查询工具")
    print("="*50)
    
    query = StreamMappingQuery()
    
    # 显示映射摘要
    query.print_mapping_summary()
    
    # 获取高置信度映射字典
    print("\n" + "="*50)
    print("📖 高置信度映射字典 (置信度 ≥ 0.85):")
    print("="*50)
    high_conf_dict = query.get_mapping_dict(min_confidence=0.85)
    for db_name, aspen_name in high_conf_dict.items():
        print(f"  '{db_name}' → '{aspen_name}'")
    
    # 展示一些查询示例
    print("\n" + "="*50)
    print("🔍 查询示例:")
    print("="*50)
    
    # 查询特定的数据库流股
    test_db_name = "BFG-FEED"
    result = query.get_mapping_by_db_name(test_db_name)
    if result:
        print(f"✅ 查询 '{test_db_name}': {result[0]} → {result[1]} (置信度: {result[2]:.2f})")
    
    # 搜索甲醇相关映射
    methanol_results = query.search_mapping("甲醇")
    if methanol_results:
        print(f"✅ 搜索 '甲醇' 相关映射:")
        for db_name, aspen_name, confidence, explanation in methanol_results:
            print(f"    {db_name} → {aspen_name} ({confidence:.2f}) - {explanation}")
    
    # 导出Excel
    print("\n💾 导出映射到Excel...")
    if query.export_to_excel():
        print("✅ Excel导出成功")
    else:
        print("❌ Excel导出失败")

if __name__ == "__main__":
    main()
