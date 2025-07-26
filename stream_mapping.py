#!/usr/bin/env python3
"""
流股名称映射工具

用于将数据库中的流股名称与Aspen Plus中的流股名称进行匹配
"""

import sqlite3
import sys
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass
import re
from difflib import SequenceMatcher
import logging

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@dataclass
class StreamMapping:
    """流股映射数据结构"""
    database_name: str
    aspen_name: str
    similarity_score: float
    mapping_reason: str
    confidence: float

class StreamNameMatcher:
    """流股名称匹配器"""
    
    def __init__(self, db_path: str = "aspen_data.db"):
        self.db_path = db_path
        self.database_streams = []
        self.aspen_streams = []
        self.mappings = []
        
    def load_database_streams(self) -> List[str]:
        """从数据库加载流股名称"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 检查表是否存在
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = [row[0] for row in cursor.fetchall()]
            logger.info(f"数据库中的表: {tables}")
            
            # 尝试不同的表名和列名组合
            possible_queries = [
                ("streams", "SELECT stream_name FROM streams ORDER BY stream_name"),
                ("aspen_streams", "SELECT stream_name FROM aspen_streams ORDER BY stream_name"),
                ("aspen_streams", "SELECT name FROM aspen_streams ORDER BY name"),
                ("aspen_streams", "SELECT stream_id FROM aspen_streams ORDER BY stream_id"),
                ("aspen_streams", "SELECT * FROM aspen_streams LIMIT 1")  # 查看结构
            ]
            
            streams = []
            for table_name, query in possible_queries:
                if table_name in tables:
                    try:
                        cursor.execute(query)
                        results = cursor.fetchall()
                        
                        if "SELECT *" in query:
                            # 查看表结构
                            if results:
                                logger.info(f"aspen_streams表第一行数据: {results[0]}")
                                # 获取列名
                                cursor.execute("PRAGMA table_info(aspen_streams)")
                                columns = cursor.fetchall()
                                column_names = [col[1] for col in columns]
                                logger.info(f"列名: {column_names}")
                                
                                # 尝试用第一个可能的名称列
                                name_columns = [col for col in column_names if 'name' in col.lower() or 'id' in col.lower()]
                                if name_columns:
                                    name_col = name_columns[0]
                                    cursor.execute(f"SELECT {name_col} FROM aspen_streams ORDER BY {name_col}")
                                    streams = [row[0] for row in cursor.fetchall()]
                                    logger.info(f"使用列 {name_col} 找到 {len(streams)} 个流股")
                                    break
                        else:
                            streams = [row[0] for row in results]
                            logger.info(f"查询成功，找到 {len(streams)} 个流股")
                            break
                            
                    except Exception as e:
                        logger.debug(f"查询失败 {query}: {e}")
                        continue
            
            conn.close()
            self.database_streams = streams
            return streams
            
        except Exception as e:
            logger.error(f"加载数据库流股时出错: {e}")
            return []
    
    def load_aspen_streams(self) -> List[str]:
        """从最近的Aspen提取中加载流股名称"""
        # 这些是我们在测试中看到的Aspen流股名称
        aspen_streams = [
            'AF-COM', 'AIR', 'BFG', 'CS1', 'FLUEGAS1', 'GASOUT1', 'H2IN',
            'LIGHTEND', 'MEOH1', 'MEOH2', 'MEOH3', 'MEOH4', 'MEOH5', 'MEOH6',
            'MEOH7', 'P1', 'PUR2', 'PUR3', 'PUR4', 'REF4', 'REF6', 'S1', 'SS2', 'SS3'
        ]
        
        self.aspen_streams = aspen_streams
        logger.info(f"加载了 {len(aspen_streams)} 个Aspen流股")
        return aspen_streams
    
    def calculate_similarity(self, str1: str, str2: str) -> float:
        """计算两个字符串的相似度"""
        return SequenceMatcher(None, str1.lower(), str2.lower()).ratio()
    
    def find_keyword_matches(self, db_name: str, aspen_name: str) -> Tuple[bool, str]:
        """基于关键词查找匹配"""
        db_lower = db_name.lower()
        aspen_lower = aspen_name.lower()
        
        # 定义关键词映射规则
        keyword_mappings = {
            # 高炉煤气相关
            'bfg': ['bfg', 'blast', 'furnace'],
            'feed': ['feed', 'input', 'in'],
            'product': ['product', 'output', 'out'],
            'methanol': ['meoh', 'methanol', 'ch3oh'],
            'water': ['water', 'h2o', 'cooling'],
            'steam': ['steam', 'vapor', 'hp', 'lp', 'mp'],
            'air': ['air'],
            'hydrogen': ['h2', 'hydrogen'],
            'purge': ['pur', 'purge'],
            'recycle': ['recycle', 'rec'],
            'flash': ['flash', 'separator'],
            'condenser': ['condenser', 'cond'],
            'makeup': ['makeup', 'make-up'],
            'light': ['light', 'lightend'],
            'reference': ['ref', 'reference'],
            'gas': ['gas', 'gasout', 'flue'],
            'liquid': ['liquid', 'liq'],
            'co2': ['co2', 'carbon', 'dioxide'],
            'reactor': ['rxn', 'reactor', 'reaction'],
            'distillation': ['t-', 'tower', 'distil'],
            'ss': ['ss', 'stainless']
        }
        
        # 检查关键词匹配
        for category, keywords in keyword_mappings.items():
            db_match = any(kw in db_lower for kw in keywords)
            aspen_match = any(kw in aspen_lower for kw in keywords)
            
            if db_match and aspen_match:
                return True, f"关键词匹配: {category}"
        
        # 特殊规则匹配
        special_rules = [
            # 高炉煤气
            (('bfg' in db_lower and 'feed' in db_lower), ('bfg' in aspen_lower), "BFG原料匹配"),
            # 二氧化碳
            (('co2' in db_lower and 'feed' in db_lower), ('co2' in aspen_lower or 'af-com' in aspen_lower), "CO2原料匹配"),
            # 甲醇产品
            (('methanol' in db_lower and 'product' in db_lower), ('meoh' in aspen_lower), "甲醇产品匹配"),
            # 水产品
            (('water' in db_lower and 'product' in db_lower), ('h2o' in aspen_lower or 'water' in aspen_lower), "水产品匹配"),
            # 冷却水
            (('cooling' in db_lower and 'water' in db_lower), ('cooling' in aspen_lower or 'water' in aspen_lower), "冷却水匹配"),
            # 蒸汽
            (('steam' in db_lower), ('steam' in aspen_lower or any(x in aspen_lower for x in ['hp', 'lp', 'mp'])), "蒸汽匹配"),
            # 氢气
            (('h2' in db_lower and 'makeup' in db_lower), ('h2' in aspen_lower), "氢气补充匹配"),
            # 吹扫气
            (('purge' in db_lower), ('pur' in aspen_lower), "吹扫气匹配"),
            # 循环气
            (('recycle' in db_lower), ('rec' in aspen_lower or 'ref' in aspen_lower), "循环气匹配"),
            # 反应器
            (('rxn' in db_lower or 'reactor' in db_lower), ('rxn' in aspen_lower or 'reactor' in aspen_lower), "反应器匹配"),
            # 分离器/闪蒸
            (('flash' in db_lower), ('flash' in aspen_lower or 'lightend' in aspen_lower), "分离器匹配"),
            # 冷凝器
            (('condenser' in db_lower), ('condenser' in aspen_lower or 'cs' in aspen_lower), "冷凝器匹配"),
        ]
        
        for db_condition, aspen_condition, reason in special_rules:
            if db_condition and aspen_condition:
                return True, reason
        
        return False, "无关键词匹配"
    
    def create_stream_mappings(self) -> List[StreamMapping]:
        """创建流股映射"""
        mappings = []
        
        if not self.database_streams or not self.aspen_streams:
            logger.warning("流股数据为空，无法创建映射")
            return mappings
        
        # 为每个数据库流股找到最佳匹配
        for db_stream in self.database_streams:
            best_match = None
            best_score = 0.0
            best_reason = ""
            
            for aspen_stream in self.aspen_streams:
                # 计算字符串相似度
                similarity = self.calculate_similarity(db_stream, aspen_stream)
                
                # 检查关键词匹配
                keyword_match, keyword_reason = self.find_keyword_matches(db_stream, aspen_stream)
                
                # 综合评分
                score = similarity
                reason = f"字符串相似度: {similarity:.2f}"
                
                if keyword_match:
                    score += 0.3  # 关键词匹配加分
                    reason += f", {keyword_reason}"
                
                # 精确匹配加分
                if db_stream.lower() == aspen_stream.lower():
                    score = 1.0
                    reason = "精确匹配"
                
                if score > best_score:
                    best_score = score
                    best_match = aspen_stream
                    best_reason = reason
            
            # 只保留置信度较高的匹配
            if best_score > 0.3:  # 阈值
                confidence = min(best_score, 1.0)
                mapping = StreamMapping(
                    database_name=db_stream,
                    aspen_name=best_match,
                    similarity_score=best_score,
                    mapping_reason=best_reason,
                    confidence=confidence
                )
                mappings.append(mapping)
        
        self.mappings = mappings
        return mappings
    
    def print_mappings(self):
        """打印映射结果"""
        if not self.mappings:
            print("❌ 没有找到匹配的流股")
            return
        
        print("\n" + "="*80)
        print("🔗 流股名称映射结果")
        print("="*80)
        
        # 按置信度排序
        sorted_mappings = sorted(self.mappings, key=lambda x: x.confidence, reverse=True)
        
        high_confidence = [m for m in sorted_mappings if m.confidence >= 0.8]
        medium_confidence = [m for m in sorted_mappings if 0.5 <= m.confidence < 0.8]
        low_confidence = [m for m in sorted_mappings if m.confidence < 0.5]
        
        if high_confidence:
            print(f"\n🟢 高置信度映射 ({len(high_confidence)} 个):")
            print("-" * 60)
            for mapping in high_confidence:
                print(f"  {mapping.database_name:20} → {mapping.aspen_name:15} "
                     f"(置信度: {mapping.confidence:.2f})")
                print(f"    📋 {mapping.mapping_reason}")
        
        if medium_confidence:
            print(f"\n🟡 中等置信度映射 ({len(medium_confidence)} 个):")
            print("-" * 60)
            for mapping in medium_confidence:
                print(f"  {mapping.database_name:20} → {mapping.aspen_name:15} "
                     f"(置信度: {mapping.confidence:.2f})")
                print(f"    📋 {mapping.mapping_reason}")
        
        if low_confidence:
            print(f"\n🔴 低置信度映射 ({len(low_confidence)} 个):")
            print("-" * 60)
            for mapping in low_confidence:
                print(f"  {mapping.database_name:20} → {mapping.aspen_name:15} "
                     f"(置信度: {mapping.confidence:.2f})")
                print(f"    📋 {mapping.mapping_reason}")
        
        print(f"\n📊 映射统计:")
        print(f"  • 数据库流股: {len(self.database_streams)}")
        print(f"  • Aspen流股: {len(self.aspen_streams)}")
        print(f"  • 成功映射: {len(self.mappings)}")
        print(f"  • 映射率: {len(self.mappings)/len(self.database_streams)*100:.1f}%")
    
    def save_mappings_to_database(self) -> bool:
        """将映射结果保存到数据库"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 创建映射表
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS stream_mappings (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    database_name TEXT NOT NULL,
                    aspen_name TEXT NOT NULL,
                    similarity_score REAL,
                    mapping_reason TEXT,
                    confidence REAL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(database_name, aspen_name)
                )
            """)
            
            # 清除旧映射
            cursor.execute("DELETE FROM stream_mappings")
            
            # 插入新映射
            for mapping in self.mappings:
                cursor.execute("""
                    INSERT INTO stream_mappings 
                    (database_name, aspen_name, similarity_score, mapping_reason, confidence)
                    VALUES (?, ?, ?, ?, ?)
                """, (
                    mapping.database_name,
                    mapping.aspen_name,
                    mapping.similarity_score,
                    mapping.mapping_reason,
                    mapping.confidence
                ))
            
            conn.commit()
            conn.close()
            
            logger.info(f"✅ 成功保存 {len(self.mappings)} 个映射到数据库")
            return True
            
        except Exception as e:
            logger.error(f"保存映射到数据库时出错: {e}")
            return False
    
    def get_mapping_dict(self) -> Dict[str, str]:
        """获取映射字典 (数据库名 -> Aspen名)"""
        return {m.database_name: m.aspen_name for m in self.mappings}

def main():
    """主函数"""
    print("🚀 流股名称映射工具")
    print("="*50)
    
    # 创建匹配器
    matcher = StreamNameMatcher()
    
    # 加载数据
    print("📥 加载流股数据...")
    db_streams = matcher.load_database_streams()
    aspen_streams = matcher.load_aspen_streams()
    
    if not db_streams:
        print("❌ 无法从数据库加载流股数据")
        return
    
    print(f"📋 数据库流股 ({len(db_streams)} 个):")
    for i, stream in enumerate(db_streams, 1):
        print(f"  {i:2d}. {stream}")
    
    print(f"\n📋 Aspen流股 ({len(aspen_streams)} 个):")
    for i, stream in enumerate(aspen_streams, 1):
        print(f"  {i:2d}. {stream}")
    
    # 创建映射
    print("\n🔄 创建流股映射...")
    mappings = matcher.create_stream_mappings()
    
    # 显示结果
    matcher.print_mappings()
    
    # 保存到数据库
    print("\n💾 保存映射到数据库...")
    if matcher.save_mappings_to_database():
        print("✅ 映射保存成功")
    else:
        print("❌ 映射保存失败")

if __name__ == "__main__":
    main()
