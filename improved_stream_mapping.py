#!/usr/bin/env python3
"""
改进的流股映射工具

基于分析结果手动优化映射关系
"""

import sqlite3
from typing import Dict, List
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ImprovedStreamMapper:
    """改进的流股映射器"""
    
    def __init__(self, db_path: str = "aspen_data.db"):
        self.db_path = db_path
        
        # 手动定义高质量映射关系
        self.manual_mappings = {
            # 高炉煤气相关
            'BFG-FEED': 'BFG',  # 高炉煤气原料
            
            # 甲醇相关
            'METHANOL-PRODUCT': 'MEOH1',  # 甲醇产品 - 选择主要的甲醇流股
            
            # 氢气相关
            'H2-MAKEUP': 'H2IN',  # 氢气补充
            
            # 二氧化碳相关
            'CO2-FEED': 'AF-COM',  # 二氧化碳原料（可能是气化副产物）
            
            # 分离器/闪蒸相关
            'FLASH-LIQUID': 'LIGHTEND',  # 闪蒸液体产物
            'FLASH-VAPOR': 'S1',  # 闪蒸气相 
            
            # 反应器相关
            'RXN-FEED': 'REF4',  # 反应器进料
            'RXN-PRODUCT': 'REF6',  # 反应器产物
            
            # 吹扫和循环
            'PURGE-GAS': 'PUR2',  # 吹扫气
            'RECYCLE-GAS': 'PUR3',  # 循环气 
            
            # 蒸汽系统
            'STEAM-HP': 'SS2',  # 高压蒸汽
            'STEAM-LP': 'SS3',  # 低压蒸汽
            'STEAM-MP': 'CS1',  # 中压蒸汽
            
            # 冷却水
            'COOLING-WATER-IN': 'AIR',  # 冷却介质入口
            'COOLING-WATER-OUT': 'GASOUT1',  # 冷却介质出口
            
            # 其他
            'CONDENSER-OUT': 'GASOUT1',  # 冷凝器出口
            'WATER-PRODUCT': 'FLUEGAS1',  # 水产品
            'T-101-FEED': 'P1'  # 塔进料
        }
        
        # 映射置信度（手动评估）
        self.confidence_scores = {
            'BFG-FEED': 0.95,
            'METHANOL-PRODUCT': 0.90,
            'H2-MAKEUP': 0.85,
            'CO2-FEED': 0.75,
            'FLASH-LIQUID': 0.80,
            'FLASH-VAPOR': 0.75,
            'RXN-FEED': 0.80,
            'RXN-PRODUCT': 0.85,
            'PURGE-GAS': 0.90,
            'RECYCLE-GAS': 0.80,
            'STEAM-HP': 0.85,
            'STEAM-LP': 0.85,
            'STEAM-MP': 0.85,
            'COOLING-WATER-IN': 0.70,
            'COOLING-WATER-OUT': 0.70,
            'CONDENSER-OUT': 0.75,
            'WATER-PRODUCT': 0.70,
            'T-101-FEED': 0.75
        }
        
        # 映射说明
        self.mapping_explanations = {
            'BFG-FEED': '高炉煤气原料直接匹配',
            'METHANOL-PRODUCT': '甲醇产品匹配主要甲醇流股',
            'H2-MAKEUP': '氢气补充流股匹配',
            'CO2-FEED': 'CO2原料可能对应AF-COM流股',
            'FLASH-LIQUID': '闪蒸液体产物匹配轻组分',
            'FLASH-VAPOR': '闪蒸气相匹配工艺流股',
            'RXN-FEED': '反应器进料匹配循环流股',
            'RXN-PRODUCT': '反应器产物匹配产品流股',
            'PURGE-GAS': '吹扫气匹配PUR系列流股',
            'RECYCLE-GAS': '循环气匹配PUR系列流股',
            'STEAM-HP': '高压蒸汽匹配SS系列',
            'STEAM-LP': '低压蒸汽匹配SS系列',
            'STEAM-MP': '中压蒸汽匹配CS系列',
            'COOLING-WATER-IN': '冷却水进口匹配冷却介质',
            'COOLING-WATER-OUT': '冷却水出口匹配气体出口',
            'CONDENSER-OUT': '冷凝器出口匹配气体产物',
            'WATER-PRODUCT': '水产品匹配烟气流股',
            'T-101-FEED': '塔进料匹配工艺中间流股'
        }
    
    def save_improved_mappings(self) -> bool:
        """保存改进的映射到数据库"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 创建改进映射表
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS improved_stream_mappings (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    database_name TEXT NOT NULL UNIQUE,
                    aspen_name TEXT NOT NULL,
                    confidence REAL,
                    explanation TEXT,
                    mapping_type TEXT DEFAULT 'manual',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # 清除旧映射
            cursor.execute("DELETE FROM improved_stream_mappings")
            
            # 插入改进的映射
            for db_name, aspen_name in self.manual_mappings.items():
                confidence = self.confidence_scores.get(db_name, 0.5)
                explanation = self.mapping_explanations.get(db_name, '手动映射')
                
                cursor.execute("""
                    INSERT INTO improved_stream_mappings 
                    (database_name, aspen_name, confidence, explanation, mapping_type)
                    VALUES (?, ?, ?, ?, ?)
                """, (db_name, aspen_name, confidence, explanation, 'manual'))
            
            conn.commit()
            conn.close()
            
            logger.info(f"✅ 成功保存 {len(self.manual_mappings)} 个改进映射到数据库")
            return True
            
        except Exception as e:
            logger.error(f"保存改进映射时出错: {e}")
            return False
    
    def print_improved_mappings(self):
        """打印改进的映射结果"""
        print("\n" + "="*80)
        print("🎯 改进的流股名称映射结果")
        print("="*80)
        
        # 按置信度分组
        high_conf = {k: v for k, v in self.manual_mappings.items() if self.confidence_scores.get(k, 0) >= 0.85}
        medium_conf = {k: v for k, v in self.manual_mappings.items() if 0.75 <= self.confidence_scores.get(k, 0) < 0.85}
        low_conf = {k: v for k, v in self.manual_mappings.items() if self.confidence_scores.get(k, 0) < 0.75}
        
        if high_conf:
            print(f"\n🟢 高置信度映射 ({len(high_conf)} 个):")
            print("-" * 70)
            for db_name, aspen_name in high_conf.items():
                conf = self.confidence_scores.get(db_name, 0)
                explanation = self.mapping_explanations.get(db_name, '')
                print(f"  {db_name:25} → {aspen_name:15} (置信度: {conf:.2f})")
                print(f"    📋 {explanation}")
        
        if medium_conf:
            print(f"\n🟡 中等置信度映射 ({len(medium_conf)} 个):")
            print("-" * 70)
            for db_name, aspen_name in medium_conf.items():
                conf = self.confidence_scores.get(db_name, 0)
                explanation = self.mapping_explanations.get(db_name, '')
                print(f"  {db_name:25} → {aspen_name:15} (置信度: {conf:.2f})")
                print(f"    📋 {explanation}")
        
        if low_conf:
            print(f"\n🔴 低置信度映射 ({len(low_conf)} 个):")
            print("-" * 70)
            for db_name, aspen_name in low_conf.items():
                conf = self.confidence_scores.get(db_name, 0)
                explanation = self.mapping_explanations.get(db_name, '')
                print(f"  {db_name:25} → {aspen_name:15} (置信度: {conf:.2f})")
                print(f"    📋 {explanation}")
        
        print(f"\n📊 改进映射统计:")
        print(f"  • 总映射数: {len(self.manual_mappings)}")
        print(f"  • 高置信度: {len(high_conf)} ({len(high_conf)/len(self.manual_mappings)*100:.1f}%)")
        print(f"  • 中等置信度: {len(medium_conf)} ({len(medium_conf)/len(self.manual_mappings)*100:.1f}%)")
        print(f"  • 低置信度: {len(low_conf)} ({len(low_conf)/len(self.manual_mappings)*100:.1f}%)")
    
    def get_mapping_dict(self) -> Dict[str, str]:
        """获取映射字典"""
        return self.manual_mappings.copy()
    
    def validate_mappings(self) -> Dict[str, List[str]]:
        """验证映射的有效性"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 获取数据库中的流股名称
            cursor.execute("SELECT name FROM aspen_streams")
            db_streams = set(row[0] for row in cursor.fetchall())
            
            conn.close()
            
            # 验证结果
            validation = {
                'valid_mappings': [],
                'invalid_db_names': [],
                'missing_aspen_names': []
            }
            
            # Aspen流股名称
            aspen_streams = {
                'AF-COM', 'AIR', 'BFG', 'CS1', 'FLUEGAS1', 'GASOUT1', 'H2IN',
                'LIGHTEND', 'MEOH1', 'MEOH2', 'MEOH3', 'MEOH4', 'MEOH5', 'MEOH6',
                'MEOH7', 'P1', 'PUR2', 'PUR3', 'PUR4', 'REF4', 'REF6', 'S1', 'SS2', 'SS3'
            }
            
            for db_name, aspen_name in self.manual_mappings.items():
                if db_name not in db_streams:
                    validation['invalid_db_names'].append(db_name)
                elif aspen_name not in aspen_streams:
                    validation['missing_aspen_names'].append(f"{db_name} → {aspen_name}")
                else:
                    validation['valid_mappings'].append(f"{db_name} → {aspen_name}")
            
            return validation
            
        except Exception as e:
            logger.error(f"验证映射时出错: {e}")
            return {'valid_mappings': [], 'invalid_db_names': [], 'missing_aspen_names': []}

def main():
    """主函数"""
    print("🎯 改进的流股名称映射工具")
    print("="*50)
    
    # 创建改进映射器
    mapper = ImprovedStreamMapper()
    
    # 验证映射
    print("🔍 验证映射有效性...")
    validation = mapper.validate_mappings()
    
    print(f"✅ 有效映射: {len(validation['valid_mappings'])} 个")
    if validation['invalid_db_names']:
        print(f"❌ 无效数据库名称: {validation['invalid_db_names']}")
    if validation['missing_aspen_names']:
        print(f"❌ 未找到的Aspen名称: {validation['missing_aspen_names']}")
    
    # 显示改进的映射
    mapper.print_improved_mappings()
    
    # 保存到数据库
    print("\n💾 保存改进映射到数据库...")
    if mapper.save_improved_mappings():
        print("✅ 改进映射保存成功")
        
        # 提供使用建议
        print("\n📋 使用建议:")
        print("  • 高置信度映射可以直接使用")
        print("  • 中等置信度映射需要工程师确认")
        print("  • 低置信度映射建议手动检查")
        print("  • 映射结果已保存到数据库表 'improved_stream_mappings'")
    else:
        print("❌ 改进映射保存失败")

if __name__ == "__main__":
    main()
