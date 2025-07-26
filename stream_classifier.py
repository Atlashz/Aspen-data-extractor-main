#!/usr/bin/env python3
"""
流股分类和标签化脚本
对Aspen Plus提取的流股进行分类：原料、过程、产品
"""

import sqlite3
import json
import pandas as pd
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass
from enum import Enum
import re

class StreamCategory(Enum):
    """流股分类枚举"""
    RAW_MATERIAL = "原料"          # Raw materials/feeds
    PROCESS = "过程"               # Process streams
    PRODUCT = "产品"               # Products
    UTILITY = "公用工程"           # Utilities
    RECYCLE = "循环"               # Recycle streams
    WASTE = "废料"                 # Waste streams
    INTERMEDIATE = "中间产物"      # Intermediate products
    HOT_UTILITY = "热公用工程"     # Hot utility streams (steam, hot oil, etc.)
    COLD_UTILITY = "冷公用工程"    # Cold utility streams (cooling water, refrigerant, etc.)

@dataclass
class StreamClassification:
    """流股分类结果"""
    name: str
    category: StreamCategory
    sub_category: str = ""
    confidence: float = 1.0
    reasoning: List[str] = None
    
    def __post_init__(self):
        if self.reasoning is None:
            self.reasoning = []

class StreamClassifier:
    """
    流股分类器
    基于流股名称、组成、工艺条件等信息进行分类
    """
    
    def __init__(self):
        # 定义分类规则
        self.classification_rules = {
            StreamCategory.RAW_MATERIAL: {
                'name_patterns': [
                    r'.*feed.*', r'.*raw.*', r'.*input.*', r'.*makeup.*',
                    r'bfg.*', r'.*co2.*feed.*', r'h2.*makeup.*', r'fresh.*'
                ],
                'composition_indicators': {
                    'high_inerts': ['N2', 'AR'],  # 高惰性气体含量
                    'raw_materials': ['CO', 'CO2', 'H2', 'CH4']
                },
                'temperature_range': (15, 100),  # 通常较低温度
                'pressure_range': (1, 10)       # 通常较低压力
            },
            
            StreamCategory.PRODUCT: {
                'name_patterns': [
                    r'.*product.*', r'.*meoh.*', r'.*methanol.*', r'.*water.*product.*',
                    r'.*outlet.*', r'.*final.*'
                ],
                'composition_indicators': {
                    'high_product': ['CH3OH', 'H2O'],  # 高产品浓度
                    'product_components': ['CH3OH']
                },
                'temperature_range': (20, 200),
                'pressure_range': (1, 60)
            },
            
            StreamCategory.RECYCLE: {
                'name_patterns': [
                    r'.*recycle.*', r'.*recirc.*', r'.*return.*', r'.*loop.*'
                ],
                'composition_indicators': {
                    'recycle_components': ['H2', 'CO', 'CO2', 'CH4'],
                    'low_inerts': ['N2']
                },
                'temperature_range': (30, 300),
                'pressure_range': (10, 60)
            },
            
            StreamCategory.PROCESS: {
                'name_patterns': [
                    r'rxn.*', r'reactor.*', r'.*out.*', r'.*in.*', r's\d+.*', r'.*mix.*'
                ],
                'composition_indicators': {
                    'mixed_components': ['CO', 'CO2', 'H2', 'CH3OH', 'H2O']
                },
                'temperature_range': (100, 350),
                'pressure_range': (20, 60)
            },
            
            StreamCategory.UTILITY: {
                'name_patterns': [
                    r'.*utility.*', r'.*service.*'
                ],
                'composition_indicators': {
                    'utility_components': ['H2O', 'STEAM']
                },
                'temperature_range': (10, 400),
                'pressure_range': (1, 100)
            },
            
            StreamCategory.HOT_UTILITY: {
                'name_patterns': [
                    r'.*steam.*', r'.*hot.*oil.*', r'.*hot.*water.*', r'.*heating.*',
                    r'.*hot.*utility.*', r'.*thermal.*oil.*', r'.*dowtherm.*'
                ],
                'composition_indicators': {
                    'steam_components': ['H2O', 'STEAM'],
                    'thermal_oil_components': ['THERMOIL', 'DOWTHERM']
                },
                'temperature_range': (120, 500),  # Hot utilities are typically high temperature
                'pressure_range': (1, 50)
            },
            
            StreamCategory.COLD_UTILITY: {
                'name_patterns': [
                    r'.*cooling.*water.*', r'.*cold.*water.*', r'.*chilled.*',
                    r'.*refrigerant.*', r'.*coolant.*', r'.*cold.*utility.*'
                ],
                'composition_indicators': {
                    'cooling_water': ['H2O'],
                    'refrigerants': ['NH3', 'R134A', 'R410A', 'PROPANE']
                },
                'temperature_range': (-50, 40),  # Cold utilities are typically low temperature
                'pressure_range': (1, 20)
            },
            
            StreamCategory.WASTE: {
                'name_patterns': [
                    r'.*waste.*', r'.*purge.*', r'.*vent.*', r'.*blow.*down.*'
                ],
                'composition_indicators': {
                    'waste_indicators': ['N2', 'CH4', 'CO2']
                }
            }
        }
    
    def classify_stream(self, stream_data: Dict) -> StreamClassification:
        """
        对单个流股进行分类
        
        Args:
            stream_data: 包含流股信息的字典
            
        Returns:
            StreamClassification对象
        """
        name = stream_data.get('name', '').lower()
        temperature = stream_data.get('temperature', 0)
        pressure = stream_data.get('pressure', 0)
        composition = stream_data.get('composition', {})
        
        # 解析组成数据
        if isinstance(composition, str):
            try:
                composition = json.loads(composition)
            except:
                composition = {}
        
        # 计算各分类的得分
        scores = {}
        detailed_reasoning = {}
        
        for category, rules in self.classification_rules.items():
            score = 0.0
            reasoning = []
            
            # 名称匹配
            name_score = self._check_name_patterns(name, rules.get('name_patterns', []))
            if name_score > 0:
                score += name_score * 0.4  # 名称权重40%
                reasoning.append(f"名称匹配 (得分: {name_score:.2f})")
            
            # 组成匹配
            comp_score = self._check_composition_indicators(composition, rules.get('composition_indicators', {}))
            if comp_score > 0:
                score += comp_score * 0.4  # 组成权重40%
                reasoning.append(f"组成匹配 (得分: {comp_score:.2f})")
            
            # 温度匹配
            temp_score = self._check_temperature_range(temperature, rules.get('temperature_range'))
            if temp_score > 0:
                score += temp_score * 0.1  # 温度权重10%
                reasoning.append(f"温度匹配 (得分: {temp_score:.2f})")
            
            # 压力匹配
            pres_score = self._check_pressure_range(pressure, rules.get('pressure_range'))
            if pres_score > 0:
                score += pres_score * 0.1  # 压力权重10%
                reasoning.append(f"压力匹配 (得分: {pres_score:.2f})")
            
            scores[category] = score
            detailed_reasoning[category] = reasoning
        
        # 选择得分最高的分类
        if scores:
            best_category = max(scores.keys(), key=lambda k: scores[k])
            confidence = scores[best_category]
            reasoning = detailed_reasoning[best_category]
        else:
            best_category = StreamCategory.PROCESS  # 默认分类
            confidence = 0.3
            reasoning = ["默认分类"]
        
        # 确定子分类
        sub_category = self._determine_sub_category(best_category, stream_data)
        
        return StreamClassification(
            name=stream_data.get('name', ''),
            category=best_category,
            sub_category=sub_category,
            confidence=confidence,
            reasoning=reasoning
        )
    
    def _check_name_patterns(self, name: str, patterns: List[str]) -> float:
        """检查名称模式匹配"""
        for pattern in patterns:
            if re.search(pattern, name, re.IGNORECASE):
                return 1.0
        return 0.0
    
    def _check_composition_indicators(self, composition: Dict[str, float], indicators: Dict[str, List[str]]) -> float:
        """检查组成指示器"""
        if not composition or not indicators:
            return 0.0
        
        total_score = 0.0
        indicator_count = 0
        
        for indicator_type, components in indicators.items():
            indicator_count += 1
            
            if indicator_type in ['high_product', 'high_inerts']:
                # 检查高浓度组分
                max_conc = max([composition.get(comp, 0) for comp in components])
                if max_conc > 0.5:  # 浓度超过50%
                    total_score += 1.0
                elif max_conc > 0.2:  # 浓度超过20%
                    total_score += 0.6
                elif max_conc > 0.05:  # 浓度超过5%
                    total_score += 0.3
            
            elif indicator_type in ['low_inerts']:
                # 检查低浓度组分
                max_conc = max([composition.get(comp, 0) for comp in components])
                if max_conc < 0.1:  # 浓度低于10%
                    total_score += 0.8
                elif max_conc < 0.3:  # 浓度低于30%
                    total_score += 0.4
            
            else:
                # 检查组分存在性
                present_count = sum([1 for comp in components if composition.get(comp, 0) > 0.01])
                if present_count > 0:
                    total_score += present_count / len(components)
        
        return total_score / max(1, indicator_count)
    
    def _check_temperature_range(self, temperature: float, temp_range: Optional[Tuple[float, float]]) -> float:
        """检查温度范围"""
        if not temp_range or temperature == 0:
            return 0.0
        
        min_temp, max_temp = temp_range
        if min_temp <= temperature <= max_temp:
            return 1.0
        elif temperature < min_temp:
            # 温度过低的惩罚
            if temperature >= min_temp - 50:
                return 0.5
        elif temperature > max_temp:
            # 温度过高的惩罚
            if temperature <= max_temp + 100:
                return 0.5
        
        return 0.0
    
    def _check_pressure_range(self, pressure: float, pres_range: Optional[Tuple[float, float]]) -> float:
        """检查压力范围"""
        if not pres_range or pressure == 0:
            return 0.0
        
        min_pres, max_pres = pres_range
        if min_pres <= pressure <= max_pres:
            return 1.0
        elif pressure < min_pres:
            if pressure >= min_pres - 5:
                return 0.5
        elif pressure > max_pres:
            if pressure <= max_pres + 20:
                return 0.5
        
        return 0.0
    
    def _determine_sub_category(self, category: StreamCategory, stream_data: Dict) -> str:
        """确定子分类"""
        name = stream_data.get('name', '').lower()
        composition = stream_data.get('composition', {})
        
        if isinstance(composition, str):
            try:
                composition = json.loads(composition)
            except:
                composition = {}
        
        if category == StreamCategory.RAW_MATERIAL:
            if 'bfg' in name or 'blast' in name:
                return "高炉煤气"
            elif 'co2' in name:
                return "二氧化碳原料"
            elif 'h2' in name:
                return "氢气补充"
            else:
                return "其他原料"
        
        elif category == StreamCategory.PRODUCT:
            if 'meoh' in name or 'methanol' in name:
                return "甲醇产品"
            elif 'water' in name:
                return "水产品"
            else:
                return "其他产品"
        
        elif category == StreamCategory.PROCESS:
            if 'rxn' in name or 'reactor' in name:
                return "反应器流股"
            elif 'mix' in name:
                return "混合流股"
            else:
                return "工艺流股"
        
        elif category == StreamCategory.RECYCLE:
            return "循环气"
        
        return ""

def classify_all_streams() -> List[StreamClassification]:
    """对数据库中的所有流股进行分类"""
    
    # 连接数据库
    conn = sqlite3.connect('aspen_data.db')
    cursor = conn.cursor()
    
    # 获取所有流股数据
    cursor.execute('''
        SELECT name, temperature, pressure, mass_flow, volume_flow, molar_flow, composition
        FROM aspen_streams
        ORDER BY name
    ''')
    
    streams_data = cursor.fetchall()
    conn.close()
    
    # 初始化分类器
    classifier = StreamClassifier()
    
    # 对每个流股进行分类
    classifications = []
    
    print("🔍 流股分类分析")
    print("=" * 60)
    
    for stream_row in streams_data:
        name, temp, pres, mass_flow, vol_flow, mol_flow, composition = stream_row
        
        stream_data = {
            'name': name,
            'temperature': temp or 0,
            'pressure': pres or 0,
            'mass_flow': mass_flow or 0,
            'volume_flow': vol_flow or 0,
            'molar_flow': mol_flow or 0,
            'composition': composition or '{}'
        }
        
        # 分类
        classification = classifier.classify_stream(stream_data)
        classifications.append(classification)
        
        # 打印分类结果
        print(f"\n📋 流股: {name}")
        print(f"   分类: {classification.category.value}")
        if classification.sub_category:
            print(f"   子分类: {classification.sub_category}")
        print(f"   置信度: {classification.confidence:.2f}")
        print(f"   条件: T={temp}°C, P={pres}bar, {mass_flow:.0f}kg/hr")
        
        # 显示主要组分
        if composition:
            try:
                comp_dict = json.loads(composition)
                main_comps = {k: f"{v:.3f}" for k, v in comp_dict.items() if v > 0.01}
                if main_comps:
                    print(f"   主要组分: {main_comps}")
            except:
                pass
        
        print(f"   分类依据: {', '.join(classification.reasoning)}")
    
    return classifications

def update_database_with_classifications(classifications: List[StreamClassification]):
    """将分类结果更新到数据库"""
    
    # 检查是否需要添加新列
    conn = sqlite3.connect('aspen_data.db')
    cursor = conn.cursor()
    
    # 检查列是否存在
    cursor.execute("PRAGMA table_info(aspen_streams)")
    columns = [column[1] for column in cursor.fetchall()]
    
    # 添加分类相关列
    if 'stream_category' not in columns:
        cursor.execute('ALTER TABLE aspen_streams ADD COLUMN stream_category TEXT')
    
    if 'stream_sub_category' not in columns:
        cursor.execute('ALTER TABLE aspen_streams ADD COLUMN stream_sub_category TEXT')
    
    if 'classification_confidence' not in columns:
        cursor.execute('ALTER TABLE aspen_streams ADD COLUMN classification_confidence REAL')
    
    if 'classification_reasoning' not in columns:
        cursor.execute('ALTER TABLE aspen_streams ADD COLUMN classification_reasoning TEXT')
    
    # 更新分类信息
    print(f"\n📝 更新数据库中的流股分类...")
    
    for classification in classifications:
        cursor.execute('''
            UPDATE aspen_streams
            SET stream_category = ?,
                stream_sub_category = ?,
                classification_confidence = ?,
                classification_reasoning = ?
            WHERE name = ?
        ''', (
            classification.category.value,
            classification.sub_category,
            classification.confidence,
            json.dumps(classification.reasoning, ensure_ascii=False),
            classification.name
        ))
    
    conn.commit()
    conn.close()
    
    print(f"✅ 已更新 {len(classifications)} 个流股的分类信息")

def generate_classification_summary(classifications: List[StreamClassification]):
    """生成分类汇总报告"""
    
    print(f"\n📊 流股分类汇总报告")
    print("=" * 60)
    
    # 按分类统计
    category_counts = {}
    for classification in classifications:
        category = classification.category.value
        category_counts[category] = category_counts.get(category, 0) + 1
    
    print(f"总流股数: {len(classifications)}")
    print(f"\n按分类统计:")
    for category, count in sorted(category_counts.items()):
        percentage = (count / len(classifications)) * 100
        print(f"  {category}: {count} ({percentage:.1f}%)")
    
    # 按子分类统计
    print(f"\n详细分类:")
    current_category = None
    for classification in sorted(classifications, key=lambda x: x.category.value):
        if classification.category.value != current_category:
            current_category = classification.category.value
            print(f"\n{current_category}:")
        
        sub_info = f" - {classification.sub_category}" if classification.sub_category else ""
        confidence_info = f" (置信度: {classification.confidence:.2f})"
        print(f"  • {classification.name}{sub_info}{confidence_info}")
    
    # 低置信度分类
    low_confidence = [c for c in classifications if c.confidence < 0.6]
    if low_confidence:
        print(f"\n⚠️  低置信度分类 (需要人工确认):")
        for classification in low_confidence:
            print(f"  • {classification.name}: {classification.category.value} "
                  f"(置信度: {classification.confidence:.2f})")

def export_classification_results(classifications: List[StreamClassification], filename: str = None):
    """导出分类结果到Excel文件"""
    
    if filename is None:
        from datetime import datetime
        filename = f"stream_classification_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    # 准备数据
    export_data = []
    for classification in classifications:
        export_data.append({
            '流股名称': classification.name,
            '分类': classification.category.value,
            '子分类': classification.sub_category,
            '置信度': classification.confidence,
            '分类依据': '; '.join(classification.reasoning)
        })
    
    # 创建DataFrame并导出
    df = pd.DataFrame(export_data)
    
    try:
        df.to_excel(filename, index=False, engine='openpyxl')
        print(f"\n📊 分类结果已导出到: {filename}")
        return True
    except Exception as e:
        print(f"❌ 导出失败: {e}")
        return False

if __name__ == "__main__":
    print("🚀 流股分类和标签化系统")
    print("=" * 50)
    
    # 执行分类
    classifications = classify_all_streams()
    
    # 生成汇总报告
    generate_classification_summary(classifications)
    
    # 询问是否更新数据库
    response = input(f"\n是否将分类结果更新到数据库? (y/n): ").lower()
    if response == 'y':
        update_database_with_classifications(classifications)
        
        # 验证更新
        print(f"\n🔍 验证数据库更新...")
        conn = sqlite3.connect('aspen_data.db')
        df = pd.read_sql_query('''
            SELECT name, stream_category, stream_sub_category, classification_confidence
            FROM aspen_streams
            ORDER BY stream_category, name
        ''', conn)
        conn.close()
        
        print(f"\n更新后的流股分类:")
        print(df.to_string(index=False))
    
    # 询问是否导出Excel
    response = input(f"\n是否导出分类结果到Excel? (y/n): ").lower()
    if response == 'y':
        export_classification_results(classifications)
    
    print(f"\n🎉 流股分类完成!")
