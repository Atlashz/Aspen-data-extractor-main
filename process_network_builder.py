#!/usr/bin/env python3
"""
流程网络构建器

整合Aspen Plus数据和Excel热交换器数据，构建完整的流程网络。
移除Aspen中的heater/cooler设备，使用Excel中的详细热交换器数据替换，
并自动修复流股连接以确保网络完整性。

主要功能:
- 从Aspen Plus提取流股和设备数据
- 从Excel提取详细的热交换器信息
- 移除热设备并用Excel数据替换
- 自动修复和创建缺失的流股连接
- 验证网络完整性和生成可视化

Author: 流程网络构建工具
Date: 2025-07-27
Version: 1.0
"""

import os
import sys
import json
import logging
from typing import Dict, List, Optional, Set, Tuple, Any
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass, field
from collections import defaultdict

import numpy as np
import pandas as pd

# Import our existing modules
from aspen_data_extractor import AspenDataExtractor, HeatExchangerDataLoader
from data_interfaces import (
    AspenProcessData, StreamData, UnitOperationData, EquipmentType
)

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


@dataclass
class NetworkStream:
    """网络中的流股对象"""
    name: str
    source_equipment: Optional[str] = None
    target_equipment: Optional[str] = None
    temperature: Optional[float] = None
    pressure: Optional[float] = None
    flow_rate: Optional[float] = None
    composition: Dict[str, float] = field(default_factory=dict)
    stream_type: str = "process"  # process, hot, cold, utility
    is_synthetic: bool = False  # 是否为合成流股
    original_aspen_name: Optional[str] = None


@dataclass
class NetworkEquipment:
    """网络中的设备对象"""
    name: str
    equipment_type: str
    aspen_type: Optional[str] = None
    inlet_streams: List[str] = field(default_factory=list)
    outlet_streams: List[str] = field(default_factory=list)
    parameters: Dict[str, Any] = field(default_factory=dict)
    is_heat_exchanger: bool = False
    excel_data: Optional[Dict] = None


@dataclass
class ProcessNetwork:
    """完整的流程网络"""
    streams: Dict[str, NetworkStream] = field(default_factory=dict)
    equipment: Dict[str, NetworkEquipment] = field(default_factory=dict)
    connections: List[Tuple[str, str, str]] = field(default_factory=list)  # (from_eq, stream, to_eq)
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def add_stream(self, stream: NetworkStream):
        """添加流股到网络"""
        self.streams[stream.name] = stream
    
    def add_equipment(self, equipment: NetworkEquipment):
        """添加设备到网络"""
        self.equipment[equipment.name] = equipment
    
    def add_connection(self, from_equipment: str, stream_name: str, to_equipment: str):
        """添加连接"""
        self.connections.append((from_equipment, stream_name, to_equipment))


class ProcessNetworkBuilder:
    """
    流程网络构建器
    
    整合Aspen Plus数据和Excel热交换器数据，构建完整的流程网络
    """
    
    def __init__(self, aspen_file: str, hex_excel_file: str):
        self.aspen_file = aspen_file
        self.hex_excel_file = hex_excel_file
        
        # 初始化数据提取器
        self.aspen_extractor = AspenDataExtractor()
        self.hex_loader = HeatExchangerDataLoader(hex_excel_file)
        
        # 网络数据
        self.network = ProcessNetwork()
        self.aspen_data = None
        self.hex_data = None
        
        # 追踪信息
        self.removed_equipment = []
        self.synthetic_streams = []
        self.connection_repairs = []
        
        logger.info(f"初始化流程网络构建器")
        logger.info(f"Aspen文件: {aspen_file}")
        logger.info(f"Excel文件: {hex_excel_file}")
    
    def build_complete_network(self) -> ProcessNetwork:
        """
        构建完整的流程网络
        
        Returns:
            ProcessNetwork: 完整的流程网络对象
        """
        logger.info("🏗️ 开始构建完整的流程网络...")
        
        try:
            # 步骤1: 提取数据
            self._extract_data()
            
            # 步骤2: 构建基础网络
            self._build_base_network()
            
            # 步骤3: 移除热设备
            self._remove_thermal_equipment()
            
            # 步骤4: 添加Excel热交换器
            self._add_excel_heat_exchangers()
            
            # 步骤5: 修复流股连接
            self._repair_stream_connections()
            
            # 步骤6: 验证网络完整性
            self._validate_network_integrity()
            
            # 步骤7: 添加元数据
            self._add_metadata()
            
            logger.info("✅ 流程网络构建完成!")
            return self.network
            
        except Exception as e:
            logger.error(f"❌ 构建流程网络失败: {e}")
            raise
    
    def _extract_data(self):
        """提取Aspen和Excel数据"""
        logger.info("📊 提取数据...")
        
        # 提取Aspen数据
        logger.info("从Aspen Plus提取数据...")
        self.aspen_data = self.aspen_extractor.extract_complete_data(self.aspen_file)
        logger.info(f"✅ 提取了 {len(self.aspen_data.streams)} 个流股, {len(self.aspen_data.units)} 个设备")
        
        # 提取Excel数据
        logger.info("从Excel提取热交换器数据...")
        self.hex_loader.load_data()
        self.hex_data = self.hex_loader.get_heat_exchanger_data_for_tea()
        logger.info(f"✅ 提取了 {self.hex_data.get('hex_count', 0)} 个热交换器")
    
    def _build_base_network(self):
        """构建基础网络结构"""
        logger.info("🔨 构建基础网络结构...")
        
        # 添加所有Aspen流股
        for stream_name, stream_data in self.aspen_data.streams.items():
            network_stream = NetworkStream(
                name=stream_name,
                temperature=stream_data.temperature,
                pressure=stream_data.pressure,
                flow_rate=stream_data.mass_flow,  # 使用mass_flow而不是flow_rate
                composition=stream_data.composition,
                original_aspen_name=stream_name
            )
            self.network.add_stream(network_stream)
        
        # 添加所有Aspen设备
        for unit_name, unit_data in self.aspen_data.units.items():
            network_equipment = NetworkEquipment(
                name=unit_name,
                equipment_type=self._map_equipment_type(unit_data.type),  # 使用type而不是equipment_type
                aspen_type=unit_data.aspen_block_type if hasattr(unit_data, 'aspen_block_type') else str(unit_data.type),
                parameters=unit_data.parameters
            )
            self.network.add_equipment(network_equipment)
        
        logger.info(f"✅ 基础网络: {len(self.network.streams)} 流股, {len(self.network.equipment)} 设备")
    
    def _remove_thermal_equipment(self):
        """移除Aspen中的加热器和冷却器"""
        logger.info("🔥 移除热设备 (heater/cooler)...")
        
        thermal_equipment_types = {'heater', 'cooler', 'heat_exchanger'}
        equipment_to_remove = []
        
        for eq_name, equipment in self.network.equipment.items():
            if equipment.equipment_type.lower() in thermal_equipment_types:
                equipment_to_remove.append(eq_name)
                self.removed_equipment.append({
                    'name': eq_name,
                    'type': equipment.equipment_type,
                    'aspen_type': equipment.aspen_type,
                    'parameters': equipment.parameters
                })
        
        # 移除设备
        for eq_name in equipment_to_remove:
            del self.network.equipment[eq_name]
            logger.info(f"  🗑️ 移除设备: {eq_name}")
        
        logger.info(f"✅ 移除了 {len(equipment_to_remove)} 个热设备")
    
    def _add_excel_heat_exchangers(self):
        """添加Excel中的热交换器到网络"""
        logger.info("🔄 添加Excel热交换器...")
        
        if not self.hex_data or not self.hex_data.get('heat_exchangers'):
            logger.warning("⚠️ 没有找到Excel热交换器数据")
            return
        
        for hex_info in self.hex_data['heat_exchangers']:
            hex_name = hex_info.get('name', f"HEX-{hex_info['index']}")
            
            # 创建热交换器设备
            hex_equipment = NetworkEquipment(
                name=hex_name,
                equipment_type="heat_exchanger",
                is_heat_exchanger=True,
                excel_data=hex_info,
                parameters={
                    'duty_kW': hex_info.get('duty', 0),
                    'area_m2': hex_info.get('area', 0),
                    'hot_stream_name': hex_info.get('hot_stream_name'),
                    'cold_stream_name': hex_info.get('cold_stream_name'),
                    'hot_inlet_temp': hex_info.get('hot_stream_inlet_temp'),
                    'hot_outlet_temp': hex_info.get('hot_stream_outlet_temp'),
                    'cold_inlet_temp': hex_info.get('cold_stream_inlet_temp'),
                    'cold_outlet_temp': hex_info.get('cold_stream_outlet_temp')
                }
            )
            
            self.network.add_equipment(hex_equipment)
            logger.info(f"  ➕ 添加热交换器: {hex_name}")
            
            # 创建热交换器相关的合成流股
            self._create_hex_synthetic_streams(hex_name, hex_info)
        
        logger.info(f"✅ 添加了 {len(self.hex_data['heat_exchangers'])} 个热交换器")
    
    def _create_hex_synthetic_streams(self, hex_name: str, hex_info: Dict):
        """为热交换器创建合成流股"""
        
        # 获取热流和冷流信息
        hot_stream_name = hex_info.get('hot_stream_name')
        cold_stream_name = hex_info.get('cold_stream_name')
        
        # 创建热流侧的入口和出口流股
        if hot_stream_name:
            # 热流入口流股
            hot_inlet_name = f"{hex_name}_HOT_IN"
            if hot_inlet_name not in self.network.streams:
                hot_inlet_stream = NetworkStream(
                    name=hot_inlet_name,
                    target_equipment=hex_name,
                    temperature=hex_info.get('hot_stream_inlet_temp'),
                    stream_type="hot",
                    is_synthetic=True,
                    original_aspen_name=hot_stream_name
                )
                self.network.add_stream(hot_inlet_stream)
                self.synthetic_streams.append(hot_inlet_name)
            
            # 热流出口流股  
            hot_outlet_name = f"{hex_name}_HOT_OUT"
            if hot_outlet_name not in self.network.streams:
                hot_outlet_stream = NetworkStream(
                    name=hot_outlet_name,
                    source_equipment=hex_name,
                    temperature=hex_info.get('hot_stream_outlet_temp'),
                    stream_type="hot",
                    is_synthetic=True,
                    original_aspen_name=hot_stream_name
                )
                self.network.add_stream(hot_outlet_stream)
                self.synthetic_streams.append(hot_outlet_name)
        
        # 创建冷流侧的入口和出口流股
        if cold_stream_name:
            # 冷流入口流股
            cold_inlet_name = f"{hex_name}_COLD_IN"
            if cold_inlet_name not in self.network.streams:
                cold_inlet_stream = NetworkStream(
                    name=cold_inlet_name,
                    target_equipment=hex_name,
                    temperature=hex_info.get('cold_stream_inlet_temp'),
                    stream_type="cold",
                    is_synthetic=True,
                    original_aspen_name=cold_stream_name
                )
                self.network.add_stream(cold_inlet_stream)
                self.synthetic_streams.append(cold_inlet_name)
            
            # 冷流出口流股
            cold_outlet_name = f"{hex_name}_COLD_OUT"
            if cold_outlet_name not in self.network.streams:
                cold_outlet_stream = NetworkStream(
                    name=cold_outlet_name,
                    source_equipment=hex_name,
                    temperature=hex_info.get('cold_stream_outlet_temp'),
                    stream_type="cold",
                    is_synthetic=True,
                    original_aspen_name=cold_stream_name
                )
                self.network.add_stream(cold_outlet_stream)
                self.synthetic_streams.append(cold_outlet_name)
    
    def _repair_stream_connections(self):
        """修复流股连接"""
        logger.info("🔗 修复流股连接...")
        
        repairs_made = 0
        
        # 分析需要修复的连接
        for hex_name, hex_equipment in self.network.equipment.items():
            if not hex_equipment.is_heat_exchanger:
                continue
                
            hex_info = hex_equipment.excel_data
            if not hex_info:
                continue
            
            hot_stream_name = hex_info.get('hot_stream_name')
            cold_stream_name = hex_info.get('cold_stream_name')
            
            # 修复热流连接
            if hot_stream_name and hot_stream_name in self.network.streams:
                repairs_made += self._repair_hex_stream_connection(
                    hex_name, hot_stream_name, "hot"
                )
            
            # 修复冷流连接
            if cold_stream_name and cold_stream_name in self.network.streams:
                repairs_made += self._repair_hex_stream_connection(
                    hex_name, cold_stream_name, "cold"
                )
        
        logger.info(f"✅ 完成了 {repairs_made} 个流股连接修复")
    
    def _repair_hex_stream_connection(self, hex_name: str, original_stream_name: str, stream_type: str) -> int:
        """修复单个热交换器的流股连接"""
        repairs = 0
        
        # 找到原始流股
        original_stream = self.network.streams.get(original_stream_name)
        if not original_stream:
            return 0
        
        # 创建连接逻辑
        if stream_type == "hot":
            inlet_synthetic = f"{hex_name}_HOT_IN"
            outlet_synthetic = f"{hex_name}_HOT_OUT"
        else:
            inlet_synthetic = f"{hex_name}_COLD_IN"
            outlet_synthetic = f"{hex_name}_COLD_OUT"
        
        # 如果原始流股有源设备，连接到热交换器入口
        if original_stream.source_equipment:
            if inlet_synthetic in self.network.streams:
                # 更新合成入口流股的源设备
                self.network.streams[inlet_synthetic].source_equipment = original_stream.source_equipment
                # 添加连接
                self.network.add_connection(original_stream.source_equipment, inlet_synthetic, hex_name)
                repairs += 1
                
                self.connection_repairs.append({
                    'type': 'inlet_connection',
                    'hex': hex_name,
                    'stream_type': stream_type,
                    'from': original_stream.source_equipment,
                    'to': hex_name,
                    'via': inlet_synthetic
                })
        
        # 如果原始流股有目标设备，从热交换器出口连接
        if original_stream.target_equipment:
            if outlet_synthetic in self.network.streams:
                # 更新合成出口流股的目标设备
                self.network.streams[outlet_synthetic].target_equipment = original_stream.target_equipment
                # 添加连接
                self.network.add_connection(hex_name, outlet_synthetic, original_stream.target_equipment)
                repairs += 1
                
                self.connection_repairs.append({
                    'type': 'outlet_connection',
                    'hex': hex_name,
                    'stream_type': stream_type,
                    'from': hex_name,  
                    'to': original_stream.target_equipment,
                    'via': outlet_synthetic
                })
        
        return repairs
    
    def _validate_network_integrity(self) -> Dict[str, Any]:
        """验证网络完整性"""
        logger.info("✅ 验证网络完整性...")
        
        validation_results = {
            'streams_count': len(self.network.streams),
            'equipment_count': len(self.network.equipment),
            'connections_count': len(self.network.connections),
            'synthetic_streams_count': len(self.synthetic_streams),
            'removed_equipment_count': len(self.removed_equipment),
            'connection_repairs_count': len(self.connection_repairs),
            'issues': []
        }
        
        # 检查孤立流股
        orphaned_streams = []
        for stream_name, stream in self.network.streams.items():
            if not stream.source_equipment and not stream.target_equipment and not stream.is_synthetic:
                orphaned_streams.append(stream_name)
        
        if orphaned_streams:
            validation_results['issues'].append(f"发现 {len(orphaned_streams)} 个孤立流股")
        
        # 检查设备连接
        unconnected_equipment = []
        for eq_name, equipment in self.network.equipment.items():
            connected = any(eq_name in conn for conn in self.network.connections)
            if not connected and not equipment.is_heat_exchanger:
                unconnected_equipment.append(eq_name)
        
        if unconnected_equipment:
            validation_results['issues'].append(f"发现 {len(unconnected_equipment)} 个未连接设备")
        
        # 打印验证结果
        logger.info("📊 网络完整性验证结果:")
        logger.info(f"  流股总数: {validation_results['streams_count']}")
        logger.info(f"  设备总数: {validation_results['equipment_count']}")
        logger.info(f"  连接总数: {validation_results['connections_count']}")
        logger.info(f"  合成流股: {validation_results['synthetic_streams_count']}")
        logger.info(f"  移除设备: {validation_results['removed_equipment_count']}")
        logger.info(f"  连接修复: {validation_results['connection_repairs_count']}")
        
        if validation_results['issues']:
            logger.warning("⚠️ 发现的问题:")
            for issue in validation_results['issues']:
                logger.warning(f"  • {issue}")
        else:
            logger.info("✅ 网络完整性验证通过!")
        
        return validation_results
    
    def _add_metadata(self):
        """添加网络元数据"""
        self.network.metadata = {
            'created_at': datetime.now().isoformat(),
            'aspen_file': self.aspen_file,
            'hex_excel_file': self.hex_excel_file,
            'original_aspen_streams': len(self.aspen_data.streams) if self.aspen_data else 0,
            'original_aspen_equipment': len(self.aspen_data.units) if self.aspen_data else 0,
            'excel_heat_exchangers': self.hex_data.get('hex_count', 0) if self.hex_data else 0,
            'removed_equipment': len(self.removed_equipment),
            'synthetic_streams': len(self.synthetic_streams),
            'connection_repairs': len(self.connection_repairs),
            'final_streams': len(self.network.streams),
            'final_equipment': len(self.network.equipment),
            'final_connections': len(self.network.connections)
        }
    
    def _map_equipment_type(self, aspen_type) -> str:
        """映射Aspen设备类型到标准类型"""
        # 如果是EquipmentType枚举，获取其值
        if hasattr(aspen_type, 'value'):
            aspen_type_str = aspen_type.value
        else:
            aspen_type_str = str(aspen_type)
            
        type_mapping = {
            'reactor': 'reactor',
            'compressor': 'compressor',
            'heat_exchanger': 'heat_exchanger',
            'distillation_column': 'distillation_column',
            'separator': 'separator',
            'mixer': 'mixer',
            'splitter': 'splitter',
            'pump': 'pump',
            'tank': 'tank',
            'valve': 'valve',
            'unknown': 'unknown',
            # Aspen specific mappings
            'ISENTROPIC': 'compressor',
            'T-SPEC': 'temperature_controller', 
            'HEATER': 'heater',
            'COOLER': 'cooler',
            'HEATX': 'heat_exchanger',
            'RADFRAC': 'distillation_column',
            'DSTWU': 'distillation_column',
            'FLASH2': 'separator',
            'SEP': 'separator',
            'MIXER': 'mixer',
            'FSPLIT': 'splitter'
        }
        
        return type_mapping.get(aspen_type_str.lower(), aspen_type_str.lower())
    
    def export_network(self, output_file: str = None) -> str:
        """导出网络数据"""
        if not output_file:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"process_network_{timestamp}.json"
        
        export_data = {
            'metadata': self.network.metadata,
            'streams': {
                name: {
                    'name': stream.name,
                    'source_equipment': stream.source_equipment,
                    'target_equipment': stream.target_equipment,
                    'temperature': stream.temperature,
                    'pressure': stream.pressure,
                    'flow_rate': stream.flow_rate,
                    'composition': stream.composition,
                    'stream_type': stream.stream_type,
                    'is_synthetic': stream.is_synthetic,
                    'original_aspen_name': stream.original_aspen_name
                }
                for name, stream in self.network.streams.items()
            },
            'equipment': {
                name: {
                    'name': eq.name,
                    'equipment_type': eq.equipment_type,
                    'aspen_type': eq.aspen_type,
                    'inlet_streams': eq.inlet_streams,
                    'outlet_streams': eq.outlet_streams,
                    'parameters': eq.parameters,
                    'is_heat_exchanger': eq.is_heat_exchanger,
                    'excel_data': eq.excel_data
                }
                for name, eq in self.network.equipment.items()
            },
            'connections': self.network.connections,
            'removed_equipment': self.removed_equipment,
            'synthetic_streams': self.synthetic_streams,
            'connection_repairs': self.connection_repairs
        }
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(export_data, f, ensure_ascii=False, indent=2)
        
        logger.info(f"✅ 网络数据已导出到: {output_file}")
        return output_file
    
    def print_network_summary(self):
        """打印网络摘要"""
        print("\n" + "="*60)
        print("🏭 完整流程网络构建摘要")
        print("="*60)
        
        print(f"📊 数据源:")
        print(f"  • Aspen文件: {Path(self.aspen_file).name}")
        print(f"  • Excel文件: {Path(self.hex_excel_file).name}")
        
        print(f"\n🔧 网络统计:")
        print(f"  • 总流股数: {len(self.network.streams)}")
        print(f"  • 总设备数: {len(self.network.equipment)}")
        print(f"  • 总连接数: {len(self.network.connections)}")
        
        print(f"\n🔄 网络构建操作:")
        print(f"  • 移除热设备: {len(self.removed_equipment)}")
        print(f"  • 添加热交换器: {self.hex_data.get('hex_count', 0) if self.hex_data else 0}")
        print(f"  • 创建合成流股: {len(self.synthetic_streams)}")
        print(f"  • 修复连接: {len(self.connection_repairs)}")
        
        # 设备类型统计
        equipment_types = defaultdict(int)
        for equipment in self.network.equipment.values():
            equipment_types[equipment.equipment_type] += 1
        
        print(f"\n⚙️ 设备类型分布:")
        for eq_type, count in sorted(equipment_types.items()):
            print(f"  • {eq_type}: {count}")
        
        # 流股类型统计
        stream_types = defaultdict(int)
        synthetic_count = 0
        for stream in self.network.streams.values():
            stream_types[stream.stream_type] += 1
            if stream.is_synthetic:
                synthetic_count += 1
        
        print(f"\n🌊 流股类型分布:")
        for stream_type, count in sorted(stream_types.items()):
            print(f"  • {stream_type}: {count}")
        print(f"  • 合成流股: {synthetic_count}")
        
        print("="*60)


def main():
    """主函数 - 运行流程网络构建测试"""
    print("🏗️ 流程网络构建器 - 测试运行")
    print("="*50)
    
    # 设置文件路径
    current_dir = Path(__file__).parent
    aspen_file = current_dir / "aspen_files" / "BFG-CO2H-MEOH V2 (purge burning).apw"
    hex_file = current_dir / "BFG-CO2H-HEX.xlsx"
    
    # 检查文件存在性
    if not aspen_file.exists():
        logger.error(f"❌ Aspen文件不存在: {aspen_file}")
        return
    
    if not hex_file.exists():
        logger.error(f"❌ Excel文件不存在: {hex_file}")
        return
    
    try:
        # 创建网络构建器
        builder = ProcessNetworkBuilder(str(aspen_file), str(hex_file))
        
        # 构建完整网络
        network = builder.build_complete_network()
        
        # 打印摘要
        builder.print_network_summary()
        
        # 导出网络数据
        output_file = builder.export_network()
        
        print(f"\n🎉 流程网络构建成功!")
        print(f"📁 网络数据已保存到: {output_file}")
        
        return network
        
    except Exception as e:
        logger.error(f"❌ 构建失败: {e}")
        import traceback
        traceback.print_exc()
        return None


if __name__ == "__main__":
    main()
