#!/usr/bin/env python3
"""
测试整合后的aspen_data_extractor.py设备流股连接功能

Author: 测试脚本
Date: 2025-07-27
"""

import logging
from aspen_data_extractor import AspenDataExtractor

# 设置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_equipment_extraction_with_connections():
    """测试带有流股连接信息的设备提取"""
    
    logger.info("=" * 80)
    logger.info("🚀 测试整合后的设备提取功能")
    logger.info("=" * 80)
    
    try:
        # 创建数据提取器
        extractor = AspenDataExtractor()
        
        # 测试COM接口连接
        logger.info("🔌 测试Aspen Plus连接...")
        com_test = extractor.com_interface.test_com_availability()
        
        if not com_test['com_objects_found']:
            logger.warning("⚠️ 未找到Aspen Plus COM对象，将使用模拟数据测试")
            test_with_mock_data(extractor)
            return
        
        # 尝试连接到活动的Aspen实例或打开文件
        aspen_file = "aspen_files/BFG-CO2H-MEOH V2 (purge burning).apw"
        connected = extractor.com_interface.connect_to_active(aspen_file)
        
        if not connected:
            logger.warning("⚠️ 无法连接到Aspen Plus，将使用模拟数据测试")
            test_with_mock_data(extractor)
            return
        
        logger.info("✅ 已连接到Aspen Plus")
        
        # 提取设备数据（现在包含流股连接信息）
        logger.info("🔧 提取设备数据...")
        equipment = extractor.extract_all_equipment()
        
        # 分析结果
        analyze_equipment_connections(equipment)
        
        # 清理连接
        extractor.com_interface.disconnect()
        logger.info("✅ 已断开Aspen连接")
        
    except Exception as e:
        logger.error(f"❌ 测试失败: {str(e)}")

def test_with_mock_data(extractor):
    """使用模拟数据测试流股连接解析功能"""
    
    logger.info("🧪 测试Excel流股连接解析功能...")
    
    # 直接测试Excel连接解析
    connections = extractor.com_interface.load_flowsheet_connections()
    
    if connections:
        logger.info(f"✅ 成功加载 {len(connections)} 个设备的连接信息")
        
        logger.info("\n📊 设备连接统计:")
        for equipment, conn_info in connections.items():
            inlet_count = len(conn_info['inlet_streams'])
            outlet_count = len(conn_info['outlet_streams'])
            logger.info(f"  {equipment}: {inlet_count} 进料, {outlet_count} 出料")
        
        # 测试单个设备查询
        logger.info("\n🔍 测试单个设备查询:")
        test_equipment = ['B1', 'MX1', 'C-301', 'DI']
        
        for eq_name in test_equipment:
            inlet_streams, outlet_streams = extractor.com_interface.get_equipment_stream_connections_from_excel(eq_name)
            logger.info(f"  {eq_name}: 进料{inlet_streams}, 出料{outlet_streams}")
    else:
        logger.warning("⚠️ 未能加载流股连接信息")

def analyze_equipment_connections(equipment):
    """分析设备连接结果"""
    
    logger.info("\n" + "=" * 80)
    logger.info("📊 设备连接分析结果")
    logger.info("=" * 80)
    
    # 统计连接信息
    total_equipment = len(equipment)
    equipment_with_connections = 0
    total_inlet_streams = 0
    total_outlet_streams = 0
    
    connection_patterns = {}
    
    for eq_name, eq_data in equipment.items():
        inlet_streams = eq_data.get('inlet_streams', [])
        outlet_streams = eq_data.get('outlet_streams', [])
        
        if inlet_streams or outlet_streams:
            equipment_with_connections += 1
        
        total_inlet_streams += len(inlet_streams)
        total_outlet_streams += len(outlet_streams)
        
        # 记录连接模式
        pattern = f"{len(inlet_streams)}→{len(outlet_streams)}"
        connection_patterns[pattern] = connection_patterns.get(pattern, 0) + 1
    
    logger.info(f"📈 总体统计:")
    logger.info(f"  总设备数: {total_equipment}")
    logger.info(f"  有连接设备: {equipment_with_connections}")
    logger.info(f"  总进料流股: {total_inlet_streams}")
    logger.info(f"  总出料流股: {total_outlet_streams}")
    
    logger.info(f"\n🔗 连接模式分布:")
    for pattern, count in sorted(connection_patterns.items()):
        logger.info(f"  {pattern}: {count} 个设备")
    
    logger.info(f"\n🔝 详细连接信息:")
    for eq_name, eq_data in equipment.items():
        inlet_streams = eq_data.get('inlet_streams', [])
        outlet_streams = eq_data.get('outlet_streams', [])
        
        if inlet_streams or outlet_streams:
            logger.info(f"  🏭 {eq_name} ({eq_data.get('type', 'Unknown')}):")
            if inlet_streams:
                logger.info(f"    📥 进料: {', '.join(inlet_streams)}")
            if outlet_streams:
                logger.info(f"    📤 出料: {', '.join(outlet_streams)}")

def main():
    """主函数"""
    test_equipment_extraction_with_connections()

if __name__ == "__main__":
    main()
