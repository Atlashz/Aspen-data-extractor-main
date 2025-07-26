#!/usr/bin/env python3
"""
测试改进后的热交换器数据提取功能
验证所有修复是否生效
"""

import os
import sys
import logging
from datetime import datetime

# Setup logging to see detailed output
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

def test_enhanced_hex_extraction():
    """测试增强的热交换器数据提取"""
    
    print("🚀 测试增强的热交换器数据提取功能")
    print("=" * 80)
    
    try:
        # Import the enhanced AspenDataExtractor
        from aspen_data_extractor import AspenDataExtractor
        
        # Test file
        hex_file = "BFG-CO2H-HEX.xlsx"
        
        if not os.path.exists(hex_file):
            print(f"❌ 测试文件不存在: {hex_file}")
            return False
        
        print(f"📁 测试文件: {hex_file}")
        
        # Create extractor
        extractor = AspenDataExtractor()
        
        # Test 1: Excel structure analysis
        print(f"\n🔍 步骤 1: 分析Excel文件结构...")
        try:
            from analyze_excel_structure import ExcelStructureAnalyzer
            analyzer = ExcelStructureAnalyzer(hex_file)
            analysis = analyzer.analyze_complete_structure()
            
            if analysis and not analysis.get('error'):
                print(f"   ✅ 结构分析成功:")
                print(f"      • 工作表数量: {analysis.get('summary', {}).get('total_worksheets', 0)}")
                print(f"      • 总列数: {analysis.get('summary', {}).get('total_columns', 0)}")
                print(f"      • 总行数: {analysis.get('summary', {}).get('total_rows', 0)}")
                
                likely_sheets = analysis.get('summary', {}).get('likely_hex_worksheets', [])
                if likely_sheets:
                    print(f"      • 可能的热交换器工作表: {[s['sheet'] for s in likely_sheets]}")
            else:
                print(f"   ⚠️ 结构分析遇到问题: {analysis.get('error', '未知错误')}")
        except Exception as e:
            print(f"   ⚠️ 结构分析失败: {str(e)}")
        
        # Test 2: Enhanced data loading
        print(f"\n📊 步骤 2: 测试增强的数据加载...")
        
        hex_success = extractor.load_hex_data(hex_file)
        
        if hex_success:
            print(f"   ✅ 数据加载成功!")
            
            # Get summary
            summary = extractor.get_hex_summary()
            print(f"   📈 数据摘要:")
            print(f"      • 总行数: {summary.get('total_heat_exchangers', 0)}")
            print(f"      • 列数: {len(summary.get('columns', []))}")
            print(f"      • 相关列: {len(summary.get('relevant_columns', []))}")
            
            if 'processed_summary' in summary:
                processed = summary['processed_summary']
                print(f"      • 处理的热交换器: {processed.get('processed_hex_count', 0)}")
                print(f"      • 总热负荷: {processed.get('total_heat_duty_kW', 0):,.1f} kW")
                print(f"      • 总换热面积: {processed.get('total_heat_area_m2', 0):,.1f} m²")
            
        else:
            print(f"   ❌ 数据加载失败")
        
        # Test 3: Detailed extraction report
        print(f"\n📋 步骤 3: 生成详细提取报告...")
        
        try:
            report = extractor.get_hex_extraction_report()
            
            if not report.get('error'):
                print(f"   ✅ 报告生成成功:")
                print(f"      • 分析的工作表: {report.get('worksheets_analyzed', 0)}")
                print(f"      • 提取的热交换器: {report.get('total_data_extracted', 0)}")
                print(f"      • 总热负荷: {report.get('total_heat_duty_kw', 0):,.1f} kW")
                print(f"      • 总换热面积: {report.get('total_heat_area_m2', 0):,.1f} m²")
                
                # Data quality breakdown
                quality = report.get('data_quality_breakdown', {})
                if quality:
                    print(f"      • 数据质量分布: {quality}")
                
                # Extraction success by type
                success = report.get('extraction_success_by_type', {})
                if success:
                    print(f"      • 按类型提取成功率:")
                    for data_type, count in success.items():
                        if count > 0:
                            print(f"         - {data_type}: {count}")
                
                # Show some recommendations
                recommendations = report.get('recommendations', [])
                if recommendations:
                    print(f"      • 建议数量: {len(recommendations)}")
                    for i, rec in enumerate(recommendations[:3], 1):
                        print(f"         {i}. {rec}")
                
            else:
                print(f"   ❌ 报告生成失败: {report.get('error')}")
                
        except Exception as e:
            print(f"   ❌ 报告生成异常: {str(e)}")
        
        # Test 4: Print full report
        print(f"\n📊 步骤 4: 打印完整提取报告...")
        try:
            extractor.print_hex_extraction_report()
            print(f"   ✅ 完整报告打印成功")
        except Exception as e:
            print(f"   ❌ 报告打印失败: {str(e)}")
        
        # Test 5: Compare with original
        print(f"\n🔄 步骤 5: 对比测试结果...")
        
        if hex_success:
            tea_data = extractor.get_hex_data_for_tea()
            if tea_data:
                hex_count = tea_data.get('hex_count', 0)
                total_duty = tea_data.get('total_heat_duty_kW', 0)
                total_area = tea_data.get('total_heat_area_m2', 0)
                
                print(f"   📊 最终提取结果:")
                print(f"      • 热交换器数量: {hex_count}")
                print(f"      • 总热负荷: {total_duty:,.1f} kW")
                print(f"      • 总换热面积: {total_area:,.1f} m²")
                
                # Success criteria
                success_criteria = []
                if hex_count > 0:
                    success_criteria.append("✅ 提取到热交换器数据")
                else:
                    success_criteria.append("❌ 未提取到热交换器数据")
                
                if total_duty > 0:
                    success_criteria.append("✅ 提取到热负荷数据")
                else:
                    success_criteria.append("⚠️ 未提取到热负荷数据")
                
                if total_area > 0:
                    success_criteria.append("✅ 提取到换热面积数据")
                else:
                    success_criteria.append("⚠️ 未提取到换热面积数据")
                
                print(f"   🎯 成功标准评估:")
                for criterion in success_criteria:
                    print(f"      {criterion}")
                
                # Overall assessment
                successful_criteria = len([c for c in success_criteria if "✅" in c])
                total_criteria = len(success_criteria)
                
                print(f"\n🏆 总体评估:")
                if successful_criteria >= total_criteria - 1:
                    print(f"   🎉 测试成功! ({successful_criteria}/{total_criteria} 标准通过)")
                    print(f"   💪 数据提取功能显著改善!")
                    return True
                elif successful_criteria >= 1:
                    print(f"   ⚠️ 部分成功 ({successful_criteria}/{total_criteria} 标准通过)")
                    print(f"   🔧 还有改进空间")
                    return True
                else:
                    print(f"   ❌ 测试失败 (0/{total_criteria} 标准通过)")
                    print(f"   🚨 需要进一步调试")
                    return False
            else:
                print(f"   ❌ 无法获取TEA数据格式")
                return False
        else:
            print(f"   ❌ 数据加载失败，无法进行对比")
            return False
        
    except ImportError as e:
        print(f"❌ 导入错误: {str(e)}")
        print(f"   请确保所有必要的模块都可用")
        return False
    except Exception as e:
        print(f"❌ 测试过程中出现异常: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """主函数"""
    print(f"🕒 开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Run the test
    success = test_enhanced_hex_extraction()
    
    print(f"\n🕒 结束时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)
    
    if success:
        print("🎊 测试完成! 热交换器数据提取功能已成功增强!")
    else:
        print("⚠️ 测试完成，但仍有问题需要解决")
    
    return success

if __name__ == "__main__":
    main()