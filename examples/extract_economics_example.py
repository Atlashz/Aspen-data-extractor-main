#!/usr/bin/env python3
"""
Aspen经济数据提取示例脚本

演示如何使用AspenEconomicsExtractor进行经济数据提取和分析的各种方法。

Author: TEA Analysis Framework
Date: 2025-07-27
Version: 1.0
"""

import os
import sys
from pathlib import Path

# 添加上级目录到路径，以便导入模块
sys.path.insert(0, str(Path(__file__).parent.parent))

try:
    from extract_aspen_economics import AspenEconomicsExtractor
    import logging
except ImportError as e:
    print(f"ERROR: 导入错误: {e}")
    print("FIX: 请先运行快速修复脚本:")
    print("   python quick_fix.py")
    print("或安装依赖:")
    print("   pip install PyYAML openpyxl pandas numpy")
    sys.exit(1)

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def example_1_extract_from_izp_file():
    """
    示例1：从IZP成本文件提取经济数据
    """
    print("\n" + "="*60)
    print("示例1：从IZP成本文件提取经济数据")
    print("="*60)
    
    try:
        # 创建提取器
        extractor = AspenEconomicsExtractor()
        
        # IZP文件路径（假设存在）
        izp_file = "aspen_files/BFG-CO2H-MEOH V2 (purge burning)Cost/Scenario1/Scenario1.izp"
        output_file = "examples/output/economics_from_izp.xlsx"
        
        if os.path.exists(izp_file):
            # 执行提取和导出
            result = extractor.extract_and_export(
                data_source=izp_file,
                output_file=output_file
            )
            
            if result['success']:
                print(f"OK: 成功生成报告: {result['report_path']}")
                print(f"CAPEX: ${result.get('total_capex', 0):,.0f}")
                print(f"OPEX: ${result.get('annual_opex', 0):,.0f}")
            else:
                print("ERROR: 提取失败:")
                for error in result['errors']:
                    print(f"   {error}")
        else:
            print(f"WARNING: IZP文件不存在: {izp_file}")
            print("   请确保测试文件存在或修改文件路径")
            
    except Exception as e:
        print(f"ERROR: 示例1执行失败: {str(e)}")


def example_2_extract_from_com_interface():
    """
    示例2：从Aspen Plus COM接口提取经济数据
    """
    print("\n" + "="*60)
    print("示例2：从Aspen Plus COM接口提取经济数据")
    print("="*60)
    
    try:
        # 创建提取器
        extractor = AspenEconomicsExtractor()
        
        output_file = "examples/output/economics_from_com.xlsx"
        
        print("NOTE: 此示例需要：")
        print("   1. Windows操作系统")
        print("   2. 已安装Aspen Plus")
        print("   3. 当前有打开的Aspen Plus仿真")
        
        # 从COM接口提取（连接到活动的Aspen Plus实例）
        result = extractor.extract_and_export(
            data_source="aspen_com",
            output_file=output_file,
            project_name="COM Interface Demo"
        )
        
        if result['success']:
            print(f"OK: 成功生成报告: {result['report_path']}")
            print(f"CAPEX: ${result.get('total_capex', 0):,.0f}")
            print(f"OPEX: ${result.get('annual_opex', 0):,.0f}")
            print(f"NPV: ${result.get('npv', 0):,.0f}")
        else:
            print("ERROR: 提取失败:")
            for error in result['errors']:
                print(f"   {error}")
                
    except Exception as e:
        print(f"ERROR: 示例2执行失败: {str(e)}")
        print("   这通常是由于没有可用的Aspen Plus COM接口")


def example_3_extract_with_custom_config():
    """
    示例3：使用自定义配置文件进行提取
    """
    print("\n" + "="*60)
    print("示例3：使用自定义配置文件进行提取")
    print("="*60)
    
    try:
        # 使用自定义配置文件创建提取器
        config_file = "config/economic_extraction_config.yaml"
        extractor = AspenEconomicsExtractor(config_file=config_file)
        
        # SZP文件路径（假设存在）
        szp_file = "aspen_files/BFG-CO2H-MEOH V2 (purge burning)Cost/Scenario1/Scenario1.szp"
        output_file = "examples/output/economics_with_config.xlsx"
        
        if os.path.exists(szp_file):
            # 执行提取和导出
            result = extractor.extract_and_export(
                data_source=szp_file,
                output_file=output_file
            )
            
            if result['success']:
                print(f"OK: 成功生成报告: {result['report_path']}")
                print(f"CAPEX: ${result.get('total_capex', 0):,.0f}")
                print(f"OPEX: ${result.get('annual_opex', 0):,.0f}")
                print(f"TIME: 耗时: {result.get('duration_seconds', 0):.1f}秒")
            else:
                print("ERROR: 提取失败:")
                for error in result['errors']:
                    print(f"   {error}")
        else:
            print(f"WARNING: SZP文件不存在: {szp_file}")
            print("   请确保测试文件存在或修改文件路径")
            
    except Exception as e:
        print(f"ERROR: 示例3执行失败: {str(e)}")


def example_4_step_by_step_extraction():
    """
    示例4：分步骤的经济数据提取和处理
    """
    print("\n" + "="*60)
    print("示例4：分步骤的经济数据提取和处理")
    print("="*60)
    
    try:
        # 创建提取器
        extractor = AspenEconomicsExtractor()
        
        # 测试文件路径
        test_file = "aspen_files/BFG-CO2H-MEOH V2 (purge burning)Cost/Scenario1/Scenario1.izp"
        
        if os.path.exists(test_file):
            print("Processing: 步骤1：提取经济数据...")
            # 第一步：提取经济数据
            economic_results = extractor.extract_from_cost_files(test_file)
            
            print(f"   项目名称: {economic_results.project_name}")
            print(f"   分析时间: {economic_results.timestamp}")
            print(f"   数据源数量: {len(economic_results.data_sources)}")
            
            print("Processing: 步骤2：生成Excel报告...")
            # 第二步：生成Excel报告
            output_file = "examples/output/step_by_step_report.xlsx"
            report_path = extractor.generate_excel_report(economic_results, output_file)
            
            print(f"OK: 报告生成完成: {report_path}")
            
            print("Processing: 步骤3：获取提取摘要...")
            # 第三步：获取提取摘要
            summary = extractor.get_extraction_summary()
            print(f"   解析错误数量: {len(summary['economic_parser_errors'])}")
            print(f"   解析警告数量: {len(summary['economic_parser_warnings'])}")
            print(f"   可用模块: {summary['available_modules']}")
            
        else:
            print(f"WARNING: 测试文件不存在: {test_file}")
            
    except Exception as e:
        print(f"ERROR: 示例4执行失败: {str(e)}")


def example_5_batch_processing():
    """
    示例5：批量处理多个文件
    """
    print("\n" + "="*60)
    print("示例5：批量处理多个文件")
    print("="*60)
    
    try:
        # 创建提取器
        extractor = AspenEconomicsExtractor()
        
        # 要处理的文件列表
        files_to_process = [
            "aspen_files/BFG-CO2H-MEOH V2 (purge burning)Cost/Scenario1/Scenario1.izp",
            "aspen_files/BFG-CO2H-MEOH V2 (purge burning)Cost/Scenario1/Scenario1.szp",
        ]
        
        results_summary = []
        
        for i, file_path in enumerate(files_to_process):
            print(f"\nProcessing: 处理文件 {i+1}/{len(files_to_process)}: {Path(file_path).name}")
            
            if os.path.exists(file_path):
                output_file = f"examples/output/batch_report_{i+1}_{Path(file_path).stem}.xlsx"
                
                # 处理文件
                result = extractor.extract_and_export(
                    data_source=file_path,
                    output_file=output_file
                )
                
                results_summary.append({
                    'file': file_path,
                    'success': result['success'],
                    'output': result.get('report_path', 'N/A'),
                    'capex': result.get('total_capex', 0),
                    'opex': result.get('annual_opex', 0)
                })
                
                if result['success']:
                    print(f"   OK: 成功处理，生成报告: {result['report_path']}")
                else:
                    print(f"   ERROR: 处理失败: {result['errors']}")
            else:
                print(f"   WARNING: 文件不存在: {file_path}")
                results_summary.append({
                    'file': file_path,
                    'success': False,
                    'output': 'File not found',
                    'capex': 0,
                    'opex': 0
                })
        
        # 输出批量处理结果摘要
        print("\nSUMMARY: 批量处理结果摘要:")
        print("-" * 60)
        for result in results_summary:
            status = "OK" if result['success'] else "ERROR"
            print(f"{status} {Path(result['file']).name}")
            if result['success']:
                print(f"   CAPEX: ${result['capex']:,.0f}, OPEX: ${result['opex']:,.0f}")
            print(f"   输出: {result['output']}")
            
    except Exception as e:
        print(f"ERROR: 示例5执行失败: {str(e)}")


def create_output_directory():
    """创建输出目录"""
    output_dir = Path("examples/output")
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Output directory created: {output_dir}")


def main():
    """主函数 - 运行所有示例"""
    print("Aspen Economic Data Extraction Examples")
    print("Author: TEA Analysis Framework")
    print("Version: 1.0")
    
    # 创建输出目录
    create_output_directory()
    
    # 运行示例
    try:
        example_1_extract_from_izp_file()
        example_2_extract_from_com_interface()
        example_3_extract_with_custom_config()
        example_4_step_by_step_extraction()
        example_5_batch_processing()
        
        print("\n" + "="*60)
        print("All examples completed successfully!")
        print("Please check the examples/output/ directory for generated files")
        print("="*60)
        
    except KeyboardInterrupt:
        print("\nUser interrupted program execution")
    except Exception as e:
        print(f"\nProgram execution failed: {str(e)}")


if __name__ == "__main__":
    main()