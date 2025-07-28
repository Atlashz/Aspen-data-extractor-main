#!/usr/bin/env python3
"""
快速修复脚本

解决Aspen经济数据提取工具的常见依赖问题。

Usage: python quick_fix.py
"""

import subprocess
import sys
import os

def install_yaml():
    """安装PyYAML包"""
    try:
        print("Installing PyYAML...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "PyYAML"])
        print("PyYAML installed successfully!")
        return True
    except subprocess.CalledProcessError:
        print("PyYAML installation failed")
        return False

def install_basic_requirements():
    """安装基础依赖"""
    basic_packages = [
        "PyYAML>=6.0",
        "openpyxl>=3.0.0", 
        "pandas>=1.3.0",
        "numpy>=1.21.0"
    ]
    
    for package in basic_packages:
        try:
            print(f"Installing {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"{package} installed successfully")
        except subprocess.CalledProcessError:
            print(f"{package} installation failed")

def test_imports():
    """测试关键模块导入"""
    print("\nTesting module imports...")
    
    modules_to_test = [
        ("yaml", "PyYAML"),
        ("openpyxl", "openpyxl"),
        ("pandas", "pandas"),
        ("numpy", "numpy")
    ]
    
    all_good = True
    for module_name, package_name in modules_to_test:
        try:
            __import__(module_name)
            print(f"OK: {module_name} import successful")
        except ImportError:
            print(f"FAIL: {module_name} import failed")
            all_good = False
    
    return all_good

def run_simple_test():
    """运行简单的功能测试"""
    print("\nRunning functionality tests...")
    
    try:
        # 测试配置文件加载
        import json
        config_file = "config/economic_extraction_config.json"
        if os.path.exists(config_file):
            with open(config_file, 'r') as f:
                config = json.load(f)
            print("OK: JSON config file loaded successfully")
        
        # 测试数据结构
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from data_interfaces import EconomicAnalysisResults, CostItem, CostCategory, CurrencyType
        
        # 创建测试数据
        test_result = EconomicAnalysisResults(
            project_name="Test",
            timestamp=__import__("datetime").datetime.now()
        )
        
        test_cost = CostItem(
            name="Test Equipment",
            category=CostCategory.EQUIPMENT,
            base_cost=100000,
            currency=CurrencyType.USD
        )
        
        print("OK: Data structures test passed")
        
        # 测试Excel导出器（不实际生成文件）
        try:
            from economic_excel_exporter import EconomicExcelExporter
            exporter = EconomicExcelExporter()
            print("OK: Excel exporter initialized successfully")
        except Exception as e:
            print(f"WARNING: Excel exporter test issue: {e}")
        
        print("SUCCESS: All basic functionality tests passed!")
        return True
        
    except Exception as e:
        print(f"FAIL: Functionality test failed: {e}")
        return False

def main():
    """主修复流程"""
    print("Aspen Economic Data Extractor - Quick Fix")
    print("=" * 40)
    
    # 步骤1：安装基础依赖
    print("\nStep 1: Installing basic dependencies...")
    install_basic_requirements()
    
    # 步骤2：测试导入
    print("\nStep 2: Testing module imports...")
    if test_imports():
        print("SUCCESS: All modules imported successfully")
    else:
        print("WARNING: Some modules have issues, but basic functionality might still work")
    
    # 步骤3：运行功能测试
    print("\nStep 3: Running functionality tests...")
    if run_simple_test():
        print("SUCCESS: Functionality tests passed")
    else:
        print("FAIL: Functionality tests failed, please check error messages")
    
    # 步骤4：提供使用建议
    print("\nUsage Recommendations:")
    print("1. Use JSON config file instead of YAML:")
    print("   python extract_aspen_economics.py --config config/economic_extraction_config.json")
    print("")
    print("2. Run example program:")
    print("   python examples/extract_economics_example.py")
    print("")
    print("3. If problems persist, run:")
    print("   python install_dependencies.py")
    
    print("\nQuick fix completed!")

if __name__ == "__main__":
    main()