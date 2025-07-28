#!/usr/bin/env python3
"""
依赖包安装脚本

自动检测和安装Aspen经济数据提取工具所需的Python包。

Author: TEA Analysis Framework
Date: 2025-07-27
Version: 1.0
"""

import subprocess
import sys
import importlib
import platform
from pathlib import Path

def check_python_version():
    """检查Python版本"""
    if sys.version_info < (3, 7):
        print("❌ 错误：需要Python 3.7或更高版本")
        print(f"   当前版本：{sys.version}")
        return False
    else:
        print(f"✅ Python版本检查通过：{sys.version}")
        return True

def check_module(module_name, package_name=None):
    """检查模块是否已安装"""
    try:
        importlib.import_module(module_name)
        return True
    except ImportError:
        return False

def install_package(package_name):
    """安装Python包"""
    try:
        print(f"🔄 正在安装 {package_name}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
        print(f"✅ {package_name} 安装成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {package_name} 安装失败: {e}")
        return False

def main():
    """主安装流程"""
    print("🚀 Aspen经济数据提取工具 - 依赖包安装")
    print("=" * 50)
    
    # 检查Python版本
    if not check_python_version():
        sys.exit(1)
    
    # 检查操作系统
    os_type = platform.system()
    print(f"📋 操作系统: {os_type}")
    
    # 必需的包列表
    required_packages = [
        ("numpy", "numpy>=1.21.0"),
        ("pandas", "pandas>=1.3.0"),
        ("openpyxl", "openpyxl>=3.0.0"),
        ("yaml", "PyYAML>=6.0"),
        ("pydantic", "pydantic>=1.8.0"),
    ]
    
    # Windows特定的包
    if os_type == "Windows":
        required_packages.append(("win32com.client", "pywin32>=227"))
    
    # 检查和安装必需的包
    print("\n📦 检查必需的依赖包...")
    missing_packages = []
    
    for module_name, package_spec in required_packages:
        if check_module(module_name):
            print(f"✅ {module_name} 已安装")
        else:
            print(f"❌ {module_name} 未安装")
            missing_packages.append(package_spec)
    
    # 安装缺失的包
    if missing_packages:
        print(f"\n🔧 需要安装 {len(missing_packages)} 个包...")
        
        for package in missing_packages:
            install_package(package)
    else:
        print("\n🎉 所有必需的依赖包已安装！")
    
    # 可选包检查
    print("\n📋 检查可选依赖包...")
    optional_packages = [
        ("pytest", "pytest>=6.2.0", "用于运行测试"),
        ("black", "black>=21.0.0", "用于代码格式化"),
        ("mypy", "mypy>=0.910", "用于类型检查")
    ]
    
    for module_name, package_spec, description in optional_packages:
        if check_module(module_name):
            print(f"✅ {module_name} 已安装 - {description}")
        else:
            print(f"⚠️ {module_name} 未安装 - {description}")
    
    # 验证安装
    print("\n🧪 验证安装...")
    
    try:
        # 测试核心模块导入
        import numpy
        import pandas
        import openpyxl
        import yaml
        import pydantic
        
        print("✅ 核心模块导入测试通过")
        
        # Windows COM测试
        if os_type == "Windows":
            try:
                import win32com.client
                print("✅ Windows COM模块可用")
            except ImportError:
                print("⚠️ Windows COM模块不可用（需要手动安装pywin32）")
        
        # 检查本地模块
        script_dir = Path(__file__).parent
        sys.path.insert(0, str(script_dir))
        
        try:
            from data_interfaces import EconomicAnalysisResults
            print("✅ 本地数据接口模块可用")
        except ImportError as e:
            print(f"⚠️ 本地模块导入问题: {e}")
        
        print("\n🎉 安装验证完成！")
        
    except ImportError as e:
        print(f"❌ 模块导入失败: {e}")
        print("请检查安装过程中是否有错误信息")
    
    # 提供使用建议
    print("\n📖 使用建议:")
    print("1. 运行示例: python examples/extract_economics_example.py")
    print("2. 查看帮助: python extract_aspen_economics.py --help")
    print("3. 阅读文档: README_ECONOMICS.md")
    
    # Windows特别说明
    if os_type == "Windows":
        print("\n🪟 Windows用户注意事项:")
        print("- 如果使用Aspen Plus COM接口，请确保以管理员身份运行Python")
        print("- 如果pywin32安装失败，请尝试从官网下载安装包")
        print("- 运行 'python -c \"import win32com.client; print('COM可用')\"' 测试COM功能")

if __name__ == "__main__":
    main()