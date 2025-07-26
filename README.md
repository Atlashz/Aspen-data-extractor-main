# Aspen Plus数据提取与数据库构建工具

专用于从Aspen Plus仿真文件和Excel热交换器数据表中提取工程数据，并构建结构化SQLite数据库的工具。

## ✨ 核心功能

- **� Aspen Plus数据提取**: 通过COM接口从Aspen Plus仿真中提取流股和设备数据
- **� Excel热交换器数据处理**: 读取和处理Excel格式的热交换器数据表
- **🗄️ SQLite数据库构建**: 将提取的数据存储到结构化的SQLite数据库
- **✅数据验证和导出**: 完整的数据验证和多格式导出功能
- **🧩 模块化设计**: 清晰、可测试、可维护的代码架构

## 🏗️ 项目结构

```
├── README.md                           # 项目说明文档
├── CLAUDE.md                           # Claude Code开发指南
├── requirements.txt                    # Python依赖包
├── 
├── 🎯 核心数据提取系统
├── aspen_data_extractor.py            # Aspen Plus数据提取器 (主要)
├── aspen_data_database.py             # Aspen数据存储和管理
├── data_interfaces.py                 # 数据结构定义
├── 
├── 📊 数据处理工具
├── stream_classifier.py               # 流股分类器
├── stream_mapping.py                  # 流股映射工具
├── improved_stream_mapping.py         # 改进的流股映射
├── equipment_model_matcher.py         # 设备模型匹配器
├── 
├── 🔧 维护和修复工具
├── check_database_completeness.py     # 数据库完整性检查
├── fix_equipment_types.py             # 设备类型修复
├── fix_hex_data.py                    # 热交换器数据修复
├── final_status_report.py             # 最终状态报告
├── query_stream_mappings.py           # 流股映射查询
├── 
├── 📁 数据文件
├── aspen_data.db                      # SQLite数据库文件
├── BFG-CO2H-HEX.xlsx                  # 热交换器数据表
├── equipment match.xlsx               # 设备匹配表
├── aspen_files/                       # Aspen仿真文件目录
│   ├── BFG-CO2H-MEOH V2 (purge burning).apw
│   └── BFG-CO2H-MEOH V2 (purge burning).ads
├── equipment match/                   # 设备匹配工具
│   ├── Equipment_Model_Functions.xlsx
│   └── equipment_model_matcher.py
├── 
├── 🧪 测试验证
├── tests/                             # 测试文件目录
│   ├── test_data_interfaces.py
│   └── test_database_manager.py
```

## 🚀 安装与设置

### 1. 环境配置

1. **下载项目文件**

2. **配置Python环境**:
   ```bash
   # 安装依赖包
   pip install -r requirements.txt
   ```

3. **Aspen Plus集成** (仅限Windows):
   ```bash
   pip install pywin32
   ```

## 🎯 Quick Start Guide

### 数据提取和处理

```bash
# 提取Aspen Plus数据
python aspen_data_extractor.py

# 检查数据库完整性
python check_database_completeness.py

# 生成状态报告
python final_status_report.py

# 查询流股映射
python query_stream_mappings.py
```

## 📖 使用方法

### 1. 从Aspen Plus提取数据

```python
from aspen_data_extractor import AspenDataExtractor

# 创建数据提取器
extractor = AspenDataExtractor()

# 从Aspen文件提取数据
process_data = extractor.extract_complete_data("path/to/your_simulation.apw")

# 数据会自动存储到aspen_data.db数据库中
```

### 2. 加载Excel热交换器数据

```python
# 加载Excel热交换器数据
extractor.load_hex_data("BFG-CO2H-HEX.xlsx")

# 获取热交换器数据摘要
hex_summary = extractor.get_hex_summary()
print(f"热交换器数量: {hex_summary['hex_count']}")
print(f"总热负荷: {hex_summary['total_heat_duty']} kW")
```

### 3. 数据处理和分析

```bash
# 流股分类和映射
python stream_classifier.py
python stream_mapping.py
python improved_stream_mapping.py

# 设备匹配
python equipment_model_matcher.py

# 数据修复和维护
python fix_equipment_types.py
python fix_hex_data.py
```

### 4. 数据库维护

```python
# 检查数据库完整性
from check_database_completeness import check_completeness
completeness_report = check_completeness("aspen_data.db")

# 查询流股映射
from query_stream_mappings import query_mappings
mappings = query_mappings()

# 生成最终报告
from final_status_report import generate_report
report = generate_report()
```

## 🔌 Aspen Plus Integration

## 🗄️ 数据库结构

工具创建的SQLite数据库(`aspen_data.db`)包含以下表结构：

### 数据表说明

1. **`streams`** - 流股数据
   - 流股名称、温度、压力、流量、组分
   
2. **`equipment`** - 设备数据  
   - 设备名称、类型、操作参数、负荷
   
3. **`heat_exchangers`** - 热交换器数据
   - 换热器名称、热负荷、面积、温差
   
4. **`sessions`** - 提取会话记录
   - 提取时间、文件信息、数据统计

### 系统要求

- **操作系统**: Windows (Aspen COM接口要求)
- **软件**: Aspen Plus V11+
- **Python包**: `pywin32` (COM接口), `pandas`, `openpyxl`

### 支持的数据类型

- ✅ 流股属性 (温度、压力、流量、组分)
- ✅ 设备操作数据 (负荷、压降)
- ✅ 热交换器参数
- ✅ Excel表格数据导入
## 🧪 测试验证

```bash
# 运行测试套件
python -m pytest tests/

# 单独运行测试
python -m pytest tests/test_data_interfaces.py
python -m pytest tests/test_database_manager.py

# 检查数据库状态
python check_database_completeness.py
```

## 📋 文件说明

### 核心文件
- `aspen_data_extractor.py` - 主要数据提取模块
- `aspen_data_database.py` - 数据库管理系统
- `data_interfaces.py` - 数据结构定义
- `stream_classifier.py` - 流股分类器
- `equipment_model_matcher.py` - 设备模型匹配

### 工具脚本
- `check_database_completeness.py` - 数据库完整性检查
- `final_status_report.py` - 生成最终状态报告
- `fix_equipment_types.py` - 修复设备类型
- `fix_hex_data.py` - 修复热交换器数据
- `query_stream_mappings.py` - 查询流股映射

### 数据文件
- `aspen_data.db` - SQLite数据库
- `BFG-CO2H-HEX.xlsx` - 热交换器数据表
- `equipment match.xlsx` - 设备匹配表
- `aspen_files/` - Aspen仿真文件目录
- `equipment match/` - 设备匹配工具和数据

## � 许可说明

本工具仅供教育和研究目的使用。

---

**版本**: 2.0 (数据提取专用版)  
**更新日期**: 2025-07-25  
**状态**: 🟢 可用