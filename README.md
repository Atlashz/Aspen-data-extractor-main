# Aspen Plus数据提取与数据库构建工具

专用于从Aspen Plus仿真文件和Excel热交换器数据表中提取工程数据，并构建结构化SQLite数据库的工具。支持经济分析(TEA)和工艺网络分析。

## ✨ 核心功能

- **🔌 Aspen Plus数据提取**: 通过COM接口从Aspen Plus仿真中提取流股和设备数据
- **📊 Excel热交换器数据处理**: 读取和处理Excel格式的热交换器数据表，支持智能列映射
- **🗄️ SQLite数据库构建**: 将提取的数据存储到结构化的SQLite数据库，支持会话管理
- **🏷️ 智能流股分类**: 自动识别原料、产品、过程流股，支持置信度评估
- **⚙️ 设备类型识别**: 基于Excel匹配表和Aspen类型的智能设备识别
- **💰 经济分析**: TEA经济分析和报告生成功能
- **🔗 工艺网络分析**: 流程连接分析和网络构建
- **✅ 数据验证和导出**: 完整的数据验证和多格式导出功能
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
├── 💰 经济分析工具
├── extract_aspen_economics.py          # Aspen经济数据提取
├── economic_file_parser.py             # 经济文件解析器
├── economic_excel_exporter.py          # 经济数据Excel导出
├── 
├── 🔗 网络分析工具  
├── process_network_builder.py          # 工艺网络构建器
├── analyze_flowsheet_connections.py    # 流程图连接分析
├── 
├── 🧪 测试验证
├── test_*.py                           # 各种功能测试文件
├── check_*.py                          # 数据检查和验证工具
```

## 🚀 快速开始

### 🎯 一键完整数据提取

```bash
# 执行完整的数据提取和存储流程
python full_extraction.py
```

这个命令会：
- 📖 从Aspen文件(`BFG-CO2H-MEOH V2 (purge burning).apw`)提取24个流股数据
- ⚙️ 从Aspen文件提取16个设备数据，包含详细参数和连接信息
- 🔥 从Excel文件(`BFG-CO2H-HEX.xlsx`)提取13个热交换器数据
- 🏷️ 自动分类流股：原料(5个)、产品(12个)、过程流股(7个)
- 💾 将所有数据存储到SQLite数据库(`aspen_data.db`)
- 📊 生成完整的提取报告和统计信息

### 📋 检查数据完整性

```bash
# 验证数据库中的数据完整性
python check_database_completeness.py
```

### 🔍 查看提取测试

```bash
# 运行综合测试套件
python aspen_data_extractor.py
```

## 📖 详细使用方法

### 1. 完整数据提取和存储

```python
from aspen_data_extractor import AspenDataExtractor

# 创建数据提取器实例
extractor = AspenDataExtractor()

# 设置文件路径
aspen_file = "aspen_files/BFG-CO2H-MEOH V2 (purge burning).apw"
hex_file = "BFG-CO2H-HEX.xlsx"

# 执行完整的数据提取和存储
result = extractor.extract_and_store_all_data(aspen_file, hex_file)

# 查看提取结果
print(f"提取成功: {result['success']}")
print(f"会话ID: {result['session_id']}")
print(f"数据统计: {result['data_counts']}")
# 输出示例:
# 提取成功: True
# 会话ID: session_20250727_162428
# 数据统计: {'heat_exchangers': 13, 'streams': 24, 'equipment': 16}
```

### 2. 单独加载Excel热交换器数据

```python
# 加载Excel热交换器数据
success = extractor.load_hex_data("BFG-CO2H-HEX.xlsx")

if success:
    # 获取热交换器数据摘要
    hex_summary = extractor.get_hex_summary()
    print(f"热交换器数量: {hex_summary['total_heat_exchangers']}")
    print(f"总热负荷: {hex_summary['total_heat_duty']:.1f} kW")
    print(f"总传热面积: {hex_summary['total_heat_area']:.1f} m²")
    
    # 获取详细的提取报告
    extractor.print_hex_extraction_report()
```

### 3. 从Aspen Plus提取流股和设备数据

```python
# 连接到Aspen Plus
if extractor.com_interface.connect(aspen_file):
    
    # 提取所有流股数据
    streams = extractor.extract_all_streams()
    print(f"提取了 {len(streams)} 个流股")
    
    # 提取所有设备数据  
    equipment = extractor.extract_all_equipment()
    print(f"提取了 {len(equipment)} 个设备")
    
    # 断开连接
    extractor.com_interface.disconnect()
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

### 4. 经济分析功能

```python
# 提取Aspen经济数据
from extract_aspen_economics import extract_economics
economics_data = extract_economics("path/to/economics_file.izp")

# 构建工艺网络
from process_network_builder import build_network
network = build_network()

# 分析流程连接
from analyze_flowsheet_connections import analyze_connections
connections = analyze_connections()

# 导出经济分析报告
from economic_excel_exporter import export_economics
export_economics("economics_report.xlsx")
```

### 5. 数据库维护

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

## �️ 维护和故障排除

### 常见问题和解决方案

#### 1. COM连接失败
```bash
# 问题: 无法连接Aspen Plus
# 解决: 检查COM组件注册状态
python -c "
from aspen_data_extractor import AspenDataExtractor
extractor = AspenDataExtractor()
com_test = extractor.com_interface.test_com_availability()
print('COM诊断结果:', com_test)
"
```

#### 2. 数据提取不完整
```bash
# 检查数据库完整性
python check_database_completeness.py

# 重新执行完整提取
python full_extraction.py
```

#### 3. Excel热交换器数据映射问题
```python
from aspen_data_extractor import AspenDataExtractor
extractor = AspenDataExtractor()
extractor.load_hex_data("BFG-CO2H-HEX.xlsx")
# 查看详细的提取报告和映射诊断
extractor.print_hex_extraction_report()
```

### 日志和调试

```python
import logging
# 启用详细日志
logging.basicConfig(level=logging.INFO, 
                   format='%(levelname)s:%(name)s:%(message)s')

# 在提取过程中会显示详细的进度信息
```

## 🧪 测试和验证

### 运行完整测试套件

```bash
# 运行AspenDataExtractor的综合测试
python aspen_data_extractor.py

# 预期输出:
# Enhanced Aspen Data Extractor - Unified Test Suite
# =====================================================
# 1. Windows COM diagnostics... ✅ COM setup OK
# 2. Heat exchanger data loading... ✅ Loaded 13 heat exchangers  
# 3. Aspen Plus data extraction... ✅ Extracted 24 streams, 16 equipment
# 4. Equipment sizing calculations... ✅ Equipment sizing OK
# Test Results: 4/4 successful (100%)
```

### 数据库完整性检查

```bash
# 检查数据库状态和数据完整性
python check_database_completeness.py

# 预期输出:
# 🔍 检查数据库完整性
# ================================================== 
# 📋 当前数据库表:
#   - extraction_sessions: 1 条记录
#   - aspen_streams: 24 条记录
#   - aspen_equipment: 16 条记录  
#   - heat_exchangers: 13 条记录
# 🔍 检查重要功能:
#   ✅ HEX换热器数据: 13 条记录
#   ✅ 流股分类功能: 24/24 个流股已分类
#   ✅ 设备类型识别: 16/16 个设备有明确类型
```

## 🗄️ 数据库结构

工具创建的SQLite数据库(`aspen_data.db`)包含以下表结构：

### 数据表详细说明

1. **`aspen_streams`** - 流股数据 (24条记录)
   - `stream_name`: 流股名称 (如: BFG, MEOH1, AIR等)
   - `temperature`, `pressure`: 温度(°C)、压力(bar)
   - `mass_flow`, `volume_flow`, `molar_flow`: 各种流量数据
   - `composition`: JSON格式的组分数据
   - `stream_category`: 自动分类 (原料/产品/过程)
   - `stream_sub_category`: 详细子分类 (如: 高炉煤气, 甲醇产品)
   - `classification_confidence`: 分类置信度 (0.0-1.0)
   
2. **`aspen_equipment`** - 设备数据 (16条记录)
   - `equipment_name`: 设备名称 (如: B1, MEOH, MC1等)
   - `equipment_type`: 设备类型 (反应器, 压缩机, 换热器等)
   - `aspen_type`: Aspen块类型 (RSTOIC, ISENTROPIC等)
   - `parameters`: JSON格式的设备参数
   - `inlet_streams`, `outlet_streams`: 进出口流股连接
   - `importance`: 设备重要性级别
   
3. **`heat_exchangers`** - 热交换器数据 (13条记录)
   - `equipment_name`: 换热器名称
   - `duty_kw`: 热负荷 (kW)
   - `area_m2`: 传热面积 (m²)
   - `hot_stream_name`, `cold_stream_name`: 热流/冷流名称
   - `hot_inlet_temp`, `hot_outlet_temp`: 热流进出口温度
   - `cold_inlet_temp`, `cold_outlet_temp`: 冷流进出口温度
   - `inlet_streams`, `outlet_streams`: 简化的流股映射
   
4. **`extraction_sessions`** - 提取会话记录
   - `session_id`: 会话标识 (如: session_20250727_162428)
   - `extraction_time`: 提取时间戳
   - `aspen_file_path`, `hex_file_path`: 源文件路径
   - `summary_stats`: JSON格式的统计摘要

### 💾 当前数据库状态

```
📋 数据库摘要 (aspen_data.db):
  - extraction_sessions: 1 条记录 ✅
  - aspen_streams: 24 条记录 ✅
  - aspen_equipment: 16 条记录 ✅
  - heat_exchangers: 13 条记录 ✅

🏷️ 流股分类统计:
  - 产品流股: 12 (50.0%) - 包含甲醇产品、轻组分等
  - 原料流股: 5 (20.8%) - 包含高炉煤气、空气等
  - 过程流股: 7 (29.2%) - 包含工艺中间流股

⚙️ 设备类型统计:
  - 反应器: 2个 (B1-RSTOIC, MEOH-T-SPEC)
  - 换热器: 3个 (COOL2, HT8, HT9)
  - 混合器: 4个 (B11, MIX3, MX1, MX2)
  - 分离器: 2个 (S2, S3)
  - 蒸馏塔: 2个 (C-301, DI)
  - 压缩机: 1个 (MC1-ISENTROPIC)
  - 其他: 2个 (F1-分流器, V3-阀门)
```

## 🔌 Aspen Plus COM接口集成

### 系统要求

- **操作系统**: Windows (Aspen Plus COM接口要求)
- **软件**: Aspen Plus V11+ (支持COM自动化)
- **Python包**: `pywin32` (COM接口), `pandas`, `openpyxl`, `sqlite3`

### 环境配置

1. **安装Python依赖**:
   ```bash
   pip install -r requirements.txt
   ```

2. **验证Aspen Plus COM可用性**:
   ```python
   from aspen_data_extractor import AspenDataExtractor
   extractor = AspenDataExtractor()
   
   # 测试COM连接
   com_test = extractor.com_interface.test_com_availability()
   print(f"COM对象可用: {com_test['com_objects_found']}")
   ```

### 支持的数据类型和提取能力

#### ✅ 流股数据提取
- **基础属性**: 温度、压力、质量流量、体积流量、摩尔流量
- **组分信息**: 完整的物料组分分析
- **智能分类**: 自动识别原料、产品、过程流股
- **置信度评估**: 基于流股属性的分类置信度
- **自定义名称**: 提取Aspen中的用户定义显示名称

#### ✅ 设备数据提取  
- **设备类型识别**: 
  - 反应器 (RSTOIC, RPLUG, RCSTR等)
  - 换热器 (HEATX, HEATER, COOLER等)
  - 分离设备 (FLASH2, SEP, RADFRAC等)
  - 压缩设备 (COMPR, MCOMPR, ISENTROPIC等)
  - 混合分流 (MIXER, FSPLIT等)
- **操作参数**: 温度、压力、负荷、效率等关键参数
- **流股连接**: 设备的进出口流股映射关系
- **Excel匹配**: 基于预定义Excel表的设备功能匹配

#### ✅ 热交换器Excel数据
- **智能列映射**: 自动识别设备名称、热负荷、面积等列
- **温度数据提取**: 热流/冷流的进出口温度
- **流股名称映射**: 热流和冷流的流股名称识别
- **数据质量评估**: 对提取数据的完整性和准确性评估

### COM接口技术细节

```python
# 支持的Aspen Plus COM对象
COM_OBJECTS = [
    "Apwn.Document",      # 主要COM对象
    "AspenPlusDocument",  # 备用COM对象
    "Aspen.Document"      # 旧版本支持
]

# 支持的初始化方法
INIT_METHODS = [
    "InitFromArchive2",   # 首选方法
    "InitFromFile2",      # 备用方法
    "InitFromFile"        # 兼容性方法
]
```

## 📋 文件说明

### 核心文件
- `aspen_data_extractor.py` - 主要数据提取模块
- `aspen_data_database.py` - 数据库管理系统
- `data_interfaces.py` - 数据结构定义
- `stream_classifier.py` - 流股分类器
- `equipment_model_matcher.py` - 设备模型匹配

### 经济分析模块
- `extract_aspen_economics.py` - Aspen经济数据提取
- `economic_file_parser.py` - 经济文件解析器
- `economic_excel_exporter.py` - 经济数据Excel导出
- `process_network_builder.py` - 工艺网络构建器
- `analyze_flowsheet_connections.py` - 流程图连接分析

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

**版本**: 2.2 (完整数据提取增强版)  
**更新日期**: 2025-01-27  
**状态**: 🟢 完全可用  
**数据库状态**: ✅ 1会话, 24流股, 16设备, 13热交换器

### 🎯 最新更新

- ✅ **完整数据提取流程**: 实现了从Aspen文件和Excel文件的完整数据提取
- ✅ **智能流股分类**: 自动识别原料、产品、过程流股，支持置信度评估
- ✅ **设备类型识别**: 基于Excel匹配表和Aspen类型的准确设备识别
- ✅ **数据库会话管理**: 完整的提取会话记录和数据溯源
- ✅ **温度数据提取**: 热交换器的完整温度数据映射
- ✅ **连接关系分析**: 设备的进出口流股连接关系
- ✅ **COM接口优化**: 稳定的Aspen Plus COM连接和错误处理

### 📊 当前项目数据概览

**BFG-CO2H-MEOH V2工艺数据**:
- 🌊 **流股**: 24个 (原料5个, 产品12个, 过程7个)
- ⚙️ **设备**: 16个 (反应器2个, 换热器3个, 蒸馏塔2个等)
- 🔥 **热交换器**: 13个 (总热负荷1443kW, 总面积56119m²)
- 💾 **数据完整性**: 100% (所有数据表完整填充)

### 🤝 技术支持

如有问题或建议，请参考：
1. `CLAUDE.md` - 详细的开发和使用指南
2. `check_database_completeness.py` - 数据库状态检查
3. `full_extraction.py` - 一键完整数据提取

---

**许可说明**: 本工具仅供教育和研究目的使用