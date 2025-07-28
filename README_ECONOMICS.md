# Aspen经济数据提取工具

这是一个专业的Aspen Plus经济数据提取和TEA（技术经济分析）工具包，能够从多种数据源提取经济参数，生成详细的Excel分析报告。

## 🌟 主要功能

### 数据源支持
- **Aspen Plus COM接口** - 实时从运行中的Aspen Plus仿真提取数据
- **IZP文件** - Aspen Icarus Cost Estimator项目文件解析
- **SZP文件** - Aspen Process Economic Analyzer数据文件解析  
- **APW文件** - Aspen Plus工作文件（通过AspenDataExtractor）

### 经济数据提取
- **CAPEX分析** - 设备成本、安装费用、间接成本、应急费用
- **OPEX分析** - 原料成本、公用工程费用、人工成本、维护费用
- **财务分析** - NPV、IRR、投资回收期、生产成本
- **设备清单** - 详细的设备尺寸、成本估算和技术参数
- **敏感性分析** - 关键参数对经济指标的影响分析

### Excel报告生成
- **多工作表报告** - 项目概览、成本分解、财务分析等8个专业工作表
- **专业图表** - 饼图、柱状图、趋势图等数据可视化
- **详细计算** - 完整的计算逻辑、参数和假设条件
- **格式化输出** - 专业的表格格式、样式和布局

## 📋 系统要求

### 基础要求
- Python 3.7+
- Windows 10/11（用于Aspen Plus COM接口）
- 必需的Python包：
  ```bash
  pip install openpyxl pandas pydantic pathlib
  ```

### Aspen Plus集成（可选）
- Aspen Plus V11+（用于COM接口功能）
- pywin32（Windows COM支持）
  ```bash
  pip install pywin32
  ```

## 🚀 快速开始

### 1. 安装依赖
```bash
# 安装基础依赖
pip install -r requirements.txt

# Windows用户安装COM支持
pip install pywin32
```

### 2. 基本使用

#### 从IZP文件提取经济数据
```bash
python extract_aspen_economics.py \
  --source "path/to/cost_file.izp" \
  --output "economic_report.xlsx" \
  --project-name "My Project"
```

#### 从Aspen Plus COM接口提取
```bash
# 确保Aspen Plus正在运行
python extract_aspen_economics.py \
  --source aspen_com \
  --output "live_analysis.xlsx" \
  --project-name "Live Simulation"
```

#### 使用配置文件
```bash
python extract_aspen_economics.py \
  --source "cost_file.szp" \
  --output "report.xlsx" \
  --config "config/economic_extraction_config.yaml"
```

### 3. Python脚本使用

```python
from extract_aspen_economics import AspenEconomicsExtractor

# 创建提取器
extractor = AspenEconomicsExtractor()

# 提取并生成报告
result = extractor.extract_and_export(
    data_source="path/to/cost_file.izp",
    output_file="economic_analysis.xlsx",
    project_name="My Economic Analysis"
)

if result['success']:
    print(f"报告生成成功: {result['report_path']}")
    print(f"总CAPEX: ${result['total_capex']:,.0f}")
    print(f"年OPEX: ${result['annual_opex']:,.0f}")
```

## 📊 输出报告结构

生成的Excel报告包含以下工作表：

### 1. Executive Summary（项目概览）
- 项目基本信息
- 关键财务指标
- 数据源摘要
- CAPEX/OPEX对比图表

### 2. CAPEX Breakdown（资本支出分解）
- 设备成本明细
- 安装和间接费用
- 成本分解柱状图
- 设备成本排序

### 3. OPEX Analysis（运营支出分析）
- 原料成本分析
- 公用工程费用
- 人工和维护成本
- 年度OPEX分解图

### 4. Equipment Details（设备详细信息）
- 设备尺寸参数
- 设计条件
- 成本估算基础
- 材料和压力等级

### 5. Financial Analysis（财务分析）
- 财务参数设置
- 经济指标计算
- 现金流分析
- 投资回收期计算

### 6. Sensitivity Analysis（敏感性分析）
- 关键参数变化影响
- NPV敏感性矩阵
- 风险评估
- 参数重要性排序

### 7. Calculation Parameters（计算参数）
- 成本估算方法
- 设备成本相关式
- 标准系数和比率
- 工程假设条件

### 8. Assumptions & Notes（假设和备注）
- 分析假设条件
- 数据来源说明
- 分析局限性
- 建议和注意事项

## ⚙️ 配置文件说明

`config/economic_extraction_config.yaml`包含详细的配置选项：

### 成本因子配置
```yaml
cost_factors:
  installation_factor: 2.5    # 设备安装因子
  engineering_rate: 0.12      # 工程设计费率
  construction_rate: 0.08     # 施工管理费率
  contingency_rate: 0.15      # 项目应急费率
```

### 财务参数
```yaml
financial_parameters:
  project_life: 20           # 项目生命周期
  discount_rate: 0.10        # 折现率
  tax_rate: 0.25            # 所得税率
  depreciation_life: 10      # 折旧年限
```

### 公用工程价格
```yaml
utility_prices:
  electricity: 0.08          # $/kWh
  steam_low_pressure: 25.0   # $/MT
  cooling_water: 0.05        # $/m³
  fuel_gas: 8.0             # $/GJ
```

### 设备成本相关式
```yaml
equipment_costing:
  reactor:
    base_cost: 50000
    scaling_factor: 0.6
    size_parameter: "volume_m3"
```

## 🔧 高级功能

### 1. 批量处理
```python
files = ["file1.izp", "file2.szp", "file3.izp"]
for file in files:
    result = extractor.extract_and_export(
        data_source=file,
        output_file=f"report_{Path(file).stem}.xlsx"
    )
```

### 2. 自定义成本模型
```python
# 通过配置文件修改设备成本相关式
custom_config = {
    'equipment_costing': {
        'reactor': {
            'base_cost': 75000,
            'scaling_factor': 0.65,
            'size_parameter': 'volume_m3'
        }
    }
}
```

### 3. 数据验证和质量检查
```python
# 获取提取摘要和质量报告
summary = extractor.get_extraction_summary()
print(f"解析错误: {summary['economic_parser_errors']}")
print(f"数据警告: {summary['economic_parser_warnings']}")
```

## 📈 应用案例

### 案例1：化工工艺TEA分析
```bash
# 从Aspen Plus仿真提取经济数据
python extract_aspen_economics.py \
  --source "methanol_plant.apw" \
  --output "methanol_tea_analysis.xlsx" \
  --project-name "甲醇工厂TEA分析" \
  --hex-file "heat_exchangers.xlsx"
```

### 案例2：多工艺方案比较
```python
scenarios = [
    {"file": "scenario_1.izp", "name": "基础工艺"},
    {"file": "scenario_2.izp", "name": "改进工艺"},
    {"file": "scenario_3.izp", "name": "新技术工艺"}
]

for scenario in scenarios:
    result = extractor.extract_and_export(
        data_source=scenario["file"],
        output_file=f"{scenario['name']}_analysis.xlsx",
        project_name=scenario["name"]
    )
```

### 案例3：实时仿真监控
```python
# 连接到运行中的Aspen Plus仿真
while simulation_running:
    result = extractor.extract_from_aspen_com(
        project_name=f"实时分析_{datetime.now().strftime('%H%M%S')}"
    )
    
    # 监控关键经济指标
    if result.npv < threshold:
        send_alert("NPV低于阈值")
```

## 🚨 故障排除

### 常见问题

#### 1. COM接口连接失败
**问题**: `无法连接到Aspen Plus`
**解决方案**:
- 确保Aspen Plus正在运行
- 以管理员身份运行Python脚本
- 重新注册COM组件：`regsvr32 apwn.exe`

#### 2. IZP文件解析失败
**问题**: `不支持的文件格式`
**解决方案**:
- 确认文件是有效的IZP/SZP格式
- 检查文件是否损坏
- 尝试从Aspen软件重新导出文件

#### 3. Excel报告生成错误
**问题**: `openpyxl模块错误`
**解决方案**:
```bash
pip install --upgrade openpyxl
pip install pandas
```

#### 4. 成本估算不合理
**问题**: 设备成本过高或过低
**解决方案**:
- 检查配置文件中的成本相关式
- 验证设备尺寸参数的合理性
- 调整成本基准年份和地区系数

### 调试模式
```bash
# 启用详细日志输出
python extract_aspen_economics.py \
  --source "file.izp" \
  --output "report.xlsx" \
  --verbose
```

## 📖 示例代码

查看 `examples/extract_economics_example.py` 获取完整的使用示例：

- 示例1：从IZP文件提取
- 示例2：COM接口实时提取
- 示例3：自定义配置使用
- 示例4：分步骤处理
- 示例5：批量文件处理

## 🤝 贡献指南

欢迎贡献代码和改进建议：

1. Fork此项目
2. 创建特性分支：`git checkout -b feature/new-feature`
3. 提交更改：`git commit -am 'Add new feature'`
4. 推送分支：`git push origin feature/new-feature`
5. 提交Pull Request

### 开发环境设置
```bash
# 克隆项目
git clone <repository-url>
cd aspen-data-extractor

# 安装开发依赖
pip install -r requirements.txt
pip install -r requirements-dev.txt

# 运行测试
python -m pytest tests/
```

## 📝 许可证

此项目采用MIT许可证 - 详见 [LICENSE](LICENSE) 文件。

## 🙏 致谢

- AspenTech公司提供的Aspen Plus软件平台
- Python开源社区的openpyxl、pandas等库
- 工程经济学和TEA分析方法论

## 📞 支持

如有问题或建议，请：

1. 查看[故障排除](#故障排除)部分
2. 提交[GitHub Issue](https://github.com/your-repo/issues)
3. 联系开发团队

---

**版本**: 1.0  
**更新时间**: 2025-07-27  
**作者**: TEA Analysis Framework