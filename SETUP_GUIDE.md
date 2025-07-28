# 安装和设置指南

## 🚨 快速解决依赖问题

如果遇到 `ModuleNotFoundError: No module named 'yaml'` 或类似错误，请按以下步骤操作：

### 方法1：快速修复（推荐）
```bash
python quick_fix.py
```

### 方法2：手动安装依赖
```bash
pip install PyYAML openpyxl pandas numpy
```

### 方法3：使用requirements.txt
```bash
pip install -r requirements.txt
```

## 📋 详细安装步骤

### 1. 检查Python版本
确保您使用的是Python 3.7或更高版本：
```bash
python --version
```

### 2. 安装必需依赖
```bash
# 核心依赖
pip install PyYAML>=6.0
pip install openpyxl>=3.0.0
pip install pandas>=1.3.0
pip install numpy>=1.21.0

# Windows用户（用于Aspen Plus COM接口）
pip install pywin32>=227
```

### 3. 验证安装
运行依赖检查脚本：
```bash
python install_dependencies.py
```

### 4. 测试功能
运行示例程序：
```bash
python examples/extract_economics_example.py
```

## 🔧 常见问题解决

### 问题1：yaml模块找不到
**错误**: `ModuleNotFoundError: No module named 'yaml'`

**解决方案**:
```bash
pip install PyYAML
```

**或者使用JSON配置文件**:
```bash
python extract_aspen_economics.py --config config/economic_extraction_config.json
```

### 问题2：openpyxl模块问题
**错误**: `ModuleNotFoundError: No module named 'openpyxl'`

**解决方案**:
```bash
pip install openpyxl
```

### 问题3：Windows COM接口问题
**错误**: COM接口相关错误

**解决方案**:
1. 安装pywin32: `pip install pywin32`
2. 以管理员身份运行Python
3. 确保Aspen Plus已安装并注册

### 问题4：权限问题
**错误**: 权限被拒绝

**解决方案**:
1. 以管理员身份运行命令提示符
2. 或使用虚拟环境：
```bash
python -m venv venv
venv\Scripts\activate  # Windows
pip install -r requirements.txt
```

## 🎯 使用建议

### 新手用户
1. 先运行 `python quick_fix.py`
2. 使用JSON配置文件：`config/economic_extraction_config.json`
3. 从示例开始：`python examples/extract_economics_example.py`

### 高级用户
1. 安装完整依赖：`pip install -r requirements.txt`
2. 使用YAML配置：`config/economic_extraction_config.yaml`
3. 自定义配置和脚本

## 📱 联系支持

如果以上方法都无法解决问题，请：

1. 检查Python和pip版本
2. 尝试在虚拟环境中安装
3. 提供完整的错误信息
4. 说明您的操作系统和Python版本

## 🔄 更新依赖

定期更新依赖包以获得最新功能：
```bash
pip install --upgrade PyYAML openpyxl pandas numpy
```

## ✅ 安装验证清单

- [ ] Python 3.7+ 已安装
- [ ] PyYAML 已安装
- [ ] openpyxl 已安装  
- [ ] pandas 已安装
- [ ] numpy 已安装
- [ ] pywin32 已安装（Windows用户）
- [ ] 示例程序可以运行
- [ ] 配置文件可以加载

完成所有检查项后，您就可以开始使用Aspen经济数据提取工具了！