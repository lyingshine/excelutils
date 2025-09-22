# 毛利表生成器 (Excel Utils)

[![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Status](https://img.shields.io/badge/Status-Production-brightgreen.svg)]()

一个专业的Excel数据处理工具，专门用于自动生成和管理商品毛利表，特别适用于自行车等商品的价格管理和毛利分析。

## 🚀 功能特性

### 核心功能
- **📊 数据导入** - 支持Excel格式的原始商品数据导入
- **🔍 智能处理** - 自动提取商品配置、尺寸(24寸/26寸/27.5寸)、速别、颜色等信息
- **📈 毛利表生成** - 按配置分组生成标准化毛利表
- **💰 价格更新** - 支持导入修改后的毛利表，自动更新原始数据价格
- **📏 多尺寸支持** - 智能处理不同规格，27.5寸商品自动加价20元

### 智能算法
- **价格匹配逻辑**：直接匹配 → 尺寸规范化匹配 → 自动价格调整
- **数据筛选规则**：优先26寸规格 → 同配置最低价格 → 自动去重排序
- **配置提取算法**：从商品简称中智能提取配置、尺寸、速别信息

## 🏗️ 项目架构

```
excelutils/
├── 📄 main.py                    # 主程序入口
├── 📄 启动毛利表生成器.bat        # 快速启动脚本
├── 📄 requirements.txt           # 依赖包列表
│
├── 📂 processors/                # 数据处理器模块
│   └── data_processor.py         # 核心数据处理器
│
├── 📂 exporters/                 # 导出器模块
│   └── excel_exporter.py         # Excel导出功能
│
├── 📂 cli/                       # 命令行工具模块
│   └── generate_profit_table.py  # 命令行版本
│
├── 📂 core/                      # 核心业务逻辑层
│   ├── data_extractor.py         # 数据提取器
│   ├── data_filter.py            # 数据筛选器
│   ├── price_matcher.py          # 价格匹配器
│   └── profit_calculator.py      # 毛利计算器
│
├── 📂 services/                  # 业务服务层
│   ├── data_service.py           # 数据服务
│   └── excel_service.py          # Excel服务
│
├── 📂 models/                    # 数据模型层
│   └── data_models.py            # 数据结构定义
│
├── 📂 app/                       # 应用程序层
│   └── application.py            # 应用程序主类
│
├── 📂 gui/                       # 用户界面层
│   └── main_window.py            # GUI界面
│
├── 📂 utils/                     # 工具类
│   └── logger.py                 # 日志工具
│
├── 📂 config/                    # 配置管理
│   └── settings.py               # 配置文件
│
└── 📂 logs/                      # 日志目录
    └── app.log                   # 应用日志
```

## 🛠️ 安装与使用

### 环境要求
- Python 3.7+
- Windows/macOS/Linux

### 快速开始

1. **克隆项目**
```bash
git clone <repository-url>
cd excelutils
```

2. **安装依赖**
```bash
pip install -r requirements.txt
```

3. **启动应用**

**GUI版本（推荐）**：
```bash
python main.py
# 或双击运行
启动毛利表生成器.bat
```

**命令行版本**：
```bash
python cli/generate_profit_table.py
```

### 使用流程

1. **准备数据** - 将商品数据整理为Excel格式，包含ID、简称、价格、成本等列
2. **导入数据** - 使用GUI或命令行导入Excel文件
3. **生成毛利表** - 系统自动处理数据并生成毛利表
4. **调整价格** - 在生成的毛利表中修改价格（可选）
5. **更新原数据** - 导入修改后的毛利表，更新原始数据价格

## 💻 开发指南

### 模块化架构

项目采用现代化的分层架构设计：

- **models/** - 数据模型定义
- **core/** - 核心业务逻辑实现
- **services/** - 业务服务层，整合核心功能
- **processors/exporters/cli/** - 功能模块，按类型分类
- **app/gui/** - 应用程序层和用户界面

### API使用示例

```python
# 使用数据处理器
from processors.data_processor import DataProcessor

processor = DataProcessor()
data = processor.import_excel_data("data.xlsx")
processed_data = processor.process_data(data)
profit_table = processor.generate_profit_table(processed_data)

# 使用服务层
from services.data_service import DataService

service = DataService()
result = service.process_complete_workflow("data.xlsx")

# 使用核心组件
from core.data_extractor import DataExtractor
from core.profit_calculator import ProfitCalculator

extractor = DataExtractor()
calculator = ProfitCalculator()
```

### 扩展开发

1. **添加新的数据提取规则** - 修改 `core/data_extractor.py`
2. **自定义筛选逻辑** - 扩展 `core/data_filter.py`
3. **新增导出格式** - 在 `exporters/` 中添加新的导出器
4. **创建新的CLI工具** - 在 `cli/` 中添加新的命令行工具

## 📊 核心算法

### 价格匹配算法
```
1. 直接用简称匹配毛利表
2. 失败时将24寸、27.5寸规范化为26寸再匹配
3. 27.5寸商品在26寸价格基础上加20元
4. 记录详细匹配日志
```

### 数据筛选规则
```
1. 优先保留26寸规格数据
2. 同配置多价格时选择最低价格
3. 自动去重和数据清洗
4. 按配置最高价格升序排列
```

## 🔧 技术栈

- **核心语言**: Python 3.7+
- **数据处理**: pandas, openpyxl
- **用户界面**: tkinter
- **日志系统**: logging
- **架构模式**: 分层架构, 服务层模式
- **设计原则**: 单一职责, 依赖注入, 开闭原则

## 📝 更新日志

### v2.0.0 (2025-09-22)
- 🎉 **重大重构**: 采用现代化模块架构
- ✨ **新增功能**: 分层服务架构，提高代码复用性
- 🔧 **优化**: 按功能类型重新组织代码结构
- 📚 **改进**: 完善文档和类型注解
- 🐛 **修复**: 解决导入路径和兼容性问题

### v1.0.0
- 🎉 初始版本发布
- ✨ 基础毛利表生成功能
- 📊 Excel数据导入导出
- 🖥️ GUI用户界面

## 🤝 贡献指南

1. Fork 项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情

## 🆘 支持与反馈

- 📧 邮箱: [your-email@example.com]
- 🐛 问题反馈: [GitHub Issues](https://github.com/your-username/excelutils/issues)
- 📖 文档: [项目Wiki](https://github.com/your-username/excelutils/wiki)

## 🙏 致谢

感谢所有为这个项目做出贡献的开发者和用户！

---

**毛利表生成器** - 让商品价格管理更简单、更智能！ 🚀