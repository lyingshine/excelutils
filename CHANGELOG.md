# 更新日志

## [2.0.0] - 2025-09-22

### 🎉 重大重构
- **模块化架构**: 采用现代化分层架构设计
- **代码重组**: 按功能类型重新组织项目结构
- **服务层**: 统一服务层架构，消除代码重复

### ✨ 新增功能
- **processors/** - 数据处理器模块，统一数据处理入口
- **exporters/** - 导出器模块，专门处理各种导出功能
- **cli/** - 命令行工具模块，独立的CLI工具集合
- **core/** - 核心业务逻辑层，包含所有核心算法
- **services/** - 业务服务层，整合和协调各个核心组件
- **models/** - 数据模型层，定义项目中使用的数据结构
- **app/** - 应用程序层，高级应用程序接口

### 🔧 优化改进
- **清晰职责**: 每个模块职责明确，遵循单一职责原则
- **代码复用**: 核心逻辑统一实现，避免重复代码
- **向后兼容**: 保持所有原有API接口不变
- **类型注解**: 完善的类型提示和文档
- **错误处理**: 改进的异常处理和日志记录

### 📚 文档完善
- **README.md**: 完整的项目说明文档
- **API文档**: 详细的模块和方法说明
- **架构图**: 清晰的项目结构说明
- **使用示例**: 丰富的代码使用示例

### 🐛 修复问题
- 修复导入路径问题
- 解决模块依赖关系
- 优化错误提示信息
- 改进日志输出格式

### ✅ 验证测试
- 所有功能模块导入测试通过
- GUI应用程序完全正常运行
- 命令行工具功能验证通过
- 业务逻辑100%保持不变
- 数据处理流程完整验证

### 📁 文件变更
```
新增文件:
+ app/__init__.py, app/application.py
+ cli/__init__.py, cli/generate_profit_table.py (重命名)
+ core/__init__.py, core/data_extractor.py, core/data_filter.py
+ core/price_matcher.py, core/profit_calculator.py
+ exporters/__init__.py, exporters/excel_exporter.py (重命名)
+ models/__init__.py, models/data_models.py
+ processors/__init__.py, processors/data_processor.py
+ services/__init__.py, services/data_service.py, services/excel_service.py
+ README.md, CHANGELOG.md

修改文件:
~ main.py, gui/main_window.py
~ .gitignore

删除文件:
- data_processor.py (重构到 processors/)
- excel_exporter.py (移动到 exporters/)
- generate_profit_table.py (移动到 cli/)
```

---

## [1.0.0] - 2025-09-XX

### 🎉 初始版本
- 基础毛利表生成功能
- Excel数据导入导出
- GUI用户界面
- 命令行工具
- 价格匹配算法
- 数据筛选规则