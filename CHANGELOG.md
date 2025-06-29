# 变更日志

所有对项目的重要更改都会记录在此文件中。

本项目的版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)，
变更日志格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.0.0/)。

## [未发布]

### 新增
- 待添加的功能

### 修改
- 待修改的功能

### 修复
- 待修复的问题

### 移除
- 待移除的功能

## [1.0.0] - 2025-06-19

### 新增
- 🎉 初始版本发布
- 📦 完整的 BasePlugin 开发模板
- 🔌 基础插件框架实现
  - IPlugin 接口完整实现
  - 功能模块架构设计
  - 错误处理机制
- 🎨 界面技术支持
  - WinForms 支持
  - WPF 支持和 MVVM 模式
- 📚 示例功能实现
  - Hello World 示例
  - 获取选择信息功能
  - 插入当前时间功能
- 🔄 自动更新配置支持
  - 完整的 updateInfo 配置
  - 更新检查 API 规范
  - 文件完整性验证
  - 版本比较机制
- 📖 完整的开发文档
  - README.md 详细开发指南
  - 开发指南.md 快速开始指南
  - 插件更新配置指南.md 专门配置说明
  - 项目总览.md 架构说明
- 🛠️ 开发工具和配置
  - Visual Studio 项目模板
  - .gitignore 文件
  - manifest.json 配置模板
- 🌐 多平台支持
  - Microsoft Excel 兼容
  - WPS Office 兼容
  - .NET Framework 4.8.1 支持

### 技术特性
- ✅ 热拔插插件架构
- ✅ COM 对象安全管理
- ✅ 模块化功能设计
- ✅ 丰富的 Excel 操作辅助方法
- ✅ 完善的异常处理
- ✅ 性能优化支持
- ✅ 内存泄漏防护

### 文档完整性
- ✅ 中文开发文档
- ✅ 代码注释完整
- ✅ 使用示例丰富
- ✅ 最佳实践指南
- ✅ 故障排除说明

---

## 版本格式说明

版本号格式：`主版本号.次版本号.修订号`

- **主版本号**：不兼容的 API 修改
- **次版本号**：向后兼容的功能性新增  
- **修订号**：向后兼容的问题修正

## 变更类型说明

- **新增** - 新功能
- **修改** - 现有功能的变更
- **弃用** - 即将移除的功能
- **移除** - 已移除的功能
- **修复** - 问题修复
- **安全** - 安全性修复