# BasePlugin 任务窗格演示说明

## 概述

本演示项目为 BasePlugin 实现了现代化的任务窗格功能，包含 WPF 和 WinForms 两种格式的任务窗格，并集成了真实的插件功能调用。

## 新增功能

### 1. 现代化 WPF 任务窗格

**文件位置**: `BasePlugin/WPF/Views/TaskPaneView.xaml` 和 `TaskPaneView.xaml.cs`

**特性**:
- 🎨 现代化界面设计，采用卡片式布局
- 🎯 多种颜色主题的按钮（主要、成功、警告、次要）
- 📊 实时工作表信息显示
- ⚡ 分类功能区域（快速操作、数据处理、格式化、工作表管理、实用工具）
- 💡 智能状态提示
- ✨ 流畅的用户体验

**界面包含**:
- 工作表信息卡片：显示工作簿、工作表、选中区域、单元格数量
- 快速操作：插入时间、获取选择信息、应用样式、数据统计
- 数据处理：排序、筛选、去重、生成随机数据、填充序列
- 格式化：数字格式、条件格式、边框、突出显示
- 工作表管理：创建、重命名、保护、列表工作表
- 实用工具：导出CSV、查找替换、工作簿信息、数据清理

### 2. WinForms 任务窗格

**文件位置**: `BasePlugin/WinForms/TaskPaneControl.cs`

**特性**:
- 🏪 传统 WinForms 界面风格
- 📋 完整的功能布局
- 🔧 与 WPF 版本功能一致
- ⚙️ 适合传统应用环境

**设计特点**:
- 分组框架构清晰
- 按钮颜色编码（蓝色主要、绿色成功、黄色警告、灰色次要）
- 状态栏实时反馈
- 紧凑的布局设计

### 3. 增强的任务窗格管理器

**文件位置**: `BasePlugin/Core/TaskPaneManager.cs`

**新增功能**:
- ✅ 支持 WPF 和 WinForms 两种类型
- 🔄 独立的创建和管理方法
- 📝 详细的日志记录
- 🛡️ 完善的错误处理

**主要方法**:
```csharp
// WPF 任务窗格
ToggleWpfTaskPane(taskPaneName)
ShowWpfTaskPane(taskPaneName, visible)
CreateWpfTaskPane(taskPaneName, width)

// WinForms 任务窗格
ToggleWinFormsTaskPane(taskPaneName)
ShowWinFormsTaskPane(taskPaneName, visible)
CreateWinFormsTaskPane(taskPaneName, width)
```

### 4. 功能集成

**真实功能调用**:
- ✅ 通过 `FeatureManager` 调用实际的插件功能
- ✅ 支持所有 BasePlugin 中定义的功能
- ✅ 完整的错误处理和用户反馈
- ✅ 自动刷新工作表信息

**支持的功能类别**:
- 基础功能 (BasicFeatures)
- 数据处理 (DataProcessingFeatures)  
- 格式化 (FormattingFeatures)
- 工作表管理 (WorksheetFeatures)
- 实用工具 (UtilityFeatures)

## 使用方法

### 在功能区中使用

通过 BasePlugin 的功能区菜单，现在可以看到三个任务窗格选项：

1. **WPF任务窗格** - 显示现代化的 WPF 界面
2. **WinForms任务窗格** - 显示传统的 WinForms 界面  
3. **默认任务窗格** - 兼容性选项（使用 WPF）

### 程序化调用

```csharp
// 创建任务窗格管理器
var taskPaneManager = new TaskPaneManager(logger, "YourPluginName");

// 显示 WPF 任务窗格
taskPaneManager.ToggleWpfTaskPane("MainPane");

// 显示 WinForms 任务窗格
taskPaneManager.ToggleWinFormsTaskPane("ToolsPane");

// 关闭所有任务窗格
taskPaneManager.CloseAllTaskPanes();
```

## 技术实现

### WPF 任务窗格技术特点

1. **XAML 样式系统**
   - 自定义按钮样式（ModernButton, SecondaryButton, SuccessButton, WarningButton）
   - 现代化 GroupBox 样式
   - 统一的颜色主题

2. **事件处理**
   - 分类的按钮事件处理器
   - 统一的命令执行模式
   - 实时状态更新

3. **资源管理**
   - 自动清理功能管理器
   - 内存安全的实现

### WinForms 任务窗格技术特点

1. **动态布局**
   - TableLayoutPanel 主布局
   - 自动滚动支持
   - 响应式设计

2. **自定义控件创建**
   - 程序化控件生成
   - 统一的样式应用
   - 事件绑定管理

3. **颜色主题**
   - 基于 Bootstrap 的颜色方案
   - 一致的视觉体验

## 日志和调试

### 日志记录
- 详细的操作日志
- 性能测量
- 错误跟踪
- 用户行为记录

### 调试功能
- 状态栏实时反馈
- 详细的错误消息
- 工作表信息实时更新

## 扩展指南

### 添加新功能

1. **在对应的 Features 类中添加新功能**
2. **在任务窗格 XAML 或代码中添加按钮**
3. **设置正确的 Tag 属性为功能 ID**
4. **测试功能调用**

### 自定义样式

1. **WPF**: 修改 `TaskPaneView.xaml` 中的样式资源
2. **WinForms**: 修改 `TaskPaneControl.cs` 中的颜色和样式常量

### 添加新的任务窗格类型

1. **在 TaskPaneManager 中添加新的创建方法**
2. **在 WindowDemoFeatures 中添加相应的功能**
3. **更新功能区按钮配置**

## 兼容性说明

- ✅ 完全向后兼容现有代码
- ✅ 支持 .NET Framework 4.8.1
- ✅ 兼容 Excel VSTO 环境
- ✅ 支持多版本 Office

## 故障排除

### 常见问题

1. **任务窗格无法显示**
   - 检查宿主应用接口连接
   - 查看日志文件中的错误信息
   - 确认 VSTO 运行时正常

2. **功能无法执行**
   - 检查功能 ID 是否正确
   - 确认 FeatureManager 初始化成功
   - 查看错误提示对话框

3. **界面显示异常**
   - 检查 DPI 缩放设置
   - 确认主题兼容性
   - 重启 Excel 应用程序

### 调试技巧

1. **启用详细日志**: 在插件配置中设置日志级别为 Debug
2. **使用断点**: 在 Visual Studio 中调试任务窗格代码
3. **检查事件**: 验证按钮点击事件是否正确绑定

---

## 总结

这个任务窗格演示项目展示了如何在 VSTO 插件中实现现代化的用户界面，同时保持与传统技术的兼容性。通过 WPF 和 WinForms 两种实现，开发者可以根据需要选择最适合的技术栈。

所有功能都经过完整测试，并提供了详细的日志记录和错误处理，确保在生产环境中的稳定性和可维护性。 