# 子插件任务窗格功能使用指南

## 概述

本文档说明如何在子插件中使用宿主（DTI_Tool.AddIn）提供的任务窗格功能。通过这个功能，子插件可以创建和管理WinForms和WPF任务窗格，而无需直接依赖VSTO。

## 架构说明

### 接口层 (DTI_Tool.IPlugin)
- `IHostApplication` 接口中定义了任务窗格管理方法
- 提供了 WinForms 和 WPF 两种类型的任务窗格支持

### 宿主层 (DTI_Tool.AddIn)
- `HostApplicationImpl` 实现了任务窗格接口
- `ExcelTaskPane` 类提供了底层的任务窗格管理功能

### 子插件层 (BasePlugin)
- `TaskPaneManager` 类封装了任务窗格操作
- `TaskPaneView` 提供了示例WPF任务窗格界面

## 使用方法

### 1. 基本使用

```csharp
// 获取宿主应用接口
var hostApp = HostApplication.Instance;

// 创建任务窗格管理器
var taskPaneManager = new TaskPaneManager(logger, "YourPluginName");

// 切换显示任务窗格
taskPaneManager.ToggleWpfTaskPane("MainTaskPane");
```

### 2. 直接使用宿主接口

```csharp
// 获取宿主应用接口
var hostApp = HostApplication.Instance;

// 创建WPF用户控件
var wpfControl = new YourWpfUserControl();

// 创建并显示WPF任务窗格
hostApp.ToggleWpfTaskPane("TaskPaneName", wpfControl, 400);

// 创建WinForms用户控件
var winformsControl = new YourWinFormsUserControl();

// 创建并显示WinForms任务窗格
hostApp.ToggleWinFormsTaskPane("TaskPaneName", winformsControl, 400);
```

### 3. 使用工厂方法

```csharp
// 使用工厂方法创建WPF任务窗格
hostApp.ToggleWpfTaskPane("TaskPaneName", () => new YourWpfUserControl(), 400);

// 使用工厂方法创建WinForms任务窗格
hostApp.ToggleWinFormsTaskPane("TaskPaneName", () => new YourWinFormsUserControl(), 400);
```

### 4. 任务窗格管理

```csharp
// 检查任务窗格是否存在
bool exists = hostApp.TaskPaneExists("TaskPaneName");

// 检查任务窗格是否可见
bool visible = hostApp.IsTaskPaneVisible("TaskPaneName");

// 显示或隐藏任务窗格
hostApp.ShowTaskPane("TaskPaneName", true);  // 显示
hostApp.ShowTaskPane("TaskPaneName", false); // 隐藏

// 关闭任务窗格
hostApp.CloseTaskPane("TaskPaneName");

// 关闭所有任务窗格
hostApp.CloseAllTaskPanes();
```

## 可用的接口方法

### WPF任务窗格方法
- `ToggleWpfTaskPane(taskPaneName, wpfControl, width)`
- `ToggleWpfTaskPane(taskPaneName, wpfControlFactory, width)`

### WinForms任务窗格方法
- `ToggleWinFormsTaskPane(taskPaneName, control, width)`
- `ToggleWinFormsTaskPane(taskPaneName, controlFactory, width)`

### 通用管理方法
- `ShowTaskPane(taskPaneName, visible)` - 显示或隐藏任务窗格
- `TaskPaneExists(taskPaneName)` - 检查任务窗格是否存在
- `IsTaskPaneVisible(taskPaneName)` - 检查任务窗格是否可见
- `CloseTaskPane(taskPaneName)` - 关闭指定任务窗格
- `CloseAllTaskPanes()` - 关闭所有任务窗格

## TaskPaneManager 使用示例

```csharp
public class YourPluginFeatures
{
    private readonly TaskPaneManager _taskPaneManager;
    
    public YourPluginFeatures(PluginLogger logger)
    {
        _taskPaneManager = new TaskPaneManager(logger, "YourPluginName");
    }
    
    public void ShowMainTaskPane()
    {
        // 显示主任务窗格
        _taskPaneManager.ToggleWpfTaskPane("Main");
    }
    
    public void ShowSettingsPane()
    {
        // 显示设置任务窗格
        _taskPaneManager.ToggleWpfTaskPane("Settings");
    }
    
    public void CloseAllPanes()
    {
        // 关闭所有任务窗格
        _taskPaneManager.CloseAllTaskPanes();
    }
}
```

## 创建自定义WPF任务窗格

### 1. 创建WPF用户控件

```xaml
<UserControl x:Class="YourPlugin.Views.CustomTaskPane"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <Grid>
        <TextBlock Text="自定义任务窗格内容" 
                   HorizontalAlignment="Center" 
                   VerticalAlignment="Center"/>
    </Grid>
</UserControl>
```

### 2. 实现代码文件

```csharp
using System.Windows.Controls;
using DTI_Tool.AddIn.Core;

namespace YourPlugin.Views
{
    public partial class CustomTaskPane : UserControl
    {
        public CustomTaskPane()
        {
            InitializeComponent();
            InitializeData();
        }
        
        private void InitializeData()
        {
            // 获取Excel应用程序对象
            var excelApp = HostApplication.Instance?.ExcelApplication;
            
            // 初始化任务窗格内容
        }
    }
}
```

### 3. 在TaskPaneManager中注册

修改 `TaskPaneManager.CreateTaskPaneContent` 方法：

```csharp
private System.Windows.Controls.UserControl CreateTaskPaneContent(string taskPaneName)
{
    switch (taskPaneName.ToLower())
    {
        case "custom":
            return new CustomTaskPane();
        case "main":
        case "default":
            return new TaskPaneView();
        default:
            return new TaskPaneView();
    }
}
```

## 注意事项

1. **任务窗格命名**：建议使用插件名称作为前缀，避免与其他插件冲突
2. **内存管理**：任务窗格会自动管理生命周期，但建议在插件卸载时主动关闭
3. **线程安全**：任务窗格操作应在主UI线程中执行
4. **错误处理**：所有任务窗格操作都应包含适当的错误处理
5. **性能考虑**：避免创建过多的任务窗格，建议重用现有实例

## 示例项目

查看 `BasePlugin/Features/WindowDemoFeatures.cs` 中的 `ShowTaskPane` 方法，了解完整的使用示例。

## 故障排除

### 常见问题

1. **任务窗格无法显示**
   - 检查宿主应用接口是否正确初始化
   - 确认WPF控件能够正常创建
   - 查看日志文件中的错误信息

2. **任务窗格内容为空**
   - 检查WPF控件的构造函数是否有异常
   - 确认XAML文件路径正确
   - 验证数据绑定是否正常

3. **多个任务窗格冲突**
   - 使用唯一的任务窗格名称
   - 添加插件前缀避免命名冲突

### 调试技巧

1. 启用详细日志记录
2. 使用调试器检查任务窗格创建过程
3. 验证宿主应用接口的可用性
4. 测试任务窗格的生命周期管理 