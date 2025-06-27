# BasePlugin 日志使用指南

## 📋 概述

BasePlugin 项目已经集成了宿主框架的日志系统，提供了完整的日志记录功能。插件开发者可以轻松地记录各种级别的日志信息，并进行性能测量。

## 🚀 快速开始

### 1. 基本日志记录

在 BasePlugin 中，日志记录器会在 `Initialize()` 方法中自动创建：

```csharp
public void Initialize()
{
    try
    {
        // 初始化日志记录器
        _logger = PluginLog.ForPlugin(Name);
        _logger.Info("开始初始化插件");
        
        // 其他初始化代码...
        
        _logger.Info("插件初始化成功");
    }
    catch (Exception ex)
    {
        _logger?.Error(ex, "插件初始化失败");
        throw;
    }
}
```

### 2. 在功能类中使用日志

在您的功能类中，可以接收并使用日志记录器：

```csharp
public class SampleFeatures
{
    private readonly Excel.Application _excelApp;
    private readonly PluginLogger _logger;
    
    public SampleFeatures(Excel.Application excelApp, PluginLogger logger)
    {
        _excelApp = excelApp;
        _logger = logger;
    }
    
    public void ProcessData()
    {
        _logger.Info("开始处理数据");
        
        try
        {
            using (_logger.MeasurePerformance("数据处理操作"))
            {
                // 处理逻辑
                for (int i = 0; i < 1000; i++)
                {
                    // 处理每一项
                    if (i % 100 == 0)
                    {
                        _logger.Debug("处理进度: {0}%", i / 10);
                    }
                }
            }
            
            _logger.Info("数据处理完成");
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "数据处理失败");
            throw;
        }
    }
}
```

## 📊 日志级别

### Debug - 调试信息
仅在开发和调试时使用，生产环境中会被过滤：

```csharp
_logger.Debug("处理单元格 {0}", cellAddress);
_logger.Debug("当前工作表: {0}", worksheet.Name);
```

### Info - 一般信息
记录重要的操作和状态信息：

```csharp
_logger.Info("开始导入数据，共 {0} 行", rowCount);
_logger.Info("用户执行了 {0} 操作", operationName);
```

### Warning - 警告信息
记录可能的问题，但不影响程序运行：

```csharp
_logger.Warning("配置文件不存在，使用默认设置");
_logger.Warning("检测到 {0} 个空单元格", emptyCount);
```

### Error - 错误信息
记录错误和异常：

```csharp
_logger.Error("保存文件失败: {0}", fileName);
_logger.Error(ex, "处理数据时发生异常");
```

## ⏱️ 性能测量

使用 `MeasurePerformance` 方法可以自动测量操作耗时：

```csharp
// 简单的性能测量
using (_logger.MeasurePerformance("Excel数据导入"))
{
    ImportExcelData();
}

// 嵌套的性能测量
using (_logger.MeasurePerformance("完整的数据处理流程"))
{
    using (_logger.MeasurePerformance("数据验证"))
    {
        ValidateData();
    }
    
    using (_logger.MeasurePerformance("数据转换"))
    {
        TransformData();
    }
    
    using (_logger.MeasurePerformance("数据保存"))
    {
        SaveData();
    }
}
```

## 🔧 最佳实践

### 1. 在构造函数中传递日志记录器

```csharp
public class DataProcessor
{
    private readonly PluginLogger _logger;
    
    public DataProcessor(PluginLogger logger)
    {
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }
}
```

### 2. 合理使用日志级别

```csharp
// ✅ 好的做法
_logger.Debug("开始处理单元格 {0}", cellAddress);
_logger.Info("成功导入 {0} 行数据", rowCount);
_logger.Warning("发现 {0} 个无效数据", invalidCount);
_logger.Error(ex, "数据保存失败");

// ❌ 避免的做法
_logger.Info("i = {0}", i);  // 过于详细，应使用Debug
_logger.Error("用户取消操作");  // 不是错误，应使用Info
```

### 3. 包含有用的上下文信息

```csharp
// ✅ 好的做法
_logger.Error("处理工作表 '{0}' 第 {1} 行时发生错误", sheetName, rowIndex);

// ❌ 避免的做法
_logger.Error("处理失败");  // 缺少上下文
```

### 4. 避免在循环中过度记录

```csharp
// ❌ 避免的做法
for (int i = 0; i < 10000; i++)
{
    _logger.Debug("处理第 {0} 项", i);  // 会产生大量日志
}

// ✅ 好的做法
_logger.Info("开始处理 {0} 个项目", totalCount);
for (int i = 0; i < totalCount; i++)
{
    // 处理逻辑
    if (i % 1000 == 0)
    {
        _logger.Debug("已处理 {0}/{1} 项", i, totalCount);
    }
}
_logger.Info("处理完成，成功: {0}, 失败: {1}", successCount, failCount);
```

## 📖 完整示例

```csharp
using System;
using BasePlugin.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasePlugin.Features
{
    public class DataImporter
    {
        private readonly Excel.Application _excelApp;
        private readonly PluginLogger _logger;
        
        public DataImporter(Excel.Application excelApp, PluginLogger logger)
        {
            _excelApp = excelApp;
            _logger = logger;
        }
        
        public void ImportData(string filePath)
        {
            using (_logger.MeasurePerformance("数据导入"))
            {
                try
                {
                    _logger.Info("开始导入数据文件: {0}", filePath);
                    
                    if (!System.IO.File.Exists(filePath))
                    {
                        _logger.Warning("文件不存在: {0}", filePath);
                        return;
                    }
                    
                    var workbook = _excelApp.Workbooks.Open(filePath);
                    _logger.Debug("成功打开工作簿: {0}", workbook.Name);
                    
                    using (_logger.MeasurePerformance("数据处理"))
                    {
                        ProcessWorkbook(workbook);
                    }
                    
                    workbook.Close(false);
                    _logger.Info("数据导入完成");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "数据导入失败: {0}", filePath);
                    throw;
                }
            }
        }
        
        private void ProcessWorkbook(Excel.Workbook workbook)
        {
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                using (_logger.MeasurePerformance($"处理工作表 {sheet.Name}"))
                {
                    var usedRange = sheet.UsedRange;
                    if (usedRange != null)
                    {
                        _logger.Debug("工作表 '{0}' 有效范围: {1}", sheet.Name, usedRange.Address);
                        ProcessRange(usedRange);
                    }
                    else
                    {
                        _logger.Warning("工作表 '{0}' 没有数据", sheet.Name);
                    }
                }
            }
        }
        
        private void ProcessRange(Excel.Range range)
        {
            var values = range.Value2;
            if (values != null)
            {
                _logger.Info("处理 {0} 行 {1} 列数据", range.Rows.Count, range.Columns.Count);
                // 处理数据的具体逻辑...
            }
        }
    }
}
```

## 🔍 查看日志

在 Excel 中：
1. 点击 **DTI Tool** 功能区
2. 点击 **调试工具** 组中的 **日志查看器**
3. 在日志查看器中可以：
   - 按插件过滤日志
   - 按级别过滤日志
   - 搜索特定内容
   - 导出日志文件

您的插件日志将显示为 "BasePlugin" 或您自定义的插件名称。

## ❓ 常见问题

### Q: 如何在不同的类中使用同一个日志记录器？
A: 通过构造函数依赖注入的方式传递日志记录器：

```csharp
public MyFeature(Excel.Application excelApp, PluginLogger logger)
{
    _excelApp = excelApp;
    _logger = logger;
}
```

### Q: 性能测量会影响性能吗？
A: 性能测量的开销很小，但在高频操作中应避免过度使用。

### Q: 日志记录失败会影响插件运行吗？
A: 不会，日志记录是安全的，即使失败也不会中断插件的正常运行。 