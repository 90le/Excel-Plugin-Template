using System;
using System.Collections.Generic;
using BasePlugin.Core;
using BasePlugin.Models;
using DTI_Tool.AddIn.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasePlugin.Features
{
    /// <summary>
    /// 日志演示功能类 - 提供日志管理相关的示例功能
    /// </summary>
    public class LoggingDemoFeatures : IFeatureProvider
    {
        #region 私有字段

        private readonly PluginLogger _logger;

        #endregion

        #region 构造函数

        public LoggingDemoFeatures(PluginLogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            
            _logger.Debug("LoggingDemoFeatures 已初始化");
        }

        #endregion

        #region 私有属性

        /// <summary>
        /// 获取Excel应用程序对象
        /// </summary>
        private Excel.Application ExcelApp => HostApplication.Instance.ExcelApplication;

        #endregion

        #region IFeatureProvider 实现

        /// <summary>
        /// 获取日志演示功能列表
        /// </summary>
        public List<PluginFeature> GetFeatures()
        {
            return new List<PluginFeature>
            {
                new PluginFeature
                {
                    Id = "logging_basic_demo",
                    Name = "基础日志演示",
                    Description = "演示各种级别的日志记录",
                    Category = "日志管理",
                    Tags = new List<string> { "日志", "调试", "信息" },
                    ImageMso = "BlogPost",
                    Action = BasicLoggingDemo
                },
                new PluginFeature
                {
                    Id = "logging_performance_demo",
                    Name = "性能测量演示",
                    Description = "演示如何测量和记录性能数据",
                    Category = "日志管理",
                    Tags = new List<string> { "性能", "测量", "计时" },
                    ImageMso = "AnimationTransition",
                    Action = PerformanceLoggingDemo
                },
                new PluginFeature
                {
                    Id = "logging_error_demo",
                    Name = "错误处理演示",
                    Description = "演示错误和异常的日志记录",
                    Category = "日志管理",
                    Tags = new List<string> { "错误", "异常", "处理" },
                    ImageMso = "WarningStyle",
                    Action = ErrorLoggingDemo
                },
                new PluginFeature
                {
                    Id = "logging_structured_demo",
                    Name = "结构化日志演示",
                    Description = "演示结构化日志记录技术",
                    Category = "日志管理",
                    Tags = new List<string> { "结构化", "格式", "数据" },
                    ImageMso = "DatabaseTable",
                    Action = StructuredLoggingDemo
                },
                new PluginFeature
                {
                    Id = "host_logger_demo",
                    Name = "宿主日志接口演示",
                    Description = "演示使用宿主应用的日志接口",
                    Category = "日志管理",
                    Tags = new List<string> { "宿主", "接口", "集成" },
                    ImageMso = "DatabaseConnectToAccess",
                    Action = HostLoggerDemo
                }
            };
        }

        public void Dispose()
        {
            _logger.Debug("LoggingDemoFeatures 已释放");
        }

        #endregion

        #region 功能实现

        /// <summary>
        /// 基础日志演示
        /// </summary>
        private void BasicLoggingDemo()
        {
            using (_logger.MeasurePerformance("基础日志演示"))
            {
                try
                {
                    _logger.Info("=== 开始基础日志演示 ===");
                    
                    // 1. Debug级别 - 用于详细的调试信息
                    _logger.Debug("这是一条Debug日志");
                    _logger.Debug("当前时间: {0}, 线程ID: {1}", 
                        DateTime.Now.ToString("HH:mm:ss.fff"), 
                        System.Threading.Thread.CurrentThread.ManagedThreadId);
                    
                    // 2. Info级别 - 用于一般信息
                    _logger.Info("这是一条Info日志");
                    _logger.Info("Excel版本: {0}", ExcelApp?.Version ?? "未知");
                    
                    // 3. Warning级别 - 用于警告信息
                    _logger.Warning("这是一条Warning日志");
                    _logger.Warning("内存使用率: {0}%", GetMemoryUsagePercent());
                    
                    // 4. Error级别 - 用于错误信息
                    _logger.Error("这是一条Error日志（非异常）");
                    
                    // 5. 带格式化的日志
                    var data = new { Name = "测试", Value = 123, Time = DateTime.Now };
                    _logger.Info("复杂数据日志: 名称={0}, 值={1}, 时间={2:yyyy-MM-dd HH:mm:ss}", 
                        data.Name, data.Value, data.Time);
                    
                    // 6. 批量日志演示
                    _logger.Info("开始批量处理演示...");
                    for (int i = 1; i <= 5; i++)
                    {
                        _logger.Debug("处理项目 {0}/5", i);
                        System.Threading.Thread.Sleep(100);
                    }
                    
                    _logger.Info("=== 基础日志演示完成 ===");
                    
                    MessageHelper.ShowInfo(
                        "基础日志演示完成！\n\n" +
                        "已生成以下类型的日志:\n" +
                        "• Debug - 调试信息\n" +
                        "• Info - 常规信息\n" +
                        "• Warning - 警告信息\n" +
                        "• Error - 错误信息\n\n" +
                        "请查看日志管理器查看详细输出。",
                        "演示完成");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "基础日志演示失败");
                    MessageHelper.ShowError($"基础日志演示失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 性能测量演示
        /// </summary>
        private void PerformanceLoggingDemo()
        {
            using (_logger.MeasurePerformance("性能测量演示"))
            {
                try
                {
                    _logger.Info("=== 开始性能测量演示 ===");
                    
                    // 1. 简单性能测量
                    using (_logger.MeasurePerformance("简单计算任务"))
                    {
                        System.Threading.Thread.Sleep(100);
                        var sum = 0;
                        for (int i = 0; i < 1000000; i++)
                        {
                            sum += i;
                        }
                        _logger.Debug("计算结果: {0}", sum);
                    }
                    
                    // 2. 嵌套性能测量
                    using (_logger.MeasurePerformance("复杂处理流程"))
                    {
                        using (_logger.MeasurePerformance("步骤1: 数据准备"))
                        {
                            System.Threading.Thread.Sleep(50);
                            _logger.Debug("数据准备完成");
                        }
                        
                        using (_logger.MeasurePerformance("步骤2: 数据处理"))
                        {
                            System.Threading.Thread.Sleep(100);
                            _logger.Debug("数据处理完成");
                        }
                        
                        using (_logger.MeasurePerformance("步骤3: 结果输出"))
                        {
                            System.Threading.Thread.Sleep(30);
                            _logger.Debug("结果输出完成");
                        }
                    }
                    
                    // 3. 并行任务性能测量
                    _logger.Info("开始并行任务性能测量");
                    var tasks = new List<System.Threading.Tasks.Task>();
                    
                    for (int i = 1; i <= 3; i++)
                    {
                        var taskId = i;
                        var task = System.Threading.Tasks.Task.Run(() =>
                        {
                            using (_logger.MeasurePerformance($"并行任务{taskId}"))
                            {
                                System.Threading.Thread.Sleep(50 * taskId);
                                _logger.Debug("任务{0}完成", taskId);
                            }
                        });
                        tasks.Add(task);
                    }
                    
                    System.Threading.Tasks.Task.WaitAll(tasks.ToArray());
                    
                    // 4. Excel操作性能测量
                    using (_logger.MeasurePerformance("Excel操作"))
                    {
                        var worksheet = ExcelApp?.ActiveSheet as Excel.Worksheet;
                        if (worksheet != null)
                        {
                            using (_logger.MeasurePerformance("写入100个单元格"))
                            {
                                for (int i = 1; i <= 100; i++)
                                {
                                    worksheet.Cells[i, 1] = $"性能测试 {i}";
                                }
                            }
                            
                            using (_logger.MeasurePerformance("清理单元格"))
                            {
                                worksheet.Range["A1:A100"].Clear();
                            }
                        }
                    }
                    
                    _logger.Info("=== 性能测量演示完成 ===");
                    
                    MessageHelper.ShowInfo(
                        "性能测量演示完成！\n\n" +
                        "演示内容:\n" +
                        "• 简单任务计时\n" +
                        "• 嵌套任务计时\n" +
                        "• 并行任务计时\n" +
                        "• Excel操作计时\n\n" +
                        "所有性能数据已记录到日志中。",
                        "演示完成");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "性能测量演示失败");
                    MessageHelper.ShowError($"性能测量演示失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 错误处理演示
        /// </summary>
        private void ErrorLoggingDemo()
        {
            using (_logger.MeasurePerformance("错误处理演示"))
            {
                _logger.Info("=== 开始错误处理演示 ===");
                
                // 1. 简单错误日志
                try
                {
                    throw new InvalidOperationException("这是一个模拟的操作异常");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "捕获到操作异常");
                }
                
                // 2. 带上下文的错误日志
                var userId = 12345;
                var operation = "数据导入";
                try
                {
                    throw new ApplicationException("模拟的应用程序异常");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "操作失败 - 用户ID: {0}, 操作: {1}", userId, operation);
                }
                
                // 3. 嵌套异常处理
                try
                {
                    try
                    {
                        throw new ArgumentException("内部异常");
                    }
                    catch (Exception inner)
                    {
                        throw new InvalidOperationException("外部异常", inner);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "嵌套异常示例");
                }
                
                // 4. 多种异常类型处理
                var exceptions = new Exception[]
                {
                    new NullReferenceException("空引用异常"),
                    new IndexOutOfRangeException("索引越界异常"),
                    new FormatException("格式化异常"),
                    new TimeoutException("超时异常")
                };
                
                foreach (var ex in exceptions)
                {
                    _logger.Error(ex, "异常类型: {0}", ex.GetType().Name);
                }
                
                // 5. 条件性错误日志
                var errorCount = 5;
                if (errorCount > 3)
                {
                    _logger.Warning("错误数量过多: {0}个错误", errorCount);
                }
                
                // 6. 错误恢复日志
                _logger.Info("尝试从错误中恢复...");
                for (int retry = 1; retry <= 3; retry++)
                {
                    try
                    {
                        if (retry < 3)
                        {
                            throw new Exception($"第{retry}次尝试失败");
                        }
                        _logger.Info("第{0}次尝试成功", retry);
                        break;
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning("重试失败: {0}", ex.Message);
                    }
                }
                
                _logger.Info("=== 错误处理演示完成 ===");
                
                MessageHelper.ShowInfo(
                    "错误处理演示完成！\n\n" +
                    "演示内容:\n" +
                    "• 简单异常日志\n" +
                    "• 带上下文的错误日志\n" +
                    "• 嵌套异常处理\n" +
                    "• 多种异常类型\n" +
                    "• 错误恢复日志\n\n" +
                    "所有错误信息已记录到日志中。",
                    "演示完成");
            }
        }

        /// <summary>
        /// 结构化日志演示
        /// </summary>
        private void StructuredLoggingDemo()
        {
            using (_logger.MeasurePerformance("结构化日志演示"))
            {
                try
                {
                    _logger.Info("=== 开始结构化日志演示 ===");
                    
                    // 1. 业务实体日志
                    var order = new
                    {
                        OrderId = "ORD-2024-001",
                        CustomerId = "CUST-123",
                        Amount = 1299.99m,
                        Items = 5,
                        Status = "Processing"
                    };
                    
                    _logger.Info("处理订单 - ID: {0}, 客户: {1}, 金额: {2:C}, 商品数: {3}, 状态: {4}",
                        order.OrderId, order.CustomerId, order.Amount, order.Items, order.Status);
                    
                    // 2. 操作上下文日志
                    var context = new
                    {
                        UserId = "USER-456",
                        SessionId = Guid.NewGuid().ToString(),
                        IpAddress = "192.168.1.100",
                        UserAgent = "Excel Plugin/1.0"
                    };
                    
                    _logger.Debug("操作上下文 - 用户: {0}, 会话: {1}, IP: {2}, 客户端: {3}",
                        context.UserId, context.SessionId, context.IpAddress, context.UserAgent);
                    
                    // 3. 性能指标日志
                    var metrics = new
                    {
                        CpuUsage = 45.2,
                        MemoryUsage = 68.5,
                        DiskUsage = 23.1,
                        ResponseTime = 125
                    };
                    
                    _logger.Info("系统指标 - CPU: {0:F1}%, 内存: {1:F1}%, 磁盘: {2:F1}%, 响应时间: {3}ms",
                        metrics.CpuUsage, metrics.MemoryUsage, metrics.DiskUsage, metrics.ResponseTime);
                    
                    // 4. 批处理结果日志
                    var batchResult = new
                    {
                        BatchId = "BATCH-789",
                        TotalItems = 1000,
                        Processed = 950,
                        Failed = 50,
                        Duration = TimeSpan.FromMinutes(5.5)
                    };
                    
                    _logger.Info("批处理完成 - ID: {0}, 总数: {1}, 成功: {2}, 失败: {3}, 耗时: {4}",
                        batchResult.BatchId, batchResult.TotalItems, 
                        batchResult.Processed, batchResult.Failed, 
                        batchResult.Duration.ToString(@"mm\:ss"));
                    
                    // 5. Excel操作日志
                    var worksheet = ExcelApp?.ActiveSheet as Excel.Worksheet;
                    if (worksheet != null)
                    {
                        var sheetInfo = new
                        {
                            Name = worksheet.Name,
                            UsedRows = worksheet.UsedRange?.Rows.Count ?? 0,
                            UsedCols = worksheet.UsedRange?.Columns.Count ?? 0,
                            ProtectionStatus = worksheet.ProtectContents ? "Protected" : "Unprotected"
                        };
                        
                        _logger.Info("工作表信息 - 名称: {0}, 行数: {1}, 列数: {2}, 保护状态: {3}",
                            sheetInfo.Name, sheetInfo.UsedRows, sheetInfo.UsedCols, sheetInfo.ProtectionStatus);
                    }
                    
                    _logger.Info("=== 结构化日志演示完成 ===");
                    
                    MessageHelper.ShowInfo(
                        "结构化日志演示完成！\n\n" +
                        "演示内容:\n" +
                        "• 业务实体日志\n" +
                        "• 操作上下文记录\n" +
                        "• 性能指标记录\n" +
                        "• 批处理结果日志\n" +
                        "• Excel操作日志\n\n" +
                        "结构化日志有助于后续的日志分析和监控。",
                        "演示完成");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "结构化日志演示失败");
                    MessageHelper.ShowError($"结构化日志演示失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 宿主日志接口演示
        /// </summary>
        private void HostLoggerDemo()
        {
            using (_logger.MeasurePerformance("宿主日志接口演示"))
            {
                try
                {
                    _logger.Info("开始演示宿主日志接口功能");
                    
                    var hostApp = HostApplication.Instance;
                    if (hostApp == null)
                    {
                        MessageHelper.ShowError("无法获取宿主应用接口");
                        return;
                    }
                    
                    // 使用宿主应用的日志功能
                    var pluginName = "BasePlugin";
                    
                    // 1. 基础日志功能
                    hostApp.LogDebug(pluginName, "通过宿主接口记录的Debug信息");
                    hostApp.LogInfo(pluginName, "通过宿主接口记录的Info信息");
                    hostApp.LogWarning(pluginName, "通过宿主接口记录的Warning信息");
                    hostApp.LogError(pluginName, "通过宿主接口记录的Error信息");
                    
                    // 2. 格式化日志
                    hostApp.LogInfoFormat(pluginName, "格式化日志 - 时间: {0}, 数值: {1}", 
                        DateTime.Now.ToString("HH:mm:ss"), 42);
                    
                    // 3. 带异常的日志
                    try
                    {
                        throw new Exception("模拟异常");
                    }
                    catch (Exception ex)
                    {
                        hostApp.LogError(pluginName, ex, "捕获异常");
                    }
                    
                    // 4. 性能测量
                    using (hostApp.StartPerformanceMeasure(pluginName, "宿主性能测量示例"))
                    {
                        System.Threading.Thread.Sleep(100);
                        hostApp.LogDebug(pluginName, "在性能测量块中执行操作");
                    }
                    
                    // 5. 条件日志
                    if (hostApp.IsDebugEnabled)
                    {
                        hostApp.LogDebug(pluginName, "Debug模式已启用，记录详细信息");
                    }
                    
                    _logger.Info("宿主日志接口演示完成");
                    
                    MessageHelper.ShowInfo(
                        "宿主日志接口演示完成！\n\n" +
                        "演示内容:\n" +
                        "• 基础日志记录\n" +
                        "• 格式化日志\n" +
                        "• 异常日志\n" +
                        "• 性能测量\n" +
                        "• 条件日志\n\n" +
                        "使用宿主日志接口可以统一管理所有插件的日志。",
                        "演示完成");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "宿主日志接口演示失败");
                    MessageHelper.ShowError($"宿主日志接口演示失败: {ex.Message}");
                }
            }
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 获取内存使用率
        /// </summary>
        private int GetMemoryUsagePercent()
        {
            // 模拟内存使用率
            var random = new Random();
            return random.Next(40, 80);
        }

        #endregion
    }
} 