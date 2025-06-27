using System;
using System.Collections.Generic;
using BasePlugin.Models;
using BasePlugin.Core;
using DTI_Tool.AddIn.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasePlugin.Features
{
    /// <summary>
    /// 示例功能类 - 提供基础功能示例
    /// </summary>
    public class SampleFeatures
    {
        private readonly Excel.Application _excelApp;
        private readonly PluginLogger _logger;
        
        public SampleFeatures(Excel.Application excelApp, PluginLogger logger)
        {
            _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            
            _logger.Debug("SampleFeatures 已初始化");
        }
        
        /// <summary>
        /// 获取所有示例功能
        /// </summary>
        public List<PluginFeature> GetFeatures()
        {
            using (_logger.MeasurePerformance("获取示例功能列表"))
            {
                var features = new List<PluginFeature>
                {
                    new PluginFeature
                    {
                        Id = "hello_world",
                        Name = "Hello World",
                        Description = "显示一个简单的问候消息",
                        Category = "示例功能",
                        Tags = new List<string> { "示例", "基础" },
                        ImageMso = "HappyFace",
                        Action = HelloWorld
                    },
                    new PluginFeature
                    {
                        Id = "get_selection_info",
                        Name = "获取选择信息",
                        Description = "显示当前选中区域的基本信息",
                        Category = "示例功能",
                        Tags = new List<string> { "选择", "信息" },
                        ImageMso = "TableExcelSelect",
                        Action = GetSelectionInfo
                    },
                    new PluginFeature
                    {
                        Id = "insert_current_time",
                        Name = "插入当前时间",
                        Description = "在活动单元格插入当前日期和时间",
                        Category = "示例功能",
                        Tags = new List<string> { "时间", "插入" },
                        ImageMso = "DateAndTimePicker",
                        Action = InsertCurrentTime
                    },
                    new PluginFeature
                    {
                        Id = "test_logging",
                        Name = "日志测试",
                        Description = "测试所有类型的日志输出功能",
                        Category = "示例功能",
                        Tags = new List<string> { "日志", "测试", "调试" },
                        ImageMso = "BlogPost",
                        Action = TestLogging
                    },
                    new PluginFeature
                    {
                        Id = "complex_data_processing",
                        Name = "复杂数据处理示例",
                        Description = "演示复杂数据处理过程中的详细日志记录",
                        Category = "数据处理",
                        Tags = new List<string> { "数据", "处理", "日志" },
                        ImageMso = "DatabaseSortDescending",
                        Action = ComplexDataProcessing
                    },
                    new PluginFeature
                    {
                        Id = "error_simulation",
                        Name = "错误模拟",
                        Description = "模拟各种错误情况并演示错误日志记录",
                        Category = "实用工具",
                        Tags = new List<string> { "错误", "模拟", "测试" },
                        ImageMso = "WarningStyle",
                        Action = ErrorSimulation
                    },
                    new PluginFeature
                    {
                        Id = "host_logger_demo",
                        Name = "宿主日志接口演示",
                        Description = "演示 IHostApplication 接口的增强日志功能",
                        Category = "实用工具",
                        Tags = new List<string> { "宿主", "日志", "接口", "演示" },
                        ImageMso = "DatabaseConnectToAccess",
                        Action = HostLoggerDemo
                    }
                };
                
                _logger.Debug("创建了 {0} 个示例功能", features.Count);
                return features;
            }
        }
        
        #region 示例功能实现
        
        /// <summary>
        /// Hello World 示例
        /// </summary>
        private void HelloWorld()
        {
            using (_logger.MeasurePerformance("Hello World 功能"))
            {
                try
                {
                    _logger.Info("执行 Hello World 功能");
                    
                    System.Windows.Forms.MessageBox.Show(
                        "Hello World! 这是一个基础插件示例。\n\n您可以基于这个模板开发自己的Excel插件。", 
                        "基础插件示例", 
                        System.Windows.Forms.MessageBoxButtons.OK, 
                        System.Windows.Forms.MessageBoxIcon.Information);
                    
                    _logger.Info("Hello World 功能执行完成");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Hello World 功能执行失败");
                    ShowError($"Hello World 功能执行失败: {ex.Message}");
                }
            }
        }
        
        /// <summary>
        /// 获取选择信息
        /// </summary>
        private void GetSelectionInfo()
        {
            using (_logger.MeasurePerformance("获取选择信息"))
            {
                try
                {
                    _logger.Info("开始获取选择信息");
                    
                    var selection = _excelApp?.Selection as Excel.Range;
                    if (selection == null)
                    {
                        _logger.Warning("用户未选择任何区域");
                        ShowMessage("请先选择一个区域");
                        return;
                    }
                    
                    _logger.Debug("选中区域地址: {0}", selection.Address);
                    
                    var info = $"选中区域信息:\n" +
                              $"地址: {selection.Address}\n" +
                              $"行数: {selection.Rows.Count}\n" +
                              $"列数: {selection.Columns.Count}\n" +
                              $"单元格数: {selection.Cells.Count}";
                    
                    ShowMessage(info);
                    _logger.Info("成功显示选择信息，区域: {0}, 单元格数: {1}", 
                        selection.Address, selection.Cells.Count);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "获取选择信息失败");
                    ShowError($"获取选择信息失败: {ex.Message}");
                }
            }
        }
        
        /// <summary>
        /// 插入当前时间
        /// </summary>
        private void InsertCurrentTime()
        {
            using (_logger.MeasurePerformance("插入当前时间"))
            {
                try
                {
                    _logger.Info("开始插入当前时间");
                    
                    var activeCell = _excelApp?.ActiveCell;
                    if (activeCell == null)
                    {
                        _logger.Warning("没有活动单元格");
                        ShowMessage("请先选择一个单元格");
                        return;
                    }
                    
                    var currentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    _logger.Debug("活动单元格地址: {0}, 插入时间: {1}", activeCell.Address, currentTime);
                    
                    activeCell.Value = currentTime;
                    ShowMessage("已在活动单元格插入当前时间");
                    _logger.Info("成功在单元格 {0} 插入时间: {1}", activeCell.Address, currentTime);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "插入时间失败");
                    ShowError($"插入时间失败: {ex.Message}");
                }
            }
        }
        
        /// <summary>
        /// 日志测试功能 - 演示所有类型的日志输出
        /// </summary>
        private void TestLogging()
        {
            using (_logger.MeasurePerformance("日志系统全面测试"))
            {
                try
                {
                    _logger.Info("=== 开始日志系统测试 ===");
                    
                    // 1. Debug 日志测试
                    _logger.Debug("这是一条调试信息，通常包含详细的程序执行信息");
                    _logger.Debug("调试信息示例 - 当前时间: {0}, 线程ID: {1}", 
                        DateTime.Now.ToString("HH:mm:ss.fff"), 
                        System.Threading.Thread.CurrentThread.ManagedThreadId);
                    
                    // 2. Info 日志测试
                    _logger.Info("这是一条信息日志，记录程序的正常运行状态");
                    _logger.Info("用户触发了日志测试功能，Excel版本: {0}", _excelApp?.Version ?? "未知");
                    
                    // 3. Warning 日志测试
                    _logger.Warning("这是一条警告信息，提示可能存在的问题");
                    _logger.Warning("检测到内存使用率较高: {0}%", GetFakeMemoryUsage());
                    
                    // 4. 模拟一些处理步骤
                    _logger.Info("开始模拟数据处理流程...");
                    for (int i = 1; i <= 5; i++)
                    {
                        _logger.Debug("处理步骤 {0}/5: 正在处理数据批次", i);
                        System.Threading.Thread.Sleep(100); // 模拟处理时间
                        
                        if (i == 3)
                        {
                            _logger.Warning("步骤 {0}: 检测到数据质量问题，但继续处理", i);
                        }
                    }
                    
                    // 5. 性能测量嵌套测试
                    using (_logger.MeasurePerformance("复杂计算操作"))
                    {
                        _logger.Debug("开始执行复杂计算...");
                        
                        // 模拟复杂计算
                        double result = 0;
                        for (int i = 0; i < 100000; i++)
                        {
                            result += Math.Sin(i) * Math.Cos(i);
                        }
                        
                        _logger.Debug("复杂计算完成，结果: {0:F6}", result);
                    }
                    
                    // 6. 错误日志测试（模拟异常）
                    try
                    {
                        _logger.Debug("尝试访问可能不存在的资源...");
                        throw new InvalidOperationException("这是一个模拟的错误，用于测试错误日志记录");
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "捕获到模拟异常，错误代码: {0}", "ERR_001");
                    }
                    
                    _logger.Info("=== 日志系统测试完成 ===");
                    
                    // 显示完成消息
                    ShowMessage("日志测试完成！\n\n" +
                              "已生成以下类型的日志:\n" +
                              "• Debug 调试信息\n" +
                              "• Info 常规信息\n" +
                              "• Warning 警告信息\n" +
                              "• Error 错误信息（含异常）\n" +
                              "• 性能测量记录\n\n" +
                              "请查看日志管理器以查看详细输出。");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "日志测试功能执行失败");
                    ShowError($"日志测试失败: {ex.Message}");
                }
            }
        }
        
        /// <summary>
        /// 复杂数据处理示例 - 演示实际工作场景中的日志记录
        /// </summary>
        private void ComplexDataProcessing()
        {
            using (_logger.MeasurePerformance("复杂数据处理流程"))
            {
                try
                {
                    _logger.Info("开始复杂数据处理流程");
                    
                    // 1. 数据验证阶段
                    using (_logger.MeasurePerformance("数据验证"))
                    {
                        _logger.Debug("检查工作簿状态...");
                        var workbook = _excelApp?.ActiveWorkbook;
                        if (workbook == null)
                        {
                            _logger.Warning("没有打开的工作簿，创建新工作簿");
                            workbook = _excelApp?.Workbooks.Add();
                        }
                        
                        var worksheet = workbook.ActiveSheet as Excel.Worksheet;
                        _logger.Debug("使用工作表: {0}", worksheet?.Name ?? "未知");
                        
                        // 检查选中区域
                        var selection = _excelApp?.Selection as Excel.Range;
                        if (selection == null || selection.Cells.Count < 2)
                        {
                            _logger.Warning("选中区域为空或数据不足，使用默认区域 A1:C10");
                            selection = worksheet?.Range["A1:C10"];
                        }
                        
                        _logger.Info("数据验证完成 - 处理区域: {0}, 单元格数: {1}", 
                            selection?.Address, selection?.Cells.Count);
                    }
                    
                    // 2. 数据生成阶段
                    using (_logger.MeasurePerformance("数据生成"))
                    {
                        _logger.Info("开始生成示例数据...");
                        var worksheet = _excelApp?.ActiveWorkbook?.ActiveSheet as Excel.Worksheet;
                        
                        if (worksheet != null)
                        {
                            _logger.Debug("在工作表中填充示例数据");
                            var random = new Random();
                            
                            // 添加表头
                            worksheet.Cells[1, 1] = "序号";
                            worksheet.Cells[1, 2] = "数值";
                            worksheet.Cells[1, 3] = "状态";
                            _logger.Debug("已创建表头");
                            
                            // 生成数据
                            int rowCount = 10;
                            for (int i = 1; i <= rowCount; i++)
                            {
                                worksheet.Cells[i + 1, 1] = i;
                                worksheet.Cells[i + 1, 2] = random.Next(1, 100);
                                worksheet.Cells[i + 1, 3] = random.NextDouble() > 0.7 ? "异常" : "正常";
                                
                                if (i % 3 == 0)
                                {
                                    _logger.Debug("已生成 {0}/{1} 行数据", i, rowCount);
                                }
                            }
                            
                            _logger.Info("数据生成完成 - 共生成 {0} 行数据", rowCount);
                        }
                    }
                    
                    // 3. 数据分析阶段
                    using (_logger.MeasurePerformance("数据分析"))
                    {
                        _logger.Info("开始数据分析...");
                        var worksheet = _excelApp?.ActiveWorkbook?.ActiveSheet as Excel.Worksheet;
                        
                        if (worksheet != null)
                        {
                            int normalCount = 0;
                            int abnormalCount = 0;
                            double totalValue = 0;
                            
                            // 分析数据
                            for (int i = 2; i <= 11; i++) // 从第2行开始（跳过表头）
                            {
                                var status = worksheet.Cells[i, 3]?.ToString();
                                var value = Convert.ToDouble(worksheet.Cells[i, 2]?? 0);
                                
                                totalValue += value;
                                
                                if (status == "正常")
                                {
                                    normalCount++;
                                }
                                else if (status == "异常")
                                {
                                    abnormalCount++;
                                    _logger.Warning("发现异常数据 - 行: {0}, 值: {1}", i, value);
                                }
                            }
                            
                            double average = totalValue / 10;
                            _logger.Info("数据分析结果 - 正常: {0}, 异常: {1}, 平均值: {2:F2}", 
                                normalCount, abnormalCount, average);
                            
                            // 添加分析结果到工作表
                            worksheet.Cells[13, 1] = "分析结果:";
                            worksheet.Cells[14, 1] = $"正常数据: {normalCount}";
                            worksheet.Cells[15, 1] = $"异常数据: {abnormalCount}";
                            worksheet.Cells[16, 1] = $"平均值: {average:F2}";
                            
                            _logger.Debug("分析结果已写入工作表");
                        }
                    }
                    
                    _logger.Info("复杂数据处理流程全部完成");
                    ShowMessage("复杂数据处理完成！\n\n" +
                              "已完成:\n" +
                              "• 数据验证\n" +
                              "• 示例数据生成\n" +
                              "• 数据分析\n" +
                              "• 结果输出\n\n" +
                              "请查看工作表和日志管理器了解详情。");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "复杂数据处理过程中发生错误");
                    ShowError($"数据处理失败: {ex.Message}");
                }
            }
        }
        
        /// <summary>
        /// 错误模拟功能 - 演示各种错误情况的日志记录
        /// </summary>
        private void ErrorSimulation()
        {
            using (_logger.MeasurePerformance("错误模拟测试"))
            {
                _logger.Info("开始错误模拟测试");
                
                try
                {
                    var choice = System.Windows.Forms.MessageBox.Show(
                        "选择要模拟的错误类型:\n\n" +
                        "确定 - 模拟可恢复的错误\n" +
                        "取消 - 模拟严重错误",
                        "错误模拟",
                        System.Windows.Forms.MessageBoxButtons.OKCancel,
                        System.Windows.Forms.MessageBoxIcon.Question);
                    
                    if (choice == System.Windows.Forms.DialogResult.OK)
                    {
                        _logger.Info("用户选择模拟可恢复错误");
                        SimulateRecoverableError();
                    }
                    else
                    {
                        _logger.Info("用户选择模拟严重错误");
                        SimulateSevereError();
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "错误模拟过程中发生意外错误");
                    ShowError($"错误模拟失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 宿主日志接口演示 - 演示 IHostApplication 接口的增强日志功能
        /// </summary>
        private void HostLoggerDemo()
        {
            try
            {
                _logger.Info("开始演示宿主日志接口功能");

                // 获取宿主应用接口
                var hostApp = HostApplication.Instance;
                if (hostApp == null)
                {
                    ShowError("无法获取宿主应用接口");
                    return;
                }

                // 创建演示对象
                var demo = new HostLoggerDemo(hostApp, "BasePlugin");

                // 显示选择对话框
                var choice = System.Windows.Forms.MessageBox.Show(
                    "选择要演示的宿主日志功能:\n\n" +
                    "确定 - 运行完整演示\n" +
                    "取消 - 仅演示基础功能",
                    "宿主日志接口演示",
                    System.Windows.Forms.MessageBoxButtons.OKCancel,
                    System.Windows.Forms.MessageBoxIcon.Question);

                if (choice == System.Windows.Forms.DialogResult.OK)
                {
                    _logger.Info("用户选择运行完整演示");
                    demo.RunAllDemos();
                }
                else
                {
                    _logger.Info("用户选择演示基础功能");
                    demo.DemoBasicLogging();
                    demo.DemoFormattedLogging();
                    demo.DemoConditionalLogging();
                }

                _logger.Info("宿主日志接口演示完成");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "宿主日志接口演示过程中发生错误");
                ShowError($"宿主日志接口演示失败: {ex.Message}");
            }
        }
        
        #endregion
        
        #region 辅助方法
        
        private void ShowMessage(string message)
        {
            _logger.Debug("显示信息消息: {0}", message);
            System.Windows.Forms.MessageBox.Show(message, "基础插件示例", 
                System.Windows.Forms.MessageBoxButtons.OK, 
                System.Windows.Forms.MessageBoxIcon.Information);
        }
        
        private void ShowError(string message)
        {
            _logger.Debug("显示错误消息: {0}", message);
            System.Windows.Forms.MessageBox.Show(message, "基础插件示例", 
                System.Windows.Forms.MessageBoxButtons.OK, 
                System.Windows.Forms.MessageBoxIcon.Error);
        }
        
        /// <summary>
        /// 获取模拟的内存使用率
        /// </summary>
        private int GetFakeMemoryUsage()
        {
            var random = new Random();
            return random.Next(60, 95); // 返回60-95%的随机值
        }
        
        /// <summary>
        /// 模拟可恢复的错误
        /// </summary>
        private void SimulateRecoverableError()
        {
            _logger.Info("开始模拟可恢复错误场景");
            
            try
            {
                // 模拟网络连接问题
                _logger.Warning("模拟网络连接超时...");
                System.Threading.Thread.Sleep(200);
                
                // 模拟重试机制
                for (int retry = 1; retry <= 3; retry++)
                {
                    _logger.Debug("尝试重新连接 - 第 {0} 次", retry);
                    
                    if (retry == 2)
                    {
                        // 第二次尝试成功
                        _logger.Info("重连成功！问题已恢复");
                        ShowMessage("模拟可恢复错误完成！\n\n" +
                                  "场景: 网络连接超时\n" +
                                  "处理: 自动重试机制\n" +
                                  "结果: 第2次重试成功");
                        return;
                    }
                    
                    _logger.Warning("第 {0} 次重试失败", retry);
                    System.Threading.Thread.Sleep(100);
                }
                
                // 如果所有重试都失败
                throw new TimeoutException("网络连接超时，所有重试都失败");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "可恢复错误模拟过程中发生异常");
                ShowError($"可恢复错误处理失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 模拟严重错误
        /// </summary>
        private void SimulateSevereError()
        {
            _logger.Info("开始模拟严重错误场景");
            
            try
            {
                // 模拟多种严重错误情况
                var random = new Random();
                var errorType = random.Next(1, 4);
                
                switch (errorType)
                {
                    case 1:
                        _logger.Error("模拟空引用异常...");
                        throw new NullReferenceException("关键对象引用为空，无法继续操作");
                        
                    case 2:
                        _logger.Error("模拟内存不足异常...");
                        throw new OutOfMemoryException("系统内存不足，无法分配所需资源");
                        
                    case 3:
                        _logger.Error("模拟文件访问异常...");
                        throw new UnauthorizedAccessException("访问被拒绝，缺少必要的权限");
                        
                    default:
                        _logger.Error("模拟未知异常...");
                        throw new InvalidOperationException("发生未知的严重错误");
                }
            }
            catch (NullReferenceException ex)
            {
                _logger.Error(ex, "捕获到空引用异常 - 系统状态: 不稳定");
                ShowError($"严重错误 - 空引用异常:\n{ex.Message}\n\n建议重启应用程序。");
            }
            catch (OutOfMemoryException ex)
            {
                _logger.Error(ex, "捕获到内存不足异常 - 系统资源: 不足");
                ShowError($"严重错误 - 内存不足:\n{ex.Message}\n\n请关闭其他程序并重试。");
            }
            catch (UnauthorizedAccessException ex)
            {
                _logger.Error(ex, "捕获到访问权限异常 - 权限检查: 失败");
                ShowError($"严重错误 - 访问权限不足:\n{ex.Message}\n\n请以管理员权限运行。");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "捕获到未知严重异常 - 错误级别: 严重");
                ShowError($"严重的未知错误:\n{ex.Message}\n\n系统可能不稳定。");
            }
        }
        
        #endregion
    }
} 