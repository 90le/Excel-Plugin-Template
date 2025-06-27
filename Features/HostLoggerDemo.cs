using System;
using DTI_Tool.AddIn.Common.Interfaces;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasePlugin.Features
{
    /// <summary>
    /// 宿主日志管理器功能演示
    /// </summary>
    public class HostLoggerDemo
    {
        private readonly IHostApplication _hostApplication;
        private readonly string _pluginName;

        public HostLoggerDemo(IHostApplication hostApplication, string pluginName)
        {
            _hostApplication = hostApplication ?? throw new ArgumentNullException(nameof(hostApplication));
            _pluginName = pluginName;
        }

        /// <summary>
        /// 演示基础日志功能
        /// </summary>
        public void DemoBasicLogging()
        {
            _hostApplication.LogInfo(_pluginName, "=== 开始演示基础日志功能 ===");

            // 基础日志记录
            _hostApplication.LogDebug(_pluginName, "这是一条调试信息");
            _hostApplication.LogInfo(_pluginName, "这是一条信息日志");
            _hostApplication.LogWarning(_pluginName, "这是一条警告信息");
            _hostApplication.LogError(_pluginName, "这是一条错误信息");

            _hostApplication.LogInfo(_pluginName, "=== 基础日志功能演示完成 ===");
        }

        /// <summary>
        /// 演示格式化日志功能
        /// </summary>
        public void DemoFormattedLogging()
        {
            _hostApplication.LogInfo(_pluginName, "=== 开始演示格式化日志功能 ===");

            var userName = "张三";
            var operationCount = 42;
            var percentage = 85.67;
            var startTime = DateTime.Now;

            // 格式化日志记录
            _hostApplication.LogInfoFormat(_pluginName, "用户 {0} 执行了 {1} 次操作", userName, operationCount);
            _hostApplication.LogDebugFormat(_pluginName, "当前进度: {0:F1}%, 开始时间: {1:HH:mm:ss}", percentage, startTime);
            _hostApplication.LogWarningFormat(_pluginName, "检测到 {0} 个潜在问题，成功率: {1:P}", 3, 0.8567);

            // 演示复杂格式化
            _hostApplication.LogInfoFormat(_pluginName, 
                "处理报告 - 用户: {0}, 操作: {1}, 成功率: {2:P}, 耗时: {3} 秒",
                userName, operationCount, percentage / 100, 2.5);

            _hostApplication.LogInfo(_pluginName, "=== 格式化日志功能演示完成 ===");
        }

        /// <summary>
        /// 演示条件日志记录
        /// </summary>
        public void DemoConditionalLogging()
        {
            _hostApplication.LogInfo(_pluginName, "=== 开始演示条件日志记录 ===");

            // 检查日志级别状态
            _hostApplication.LogInfoFormat(_pluginName, "Debug级别启用: {0}", _hostApplication.IsDebugEnabled);
            _hostApplication.LogInfoFormat(_pluginName, "Info级别启用: {0}", _hostApplication.IsInfoEnabled);
            _hostApplication.LogInfoFormat(_pluginName, "Warning级别启用: {0}", _hostApplication.IsWarningEnabled);
            _hostApplication.LogInfoFormat(_pluginName, "Error级别启用: {0}", _hostApplication.IsErrorEnabled);

            // 条件日志记录 - 仅在Debug启用时执行复杂操作
            if (_hostApplication.IsDebugEnabled)
            {
                var complexDebugInfo = GenerateComplexDebugInfo();
                _hostApplication.LogDebugFormat(_pluginName, "复杂调试信息: {0}", complexDebugInfo);
            }
            else
            {
                _hostApplication.LogInfo(_pluginName, "Debug级别未启用，跳过复杂调试信息生成");
            }

            _hostApplication.LogInfo(_pluginName, "=== 条件日志记录演示完成 ===");
        }

        /// <summary>
        /// 演示异常处理和日志记录
        /// </summary>
        public void DemoExceptionLogging()
        {
            _hostApplication.LogInfo(_pluginName, "=== 开始演示异常处理日志 ===");

            try
            {
                // 模拟一个可能出错的操作
                RiskyOperation();
            }
            catch (ArgumentException ex)
            {
                _hostApplication.LogErrorFormat(_pluginName, ex, "参数错误: {0}", ex.Message);
            }
            catch (InvalidOperationException ex)
            {
                _hostApplication.LogErrorFormat(_pluginName, ex, "操作无效: 操作ID={0}, 状态={1}", "OP001", "失败");
            }
            catch (Exception ex)
            {
                _hostApplication.LogErrorFormat(_pluginName, ex, "未预期的错误: {0}", ex.GetType().Name);
            }

            _hostApplication.LogInfo(_pluginName, "=== 异常处理日志演示完成 ===");
        }

        /// <summary>
        /// 演示性能测量功能
        /// </summary>
        public void DemoPerformanceMeasurement()
        {
            _hostApplication.LogInfo(_pluginName, "=== 开始演示性能测量功能 ===");

            // 简单性能测量
            using (_hostApplication.StartPerformanceMeasure(_pluginName, "简单计算操作"))
            {
                // 模拟一些计算
                System.Threading.Thread.Sleep(100);
                for (int i = 0; i < 1000; i++)
                {
                    Math.Sqrt(i);
                }
            }

            // 嵌套性能测量
            using (_hostApplication.StartPerformanceMeasure(_pluginName, "复杂数据处理流程"))
            {
                using (_hostApplication.StartPerformanceMeasure(_pluginName, "数据验证"))
                {
                    System.Threading.Thread.Sleep(50);
                    _hostApplication.LogDebugFormat(_pluginName, "验证了 {0} 个数据项", 100);
                }

                using (_hostApplication.StartPerformanceMeasure(_pluginName, "数据转换"))
                {
                    System.Threading.Thread.Sleep(75);
                    _hostApplication.LogDebugFormat(_pluginName, "转换了 {0} 个数据项", 100);
                }

                using (_hostApplication.StartPerformanceMeasure(_pluginName, "数据输出"))
                {
                    System.Threading.Thread.Sleep(25);
                    _hostApplication.LogDebugFormat(_pluginName, "输出了 {0} 个结果", 100);
                }
            }

            _hostApplication.LogInfo(_pluginName, "=== 性能测量功能演示完成 ===");
        }

        /// <summary>
        /// 演示Excel操作的日志记录
        /// </summary>
        public void DemoExcelOperationLogging()
        {
            _hostApplication.LogInfo(_pluginName, "=== 开始演示Excel操作日志记录 ===");

            try
            {
                using (_hostApplication.StartPerformanceMeasure(_pluginName, "Excel工作簿操作"))
                {
                    var workbook = _hostApplication.GetActiveWorkbook();
                    if (workbook != null)
                    {
                        _hostApplication.LogInfoFormat(_pluginName, "当前工作簿: {0}", workbook.Name);

                        var worksheet = _hostApplication.GetActiveWorksheet();
                        if (worksheet != null)
                        {
                            _hostApplication.LogInfoFormat(_pluginName, "当前工作表: {0}", worksheet.Name);

                            var selection = _hostApplication.GetSelection();
                            if (selection != null)
                            {
                                _hostApplication.LogInfoFormat(_pluginName, 
                                    "选中区域: {0}, 行数: {1}, 列数: {2}",
                                    selection.Address, selection.Rows.Count, selection.Columns.Count);

                                // 条件性详细日志
                                if (_hostApplication.IsDebugEnabled)
                                {
                                    _hostApplication.LogDebugFormat(_pluginName, "选中区域详细信息:");
                                    _hostApplication.LogDebugFormat(_pluginName, "- 起始行: {0}", selection.Row);
                                    _hostApplication.LogDebugFormat(_pluginName, "- 起始列: {0}", selection.Column);
                                    _hostApplication.LogDebugFormat(_pluginName, "- 工作表名: {0}", selection.Worksheet.Name);
                                }
                            }
                            else
                            {
                                _hostApplication.LogWarningFormat(_pluginName, "无法获取选中区域");
                            }
                        }
                        else
                        {
                            _hostApplication.LogWarningFormat(_pluginName, "无法获取当前工作表");
                        }
                    }
                    else
                    {
                        _hostApplication.LogWarningFormat(_pluginName, "无法获取当前工作簿");
                    }
                }
            }
            catch (Exception ex)
            {
                _hostApplication.LogErrorFormat(_pluginName, ex, "Excel操作失败");
            }

            _hostApplication.LogInfo(_pluginName, "=== Excel操作日志记录演示完成 ===");
        }

        /// <summary>
        /// 运行所有演示
        /// </summary>
        public void RunAllDemos()
        {
            _hostApplication.LogInfo(_pluginName, "开始运行宿主日志管理器完整演示");

            try
            {
                using (_hostApplication.StartPerformanceMeasure(_pluginName, "完整日志演示"))
                {
                    DemoBasicLogging();
                    DemoFormattedLogging();
                    DemoConditionalLogging();
                    DemoExceptionLogging();
                    DemoPerformanceMeasurement();
                    DemoExcelOperationLogging();
                }

                _hostApplication.LogInfo(_pluginName, "所有演示完成！请查看日志管理器以查看详细输出。");
            }
            catch (Exception ex)
            {
                _hostApplication.LogErrorFormat(_pluginName, ex, "演示过程中发生错误");
            }
        }

        #region 辅助方法

        private string GenerateComplexDebugInfo()
        {
            // 模拟生成复杂的调试信息
            var info = $"系统状态 - 时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}, " +
                       $"内存: {GC.GetTotalMemory(false) / 1024 / 1024} MB, " +
                       $"线程ID: {System.Threading.Thread.CurrentThread.ManagedThreadId}";
            return info;
        }

        private void RiskyOperation()
        {
            // 模拟一个可能抛出异常的操作
            var random = new Random();
            var value = random.Next(1, 4);

            switch (value)
            {
                case 1:
                    throw new ArgumentException("模拟参数异常", "testParameter");
                case 2:
                    throw new InvalidOperationException("模拟操作异常");
                case 3:
                    throw new NotSupportedException("模拟不支持异常");
                default:
                    // 正常执行
                    break;
            }
        }

        #endregion
    }
} 