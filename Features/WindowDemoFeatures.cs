using System;
using System.Collections.Generic;
using System.Windows;
using BasePlugin.Core;
using BasePlugin.Models;
using BasePlugin.WPF.Views;
using DTI_Tool.AddIn.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasePlugin.Features
{
    /// <summary>
    /// 窗口演示功能类 - 提供WPF窗口相关的示例功能
    /// </summary>
    public class WindowDemoFeatures : IFeatureProvider
    {
        #region 私有字段

        private readonly PluginLogger _logger;

        #endregion

        #region 构造函数

        public WindowDemoFeatures(PluginLogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            
            _logger.Debug("WindowDemoFeatures 已初始化");
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
        /// 获取窗口演示功能列表
        /// </summary>
        public List<PluginFeature> GetFeatures()
        {
            return new List<PluginFeature>
            {
                new PluginFeature
                {
                    Id = "show_simple_window",
                    Name = "简单WPF窗口",
                    Description = "显示一个简单的WPF窗口示例",
                    Category = "窗口演示",
                    Tags = new List<string> { "WPF", "窗口", "界面" },
                    ImageMso = "WindowNew",
                    Action = ShowSimpleWindow
                },
                new PluginFeature
                {
                    Id = "show_data_entry_form",
                    Name = "数据录入窗口",
                    Description = "显示数据录入表单窗口",
                    Category = "窗口演示",
                    Tags = new List<string> { "WPF", "表单", "数据录入" },
                    ImageMso = "FormControlEditBox",
                    Action = ShowDataEntryForm
                },
                new PluginFeature
                {
                    Id = "show_settings_window",
                    Name = "设置窗口",
                    Description = "显示插件设置窗口",
                    Category = "窗口演示",
                    Tags = new List<string> { "WPF", "设置", "配置" },
                    ImageMso = "AdpDiagrammer",
                    Action = ShowSettingsWindow
                },
                new PluginFeature
                {
                    Id = "show_progress_window",
                    Name = "进度窗口",
                    Description = "显示任务进度窗口",
                    Category = "窗口演示",
                    Tags = new List<string> { "WPF", "进度", "任务" },
                    ImageMso = "AnimationTransition",
                    Action = ShowProgressWindow
                },
                new PluginFeature
                {
                    Id = "show_wpf_task_pane",
                    Name = "WPF任务窗格",
                    Description = "显示/隐藏现代化WPF任务窗格",
                    Category = "窗口演示",
                    Tags = new List<string> { "WPF", "任务窗格", "现代化", "界面" },
                    ImageMso = "TaskPaneInsert",
                    Action = ShowWpfTaskPane
                },
                new PluginFeature
                {
                    Id = "show_winforms_task_pane",
                    Name = "WinForms任务窗格",
                    Description = "显示/隐藏WinForms任务窗格",
                    Category = "窗口演示",
                    Tags = new List<string> { "WinForms", "任务窗格", "传统", "界面" },
                    ImageMso = "TaskPaneProperties",
                    Action = ShowWinFormsTaskPane
                },
                new PluginFeature
                {
                    Id = "show_task_pane",
                    Name = "默认任务窗格",
                    Description = "显示/隐藏默认任务窗格（兼容性）",
                    Category = "窗口演示",
                    Tags = new List<string> { "任务窗格", "兼容性", "默认" },
                    ImageMso = "TaskPaneLegacy",
                    Action = ShowTaskPane
                }
            };
        }

        public void Dispose()
        {
            _logger.Debug("WindowDemoFeatures 已释放");
        }

        #endregion

        #region 功能实现

        /// <summary>
        /// 显示简单WPF窗口
        /// </summary>
        private void ShowSimpleWindow()
        {
            using (_logger.MeasurePerformance("显示简单WPF窗口"))
            {
                try
                {
                    _logger.Info("显示简单WPF窗口");
                    
                    var window = new SimpleWindow
                    {
                        Owner = GetExcelWindow()
                    };
                    
                    var result = window.ShowDialog();
                    _logger.Debug("窗口关闭，返回结果: {0}", result);
                    
                    if (result == true)
                    {
                        MessageHelper.ShowInfo("您点击了确定按钮", "窗口结果");
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "显示简单WPF窗口失败");
                    MessageHelper.ShowError($"显示窗口失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 显示数据录入表单
        /// </summary>
        private void ShowDataEntryForm()
        {
            using (_logger.MeasurePerformance("显示数据录入表单"))
            {
                try
                {
                    _logger.Info("显示数据录入表单");
                    
                    var form = new DataEntryForm
                    {
                        Owner = GetExcelWindow()
                    };
                    
                    if (form.ShowDialog() == true)
                    {
                        // 获取录入的数据
                        var viewModel = form.DataContext as DataEntryViewModel;
                        if (viewModel != null)
                        {
                            // 将数据写入Excel
                            WriteDataToExcel(viewModel);
                            MessageHelper.ShowInfo("数据已成功写入Excel", "录入成功");
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "显示数据录入表单失败");
                    MessageHelper.ShowError($"显示数据录入表单失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 显示设置窗口
        /// </summary>
        private void ShowSettingsWindow()
        {
            using (_logger.MeasurePerformance("显示设置窗口"))
            {
                try
                {
                    _logger.Info("显示设置窗口");
                    
                    var settings = new SettingsWindow
                    {
                        Owner = GetExcelWindow()
                    };
                    
                    if (settings.ShowDialog() == true)
                    {
                        _logger.Info("设置已保存");
                        MessageHelper.ShowInfo("设置已成功保存", "保存成功");
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "显示设置窗口失败");
                    MessageHelper.ShowError($"显示设置窗口失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 显示进度窗口
        /// </summary>
        private void ShowProgressWindow()
        {
            using (_logger.MeasurePerformance("显示进度窗口"))
            {
                try
                {
                    _logger.Info("显示进度窗口");
                    
                    var progress = new ProgressWindow
                    {
                        Owner = GetExcelWindow()
                    };
                    
                    // 启动后台任务
                    progress.StartTask(async (progressReporter, cancellationToken) =>
                    {
                        const int totalSteps = 100;
                        
                        for (int i = 0; i <= totalSteps; i++)
                        {
                            if (cancellationToken.IsCancellationRequested)
                            {
                                _logger.Info("任务被用户取消");
                                break;
                            }
                            
                            // 模拟工作
                            await System.Threading.Tasks.Task.Delay(50, cancellationToken);
                            
                            // 报告进度
                            progressReporter.Report(new ProgressInfo
                            {
                                Progress = i,
                                Message = $"正在处理... {i}/{totalSteps}"
                            });
                        }
                    });
                    
                    progress.ShowDialog();
                    _logger.Info("进度窗口已关闭");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "显示进度窗口失败");
                    MessageHelper.ShowError($"显示进度窗口失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 显示/隐藏WPF任务窗格
        /// </summary>
        private void ShowWpfTaskPane()
        {
            using (_logger.MeasurePerformance("显示WPF任务窗格"))
            {
                try
                {
                    _logger.Info("切换WPF任务窗格显示状态");
                    
                    // 使用宿主提供的任务窗格接口
                    var hostApp = HostApplication.Instance;
                    if (hostApp == null)
                    {
                        MessageHelper.ShowError("无法获取宿主应用接口");
                        return;
                    }
                    
                    // 创建任务窗格管理器
                    var taskPaneManager = new TaskPaneManager(_logger, "BasePluginDemo");

                    // 检查任务窗格是否存在
                    if (taskPaneManager.TaskPaneExists("WpfTaskPane"))
                    {
                        // 如果存在，切换显示状态
                        taskPaneManager.ToggleWpfTaskPane("WpfTaskPane");
                    }
                    else
                    {
                        // 如果不存在，创建并设置宽度为380像素
                        taskPaneManager.CreateWpfTaskPane("WpfTaskPane", 380);
                    }
                    
                    _logger.Info("WPF任务窗格操作完成");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "显示WPF任务窗格失败");
                    MessageHelper.ShowError($"显示WPF任务窗格失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 显示/隐藏WinForms任务窗格
        /// </summary>
        private void ShowWinFormsTaskPane()
        {
            using (_logger.MeasurePerformance("显示WinForms任务窗格"))
            {
                try
                {
                    _logger.Info("切换WinForms任务窗格显示状态");
                    
                    // 使用宿主提供的任务窗格接口
                    var hostApp = HostApplication.Instance;
                    if (hostApp == null)
                    {
                        MessageHelper.ShowError("无法获取宿主应用接口");
                        return;
                    }
                    
                    // 创建任务窗格管理器
                    var taskPaneManager = new TaskPaneManager(_logger, "BasePluginDemo");

                    // 检查任务窗格是否存在
                    if (taskPaneManager.TaskPaneExists("WinFormsTaskPane"))
                    {
                        // 如果存在，切换显示状态
                        taskPaneManager.ToggleWinFormsTaskPane("WinFormsTaskPane");
                    }
                    else
                    {
                        // 如果不存在，创建并设置宽度为360像素
                        taskPaneManager.CreateWinFormsTaskPane("WinFormsTaskPane", 360);
                    }
                    
                    _logger.Info("WinForms任务窗格操作完成");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "显示WinForms任务窗格失败");
                    MessageHelper.ShowError($"显示WinForms任务窗格失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 显示/隐藏任务窗格（兼容性方法）
        /// </summary>
        private void ShowTaskPane()
        {
            using (_logger.MeasurePerformance("显示任务窗格"))
            {
                try
                {
                    _logger.Info("切换任务窗格显示状态");
                    
                    // 使用宿主提供的任务窗格接口
                    var hostApp = HostApplication.Instance;
                    if (hostApp == null)
                    {
                        MessageHelper.ShowError("无法获取宿主应用接口");
                        return;
                    }
                    
                    // 创建任务窗格管理器
                    var taskPaneManager = new TaskPaneManager(_logger, "BasePluginDemo");

                    // 检查任务窗格是否存在
                    if (taskPaneManager.TaskPaneExists("DemoTaskPane"))
                    {
                        // 如果存在，切换显示状态
                        taskPaneManager.ToggleWpfTaskPane("DemoTaskPane");
                    }
                    else
                    {
                        // 如果不存在，创建并设置宽度为380像素（默认使用WPF）
                        taskPaneManager.CreateWpfTaskPane("DemoTaskPane", 380);
                    }
                    
                    _logger.Info("任务窗格操作完成");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "显示任务窗格失败");
                    MessageHelper.ShowError($"显示任务窗格失败: {ex.Message}");
                }
            }
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 获取Excel主窗口
        /// </summary>
        private Window GetExcelWindow()
        {
            try
            {
                var hwnd = new IntPtr(ExcelApp.Hwnd);
                // 从Excel窗口句柄获取WPF窗口
                return System.Windows.Interop.HwndSource.FromHwnd(hwnd)?.RootVisual as Window;
            }
            catch
            {
                // 如果无法获取Excel窗口，返回null
                return null;
            }
        }

        /// <summary>
        /// 将数据写入Excel
        /// </summary>
        private void WriteDataToExcel(DataEntryViewModel data)
        {
            var worksheet = ExcelApp?.ActiveSheet as Excel.Worksheet;
            if (worksheet == null) return;
            
            var activeCell = ExcelApp.ActiveCell;
            if (activeCell == null) return;
            
            var row = activeCell.Row;
            var col = activeCell.Column;
            
            // 写入数据
            worksheet.Cells[row, col] = data.Name;
            worksheet.Cells[row, col + 1] = data.Email;
            worksheet.Cells[row, col + 2] = data.Phone;
            worksheet.Cells[row, col + 3] = data.Department;
            worksheet.Cells[row, col + 4] = data.JoinDate?.ToString("yyyy-MM-dd");
            
            _logger.Debug("数据已写入Excel，起始位置: 行{0}, 列{1}", row, col);
        }

        #endregion
    }

    #region 辅助类

    /// <summary>
    /// 数据录入视图模型
    /// </summary>
    public class DataEntryViewModel
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string Department { get; set; }
        public DateTime? JoinDate { get; set; }
    }

    /// <summary>
    /// 进度信息
    /// </summary>
    public class ProgressInfo
    {
        public int Progress { get; set; }
        public string Message { get; set; }
    }

    #endregion
} 