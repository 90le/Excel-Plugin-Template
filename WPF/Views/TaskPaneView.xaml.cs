using System;
using System.Windows;
using System.Windows.Controls;
using DTI_Tool.AddIn.Core;
using BasePlugin.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasePlugin.WPF.Views
{
    /// <summary>
    /// TaskPaneView.xaml 的交互逻辑 - 现代化任务窗格
    /// </summary>
    public partial class TaskPaneView : UserControl
    {
        #region 私有字段

        private readonly PluginLogger _logger;
        private readonly FeatureManager _featureManager;

        #endregion

        #region 私有属性

        /// <summary>
        /// 获取Excel应用程序对象
        /// </summary>
        private Excel.Application ExcelApp => HostApplication.Instance?.ExcelApplication;

        #endregion

        #region 构造函数

        public TaskPaneView()
        {
            InitializeComponent();
            
            // 初始化日志记录器
            _logger = PluginLog.ForPlugin("BasePlugin.TaskPane");
            
            // 初始化功能管理器
            _featureManager = new FeatureManager(_logger);
            try
            {
                _featureManager.Initialize();
                _logger.Info("TaskPaneView 功能管理器初始化成功");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "TaskPaneView 功能管理器初始化失败");
            }
            
            // 初始化界面
            RefreshInfo();
            UpdateStatus("任务窗格已加载，准备就绪");
        }

        #endregion

        #region 事件处理

        /// <summary>
        /// 关闭任务窗格
        /// </summary>
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var hostApp = HostApplication.Instance;
                if (hostApp != null)
                {
                    // 关闭当前任务窗格
                    hostApp.CloseTaskPane("BasePluginDemo_DemoTaskPane");
                    _logger.Info("用户手动关闭任务窗格");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "关闭任务窗格时发生错误");
            }
        }

        /// <summary>
        /// 刷新信息按钮点击事件
        /// </summary>
        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            RefreshInfo();
            UpdateStatus("信息已刷新");
        }

        /// <summary>
        /// 快速操作按钮点击事件
        /// </summary>
        private void btnQuickAction_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var commandId = button?.Tag as string;
            
            if (string.IsNullOrEmpty(commandId))
            {
                UpdateStatus("错误：未找到命令标识", true);
                return;
            }

            ExecuteCommand(commandId, "快速操作");
        }

        /// <summary>
        /// 数据处理操作按钮点击事件
        /// </summary>
        private void btnDataAction_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var commandId = button?.Tag as string;
            
            if (string.IsNullOrEmpty(commandId))
            {
                UpdateStatus("错误：未找到命令标识", true);
                return;
            }

            ExecuteCommand(commandId, "数据处理");
        }

        /// <summary>
        /// 格式化操作按钮点击事件
        /// </summary>
        private void btnFormatAction_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var commandId = button?.Tag as string;
            
            if (string.IsNullOrEmpty(commandId))
            {
                UpdateStatus("错误：未找到命令标识", true);
                return;
            }

            ExecuteCommand(commandId, "格式化");
        }

        /// <summary>
        /// 工作表管理操作按钮点击事件
        /// </summary>
        private void btnWorksheetAction_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var commandId = button?.Tag as string;
            
            if (string.IsNullOrEmpty(commandId))
            {
                UpdateStatus("错误：未找到命令标识", true);
                return;
            }

            ExecuteCommand(commandId, "工作表管理");
        }

        /// <summary>
        /// 实用工具操作按钮点击事件
        /// </summary>
        private void btnUtilityAction_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            var commandId = button?.Tag as string;
            
            if (string.IsNullOrEmpty(commandId))
            {
                UpdateStatus("错误：未找到命令标识", true);
                return;
            }

            ExecuteCommand(commandId, "实用工具");
        }

        #endregion

        #region 私有方法

        /// <summary>
        /// 刷新工作表信息
        /// </summary>
        private void RefreshInfo()
        {
            try
            {
                if (ExcelApp == null)
                {
                    txtWorkbookName.Text = "Excel 未连接";
                    txtWorksheetName.Text = "无";
                    txtSelection.Text = "无";
                    txtCellCount.Text = "0";
                    return;
                }

                var workbook = ExcelApp.ActiveWorkbook;
                var worksheet = ExcelApp.ActiveSheet as Excel.Worksheet;
                var selection = ExcelApp.Selection as Excel.Range;

                txtWorkbookName.Text = workbook?.Name ?? "无工作簿";
                txtWorksheetName.Text = worksheet?.Name ?? "无工作表";
                
                if (selection != null)
                {
                    txtSelection.Text = selection.Address;
                    txtCellCount.Text = selection.Cells.Count.ToString();
                }
                else
                {
                    txtSelection.Text = "无选择";
                    txtCellCount.Text = "0";
                }

                _logger.Debug("工作表信息已刷新");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "刷新工作表信息时发生错误");
                txtWorkbookName.Text = "错误";
                txtWorksheetName.Text = "错误";
                txtSelection.Text = "错误";
                txtCellCount.Text = "0";
            }
        }

        /// <summary>
        /// 执行命令
        /// </summary>
        /// <param name="commandId">命令ID</param>
        /// <param name="category">操作类别</param>
        private void ExecuteCommand(string commandId, string category)
        {
            try
            {
                if (_featureManager == null)
                {
                    UpdateStatus("错误：功能管理器未初始化", true);
                    return;
                }

                _logger.Info("执行命令: {0} (类别: {1})", commandId, category);
                UpdateStatus($"正在执行 {category} 操作...");

                // 通过功能管理器执行命令
                _featureManager.ExecuteCommand(commandId);

                // 刷新信息
                RefreshInfo();
                
                UpdateStatus($"{category} 操作已完成");
                _logger.Info("命令执行成功: {0}", commandId);
            }
            catch (ArgumentException)
            {
                var errorMsg = $"未找到命令: {commandId}";
                UpdateStatus(errorMsg, true);
                _logger.Warning("命令未找到: {0}", commandId);
                
                // 显示友好的错误提示
                MessageBox.Show($"功能 '{commandId}' 暂未实现或不可用。\n\n请检查插件配置。", 
                    "功能不可用", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                var errorMsg = $"{category} 操作失败: {ex.Message}";
                UpdateStatus(errorMsg, true);
                _logger.Error(ex, "执行命令失败: {0}", commandId);
                
                // 显示详细错误信息
                MessageBox.Show($"执行 {category} 操作时发生错误：\n\n{ex.Message}\n\n请查看日志文件了解详细信息。", 
                    "操作失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 更新状态信息
        /// </summary>
        /// <param name="message">状态消息</param>
        /// <param name="isError">是否为错误信息</param>
        private void UpdateStatus(string message, bool isError = false)
        {
            try
            {
                if (txtStatus != null)
                {
                    txtStatus.Text = $"{DateTime.Now:HH:mm:ss} - {message}";
                    
                    // 根据消息类型设置颜色
                    if (isError)
                    {
                        txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Colors.Red);
                    }
                    else
                    {
                        txtStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(0x42, 0x42, 0x42));
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "更新状态信息时发生错误");
            }
        }

        #endregion

        #region 清理资源

        /// <summary>
        /// 清理资源
        /// </summary>
        public void Cleanup()
        {
            try
            {
                _featureManager?.Dispose();
                _logger?.Info("TaskPaneView 资源已清理");
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "清理 TaskPaneView 资源时发生错误");
            }
        }

        #endregion
    }
} 