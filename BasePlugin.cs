using System;
using System.Collections.Generic;
using DTI_Tool.AddIn.Common.Interfaces;
using DTI_Tool.AddIn.Common.Models;
using DTI_Tool.AddIn.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using BasePlugin.Features;
using BasePlugin.Models;

namespace BasePlugin
{
    /// <summary>
    /// 基础插件模板 - 开发者可基于此模板快速创建插件
    /// </summary>
    public class BasePlugin : IPlugin
    {
        private Excel.Application _excelApp;
        private bool _isInitialized = false;
        
        // 功能类实例 - 根据需要添加更多功能类
        private SampleFeatures _sampleFeatures;

        #region IPlugin 接口实现

        public string Name => "插件名称"; // 修改为您的插件名称
        public string Version => "1.0.0";
        public string Description => "插件描述信息"; // 修改为您的插件描述
        public string Author => "开发者姓名"; // 修改为您的姓名

        public void Initialize()
        {
            try
            {
                // 获取Excel应用程序对象
                _excelApp = HostApplication.Instance.ExcelApplication;
                
                // 初始化功能类 - 根据需要添加更多功能类
                _sampleFeatures = new SampleFeatures(_excelApp);
                
                _isInitialized = true;
                
                LogInfo("插件初始化成功");
            }
            catch (Exception ex)
            {
                LogError($"插件初始化失败: {ex.Message}");
                throw;
            }
        }

        public void Load()
        {
            if (!_isInitialized)
                throw new InvalidOperationException("插件未初始化");

            LogInfo("插件加载成功");
        }

        public void Unload()
        {
            try
            {
                // 清理资源
                _sampleFeatures = null;
                _excelApp = null;
                
                LogInfo("插件卸载成功");
            }
            catch (Exception ex)
            {
                LogError($"插件卸载时发生错误: {ex.Message}");
            }
        }

        public void Dispose()
        {
            Unload();
        }

        /// <summary>
        /// 获取功能区按钮列表
        /// </summary>
        public List<RibbonButton> GetRibbonButtons()
        {
            var allFeatures = GetAllFeatures();
            var ribbonButtons = new List<RibbonButton>();

            // 按类别分组创建菜单
            var categories = allFeatures.GroupBy(f => f.Category).ToList();

            foreach (var category in categories)
            {
                var menuButton = new RibbonButton
                {
                    Type = "menu",
                    Id = $"menu_{category.Key.Replace(" ", "_")}",
                    Label = category.Key,
                    ImageMso = GetCategoryImageMso(category.Key),
                    Tooltip = $"包含{category.Count()}个{category.Key}功能",
                    Enabled = true,
                    Items = new List<RibbonButton>()
                };

                // 为每个类别添加功能按钮
                foreach (var feature in category)
                {
                    menuButton.Items.Add(feature.ToRibbonButton());
                }

                ribbonButtons.Add(menuButton);
            }

            return ribbonButtons;
        }

        /// <summary>
        /// 获取所有可执行命令
        /// </summary>
        public List<PluginCommand> GetCommands()
        {
            var allFeatures = GetAllFeatures();
            return allFeatures.Select(f => f.ToPluginCommand()).ToList();
        }

        /// <summary>
        /// 搜索命令
        /// </summary>
        public List<PluginCommand> SearchCommands(string keyword)
        {
            if (string.IsNullOrEmpty(keyword))
                return GetCommands();

            var commands = GetCommands();
            return commands.Where(cmd =>
                cmd.Name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                cmd.Description.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                cmd.Category.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                cmd.Tags.Any(tag => tag.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
            ).ToList();
        }

        public void ExecuteCommand(string commandId, object[] parameters)
        {
            try
            {
                var allFeatures = GetAllFeatures();
                var feature = allFeatures.FirstOrDefault(f => f.Id == commandId);
                
                if (feature?.Action != null)
                {
                    feature.Action.Invoke();
                    LogInfo($"执行命令成功: {commandId}");
                }
                else
                {
                    LogError($"未找到命令: {commandId}");
                    throw new ArgumentException($"未知命令: {commandId}");
                }
            }
            catch (Exception ex)
            {
                LogError($"执行命令失败 {commandId}: {ex.Message}");
                throw;
            }
        }

        #endregion

        #region 私有方法

        /// <summary>
        /// 获取所有功能
        /// </summary>
        private List<PluginFeature> GetAllFeatures()
        {
            var allFeatures = new List<PluginFeature>();
            
            // 添加功能类的功能 - 根据需要添加更多功能类
            if (_sampleFeatures != null)
                allFeatures.AddRange(_sampleFeatures.GetFeatures());
            
            return allFeatures;
        }

        /// <summary>
        /// 根据类别获取图标
        /// </summary>
        private string GetCategoryImageMso(string category)
        {
            return category switch
            {
                "示例功能" => "TableExcelSelect",
                "数据处理" => "DatabaseSortDescending",
                "格式化" => "FontColorPicker",
                "图表" => "ChartColumnChart",
                "统计分析" => "FunctionWizard",
                "工作表管理" => "WorksheetInsert",
                "财务计算" => "AcceptInvitationExcel",
                "实用工具" => "ToolsOptions",
                _ => "FileNew"
            };
        }

        /// <summary>
        /// Excel操作辅助方法
        /// </summary>
        private Excel.Range GetActiveCell()
        {
            return _excelApp?.ActiveCell;
        }

        private Excel.Range GetSelection()
        {
            return _excelApp?.Selection as Excel.Range;
        }

        private Excel.Workbook GetActiveWorkbook()
        {
            return _excelApp?.ActiveWorkbook;
        }

        private Excel.Worksheet GetActiveWorksheet()
        {
            return _excelApp?.ActiveSheet as Excel.Worksheet;
        }

        /// <summary>
        /// 工具方法
        /// </summary>
        private bool IsNumeric(object value)
        {
            return double.TryParse(value?.ToString(), out _);
        }

        private void ShowMessage(string message)
        {
            System.Windows.Forms.MessageBox.Show(message, Name, 
                System.Windows.Forms.MessageBoxButtons.OK, 
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        private void ShowError(string message)
        {
            System.Windows.Forms.MessageBox.Show(message, Name, 
                System.Windows.Forms.MessageBoxButtons.OK, 
                System.Windows.Forms.MessageBoxIcon.Error);
        }

        private void LogInfo(string message)
        {
            // 这里可以添加日志记录逻辑
            System.Diagnostics.Debug.WriteLine($"[INFO] {Name}: {message}");
        }

        private void LogError(string message)
        {
            // 这里可以添加错误日志记录逻辑
            System.Diagnostics.Debug.WriteLine($"[ERROR] {Name}: {message}");
        }

        #endregion
    }
} 