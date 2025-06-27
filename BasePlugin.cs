using System;
using System.Collections.Generic;
using DTI_Tool.AddIn.Common.Interfaces;
using DTI_Tool.AddIn.Common.Models;
using DTI_Tool.AddIn.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using BasePlugin.Features;
using BasePlugin.Models;
using BasePlugin.Core;

namespace BasePlugin
{
    /// <summary>
    /// 基础插件模板 - 开发者可基于此模板快速创建插件
    /// </summary>
    public class BasePlugin : IPlugin
    {
        private Excel.Application _excelApp;
        private bool _isInitialized = false;
        private PluginLogger _logger;
        
        // 功能类实例 - 根据需要添加更多功能类
        private SampleFeatures _sampleFeatures;

        #region IPlugin 接口实现

        public string Name => "BasePlugin"; // 修改为您的插件名称，请和manifest.json文件的name一致，便于日志管理。
        public string Version => "1.0.0";
        public string Description => "插件描述信息"; // 修改为您的插件描述
        public string Author => "开发者姓名"; // 修改为您的姓名

        public void Initialize()
        {
            try
            {
                // 初始化日志记录器
                _logger = PluginLog.ForPlugin(Name);
                _logger.Info("=== 开始初始化 BasePlugin 插件 ===");
                _logger.Debug("插件版本: {0}, 作者: {1}", Version, Author);

                // 获取Excel应用程序对象
                _logger.Debug("获取Excel应用程序对象...");
                _excelApp = HostApplication.Instance.ExcelApplication;
                _logger.Info("Excel应用程序对象获取成功，版本: {0}", _excelApp?.Version ?? "未知");
                
                // 初始化功能类 - 根据需要添加更多功能类
                _logger.Debug("初始化示例功能类...");
                _sampleFeatures = new SampleFeatures(_excelApp, _logger);
                
                // 获取功能统计信息
                var features = GetAllFeatures();
                var categories = features.GroupBy(f => f.Category).ToList();
                _logger.Info("插件功能加载完成 - 共 {0} 个功能，分为 {1} 个类别", 
                    features.Count, categories.Count);
                
                foreach (var category in categories)
                {
                    _logger.Debug("类别 '{0}': {1} 个功能", category.Key, category.Count());
                }
                
                _isInitialized = true;
                
                _logger.Info("=== BasePlugin 插件初始化完成 ===");
            }
            catch (Exception ex)
            {
                if (_logger != null)
                {
                    _logger.Error(ex, "插件初始化失败");
                }
                else
                {
                    // 如果日志记录器创建失败，回退到原来的方式
                    System.Diagnostics.Debug.WriteLine($"[ERROR] {Name}: 插件初始化失败: {ex.Message}");
                }
                throw;
            }
        }

        public void Load()
        {
            using (_logger?.MeasurePerformance("插件加载"))
            {
                if (!_isInitialized)
                {
                    _logger?.Error("尝试加载未初始化的插件");
                    throw new InvalidOperationException("插件未初始化");
                }

                _logger?.Info("插件加载成功，状态：已就绪");
                _logger?.Debug("插件可用功能数：{0}", GetAllFeatures().Count);
            }
        }

        public void Unload()
        {
            using (_logger?.MeasurePerformance("插件卸载"))
            {
                try
                {
                    _logger?.Info("=== 开始卸载 BasePlugin 插件 ===");
                    _logger?.Debug("清理功能类实例...");
                    
                    // 清理资源
                    _sampleFeatures = null;
                    _logger?.Debug("示例功能类已清理");
                    
                    _excelApp = null;
                    _logger?.Debug("Excel应用程序引用已清理");
                    
                    _isInitialized = false;
                    _logger?.Info("=== BasePlugin 插件卸载完成 ===");
                }
                catch (Exception ex)
                {
                    _logger?.Error(ex, "插件卸载时发生错误");
                }
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
            using (_logger?.MeasurePerformance("获取功能区按钮"))
            {
                var allFeatures = GetAllFeatures();
                var ribbonButtons = new List<RibbonButton>();

                // 按类别分组创建菜单
                var categories = allFeatures.GroupBy(f => f.Category).ToList();
                _logger?.Debug("找到 {0} 个功能类别", categories.Count);

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
                    _logger?.Debug("创建类别菜单: {0}，包含 {1} 个功能", category.Key, category.Count());
                }

                _logger?.Info("成功创建 {0} 个功能区按钮", ribbonButtons.Count);
                return ribbonButtons;
            }
        }

        /// <summary>
        /// 获取所有可执行命令
        /// </summary>
        public List<PluginCommand> GetCommands()
        {
            using (_logger?.MeasurePerformance("获取可执行命令"))
            {
                var allFeatures = GetAllFeatures();
                var commands = allFeatures.Select(f => f.ToPluginCommand()).ToList();
                _logger?.Debug("获取到 {0} 个可执行命令", commands.Count);
                return commands;
            }
        }

        /// <summary>
        /// 搜索命令
        /// </summary>
        public List<PluginCommand> SearchCommands(string keyword)
        {
            using (_logger?.MeasurePerformance("搜索命令"))
            {
                if (string.IsNullOrEmpty(keyword))
                {
                    _logger?.Debug("搜索关键词为空，返回所有命令");
                    return GetCommands();
                }

                var commands = GetCommands();
                var results = commands.Where(cmd =>
                    cmd.Name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    cmd.Description.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    cmd.Category.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    cmd.Tags.Any(tag => tag.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                ).ToList();

                _logger?.Info("搜索关键词 '{0}' 找到 {1} 个匹配命令", keyword, results.Count);
                return results;
            }
        }

        public void ExecuteCommand(string commandId, object[] parameters)
        {
            using (_logger?.MeasurePerformance($"执行命令 {commandId}"))
            {
                try
                {
                    _logger?.Debug("开始执行命令: {0}，参数个数: {1}", commandId, parameters?.Length ?? 0);
                    
                    var allFeatures = GetAllFeatures();
                    var feature = allFeatures.FirstOrDefault(f => f.Id == commandId);
                    
                    if (feature?.Action != null)
                    {
                        feature.Action.Invoke();
                        _logger?.Info("成功执行命令: {0}", commandId);
                    }
                    else
                    {
                        _logger?.Warning("未找到命令: {0}", commandId);
                        throw new ArgumentException($"未知命令: {commandId}");
                    }
                }
                catch (Exception ex)
                {
                    _logger?.Error(ex, "执行命令失败: {0}", commandId);
                    throw;
                }
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
            _logger?.Info("显示消息: {0}", message);
            System.Windows.Forms.MessageBox.Show(message, Name, 
                System.Windows.Forms.MessageBoxButtons.OK, 
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        private void ShowError(string message)
        {
            _logger?.Error("显示错误: {0}", message);
            System.Windows.Forms.MessageBox.Show(message, Name, 
                System.Windows.Forms.MessageBoxButtons.OK, 
                System.Windows.Forms.MessageBoxIcon.Error);
        }

        #endregion
    }
} 