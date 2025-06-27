using System;
using System.Collections.Generic;
using System.Linq;
using DTI_Tool.AddIn.Common.Models;
using BasePlugin.Features;
using BasePlugin.Models;

namespace BasePlugin.Core
{
    /// <summary>
    /// 功能管理器 - 统一管理插件的所有功能
    /// </summary>
    public class FeatureManager : IDisposable
    {
        #region 私有字段

        private readonly PluginLogger _logger;
        private readonly Dictionary<string, IFeatureProvider> _featureProviders;
        private List<PluginFeature> _allFeatures;

        #endregion

        #region 构造函数

        /// <summary>
        /// 初始化功能管理器
        /// </summary>
        /// <param name="logger">日志记录器</param>
        public FeatureManager(PluginLogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _featureProviders = new Dictionary<string, IFeatureProvider>();
            _allFeatures = new List<PluginFeature>();
        }

        #endregion

        #region 公共方法

        /// <summary>
        /// 初始化功能管理器
        /// </summary>
        public void Initialize()
        {
            try
            {
                _logger.Debug("正在初始化功能管理器...");

                // 注册所有功能提供者
                RegisterFeatureProviders();

                // 加载所有功能
                LoadAllFeatures();

                _logger.Info("功能管理器初始化完成，共加载 {0} 个功能", _allFeatures.Count);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "功能管理器初始化失败");
                throw;
            }
        }

        /// <summary>
        /// 加载功能
        /// </summary>
        public void Load()
        {
            _logger.Debug("功能管理器已加载");
        }

        /// <summary>
        /// 卸载功能
        /// </summary>
        public void Unload()
        {
            _logger.Debug("正在卸载功能管理器...");
            
            foreach (var provider in _featureProviders.Values)
            {
                try
                {
                    provider.Dispose();
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "卸载功能提供者 {0} 时发生错误", provider.GetType().Name);
                }
            }
            
            _featureProviders.Clear();
            _allFeatures.Clear();
            
            _logger.Debug("功能管理器已卸载");
        }

        /// <summary>
        /// 获取功能区按钮列表
        /// </summary>
        public List<RibbonButton> GetRibbonButtons()
        {
            var ribbonButtons = new List<RibbonButton>();

            // 按类别分组创建菜单
            var categories = _allFeatures.GroupBy(f => f.Category).ToList();
            _logger.Debug("找到 {0} 个功能类别", categories.Count);

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
                _logger.Debug("创建类别菜单: {0}，包含 {1} 个功能", category.Key, category.Count());
            }

            _logger.Info("成功创建 {0} 个功能区按钮", ribbonButtons.Count);
            return ribbonButtons;
        }

        /// <summary>
        /// 获取所有可执行命令
        /// </summary>
        public List<DTI_Tool.AddIn.Common.Interfaces.PluginCommand> GetCommands()
        {
            var commands = _allFeatures.Select(f => f.ToPluginCommand()).ToList();
            _logger.Debug("获取到 {0} 个可执行命令", commands.Count);
            return commands;
        }

        /// <summary>
        /// 搜索命令
        /// </summary>
        public List<DTI_Tool.AddIn.Common.Interfaces.PluginCommand> SearchCommands(string keyword)
        {
            if (string.IsNullOrEmpty(keyword))
            {
                _logger.Debug("搜索关键词为空，返回所有命令");
                return GetCommands();
            }

            var commands = GetCommands();
            var results = commands.Where(cmd =>
                cmd.Name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                cmd.Description.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                cmd.Category.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                cmd.Tags.Any(tag => tag.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
            ).ToList();

            _logger.Info("搜索关键词 '{0}' 找到 {1} 个匹配命令", keyword, results.Count);
            return results;
        }

        /// <summary>
        /// 执行命令
        /// </summary>
        public void ExecuteCommand(string commandId)
        {
            ExecuteCommand(commandId, null);
        }

        /// <summary>
        /// 执行命令
        /// </summary>
        public void ExecuteCommand(string commandId, object[] parameters)
        {
            var feature = _allFeatures.FirstOrDefault(f => f.Id == commandId);
            
            if (feature?.Action != null)
            {
                feature.Action.Invoke();
            }
            else
            {
                _logger.Warning("未找到命令: {0}", commandId);
                throw new ArgumentException($"未知命令: {commandId}");
            }
        }

        #endregion

        #region 私有方法

        /// <summary>
        /// 注册所有功能提供者
        /// </summary>
        private void RegisterFeatureProviders()
        {
            _logger.Debug("正在注册功能提供者...");

            // 基础功能
            RegisterProvider("basic", new BasicFeatures(_logger));
            
            // 数据处理功能
            RegisterProvider("data", new DataProcessingFeatures(_logger));
            
            // 格式化功能
            RegisterProvider("format", new FormattingFeatures(_logger));
            
            // 工作表管理功能
            RegisterProvider("worksheet", new WorksheetFeatures(_logger));
            
            // 实用工具功能
            RegisterProvider("utility", new UtilityFeatures(_logger));
            
            // 窗口演示功能
            RegisterProvider("window", new WindowDemoFeatures(_logger));
            
            // 日志演示功能
            RegisterProvider("logging", new LoggingDemoFeatures(_logger));

            _logger.Debug("已注册 {0} 个功能提供者", _featureProviders.Count);
        }

        /// <summary>
        /// 注册功能提供者
        /// </summary>
        private void RegisterProvider(string key, IFeatureProvider provider)
        {
            _featureProviders[key] = provider;
            _logger.Debug("已注册功能提供者: {0}", provider.GetType().Name);
        }

        /// <summary>
        /// 加载所有功能
        /// </summary>
        private void LoadAllFeatures()
        {
            _allFeatures.Clear();

            foreach (var provider in _featureProviders.Values)
            {
                try
                {
                    var features = provider.GetFeatures();
                    _allFeatures.AddRange(features);
                    _logger.Debug("从 {0} 加载了 {1} 个功能", provider.GetType().Name, features.Count);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "从 {0} 加载功能时发生错误", provider.GetType().Name);
                }
            }

            // 按类别统计功能
            var categories = _allFeatures.GroupBy(f => f.Category).ToList();
            foreach (var category in categories)
            {
                _logger.Debug("类别 '{0}': {1} 个功能", category.Key, category.Count());
            }
        }

        /// <summary>
        /// 根据类别获取图标
        /// </summary>
        private string GetCategoryImageMso(string category)
        {
            return category switch
            {
                "基础功能" => "TableExcelSelect",
                "数据处理" => "DatabaseSortDescending",
                "格式化" => "FontColorPicker",
                "图表" => "ChartColumnChart",
                "统计分析" => "FunctionWizard",
                "工作表管理" => "WorksheetInsert",
                "财务计算" => "AcceptInvitationExcel",
                "实用工具" => "ToolsOptions",
                "窗口演示" => "WindowNew",
                "日志管理" => "BlogPost",
                _ => "FileNew"
            };
        }

        #endregion

        #region IDisposable 实现

        private bool _disposed = false;

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        /// <param name="disposing">是否释放托管资源</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed && disposing)
            {
                Unload();
                _disposed = true;
            }
        }

        #endregion
    }

    /// <summary>
    /// 功能提供者接口
    /// </summary>
    public interface IFeatureProvider : IDisposable
    {
        /// <summary>
        /// 获取功能列表
        /// </summary>
        List<PluginFeature> GetFeatures();
    }
} 