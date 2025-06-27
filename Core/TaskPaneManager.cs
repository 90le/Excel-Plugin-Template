using System;
using System.Reflection;
using BasePlugin.WPF.Views;
using BasePlugin.WinForms;
using DTI_Tool.AddIn.Core;
using DTI_Tool.AddIn.Common.Interfaces;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace BasePlugin.Core
{
    /// <summary>
    /// 任务窗格管理器 - 管理Excel任务窗格
    /// </summary>
    public class TaskPaneManager
    {
        #region 私有字段

        private readonly PluginLogger _logger;
        private readonly IHostApplication _hostApp;
        private readonly string _pluginName;

        #endregion

        #region 构造函数

        /// <summary>
        /// 初始化任务窗格管理器
        /// </summary>
        /// <param name="logger">日志记录器</param>
        /// <param name="pluginName">插件名称（用于区分不同插件的任务窗格）</param>
        public TaskPaneManager(PluginLogger logger, string pluginName = "BasePlugin")
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _pluginName = pluginName;
            
            try
            {
                _hostApp = HostApplication.Instance;
                _logger.Debug($"任务窗格管理器已初始化 - 插件: {_pluginName}");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "获取宿主应用接口失败");
                throw;
            }
        }

        #endregion

        #region 公共方法

        /// <summary>
        /// 初始化任务窗格管理器
        /// </summary>
        public void Initialize()
        {
            try
            {
                _logger.Debug("正在初始化任务窗格管理器...");
                // 检查宿主应用接口是否可用
                if (_hostApp == null)
                {
                    throw new InvalidOperationException("宿主应用接口不可用");
                }
                _logger.Debug("任务窗格管理器初始化完成");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "任务窗格管理器初始化失败");
                throw;
            }
        }

        /// <summary>
        /// 加载任务窗格
        /// </summary>
        public void Load()
        {
            _logger.Debug("任务窗格管理器已加载");
        }

        /// <summary>
        /// 卸载任务窗格
        /// </summary>
        public void Unload()
        {
            try
            {
                _logger.Debug("正在卸载任务窗格管理器...");
                
                // 关闭所有相关的任务窗格
                CloseAllTaskPanes();
                
                _logger.Debug("任务窗格管理器已卸载");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "卸载任务窗格管理器时发生错误");
            }
        }

        /// <summary>
        /// 显示或隐藏WPF任务窗格
        /// </summary>
        /// <param name="taskPaneName">任务窗格名称</param>
        /// <param name="visible">是否显示</param>
        public void ShowWpfTaskPane(string taskPaneName, bool visible)
        {
            try
            {
                if (_hostApp == null)
                {
                    _logger.Error("宿主应用接口不可用");
                    return;
                }

                var fullTaskPaneName = GetFullTaskPaneName(taskPaneName);
                
                if (!_hostApp.TaskPaneExists(fullTaskPaneName) && visible)
                {
                    // 如果任务窗格不存在且需要显示，则创建
                    CreateWpfTaskPane(taskPaneName);
                }
                else
                {
                    // 设置显示状态
                    _hostApp.ShowTaskPane(fullTaskPaneName, visible);
                    _logger.Info("WPF任务窗格 {0} 已{1}", taskPaneName, visible ? "显示" : "隐藏");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "显示/隐藏WPF任务窗格时发生错误: {0}", taskPaneName);
            }
        }

        /// <summary>
        /// 显示或隐藏WinForms任务窗格
        /// </summary>
        /// <param name="taskPaneName">任务窗格名称</param>
        /// <param name="visible">是否显示</param>
        public void ShowWinFormsTaskPane(string taskPaneName, bool visible)
        {
            try
            {
                if (_hostApp == null)
                {
                    _logger.Error("宿主应用接口不可用");
                    return;
                }

                var fullTaskPaneName = GetFullTaskPaneName(taskPaneName);
                
                if (!_hostApp.TaskPaneExists(fullTaskPaneName) && visible)
                {
                    // 如果任务窗格不存在且需要显示，则创建
                    CreateWinFormsTaskPane(taskPaneName);
                }
                else
                {
                    // 设置显示状态
                    _hostApp.ShowTaskPane(fullTaskPaneName, visible);
                    _logger.Info("WinForms任务窗格 {0} 已{1}", taskPaneName, visible ? "显示" : "隐藏");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "显示/隐藏WinForms任务窗格时发生错误: {0}", taskPaneName);
            }
        }

        /// <summary>
        /// 切换WPF任务窗格的显示状态
        /// </summary>
        /// <param name="taskPaneName">任务窗格名称</param>
        public void ToggleWpfTaskPane(string taskPaneName)
        {
            try
            {
                if (_hostApp == null)
                {
                    _logger.Error("宿主应用接口不可用");
                    return;
                }

                var fullTaskPaneName = GetFullTaskPaneName(taskPaneName);
                
                if (!_hostApp.TaskPaneExists(fullTaskPaneName))
                {
                    // 任务窗格不存在，创建并显示
                    CreateWpfTaskPane(taskPaneName);
                }
                else
                {
                    // 切换显示状态
                    var currentVisible = _hostApp.IsTaskPaneVisible(fullTaskPaneName);
                    _hostApp.ShowTaskPane(fullTaskPaneName, !currentVisible);
                    _logger.Info("WPF任务窗格 {0} 已{1}", taskPaneName, !currentVisible ? "显示" : "隐藏");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "切换WPF任务窗格显示状态时发生错误: {0}", taskPaneName);
            }
        }

        /// <summary>
        /// 切换WinForms任务窗格的显示状态
        /// </summary>
        /// <param name="taskPaneName">任务窗格名称</param>
        public void ToggleWinFormsTaskPane(string taskPaneName)
        {
            try
            {
                if (_hostApp == null)
                {
                    _logger.Error("宿主应用接口不可用");
                    return;
                }

                var fullTaskPaneName = GetFullTaskPaneName(taskPaneName);
                
                if (!_hostApp.TaskPaneExists(fullTaskPaneName))
                {
                    // 任务窗格不存在，创建并显示
                    CreateWinFormsTaskPane(taskPaneName);
                }
                else
                {
                    // 切换显示状态
                    var currentVisible = _hostApp.IsTaskPaneVisible(fullTaskPaneName);
                    _hostApp.ShowTaskPane(fullTaskPaneName, !currentVisible);
                    _logger.Info("WinForms任务窗格 {0} 已{1}", taskPaneName, !currentVisible ? "显示" : "隐藏");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "切换WinForms任务窗格显示状态时发生错误: {0}", taskPaneName);
            }
        }

        /// <summary>
        /// 创建WPF任务窗格
        /// </summary>
        /// <param name="taskPaneName">任务窗格名称</param>
        /// <param name="width">任务窗格宽度</param>
        public void CreateWpfTaskPane(string taskPaneName, int width = 350)
        {
            try
            {
                if (_hostApp == null)
                {
                    _logger.Error("宿主应用接口不可用");
                    return;
                }

                var fullTaskPaneName = GetFullTaskPaneName(taskPaneName);
                
                // 创建WPF任务窗格内容
                var taskPaneView = CreateWpfTaskPaneContent(taskPaneName);
                if (taskPaneView != null)
                {
                    _hostApp.ToggleWpfTaskPane(fullTaskPaneName, taskPaneView, width);
                    _logger.Info("WPF任务窗格已创建: {0}", taskPaneName);
                }
                else
                {
                    _logger.Error("无法创建WPF任务窗格内容: {0}", taskPaneName);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "创建WPF任务窗格失败: {0}", taskPaneName);
            }
        }

        /// <summary>
        /// 创建WinForms任务窗格
        /// </summary>
        /// <param name="taskPaneName">任务窗格名称</param>
        /// <param name="width">任务窗格宽度</param>
        public void CreateWinFormsTaskPane(string taskPaneName, int width = 320)
        {
            try
            {
                if (_hostApp == null)
                {
                    _logger.Error("宿主应用接口不可用");
                    return;
                }

                var fullTaskPaneName = GetFullTaskPaneName(taskPaneName);
                
                // 创建WinForms任务窗格内容
                var taskPaneControl = CreateWinFormsTaskPaneContent(taskPaneName);
                if (taskPaneControl != null)
                {
                    _hostApp.ToggleWinFormsTaskPane(fullTaskPaneName, taskPaneControl, width);
                    _logger.Info("WinForms任务窗格已创建: {0}", taskPaneName);
                }
                else
                {
                    _logger.Error("无法创建WinForms任务窗格内容: {0}", taskPaneName);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "创建WinForms任务窗格失败: {0}", taskPaneName);
            }
        }

        /// <summary>
        /// 检查任务窗格是否存在
        /// </summary>
        /// <param name="taskPaneName">任务窗格名称</param>
        /// <returns>如果任务窗格存在返回true</returns>
        public bool TaskPaneExists(string taskPaneName)
        {
            try
            {
                if (_hostApp == null) return false;
                
                var fullTaskPaneName = GetFullTaskPaneName(taskPaneName);
                return _hostApp.TaskPaneExists(fullTaskPaneName);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "检查任务窗格存在性失败: {0}", taskPaneName);
                return false;
            }
        }

        /// <summary>
        /// 检查任务窗格是否可见
        /// </summary>
        /// <param name="taskPaneName">任务窗格名称</param>
        /// <returns>如果任务窗格可见返回true</returns>
        public bool IsTaskPaneVisible(string taskPaneName)
        {
            try
            {
                if (_hostApp == null) return false;
                
                var fullTaskPaneName = GetFullTaskPaneName(taskPaneName);
                return _hostApp.IsTaskPaneVisible(fullTaskPaneName);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "检查任务窗格可见性失败: {0}", taskPaneName);
                return false;
            }
        }

        /// <summary>
        /// 关闭指定的任务窗格
        /// </summary>
        /// <param name="taskPaneName">任务窗格名称</param>
        public void CloseTaskPane(string taskPaneName)
        {
            try
            {
                if (_hostApp == null)
                {
                    _logger.Error("宿主应用接口不可用");
                    return;
                }

                var fullTaskPaneName = GetFullTaskPaneName(taskPaneName);
                _hostApp.CloseTaskPane(fullTaskPaneName);
                _logger.Info("任务窗格已关闭: {0}", taskPaneName);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "关闭任务窗格失败: {0}", taskPaneName);
            }
        }

        /// <summary>
        /// 关闭所有相关的任务窗格
        /// </summary>
        public void CloseAllTaskPanes()
        {
            try
            {
                if (_hostApp == null)
                {
                    _logger.Error("宿主应用接口不可用");
                    return;
                }

                // 这里可以实现更精确的关闭逻辑，只关闭属于此插件的任务窗格
                // 暂时关闭所有任务窗格
                _hostApp.CloseAllTaskPanes();
                _logger.Info("所有任务窗格已关闭");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "关闭所有任务窗格失败");
            }
        }

        #endregion

        #region 私有方法

        /// <summary>
        /// 获取完整的任务窗格名称（包含插件前缀）
        /// </summary>
        /// <param name="taskPaneName">任务窗格名称</param>
        /// <returns>完整的任务窗格名称</returns>
        private string GetFullTaskPaneName(string taskPaneName)
        {
            return $"{_pluginName}_{taskPaneName}";
        }

        /// <summary>
        /// 创建WPF任务窗格内容
        /// </summary>
        /// <param name="taskPaneName">任务窗格名称</param>
        /// <returns>WPF用户控件</returns>
        private System.Windows.Controls.UserControl CreateWpfTaskPaneContent(string taskPaneName)
        {
            try
            {
                _logger.Debug("正在创建WPF任务窗格内容: {0}", taskPaneName);

                // 根据任务窗格名称创建不同的内容
                switch (taskPaneName.ToLower())
                {
                    case "main":
                    case "default":
                    case "demotaskpane":
                    case "wpf":
                    case "wpftaskpane":
                        return new TaskPaneView();
                    
                    case "demo":
                        // 创建演示任务窗格
                        var demoPane = new TaskPaneView();
                        _logger.Info("创建了WPF演示任务窗格");
                        return demoPane;
                    
                    case "settings":
                        // 可以在未来添加设置界面
                        return new TaskPaneView(); // 暂时使用默认视图
                    
                    default:
                        _logger.Warning("未知的WPF任务窗格类型: {0}, 使用默认视图", taskPaneName);
                        return new TaskPaneView();
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "创建WPF任务窗格内容失败: {0}", taskPaneName);
                return null;
            }
        }

        /// <summary>
        /// 创建WinForms任务窗格内容
        /// </summary>
        /// <param name="taskPaneName">任务窗格名称</param>
        /// <returns>WinForms用户控件</returns>
        private System.Windows.Forms.UserControl CreateWinFormsTaskPaneContent(string taskPaneName)
        {
            try
            {
                _logger.Debug("正在创建WinForms任务窗格内容: {0}", taskPaneName);

                // 根据任务窗格名称创建不同的内容
                switch (taskPaneName.ToLower())
                {
                    case "winforms":
                    case "winformstaskpane":
                    case "forms":
                        return new TaskPaneControl();
                    
                    case "demo":
                        // 创建演示任务窗格
                        var demoPane = new TaskPaneControl();
                        _logger.Info("创建了WinForms演示任务窗格");
                        return demoPane;
                    
                    case "settings":
                        // 可以在未来添加设置界面
                        return new TaskPaneControl(); // 暂时使用默认视图
                    
                    default:
                        _logger.Warning("未知的WinForms任务窗格类型: {0}, 使用默认视图", taskPaneName);
                        return new TaskPaneControl();
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "创建WinForms任务窗格内容失败: {0}", taskPaneName);
                return null;
            }
        }

        #endregion

        #region 兼容性方法（向后兼容）

        /// <summary>
        /// 显示或隐藏默认任务窗格（向后兼容）
        /// </summary>
        /// <param name="visible">是否显示</param>
        public void ShowTaskPane(bool visible)
        {
            ShowWpfTaskPane("Default", visible);
        }

        /// <summary>
        /// 切换默认任务窗格的显示状态（向后兼容）
        /// </summary>
        public void ToggleTaskPane()
        {
            ToggleWpfTaskPane("Default");
        }

        #endregion
    }
} 