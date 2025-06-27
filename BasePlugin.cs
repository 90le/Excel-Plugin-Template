using System;
using System.Collections.Generic;
using DTI_Tool.AddIn.Common.Interfaces;
using DTI_Tool.AddIn.Common.Models;
using BasePlugin.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasePlugin
{
    /// <summary>
    /// 基础插件模板 - 演示如何开发DTI Tool插件
    /// </summary>
    public class BasePlugin : IPlugin
    {
        #region 私有字段

        private PluginLogger _logger;
        private FeatureManager _featureManager;
        private TaskPaneManager _taskPaneManager;

        #endregion

        #region IPlugin 接口属性

        public string Name => "BasePlugin";
        public string Version => "1.0.0";
        public string Description => "基础插件开发模板 - 为开发者提供完整的示例代码";
        public string Author => "开发者姓名";

        #endregion

        #region IPlugin 接口方法

        /// <summary>
        /// 初始化插件
        /// </summary>
        public void Initialize()
        {
            try
            {
                // 初始化日志记录器
                _logger = PluginLog.ForPlugin(Name);
                _logger.Info("=== 正在初始化 BasePlugin 插件 ===");
                _logger.Debug("插件版本: {0}, 作者: {1}", Version, Author);

                // 初始化功能管理器
                _featureManager = new FeatureManager(_logger);
                _featureManager.Initialize();

                // 初始化任务窗格管理器
                _taskPaneManager = new TaskPaneManager(_logger);
                _taskPaneManager.Initialize();

                _logger.Info("=== BasePlugin 插件初始化完成 ===");
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "插件初始化失败");
                throw;
            }
        }

        /// <summary>
        /// 加载插件
        /// </summary>
        public void Load()
        {
            using (_logger?.MeasurePerformance("插件加载"))
            {
                try
                {
                    _logger?.Info("正在加载插件...");
                    
                    // 加载功能
                    _featureManager?.Load();
                    
                    // 加载任务窗格
                    _taskPaneManager?.Load();

                    _logger?.Info("插件加载完成，状态：已就绪");
                }
                catch (Exception ex)
                {
                    _logger?.Error(ex, "插件加载失败");
                    throw;
                }
            }
        }

        /// <summary>
        /// 卸载插件
        /// </summary>
        public void Unload()
        {
            using (_logger?.MeasurePerformance("插件卸载"))
            {
                try
                {
                    _logger?.Info("=== 正在卸载 BasePlugin 插件 ===");

                    // 卸载任务窗格
                    _taskPaneManager?.Unload();

                    // 卸载功能
                    _featureManager?.Unload();

                    _logger?.Info("=== BasePlugin 插件卸载完成 ===");
                }
                catch (Exception ex)
                {
                    _logger?.Error(ex, "插件卸载时发生错误");
                }
            }
        }

        /// <summary>
        /// 释放资源
        /// </summary>
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
                return _featureManager?.GetRibbonButtons() ?? new List<RibbonButton>();
            }
        }

        /// <summary>
        /// 获取所有可执行命令
        /// </summary>
        public List<DTI_Tool.AddIn.Common.Interfaces.PluginCommand> GetCommands()
        {
            using (_logger?.MeasurePerformance("获取可执行命令"))
            {
                return _featureManager?.GetCommands() ?? new List<DTI_Tool.AddIn.Common.Interfaces.PluginCommand>();
            }
        }

        /// <summary>
        /// 搜索命令
        /// </summary>
        public List<DTI_Tool.AddIn.Common.Interfaces.PluginCommand> SearchCommands(string keyword)
        {
            using (_logger?.MeasurePerformance("搜索命令"))
            {
                return _featureManager?.SearchCommands(keyword) ?? new List<DTI_Tool.AddIn.Common.Interfaces.PluginCommand>();
            }
        }

        /// <summary>
        /// 执行命令
        /// </summary>
        public void ExecuteCommand(string commandId, object[] parameters)
        {
            using (_logger?.MeasurePerformance($"执行命令 {commandId}"))
            {
                try
                {
                    _logger?.Debug("开始执行命令: {0}，参数个数: {1}", commandId, parameters?.Length ?? 0);
                    _featureManager?.ExecuteCommand(commandId, parameters);
                    _logger?.Info("成功执行命令: {0}", commandId);
                }
                catch (Exception ex)
                {
                    _logger?.Error(ex, "执行命令失败: {0}", commandId);
                    throw;
                }
            }
        }

        #endregion
    }
} 