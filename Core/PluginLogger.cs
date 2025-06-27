using System;
using DTI_Tool.AddIn.Common.Interfaces;
using DTI_Tool.AddIn.Core;

namespace BasePlugin.Core
{
    /// <summary>
    /// 插件日志记录器 - 封装宿主日志功能
    /// </summary>
    public class PluginLogger
    {
        private readonly string _pluginName;
        private readonly IHostApplication _hostApp;

        /// <summary>
        /// 初始化插件日志记录器
        /// </summary>
        /// <param name="pluginName">插件名称</param>
        public PluginLogger(string pluginName)
        {
            _pluginName = pluginName ?? throw new ArgumentNullException(nameof(pluginName));
            _hostApp = HostApplication.Instance ?? throw new InvalidOperationException("宿主应用程序未初始化");
        }

        /// <summary>
        /// 记录调试信息（仅在Debug模式下记录）
        /// </summary>
        /// <param name="message">调试消息</param>
        public void Debug(string message)
        {
            if (string.IsNullOrEmpty(message)) return;
            _hostApp.LogDebug(_pluginName, message);
        }

        /// <summary>
        /// 记录调试信息（格式化）
        /// </summary>
        /// <param name="format">格式字符串</param>
        /// <param name="args">参数</param>
        public void Debug(string format, params object[] args)
        {
            if (string.IsNullOrEmpty(format)) return;
            Debug(string.Format(format, args));
        }

        /// <summary>
        /// 记录一般信息
        /// </summary>
        /// <param name="message">信息消息</param>
        public void Info(string message)
        {
            if (string.IsNullOrEmpty(message)) return;
            _hostApp.LogInfo(_pluginName, message);
        }

        /// <summary>
        /// 记录一般信息（格式化）
        /// </summary>
        /// <param name="format">格式字符串</param>
        /// <param name="args">参数</param>
        public void Info(string format, params object[] args)
        {
            if (string.IsNullOrEmpty(format)) return;
            Info(string.Format(format, args));
        }

        /// <summary>
        /// 记录警告信息
        /// </summary>
        /// <param name="message">警告消息</param>
        public void Warning(string message)
        {
            if (string.IsNullOrEmpty(message)) return;
            _hostApp.LogWarning(_pluginName, message);
        }

        /// <summary>
        /// 记录警告信息（格式化）
        /// </summary>
        /// <param name="format">格式字符串</param>
        /// <param name="args">参数</param>
        public void Warning(string format, params object[] args)
        {
            if (string.IsNullOrEmpty(format)) return;
            Warning(string.Format(format, args));
        }

        /// <summary>
        /// 记录错误信息
        /// </summary>
        /// <param name="message">错误消息</param>
        public void Error(string message)
        {
            if (string.IsNullOrEmpty(message)) return;
            _hostApp.LogError(_pluginName, message);
        }

        /// <summary>
        /// 记录错误信息（格式化）
        /// </summary>
        /// <param name="format">格式字符串</param>
        /// <param name="args">参数</param>
        public void Error(string format, params object[] args)
        {
            if (string.IsNullOrEmpty(format)) return;
            Error(string.Format(format, args));
        }

        /// <summary>
        /// 记录错误信息（带异常）
        /// </summary>
        /// <param name="exception">异常对象</param>
        /// <param name="message">错误消息</param>
        public void Error(Exception exception, string message)
        {
            if (exception == null && string.IsNullOrEmpty(message)) return;
            _hostApp.LogError(_pluginName, exception, message ?? "");
        }

        /// <summary>
        /// 记录错误信息（带异常，格式化）
        /// </summary>
        /// <param name="exception">异常对象</param>
        /// <param name="format">格式字符串</param>
        /// <param name="args">参数</param>
        public void Error(Exception exception, string format, params object[] args)
        {
            if (exception == null && string.IsNullOrEmpty(format)) return;
            Error(exception, string.Format(format ?? "", args));
        }

        /// <summary>
        /// 开始性能测量
        /// </summary>
        /// <param name="operationName">操作名称</param>
        /// <returns>性能测量句柄（需要在using语句中使用）</returns>
        public IDisposable MeasurePerformance(string operationName)
        {
            if (string.IsNullOrEmpty(operationName))
                throw new ArgumentException("操作名称不能为空", nameof(operationName));

            return _hostApp.StartPerformanceMeasure(_pluginName, operationName);
        }
    }

    /// <summary>
    /// 插件日志静态工厂类
    /// </summary>
    public static class PluginLog
    {
        /// <summary>
        /// 为指定插件创建日志记录器
        /// </summary>
        /// <param name="pluginName">插件名称</param>
        /// <returns>日志记录器实例</returns>
        public static PluginLogger ForPlugin(string pluginName)
        {
            return new PluginLogger(pluginName);
        }

        /// <summary>
        /// 为当前调用程序集创建日志记录器（自动检测插件名称）
        /// </summary>
        /// <returns>日志记录器实例</returns>
        public static PluginLogger ForCurrentPlugin()
        {
            var assembly = System.Reflection.Assembly.GetCallingAssembly();
            var pluginName = assembly.GetName().Name ?? "UnknownPlugin";
            return new PluginLogger(pluginName);
        }
    }
} 