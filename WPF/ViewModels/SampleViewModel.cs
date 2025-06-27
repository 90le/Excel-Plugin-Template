using System;
using System.Windows.Input;
using BasePlugin.WPF.Common;

namespace BasePlugin.WPF.ViewModels
{
    /// <summary>
    /// 示例ViewModel - 展示MVVM模式的使用
    /// </summary>
    public class SampleViewModel : BaseViewModel
    {
        #region 私有字段

        private string _title = "示例窗口";
        private string _message = "这是一个MVVM示例";
        private bool _isEnabled = true;
        private DateTime _selectedDate = DateTime.Today;

        #endregion

        #region 属性

        /// <summary>
        /// 窗口标题
        /// </summary>
        public string Title
        {
            get => _title;
            set => SetProperty(ref _title, value);
        }

        /// <summary>
        /// 显示消息
        /// </summary>
        public string Message
        {
            get => _message;
            set => SetProperty(ref _message, value);
        }

        /// <summary>
        /// 是否启用
        /// </summary>
        public bool IsEnabled
        {
            get => _isEnabled;
            set => SetProperty(ref _isEnabled, value);
        }

        /// <summary>
        /// 选中的日期
        /// </summary>
        public DateTime SelectedDate
        {
            get => _selectedDate;
            set => SetProperty(ref _selectedDate, value);
        }

        #endregion

        #region 命令

        /// <summary>
        /// 确定命令
        /// </summary>
        public ICommand OkCommand { get; }

        /// <summary>
        /// 取消命令
        /// </summary>
        public ICommand CancelCommand { get; }

        /// <summary>
        /// 重置命令
        /// </summary>
        public ICommand ResetCommand { get; }

        #endregion

        #region 构造函数

        public SampleViewModel()
        {
            // 初始化命令
            OkCommand = new RelayCommand(ExecuteOk, CanExecuteOk);
            CancelCommand = new RelayCommand(ExecuteCancel);
            ResetCommand = new RelayCommand(ExecuteReset);
        }

        #endregion

        #region 命令实现

        private bool CanExecuteOk()
        {
            // 验证逻辑
            return IsEnabled && !string.IsNullOrWhiteSpace(Message);
        }

        private void ExecuteOk()
        {
            // 确定按钮逻辑
            Message = $"确定执行于: {DateTime.Now:yyyy-MM-dd HH:mm:ss}";
        }

        private void ExecuteCancel()
        {
            // 取消按钮逻辑
            Message = "操作已取消";
        }

        private void ExecuteReset()
        {
            // 重置到默认值
            Title = "示例窗口";
            Message = "这是一个MVVM示例";
            IsEnabled = true;
            SelectedDate = DateTime.Today;
        }

        #endregion
    }
} 