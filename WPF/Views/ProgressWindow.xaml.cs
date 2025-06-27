using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using BasePlugin.Features;

namespace BasePlugin.WPF.Views
{
    /// <summary>
    /// ProgressWindow.xaml 的交互逻辑
    /// </summary>
    public partial class ProgressWindow : Window
    {
        private CancellationTokenSource _cancellationTokenSource;
        private IProgress<ProgressInfo> _progressReporter;

        public ProgressWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 启动异步任务
        /// </summary>
        /// <param name="taskAction">要执行的任务</param>
        public void StartTask(Func<IProgress<ProgressInfo>, CancellationToken, Task> taskAction)
        {
            _cancellationTokenSource = new CancellationTokenSource();
            _progressReporter = new Progress<ProgressInfo>(UpdateProgress);
            
            // 在后台线程执行任务
            Task.Run(async () =>
            {
                try
                {
                    await taskAction(_progressReporter, _cancellationTokenSource.Token);
                    
                    // 任务完成，关闭窗口
                    Dispatcher.Invoke(() =>
                    {
                        DialogResult = true;
                        Close();
                    });
                }
                catch (OperationCanceledException)
                {
                    // 任务被取消
                    Dispatcher.Invoke(() =>
                    {
                        DialogResult = false;
                        Close();
                    });
                }
                catch (Exception ex)
                {
                    // 发生错误
                    Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show($"任务执行失败: {ex.Message}", "错误", 
                            MessageBoxButton.OK, MessageBoxImage.Error);
                        DialogResult = false;
                        Close();
                    });
                }
            });
        }

        /// <summary>
        /// 更新进度
        /// </summary>
        private void UpdateProgress(ProgressInfo info)
        {
            progressBar.Value = info.Progress;
            txtProgress.Text = $"{info.Progress}%";
            txtDetails.Text = info.Message;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            _cancellationTokenSource?.Cancel();
            btnCancel.IsEnabled = false;
            txtDetails.Text = "正在取消任务...";
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            _cancellationTokenSource?.Cancel();
            _cancellationTokenSource?.Dispose();
        }
    }
} 