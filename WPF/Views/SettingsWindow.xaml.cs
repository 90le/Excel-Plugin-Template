using System.Windows;
using Microsoft.Win32;

namespace BasePlugin.WPF.Views
{
    /// <summary>
    /// SettingsWindow.xaml 的交互逻辑
    /// </summary>
    public partial class SettingsWindow : Window
    {
        public SettingsWindow()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            SaveSettings();
            DialogResult = true;
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void btnApply_Click(object sender, RoutedEventArgs e)
        {
            SaveSettings();
            MessageBox.Show("设置已应用", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void btnBrowseLog_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "日志文件 (*.log)|*.log|所有文件 (*.*)|*.*",
                DefaultExt = "log",
                FileName = "BasePlugin.log"
            };
            
            if (dialog.ShowDialog() == true)
            {
                txtLogPath.Text = dialog.FileName;
            }
        }
        
        private void LoadSettings()
        {
            // 这里应该从配置文件或注册表加载设置
            // 示例代码仅设置默认值
        }
        
        private void SaveSettings()
        {
            // 这里应该将设置保存到配置文件或注册表
            // 示例代码仅演示获取控件值
            
            var settings = new
            {
                EnableLogging = chkEnableLogging.IsChecked ?? true,
                AutoSave = chkAutoSave.IsChecked ?? false,
                ShowNotifications = chkShowNotifications.IsChecked ?? true,
                DateFormat = (cmbDateFormat.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content?.ToString(),
                NumberFormat = (cmbNumberFormat.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content?.ToString(),
                MaxRows = txtMaxRows.Text,
                BatchSize = txtBatchSize.Text,
                Timeout = txtTimeout.Text,
                LogLevel = (cmbLogLevel.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Content?.ToString(),
                LogToFile = chkLogToFile.IsChecked ?? false,
                LogPath = txtLogPath.Text
            };
            
            // TODO: 实际保存设置
        }
    }
} 