using System.Windows;

namespace BasePlugin.WPF.Views
{
    /// <summary>
    /// SimpleWindow.xaml 的交互逻辑
    /// </summary>
    public partial class SimpleWindow : Window
    {
        public SimpleWindow()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
} 