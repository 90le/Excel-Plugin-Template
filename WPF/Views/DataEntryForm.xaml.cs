using System;
using System.Windows;
using System.Windows.Controls;
using BasePlugin.Features;

namespace BasePlugin.WPF.Views
{
    /// <summary>
    /// DataEntryForm.xaml 的交互逻辑
    /// </summary>
    public partial class DataEntryForm : Window
    {
        public DataEntryForm()
        {
            InitializeComponent();
            
            // 设置数据上下文
            DataContext = new DataEntryViewModel
            {
                JoinDate = DateTime.Today
            };
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as DataEntryViewModel;
            if (viewModel == null) return;
            
            // 验证数据
            if (string.IsNullOrWhiteSpace(viewModel.Name))
            {
                MessageBox.Show("请输入姓名", "验证错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtName.Focus();
                return;
            }
            
            if (string.IsNullOrWhiteSpace(viewModel.Email))
            {
                MessageBox.Show("请输入邮箱", "验证错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtEmail.Focus();
                return;
            }
            
            if (string.IsNullOrWhiteSpace(viewModel.Phone))
            {
                MessageBox.Show("请输入电话", "验证错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtPhone.Focus();
                return;
            }
            
            if (cmbDepartment.SelectedItem == null)
            {
                MessageBox.Show("请选择部门", "验证错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                cmbDepartment.Focus();
                return;
            }
            
            // 从ComboBoxItem获取内容
            var selectedItem = cmbDepartment.SelectedItem as ComboBoxItem;
            viewModel.Department = selectedItem?.Content?.ToString();
            
            if (!viewModel.JoinDate.HasValue)
            {
                MessageBox.Show("请选择入职日期", "验证错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                dpJoinDate.Focus();
                return;
            }
            
            // 验证邮箱格式
            if (!IsValidEmail(viewModel.Email))
            {
                MessageBox.Show("邮箱格式不正确", "验证错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtEmail.Focus();
                return;
            }
            
            DialogResult = true;
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        
        private bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }
    }
} 