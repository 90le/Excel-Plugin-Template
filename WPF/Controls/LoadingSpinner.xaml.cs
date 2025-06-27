using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;

namespace BasePlugin.WPF.Controls
{
    /// <summary>
    /// LoadingSpinner.xaml 的交互逻辑
    /// </summary>
    public partial class LoadingSpinner : UserControl
    {
        private Storyboard _spinAnimation;

        public LoadingSpinner()
        {
            InitializeComponent();
            _spinAnimation = FindResource("SpinAnimation") as Storyboard;
            Loaded += OnLoaded;
        }

        /// <summary>
        /// 加载文本依赖属性
        /// </summary>
        public static readonly DependencyProperty LoadingTextProperty =
            DependencyProperty.Register(
                nameof(LoadingMessage), 
                typeof(string), 
                typeof(LoadingSpinner), 
                new PropertyMetadata("加载中...", OnLoadingTextChanged));

        /// <summary>
        /// 获取或设置加载文本
        /// </summary>
        public string LoadingMessage
        {
            get { return (string)GetValue(LoadingTextProperty); }
            set { SetValue(LoadingTextProperty, value); }
        }

        /// <summary>
        /// 是否正在旋转依赖属性
        /// </summary>
        public static readonly DependencyProperty IsSpinningProperty =
            DependencyProperty.Register(
                nameof(IsSpinning), 
                typeof(bool), 
                typeof(LoadingSpinner), 
                new PropertyMetadata(true, OnIsSpinningChanged));

        /// <summary>
        /// 获取或设置是否正在旋转
        /// </summary>
        public bool IsSpinning
        {
            get { return (bool)GetValue(IsSpinningProperty); }
            set { SetValue(IsSpinningProperty, value); }
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            if (IsSpinning)
            {
                StartAnimation();
            }
        }

        private static void OnLoadingTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var spinner = d as LoadingSpinner;
            if (spinner?.LoadingText != null)
            {
                spinner.LoadingText.Text = e.NewValue?.ToString() ?? "";
            }
        }

        private static void OnIsSpinningChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var spinner = d as LoadingSpinner;
            if (spinner == null) return;

            if ((bool)e.NewValue)
            {
                spinner.StartAnimation();
            }
            else
            {
                spinner.StopAnimation();
            }
        }

        /// <summary>
        /// 开始动画
        /// </summary>
        public void StartAnimation()
        {
            _spinAnimation?.Begin();
        }

        /// <summary>
        /// 停止动画
        /// </summary>
        public void StopAnimation()
        {
            _spinAnimation?.Stop();
        }
    }
} 