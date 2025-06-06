# WPF 视图模型目录

本目录用于存放 MVVM 模式的视图模型类。

## 命名规范

- 视图模型以 `ViewModel` 结尾：`SampleViewModel.cs`
- 继承 `INotifyPropertyChanged` 接口

## 示例结构

```csharp
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace BasePlugin.WPF.ViewModels
{
    public class SampleViewModel : INotifyPropertyChanged
    {
        private string _title;
        
        public string Title
        {
            get => _title;
            set
            {
                _title = value;
                OnPropertyChanged();
            }
        }
        
        public ICommand SampleCommand { get; }
        
        public SampleViewModel()
        {
            SampleCommand = new RelayCommand(ExecuteSample);
        }
        
        private void ExecuteSample()
        {
            // 命令执行逻辑
        }
        
        public event PropertyChangedEventHandler PropertyChanged;
        
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
```

## 最佳实践

1. 实现 `INotifyPropertyChanged` 接口
2. 使用命令模式处理用户交互
3. 避免在 ViewModel 中直接操作 UI 元素
4. 通过依赖注入获取服务 