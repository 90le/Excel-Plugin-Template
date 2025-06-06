# WPF 自定义控件目录

本目录用于存放自定义的 WPF 用户控件。

## 命名规范

- 控件文件以 `Control` 结尾：`DataGridControl.xaml`
- 对应的代码文件：`DataGridControl.xaml.cs`

## 示例结构

```xml
<UserControl x:Class="BasePlugin.WPF.Controls.SampleControl"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <Grid>
        <!-- 控件内容 -->
    </Grid>
</UserControl>
```

```csharp
using System.Windows.Controls;

namespace BasePlugin.WPF.Controls
{
    public partial class SampleControl : UserControl
    {
        public SampleControl()
        {
            InitializeComponent();
        }
    }
}
```

## 使用场景

- 可复用的复杂UI组件
- 特定业务逻辑的封装控件
- 第三方控件的包装
- 自定义输入控件

## 最佳实践

1. 封装可复用的功能
2. 提供必要的依赖属性
3. 保持控件的独立性
4. 提供清晰的公共接口 