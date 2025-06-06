# WPF 视图目录

本目录用于存放 WPF 视图文件 (.xaml)。

## 命名规范

- 视图文件以 `View` 结尾：`SampleView.xaml`
- 对应的代码文件：`SampleView.xaml.cs`

## 示例结构

```xml
<UserControl x:Class="BasePlugin.WPF.Views.SampleView"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <Grid>
        <!-- 您的界面内容 -->
    </Grid>
</UserControl>
```

## 最佳实践

1. 使用 UserControl 作为基础控件
2. 遵循 MVVM 模式，通过 DataContext 绑定 ViewModel
3. 使用数据绑定而不是代码后置处理业务逻辑
4. 保持 View 的纯净，只处理界面相关的逻辑 