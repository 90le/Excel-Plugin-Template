# WPF 通用组件目录

本目录用于存放 WPF 相关的通用组件和工具类。

## 建议的子目录结构

```
Common/
├── Converters/           # 值转换器
├── Behaviors/            # 行为类
├── Styles/               # 样式和模板
├── Commands/             # 通用命令
└── Helpers/              # 辅助工具类
```

## 使用说明

- **Converters/**: 实现 IValueConverter 接口的转换器
- **Behaviors/**: 实现附加行为的类
- **Styles/**: XAML 样式和控件模板
- **Commands/**: 实现 ICommand 接口的命令类
- **Helpers/**: WPF 相关的辅助方法和工具类 