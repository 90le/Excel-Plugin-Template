# 示例插件配置文件说明

本文档详细说明了插件配置文件 `sample-plugin-config.json` 中各个字段的含义和用法。

## 插件基本信息

```json
{
    "name": "SamplePlugin",                    // 插件名称，必须唯一，用于标识插件
    "version": "1.0.0",                        // 插件版本号，建议使用语义化版本控制 (SemVer)
    "description": "示例插件 - 演示插件开发的各种功能特性",  // 插件描述，简短说明插件功能
    "author": "示例开发者",                     // 插件作者姓名
    "email": "sample@developer.com",           // 作者联系邮箱
    "website": "https://github.com/sample/plugin"  // 插件官方网站或项目地址
}
```

### 字段说明

- **name**: 插件的唯一标识符，不能与其他插件重复
- **version**: 版本号，推荐使用 X.Y.Z 格式（主版本.次版本.修订版本）
- **description**: 插件功能描述，显示在插件管理界面中
- **author**: 插件开发者姓名或团队名称
- **email**: 联系邮箱，用于问题反馈和技术支持
- **website**: 插件官网或源码仓库地址

## 插件入口配置

```json
{
    "entry": "SamplePlugin.dll",               // 插件主程序集文件名
    "mainClass": "SamplePlugin.MainPlugin",    // 插件主类的完全限定名
    "minimumHostVersion": "1.0.0"              // 支持的最低宿主应用版本
}
```

### 字段说明

- **entry**: 插件的主要DLL文件名
- **mainClass**: 实现插件接口的主类，包含命名空间
- **minimumHostVersion**: 插件需要的最低宿主应用版本

## 版权和许可证信息

```json
{
    "copyright": "Copyright © 2024 示例开发者. All rights reserved.",  // 版权声明
    "license": "MIT License",                  // 许可证类型
    "icon": "sample-icon.png"                 // 插件图标文件名（相对于插件目录）
}
```

### 字段说明

- **copyright**: 版权声明文本
- **license**: 许可证类型（如MIT、Apache 2.0、GPL等）
- **icon**: 插件图标文件，相对于插件根目录的路径

## 自动更新配置

```json
{
    "updateInfo": {
        "updateUrl": "https://api.example.com/plugins/sample/updateInfo.json",  // 更新检查URL
        "checkIntervalHours": 24,              // 更新检查间隔（小时）
        "autoCheck": true,                     // 是否启用自动检查更新
        "supportedUpdateModes": ["full", "incremental"],  // 支持的更新模式
        "backupBeforeUpdate": true,            // 更新前是否备份
        "restartRequired": true,               // 更新后是否需要重启
        "verifySignature": true,               // 是否验证更新包签名
        "publicKeyPath": "public.key"          // 公钥文件路径（用于签名验证）
    }
}
```

### 字段说明

- **updateUrl**: 更新信息检查的服务端URL
- **checkIntervalHours**: 自动检查更新的时间间隔（小时）
- **autoCheck**: 是否启用自动更新检查
- **supportedUpdateModes**: 支持的更新模式
  - `full`: 完整更新（替换整个插件）
  - `incremental`: 增量更新（只更新变化的文件）
- **backupBeforeUpdate**: 更新前是否创建备份
- **restartRequired**: 更新完成后是否需要重启应用
- **verifySignature**: 是否验证更新包的数字签名
- **publicKeyPath**: 用于验证签名的公钥文件路径

## 版本更新日志

```json
{
    "changeLog": [
        {
            "version": "1.0.0",                // 版本号
            "date": "2024-12-24",              // 发布日期
            "type": "major",                   // 更新类型
            "changes": [                       // 更新内容列表
                "初始版本发布",
                "实现基础数据处理功能",
                "添加用户界面",
                "支持Excel数据导入导出",
                "集成日志记录系统"
            ]
        }
    ]
}
```

### 字段说明

- **version**: 对应的版本号
- **date**: 版本发布日期（YYYY-MM-DD格式）
- **type**: 更新类型
  - `major`: 主要更新（不兼容的API变更）
  - `minor`: 次要更新（向后兼容的功能增加）
  - `patch`: 补丁更新（向后兼容的问题修复）
- **changes**: 该版本的更新内容列表

## 依赖项配置

```json
{
    "dependencies": [
        {
            "name": "Microsoft.Office.Interop.Excel",  // 依赖包名称
            "version": "15.0.0.0",             // 所需版本
            "required": true                   // 是否为必需依赖
        }
    ]
}
```

### 字段说明

- **name**: 依赖项的名称（通常是NuGet包名或程序集名）
- **version**: 所需的版本号
- **required**: 是否为必需依赖
  - `true`: 必需依赖，缺少时插件无法加载
  - `false`: 可选依赖，缺少时插件仍可运行，但某些功能可能不可用

## 权限配置

```json
{
    "permissions": [
        "Excel.Read",                          // Excel读取权限
        "Excel.Write",                         // Excel写入权限
        "File.Read",                           // 文件读取权限
        "File.Write",                          // 文件写入权限
        "Network.Access"                       // 网络访问权限
    ]
}
```

### 常用权限类型

- **Excel.Read**: 读取Excel文档
- **Excel.Write**: 修改Excel文档
- **File.Read**: 读取文件系统
- **File.Write**: 写入文件系统
- **Network.Access**: 访问网络资源
- **Registry.Read**: 读取注册表
- **Registry.Write**: 修改注册表

## 插件设置配置

```json
{
    "settings": [
        {
            "key": "enableLogging",               // 设置键名
            "type": "boolean",                    // 数据类型
            "default": true,                      // 默认值
            "description": "是否启用详细日志记录"  // 设置说明
        }
    ]
}
```

### 支持的数据类型

- **boolean**: 布尔值（true/false）
- **string**: 字符串
- **number**: 数字（整数或浮点数）
- **array**: 数组

### 字段说明

- **key**: 设置项的唯一标识符
- **type**: 数据类型
- **default**: 默认值
- **description**: 设置项的说明文字

## 插件功能模块配置（可选）

```json
{
    "modules": [
        {
            "name": "DataProcessor",              // 模块名称
            "description": "数据处理模块",        // 模块描述
            "enabled": true,                      // 是否启用
            "assembly": "SamplePlugin.dll",      // 模块程序集
            "className": "SamplePlugin.Modules.DataProcessor"  // 模块类名
        }
    ]
}
```

### 字段说明

- **name**: 模块的唯一名称
- **description**: 模块功能描述
- **enabled**: 是否启用该模块
- **assembly**: 模块所在的程序集文件
- **className**: 模块实现类的完全限定名

## 国际化支持（可选）

```json
{
    "localization": {
        "defaultLanguage": "zh-CN",            // 默认语言
        "supportedLanguages": ["zh-CN", "en-US"],  // 支持的语言列表
        "resourcePath": "./Resources/Languages"     // 语言资源文件路径
    }
}
```

### 字段说明

- **defaultLanguage**: 默认语言代码（如zh-CN、en-US）
- **supportedLanguages**: 插件支持的所有语言列表
- **resourcePath**: 语言资源文件的存储路径

## 插件分类和标签（可选）

```json
{
    "category": "数据处理",                    // 插件分类
    "tags": ["Excel", "数据处理", "自动化", "办公"]  // 插件标签
}
```

### 字段说明

- **category**: 插件所属的主要分类
- **tags**: 插件的标签列表，用于搜索和筛选

## 兼容性信息（可选）

```json
{
    "compatibility": {
        "minDotNetVersion": "4.7.2",          // 最低.NET Framework版本
        "supportedOfficeVersions": ["2016", "2019", "365"],  // 支持的Office版本
        "operatingSystems": ["Windows 10", "Windows 11"]     // 支持的操作系统
    }
}
```

### 字段说明

- **minDotNetVersion**: 所需的最低.NET Framework版本
- **supportedOfficeVersions**: 支持的Microsoft Office版本
- **operatingSystems**: 支持的操作系统版本

## 最佳实践

1. **版本管理**: 使用语义化版本控制，确保版本号的一致性
2. **依赖管理**: 明确标注必需和可选依赖，避免版本冲突
3. **权限最小化**: 只申请插件实际需要的权限
4. **设置合理化**: 提供合理的默认值，减少用户配置负担
5. **文档完整性**: 确保所有字段都有清晰的描述
6. **向后兼容**: 在更新配置格式时保持向后兼容性

## 示例使用场景

### 数据处理插件
适用于需要处理Excel数据、进行计算分析的插件。

### UI增强插件
适用于为Office应用添加自定义界面和交互功能的插件。

### 自动化插件
适用于实现重复任务自动化的插件。

### 集成插件
适用于与外部系统集成的插件，如ERP、CRM系统对接。 