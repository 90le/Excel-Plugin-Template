# BasePlugin 基础插件开发模板

## 项目概述

BasePlugin 是一个基于 DTI_Tool.AddIn 框架的 Excel 插件开发基础模板，为开发者提供快速开发框架和最佳实践示例。

> **关于 DTI_Tool.AddIn 框架**
>
> DTI_Tool.AddIn 是一个强大的热拔插插件开发框架，支持 Excel 和 WPS Office 双平台。通过该框架，您可以：
> - 🔌 **热拔插支持**：无需重启应用程序即可加载/卸载插件
> - 🎯 **双平台兼容**：同时支持 Microsoft Excel 和 WPS Office
> - 📦 **简单部署**：将编译后的插件文件放置到指定目录即可自动加载
> - 🛡️ **安全稳定**：插件运行在独立的应用程序域中，确保宿主应用的稳定性

**插件信息：**
- 名称：基础插件开发模板
- 版本：1.0.0
- 作者：开发者姓名（请修改为您的信息）
- 邮箱：developer@example.com（请修改为您的邮箱）
- 官网：https://example.com（请修改为您的官网）

## 快速开始

### 1. 获取模板
将此模板复制到您的开发目录，并重命名为您的插件名称。

### 2. 自定义插件信息
修改以下文件中的基本信息：

**BasePlugin.cs**
```csharp
public string Name => "您的插件名称";
public string Description => "您的插件描述";
public string Author => "您的姓名";
```

**manifest.json**
```json
{
    "name": "YourPluginName",
    "description": "您的插件描述",
    "author": "您的姓名",
    "email": "您的邮箱",
    "website": "您的官网"
}
```

**项目文件（.csproj）**
- 重命名 `BasePlugin.csproj` 为 `YourPluginName.csproj`
- 修改项目文件中的 `AssemblyName` 和 `RootNamespace`

### 3. 添加您的功能
在 `Features/` 目录下创建新的功能类，参考 `SampleFeatures.cs` 的实现方式。

## 目录结构说明

```
BasePlugin/                            # 插件根目录
├── .vs/                               # Visual Studio 配置目录（自动生成）
├── bin/                               # 编译输出目录（自动生成）
├── obj/                               # 编译缓存目录（自动生成）
├── Features/                          # 功能实现目录 ⭐
│   └── SampleFeatures.cs              # 示例功能类
├── Models/                            # 数据模型目录 ⭐
│   └── PluginFeature.cs               # 插件功能基础模型
├── WPF/                               # WPF界面目录 ⭐
│   ├── Common/                        # WPF通用组件
│   ├── Views/                         # WPF视图
│   ├── ViewModels/                    # MVVM视图模型
│   └── Controls/                      # 自定义控件
├── manifest.json                      # 插件配置文件 ⭐
├── BasePlugin.cs                      # 插件主入口文件 ⭐
├── BasePlugin.csproj                  # 项目文件 ⭐
└── README.md                          # 项目说明文档
```

## 开发指南

### 1. 添加新功能

#### 1.1 创建功能类
在 `Features/` 目录下创建新的功能类：

```csharp
using System;
using System.Collections.Generic;
using BasePlugin.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasePlugin.Features
{
    public class YourFeatures
    {
        private Excel.Application _excelApp;
        
        public YourFeatures(Excel.Application excelApp)
        {
            _excelApp = excelApp;
        }
        
        public List<PluginFeature> GetFeatures()
        {
            return new List<PluginFeature>
            {
                new PluginFeature
                {
                    Id = "your_feature_id",
                    Name = "您的功能名称",
                    Description = "功能描述",
                    Category = "功能类别",
                    Tags = new List<string> { "标签1", "标签2" },
                    ImageMso = "FileNew", // Office图标
                    Action = YourFeatureMethod
                }
            };
        }
        
        private void YourFeatureMethod()
        {
            // 实现您的功能逻辑
        }
    }
}
```

#### 1.2 注册功能类
在 `BasePlugin.cs` 中注册新功能类：

```csharp
// 在类的顶部添加字段
private YourFeatures _yourFeatures;

// 在 Initialize() 方法中初始化
_yourFeatures = new YourFeatures(_excelApp);

// 在 GetAllFeatures() 方法中添加功能
if (_yourFeatures != null)
    allFeatures.AddRange(_yourFeatures.GetFeatures());
```

### 2. 支持的界面技术

#### 2.1 WinForms
- 适用于简单的对话框和工具窗口
- 直接使用 `System.Windows.Forms` 命名空间
- 示例：消息框、输入对话框

#### 2.2 WPF
- 适用于复杂的现代界面
- 遵循MVVM模式
- 支持数据绑定和命令模式

### 3. Excel 操作常用方法

```csharp
// 获取当前活动单元格
var activeCell = _excelApp?.ActiveCell;

// 获取当前选择区域
var selection = _excelApp?.Selection as Excel.Range;

// 获取当前工作簿
var workbook = _excelApp?.ActiveWorkbook;

// 获取当前工作表
var worksheet = _excelApp?.ActiveSheet as Excel.Worksheet;
```

### 4. 错误处理最佳实践

```csharp
private void YourFeatureMethod()
{
    try
    {
        // 您的功能逻辑
    }
    catch (Exception ex)
    {
        ShowError($"功能执行失败: {ex.Message}");
    }
}
```

## 内置示例功能

模板包含以下示例功能：

1. **Hello World** - 显示问候消息
2. **获取选择信息** - 显示当前选中区域的基本信息
3. **插入当前时间** - 在活动单元格插入当前日期和时间

这些示例展示了基本的插件开发模式，您可以参考这些实现来开发自己的功能。

## 配置文件说明

### manifest.json 基础配置

```json
{
    "name": "插件名称",
    "version": "版本号",
    "description": "插件描述",
    "author": "作者姓名",
    "email": "联系邮箱",
    "website": "官方网站",
    "entry": "DLL文件名",
    "mainClass": "主类完整名称",
    "minimumHostVersion": "最低宿主版本",
    "permissions": ["Excel.Read", "Excel.Write"],
    "settings": [
        {
            "key": "设置键名",
            "type": "数据类型",
            "default": "默认值",
            "description": "设置描述"
        }
    ]
}
```

## 插件自动更新配置

### 更新配置字段说明

在 `manifest.json` 中添加更新配置：

```json
{
  "updateInfo": {
    "updateUrl": "https://example.com/api/plugins/update-check",
    "downloadUrl": "https://example.com/api/plugins/download",
    "checkIntervalHours": 24,
    "autoCheck": true,
    "supportedUpdateModes": ["full", "incremental"],
    "backupBeforeUpdate": true,
    "restartRequired": false,
    "verifySignature": false,
    "publicKeyPath": "public.key"
  }
}
```

#### 配置字段详解

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `updateUrl` | string | ✅ | 更新检查API地址 |
| `downloadUrl` | string | ❌ | 文件下载地址（可选，如果与更新检查地址不同） |
| `checkIntervalHours` | number | ❌ | 检查间隔时间（小时），默认24 |
| `autoCheck` | boolean | ❌ | 是否自动检查更新，默认true |
| `supportedUpdateModes` | array | ❌ | 支持的更新模式，["full", "incremental"] |
| `backupBeforeUpdate` | boolean | ❌ | 更新前是否备份，默认true |
| `restartRequired` | boolean | ❌ | 更新后是否需要重启，默认false |
| `verifySignature` | boolean | ❌ | 是否验证数字签名，默认false |
| `publicKeyPath` | string | ❌ | 公钥文件路径（验证签名时需要） |

### 更新服务器开发

#### 更新检查接口

更新检查服务器需要返回以下格式的 JSON 响应：

```json
{
  "latestVersion": "1.1.0",
  "downloadUrl": "https://example.com/plugins/plugin-v1.1.0.zip",
  "updateMode": "full",
  "fileSize": 1048576,
  "fileHash": "sha256:e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
  "isForced": false,
  "releaseNotes": "版本更新说明\n\n新增功能:\n- 功能1\n- 功能2\n\n修复问题:\n- 问题1\n- 问题2",
  "releaseDate": "2024-12-24",
  "minCompatibleVersion": "1.0.0"
}
```

#### 响应字段说明

| 字段 | 类型 | 必需 | 说明 |
|------|------|------|------|
| `latestVersion` | string | ✅ | 最新版本号 |
| `downloadUrl` | string | ✅ | 更新文件下载地址 |
| `updateMode` | string | ❌ | 更新模式（"full" 或 "incremental"） |
| `fileSize` | number | ❌ | 文件大小（字节） |
| `fileHash` | string | ❌ | 文件哈希值，格式为 "算法:哈希值" |
| `isForced` | boolean | ❌ | 是否强制更新 |
| `releaseNotes` | string | ❌ | 更新说明，支持多行文本 |
| `releaseDate` | string | ❌ | 发布日期 |
| `minCompatibleVersion` | string | ❌ | 最低兼容版本 |

#### 版本号格式支持

- **语义化版本号**：`1.0.0`, `1.2.3`, `2.0.0-beta.1`
- **简化格式**：`1.0`, `1`（自动补全为 `1.0.0`, `1.0.0`）
- **前缀格式**：`v1.0.0`（自动移除前缀）

#### 哈希算法支持

系统支持以下哈希算法：
- `sha256`: SHA-256（推荐）
- `sha1`: SHA-1  
- `md5`: MD5

哈希值格式：
- **带前缀**：`sha256:e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`
- **纯哈希**：`e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`（默认使用SHA-256）

#### 更新包制作

1. **全量更新**：包含插件的所有文件
2. **增量更新**：仅包含修改的文件
3. 使用 ZIP 格式压缩
4. 保持目录结构与插件目录一致
5. 生成文件哈希值确保完整性

#### 哈希值生成示例

```bash
# 生成 SHA-256 哈希
sha256sum plugin.zip

# 生成 MD5 哈希  
md5sum plugin.zip

# Windows PowerShell
Get-FileHash plugin.zip -Algorithm SHA256
```

#### 更新服务器示例实现

```csharp
// ASP.NET Core Web API 示例
[ApiController]
[Route("api/plugins")]
public class UpdateController : ControllerBase
{
    [HttpGet("update-check")]
    public IActionResult CheckUpdate([FromQuery] string pluginName, [FromQuery] string currentVersion)
    {
        // 检查是否有新版本
        var latestVersion = GetLatestVersion(pluginName);
        
        if (IsNewerVersion(latestVersion, currentVersion))
        {
            return Ok(new
            {
                latestVersion = latestVersion,
                downloadUrl = $"https://example.com/plugins/{pluginName}-v{latestVersion}.zip",
                updateMode = "full",
                fileSize = GetFileSize(pluginName, latestVersion),
                fileHash = $"sha256:{GetFileHash(pluginName, latestVersion)}",
                isForced = false,
                releaseNotes = GetReleaseNotes(pluginName, latestVersion),
                releaseDate = GetReleaseDate(pluginName, latestVersion),
                minCompatibleVersion = "1.0.0"
            });
        }
        
        return Ok(new { latestVersion = currentVersion });
    }
}
```

## 构建和部署

### 开发调试
1. 按 F5 启动调试，插件会自动加载到 Excel 中
2. 在 Excel 的功能区中找到您的插件按钮

### 发布版本
1. 使用 Release 配置构建项目
2. 输出文件位于 `bin/Release/net481/`
3. 将编译输出复制到 DTI_Tool.AddIn 的插件目录

## 开发环境要求

- Visual Studio 2019/2022
- .NET Framework 4.8.1 SDK
- Microsoft Office Excel 2016 或更高版本

## 注意事项

1. **COM 对象释放**：及时释放 Excel COM 对象，避免内存泄漏
2. **线程安全**：Excel 操作必须在主线程进行
3. **异常处理**：包装所有 Excel 操作，提供用户友好的错误信息
4. **性能优化**：批量操作时考虑关闭 Excel 的屏幕更新

## 技术支持

如有开发问题，请：
1. 查看 DTI_Tool.AddIn 框架文档
2. 参考示例代码实现
3. 访问 [博客官网](https://www.90le.cn)
4. 发送邮件至 767759678@qq.com
5. 添加作者微信交流 binStudy

## 许可证

本模板采用 MIT 许可证，您可以自由使用、修改和分发。 
