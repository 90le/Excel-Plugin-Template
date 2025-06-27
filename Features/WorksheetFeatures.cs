using System;
using System.Collections.Generic;
using System.Linq;
using BasePlugin.Core;
using BasePlugin.Models;
using DTI_Tool.AddIn.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasePlugin.Features
{
    /// <summary>
    /// 工作表管理功能类 - 提供工作表操作相关的示例功能
    /// </summary>
    public class WorksheetFeatures : IFeatureProvider
    {
        #region 私有字段

        private readonly PluginLogger _logger;

        #endregion

        #region 构造函数

        public WorksheetFeatures(PluginLogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            
            _logger.Debug("WorksheetFeatures 已初始化");
        }

        #endregion

        #region 私有属性

        /// <summary>
        /// 获取Excel应用程序对象
        /// </summary>
        private Excel.Application ExcelApp => HostApplication.Instance.ExcelApplication;

        #endregion

        #region IFeatureProvider 实现

        /// <summary>
        /// 获取工作表管理功能列表
        /// </summary>
        public List<PluginFeature> GetFeatures()
        {
            return new List<PluginFeature>
            {
                new PluginFeature
                {
                    Id = "create_worksheet",
                    Name = "创建工作表",
                    Description = "创建新的工作表并设置基本属性",
                    Category = "工作表管理",
                    Tags = new List<string> { "工作表", "创建", "新建" },
                    ImageMso = "ReviewNewComment",
                    Action = CreateWorksheet
                },
                new PluginFeature
                {
                    Id = "rename_worksheet",
                    Name = "重命名工作表",
                    Description = "批量重命名工作表",
                    Category = "工作表管理",
                    Tags = new List<string> { "工作表", "重命名", "批量" },
                    ImageMso = "RenameLayoutCustom",
                    Action = RenameWorksheet
                },
                new PluginFeature
                {
                    Id = "copy_worksheet",
                    Name = "复制工作表",
                    Description = "复制当前工作表",
                    Category = "工作表管理",
                    Tags = new List<string> { "工作表", "复制", "克隆" },
                    ImageMso = "Copy",
                    Action = CopyWorksheet
                },
                new PluginFeature
                {
                    Id = "worksheet_navigation",
                    Name = "工作表导航",
                    Description = "快速导航到指定工作表",
                    Category = "工作表管理",
                    Tags = new List<string> { "工作表", "导航", "跳转" },
                    ImageMso = "WorksheetInsert",
                    Action = WorksheetNavigation
                },
                new PluginFeature
                {
                    Id = "protect_worksheet",
                    Name = "保护工作表",
                    Description = "设置工作表保护",
                    Category = "工作表管理",
                    Tags = new List<string> { "工作表", "保护", "锁定" },
                    ImageMso = "ReviewProtectSheet",
                    Action = ProtectWorksheet
                }
            };
        }

        public void Dispose()
        {
            _logger.Debug("WorksheetFeatures 已释放");
        }

        #endregion

        #region 功能实现

        /// <summary>
        /// 创建工作表
        /// </summary>
        private void CreateWorksheet()
        {
            using (_logger.MeasurePerformance("创建工作表"))
            {
                try
                {
                    _logger.Info("开始创建工作表");
                    
                    var workbook = ExcelApp?.ActiveWorkbook;
                    if (workbook == null)
                    {
                        MessageHelper.ShowWarning("请先打开一个工作簿");
                        return;
                    }
                    
                    // 获取工作表名称
                    var sheetName = Microsoft.VisualBasic.Interaction.InputBox(
                        "请输入新工作表名称:",
                        "创建工作表",
                        $"Sheet{workbook.Worksheets.Count + 1}");
                    
                    if (string.IsNullOrEmpty(sheetName))
                    {
                        _logger.Info("用户取消了创建工作表");
                        return;
                    }
                    
                    // 检查名称是否已存在
                    if (WorksheetExists(workbook, sheetName))
                    {
                        MessageHelper.ShowWarning($"工作表 '{sheetName}' 已存在");
                        return;
                    }
                    
                    // 创建新工作表
                    var newSheet = workbook.Worksheets.Add() as Excel.Worksheet;
                    newSheet.Name = sheetName;
                    
                    // 设置一些基本内容
                    newSheet.Range["A1"].Value = $"工作表 {sheetName}";
                    newSheet.Range["A1"].Font.Bold = true;
                    newSheet.Range["A1"].Font.Size = 14;
                    
                    newSheet.Range["A3"].Value = "创建时间:";
                    newSheet.Range["B3"].Value = DateTime.Now;
                    newSheet.Range["B3"].NumberFormat = "yyyy-mm-dd hh:mm:ss";
                    
                    // 设置列宽
                    var columnA = newSheet.Columns["A:A"] as Excel.Range;
                    var columnB = newSheet.Columns["B:B"] as Excel.Range;
                    if (columnA != null) columnA.ColumnWidth = 15;
                    if (columnB != null) columnB.ColumnWidth = 20;
                    
                    // 激活新工作表
                    newSheet.Activate();
                    
                    MessageHelper.ShowInfo($"工作表 '{sheetName}' 创建成功", "创建成功");
                    _logger.Info("工作表创建成功: {0}", sheetName);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "创建工作表失败");
                    MessageHelper.ShowError($"创建工作表失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 重命名工作表
        /// </summary>
        private void RenameWorksheet()
        {
            using (_logger.MeasurePerformance("重命名工作表"))
            {
                try
                {
                    _logger.Info("开始重命名工作表");
                    
                    var workbook = ExcelApp?.ActiveWorkbook;
                    if (workbook == null)
                    {
                        MessageHelper.ShowWarning("请先打开一个工作簿");
                        return;
                    }
                    
                    var worksheet = ExcelApp.ActiveSheet as Excel.Worksheet;
                    if (worksheet == null)
                    {
                        MessageHelper.ShowWarning("请先选择一个工作表");
                        return;
                    }
                    
                    // 显示重命名选项
                    var choice = MessageHelper.ShowYesNoCancel(
                        "重命名选项:\n\n" +
                        "是 - 重命名当前工作表\n" +
                        "否 - 批量重命名所有工作表\n" +
                        "取消 - 取消操作",
                        "重命名工作表");
                    
                    if (choice == System.Windows.Forms.DialogResult.Cancel)
                    {
                        _logger.Info("用户取消了重命名操作");
                        return;
                    }
                    
                    if (choice == System.Windows.Forms.DialogResult.Yes)
                    {
                        // 重命名当前工作表
                        RenameSingleWorksheet(worksheet);
                    }
                    else
                    {
                        // 批量重命名
                        BatchRenameWorksheets(workbook);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "重命名工作表失败");
                    MessageHelper.ShowError($"重命名工作表失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 复制工作表
        /// </summary>
        private void CopyWorksheet()
        {
            using (_logger.MeasurePerformance("复制工作表"))
            {
                try
                {
                    _logger.Info("开始复制工作表");
                    
                    var worksheet = ExcelApp?.ActiveSheet as Excel.Worksheet;
                    if (worksheet == null)
                    {
                        MessageHelper.ShowWarning("请先选择要复制的工作表");
                        return;
                    }
                    
                    var originalName = worksheet.Name;
                    _logger.Debug("复制工作表: {0}", originalName);
                    
                    // 获取复制选项
                    var copyCount = Microsoft.VisualBasic.Interaction.InputBox(
                        "请输入要创建的副本数量:",
                        "复制工作表",
                        "1");
                    
                    if (string.IsNullOrEmpty(copyCount) || !int.TryParse(copyCount, out int count) || count < 1)
                    {
                        _logger.Info("用户取消了复制或输入无效");
                        return;
                    }
                    
                    // 执行复制
                    for (int i = 1; i <= count; i++)
                    {
                        var workbook = worksheet.Parent as Excel.Workbook;
                        if (workbook != null)
                        {
                            worksheet.Copy(After: workbook.Worksheets[workbook.Worksheets.Count]);
                            
                            // 重命名复制的工作表
                            var newSheet = ExcelApp.ActiveSheet as Excel.Worksheet;
                            if (newSheet != null)
                            {
                                var newName = GetUniqueWorksheetName(workbook, $"{originalName}_副本{i}");
                                newSheet.Name = newName;
                                _logger.Debug("创建副本: {0}", newName);
                            }
                        }
                    }
                    
                    MessageHelper.ShowInfo($"已创建 {count} 个工作表副本", "复制成功");
                    _logger.Info("成功创建 {0} 个工作表副本", count);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "复制工作表失败");
                    MessageHelper.ShowError($"复制工作表失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 工作表导航
        /// </summary>
        private void WorksheetNavigation()
        {
            using (_logger.MeasurePerformance("工作表导航"))
            {
                try
                {
                    _logger.Info("开始工作表导航");
                    
                    var workbook = ExcelApp?.ActiveWorkbook;
                    if (workbook == null)
                    {
                        MessageHelper.ShowWarning("请先打开一个工作簿");
                        return;
                    }
                    
                    // 获取所有工作表列表
                    var sheets = new List<string>();
                    foreach (Excel.Worksheet sheet in workbook.Worksheets)
                    {
                        sheets.Add($"{sheets.Count + 1}. {sheet.Name}");
                    }
                    
                    if (sheets.Count == 0)
                    {
                        MessageHelper.ShowWarning("工作簿中没有工作表");
                        return;
                    }
                    
                    // 显示工作表列表
                    var sheetList = string.Join("\n", sheets);
                    var choice = Microsoft.VisualBasic.Interaction.InputBox(
                        $"请选择要导航到的工作表（输入序号）:\n\n{sheetList}",
                        "工作表导航",
                        "1");
                    
                    if (string.IsNullOrEmpty(choice) || !int.TryParse(choice, out int index) || index < 1 || index > sheets.Count)
                    {
                        _logger.Info("用户取消了导航或选择无效");
                        return;
                    }
                    
                    // 导航到选定的工作表
                    var targetSheet = workbook.Worksheets[index] as Excel.Worksheet;
                    targetSheet.Activate();
                    
                    MessageHelper.ShowInfo($"已导航到工作表: {targetSheet.Name}", "导航成功");
                    _logger.Info("已导航到工作表: {0}", targetSheet.Name);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "工作表导航失败");
                    MessageHelper.ShowError($"工作表导航失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 保护工作表
        /// </summary>
        private void ProtectWorksheet()
        {
            using (_logger.MeasurePerformance("保护工作表"))
            {
                try
                {
                    _logger.Info("开始设置工作表保护");
                    
                    var worksheet = ExcelApp?.ActiveSheet as Excel.Worksheet;
                    if (worksheet == null)
                    {
                        MessageHelper.ShowWarning("请先选择要保护的工作表");
                        return;
                    }
                    
                    // 检查当前保护状态
                    if (worksheet.ProtectContents)
                    {
                        // 已保护，询问是否取消保护
                        if (MessageHelper.ShowConfirm($"工作表 '{worksheet.Name}' 已被保护。\n是否取消保护？", "工作表已保护"))
                        {
                            var password = Microsoft.VisualBasic.Interaction.InputBox(
                                "请输入密码（如果有）:",
                                "取消保护",
                                "");
                            
                            try
                            {
                                worksheet.Unprotect(password);
                                MessageHelper.ShowInfo("工作表保护已取消", "操作成功");
                                _logger.Info("已取消工作表保护: {0}", worksheet.Name);
                            }
                            catch
                            {
                                MessageHelper.ShowError("密码错误或取消保护失败");
                            }
                        }
                    }
                    else
                    {
                        // 未保护，设置保护
                        var password = Microsoft.VisualBasic.Interaction.InputBox(
                            "请输入保护密码（可留空）:",
                            "设置保护",
                            "");
                        
                        // 设置保护选项
                        worksheet.Protect(
                            Password: password,
                            DrawingObjects: true,
                            Contents: true,
                            Scenarios: true,
                            AllowFormattingCells: true,
                            AllowFormattingColumns: false,
                            AllowFormattingRows: false,
                            AllowInsertingColumns: false,
                            AllowInsertingRows: false,
                            AllowDeletingColumns: false,
                            AllowDeletingRows: false,
                            AllowSorting: true,
                            AllowFiltering: true
                        );
                        
                        var protectionInfo = "工作表已保护！\n\n" +
                                           "允许的操作:\n" +
                                           "• 格式化单元格\n" +
                                           "• 排序\n" +
                                           "• 筛选\n\n" +
                                           "禁止的操作:\n" +
                                           "• 修改内容\n" +
                                           "• 插入/删除行列";
                        
                        MessageHelper.ShowInfo(protectionInfo, "保护成功");
                        _logger.Info("已设置工作表保护: {0}", worksheet.Name);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "设置工作表保护失败");
                    MessageHelper.ShowError($"设置工作表保护失败: {ex.Message}");
                }
            }
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 检查工作表是否存在
        /// </summary>
        private bool WorksheetExists(Excel.Workbook workbook, string name)
        {
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 获取唯一的工作表名称
        /// </summary>
        private string GetUniqueWorksheetName(Excel.Workbook workbook, string baseName)
        {
            var name = baseName;
            var counter = 1;
            
            while (WorksheetExists(workbook, name))
            {
                name = $"{baseName}_{counter}";
                counter++;
            }
            
            return name;
        }

        /// <summary>
        /// 重命名单个工作表
        /// </summary>
        private void RenameSingleWorksheet(Excel.Worksheet worksheet)
        {
            var newName = Microsoft.VisualBasic.Interaction.InputBox(
                $"请输入新的工作表名称:\n当前名称: {worksheet.Name}",
                "重命名工作表",
                worksheet.Name);
            
            if (string.IsNullOrEmpty(newName) || newName == worksheet.Name)
            {
                _logger.Info("用户取消了重命名或名称未更改");
                return;
            }
            
            // 检查名称是否已存在
            var workbook = worksheet.Parent as Excel.Workbook;
            if (workbook != null && WorksheetExists(workbook, newName))
            {
                MessageHelper.ShowWarning($"工作表名称 '{newName}' 已存在");
                return;
            }
            
            var oldName = worksheet.Name;
            worksheet.Name = newName;
            MessageHelper.ShowInfo($"工作表已重命名:\n{oldName} → {newName}", "重命名成功");
            _logger.Info("工作表重命名: {0} → {1}", oldName, newName);
        }

        /// <summary>
        /// 批量重命名工作表
        /// </summary>
        private void BatchRenameWorksheets(Excel.Workbook workbook)
        {
            var prefix = Microsoft.VisualBasic.Interaction.InputBox(
                "请输入工作表名称前缀:",
                "批量重命名",
                "Sheet");
            
            if (string.IsNullOrEmpty(prefix))
            {
                _logger.Info("用户取消了批量重命名");
                return;
            }
            
            if (!MessageHelper.ShowConfirm($"将把所有工作表重命名为:\n{prefix}1, {prefix}2, {prefix}3...\n\n是否继续？", "确认批量重命名"))
            {
                return;
            }
            
            var index = 1;
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    var oldName = sheet.Name;
                    sheet.Name = $"{prefix}{index}";
                    _logger.Debug("重命名: {0} → {1}", oldName, sheet.Name);
                    index++;
                }
                catch (Exception ex)
                {
                    _logger.Warning("无法重命名工作表: {0}", ex.Message);
                }
            }
            
            MessageHelper.ShowInfo($"已重命名 {index - 1} 个工作表", "批量重命名完成");
            _logger.Info("批量重命名完成，共重命名 {0} 个工作表", index - 1);
        }

        #endregion
    }
} 