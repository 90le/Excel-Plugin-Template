using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using BasePlugin.Core;
using BasePlugin.Models;
using DTI_Tool.AddIn.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasePlugin.Features
{
    /// <summary>
    /// 实用工具功能类 - 提供各种实用工具功能
    /// </summary>
    public class UtilityFeatures : IFeatureProvider
    {
        #region 私有字段

        private readonly PluginLogger _logger;

        #endregion

        #region 构造函数

        public UtilityFeatures(PluginLogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            
            _logger.Debug("UtilityFeatures 已初始化");
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
        /// 获取实用工具功能列表
        /// </summary>
        public List<PluginFeature> GetFeatures()
        {
            return new List<PluginFeature>
            {
                new PluginFeature
                {
                    Id = "export_csv",
                    Name = "导出为CSV",
                    Description = "将选中区域导出为CSV文件",
                    Category = "实用工具",
                    Tags = new List<string> { "导出", "CSV", "文件" },
                    ImageMso = "ExportTextFile",
                    Action = ExportToCSV
                },
                new PluginFeature
                {
                    Id = "generate_random_data",
                    Name = "生成随机数据",
                    Description = "在选中区域生成随机测试数据",
                    Category = "实用工具",
                    Tags = new List<string> { "随机", "测试", "数据" },
                    ImageMso = "RandomNumber",
                    Action = GenerateRandomData
                },
                new PluginFeature
                {
                    Id = "find_replace_advanced",
                    Name = "高级查找替换",
                    Description = "批量查找替换多个文本",
                    Category = "实用工具",
                    Tags = new List<string> { "查找", "替换", "批量" },
                    ImageMso = "FindDialog",
                    Action = FindReplaceAdvanced
                },
                new PluginFeature
                {
                    Id = "workbook_info",
                    Name = "工作簿信息",
                    Description = "显示当前工作簿的详细信息",
                    Category = "实用工具",
                    Tags = new List<string> { "信息", "统计", "工作簿" },
                    ImageMso = "Info",
                    Action = ShowWorkbookInfo
                },
                new PluginFeature
                {
                    Id = "clean_data",
                    Name = "数据清理",
                    Description = "清理选中区域的数据（去空格、统一格式等）",
                    Category = "实用工具",
                    Tags = new List<string> { "清理", "格式", "空格" },
                    ImageMso = "DataCleansingWizard",
                    Action = CleanData
                }
            };
        }

        public void Dispose()
        {
            _logger.Debug("UtilityFeatures 已释放");
        }

        #endregion

        #region 功能实现

        /// <summary>
        /// 导出为CSV
        /// </summary>
        private void ExportToCSV()
        {
            using (_logger.MeasurePerformance("导出为CSV"))
            {
                try
                {
                    _logger.Info("开始导出CSV");
                    
                    var selection = ExcelApp?.Selection as Excel.Range;
                    if (selection == null || selection.Cells.Count == 0)
                    {
                        MessageHelper.ShowWarning("请选择要导出的数据区域");
                        return;
                    }
                    
                    // 选择保存位置
                    using (var saveDialog = new System.Windows.Forms.SaveFileDialog())
                    {
                        saveDialog.Filter = "CSV文件 (*.csv)|*.csv|所有文件 (*.*)|*.*";
                        saveDialog.DefaultExt = "csv";
                        saveDialog.FileName = $"Export_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                        
                        if (saveDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                        {
                            _logger.Info("用户取消了导出");
                            return;
                        }
                        
                        var filePath = saveDialog.FileName;
                        _logger.Debug("导出路径: {0}", filePath);
                        
                        // 构建CSV内容
                        var csvContent = new StringBuilder();
                        
                        foreach (Excel.Range row in selection.Rows)
                        {
                            var rowValues = new List<string>();
                            foreach (Excel.Range cell in row.Cells)
                            {
                                var value = cell.Value?.ToString() ?? "";
                                // 如果值包含逗号或引号，需要用引号包围
                                if (value.Contains(",") || value.Contains("\""))
                                {
                                    value = $"\"{value.Replace("\"", "\"\"")}\"";
                                }
                                rowValues.Add(value);
                            }
                            csvContent.AppendLine(string.Join(",", rowValues));
                        }
                        
                        // 写入文件
                        File.WriteAllText(filePath, csvContent.ToString(), Encoding.UTF8);
                        
                        MessageHelper.ShowInfo($"数据已成功导出到:\n{filePath}\n\n导出行数: {selection.Rows.Count}", "导出成功");
                        _logger.Info("CSV导出成功，行数: {0}", selection.Rows.Count);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "导出CSV失败");
                    MessageHelper.ShowError($"导出CSV失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 生成随机数据
        /// </summary>
        private void GenerateRandomData()
        {
            using (_logger.MeasurePerformance("生成随机数据"))
            {
                try
                {
                    _logger.Info("开始生成随机数据");
                    
                    var selection = ExcelApp?.Selection as Excel.Range;
                    if (selection == null || selection.Cells.Count == 0)
                    {
                        MessageHelper.ShowWarning("请选择要填充随机数据的区域");
                        return;
                    }
                    
                    var dataTypes = new[]
                    {
                        "1. 随机数字 (0-100)",
                        "2. 随机日期 (最近一年)",
                        "3. 随机姓名",
                        "4. 随机邮箱",
                        "5. 随机手机号",
                        "6. 随机产品名称"
                    };
                    
                    var choice = Microsoft.VisualBasic.Interaction.InputBox(
                        $"选择要生成的数据类型:\n\n{string.Join("\n", dataTypes)}",
                        "生成随机数据",
                        "1");
                    
                    if (string.IsNullOrEmpty(choice))
                    {
                        _logger.Info("用户取消了生成随机数据");
                        return;
                    }
                    
                    _logger.Debug("生成随机数据类型: {0}", choice);
                    var random = new Random();
                    
                    foreach (Excel.Range cell in selection.Cells)
                    {
                        switch (choice)
                        {
                            case "1": // 随机数字
                                cell.Value = random.Next(0, 101);
                                break;
                            case "2": // 随机日期
                                var daysBack = random.Next(0, 365);
                                cell.Value = DateTime.Now.AddDays(-daysBack);
                                cell.NumberFormat = "yyyy-mm-dd";
                                break;
                            case "3": // 随机姓名
                                cell.Value = GenerateRandomName(random);
                                break;
                            case "4": // 随机邮箱
                                cell.Value = GenerateRandomEmail(random);
                                break;
                            case "5": // 随机手机号
                                cell.Value = GenerateRandomPhone(random);
                                break;
                            case "6": // 随机产品名称
                                cell.Value = GenerateRandomProduct(random);
                                break;
                            default:
                                MessageHelper.ShowWarning("无效的选择");
                                return;
                        }
                    }
                    
                    MessageHelper.ShowInfo($"已在 {selection.Cells.Count} 个单元格中生成随机数据", "生成成功");
                    _logger.Info("随机数据生成完成，单元格数: {0}", selection.Cells.Count);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "生成随机数据失败");
                    MessageHelper.ShowError($"生成随机数据失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 高级查找替换
        /// </summary>
        private void FindReplaceAdvanced()
        {
            using (_logger.MeasurePerformance("高级查找替换"))
            {
                try
                {
                    _logger.Info("开始高级查找替换");
                    
                    var worksheet = ExcelApp?.ActiveSheet as Excel.Worksheet;
                    if (worksheet == null)
                    {
                        MessageHelper.ShowWarning("请先选择一个工作表");
                        return;
                    }
                    
                    // 获取查找替换对
                    var pairs = new List<(string find, string replace)>();
                    
                    while (true)
                    {
                        var findText = Microsoft.VisualBasic.Interaction.InputBox(
                            $"输入要查找的文本 (第{pairs.Count + 1}个，留空结束):",
                            "查找文本",
                            "");
                        
                        if (string.IsNullOrEmpty(findText))
                            break;
                        
                        var replaceText = Microsoft.VisualBasic.Interaction.InputBox(
                            $"将 '{findText}' 替换为:",
                            "替换文本",
                            "");
                        
                        pairs.Add((findText, replaceText));
                    }
                    
                    if (pairs.Count == 0)
                    {
                        _logger.Info("没有输入查找替换对");
                        return;
                    }
                    
                    // 询问替换范围
                    var scope = MessageHelper.ShowYesNoCancel(
                        "选择替换范围:\n\n" +
                        "是 - 整个工作表\n" +
                        "否 - 仅选中区域\n" +
                        "取消 - 取消操作",
                        "替换范围");
                    
                    if (scope == System.Windows.Forms.DialogResult.Cancel)
                    {
                        _logger.Info("用户取消了查找替换");
                        return;
                    }
                    
                    Excel.Range searchRange = scope == System.Windows.Forms.DialogResult.Yes 
                        ? worksheet.UsedRange 
                        : ExcelApp.Selection as Excel.Range;
                    
                    if (searchRange == null)
                    {
                        MessageHelper.ShowWarning("无效的搜索范围");
                        return;
                    }
                    
                    _logger.Debug("开始执行 {0} 个查找替换对", pairs.Count);
                    
                    // 执行批量替换
                    var totalReplacements = 0;
                    var results = new StringBuilder();
                    
                    foreach (var (find, replace) in pairs)
                    {
                        var count = 0;
                        Excel.Range foundCell = searchRange.Find(find);
                        var firstAddress = foundCell?.Address;
                        
                        while (foundCell != null)
                        {
                            foundCell.Value = foundCell.Value.ToString().Replace(find, replace);
                            count++;
                            
                            foundCell = searchRange.FindNext(foundCell);
                            if (foundCell?.Address == firstAddress)
                                break;
                        }
                        
                        totalReplacements += count;
                        results.AppendLine($"'{find}' → '{replace}': {count} 处");
                        _logger.Debug("替换 '{0}' 为 '{1}'，共 {2} 处", find, replace, count);
                    }
                    
                    MessageHelper.ShowInfo(
                        $"批量查找替换完成！\n\n" +
                        $"总计替换: {totalReplacements} 处\n\n" +
                        $"详细结果:\n{results}",
                        "替换完成");
                    
                    _logger.Info("高级查找替换完成，总计替换 {0} 处", totalReplacements);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "高级查找替换失败");
                    MessageHelper.ShowError($"高级查找替换失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 显示工作簿信息
        /// </summary>
        private void ShowWorkbookInfo()
        {
            using (_logger.MeasurePerformance("显示工作簿信息"))
            {
                try
                {
                    _logger.Info("开始收集工作簿信息");
                    
                    var workbook = ExcelApp?.ActiveWorkbook;
                    if (workbook == null)
                    {
                        MessageHelper.ShowWarning("请先打开一个工作簿");
                        return;
                    }
                    
                    // 收集基本信息
                    var info = new StringBuilder();
                    info.AppendLine($"工作簿名称: {workbook.Name}");
                    info.AppendLine($"完整路径: {workbook.FullName}");
                    info.AppendLine($"文件大小: {GetFileSize(workbook.FullName)}");
                    info.AppendLine($"创建时间: {GetFileCreationTime(workbook.FullName)}");
                    info.AppendLine($"修改时间: {GetFileModifiedTime(workbook.FullName)}");
                    info.AppendLine($"只读状态: {(workbook.ReadOnly ? "是" : "否")}");
                    info.AppendLine();
                    
                    // 统计工作表信息
                    info.AppendLine($"工作表数量: {workbook.Worksheets.Count}");
                    var totalUsedCells = 0;
                    
                    foreach (Excel.Worksheet sheet in workbook.Worksheets)
                    {
                        var usedRange = sheet.UsedRange;
                        if (usedRange != null)
                        {
                            totalUsedCells += usedRange.Cells.Count;
                        }
                    }
                    
                    info.AppendLine($"已使用单元格总数: {totalUsedCells:N0}");
                    info.AppendLine();
                    
                    // 工作表详情
                    info.AppendLine("工作表列表:");
                    foreach (Excel.Worksheet sheet in workbook.Worksheets)
                    {
                        var usedRange = sheet.UsedRange;
                        if (usedRange != null)
                        {
                            info.AppendLine($"  • {sheet.Name} - 使用区域: {usedRange.Address}, 单元格数: {usedRange.Cells.Count:N0}");
                        }
                        else
                        {
                            info.AppendLine($"  • {sheet.Name} - (空)");
                        }
                    }
                    
                    // 显示信息
                    using (var form = new System.Windows.Forms.Form())
                    {
                        form.Text = "工作簿信息";
                        form.Size = new System.Drawing.Size(600, 500);
                        form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                        
                        var textBox = new System.Windows.Forms.TextBox
                        {
                            Multiline = true,
                            ScrollBars = System.Windows.Forms.ScrollBars.Vertical,
                            ReadOnly = true,
                            Dock = System.Windows.Forms.DockStyle.Fill,
                            Font = new System.Drawing.Font("Consolas", 10),
                            Text = info.ToString()
                        };
                        
                        form.Controls.Add(textBox);
                        form.ShowDialog();
                    }
                    
                    _logger.Info("工作簿信息显示完成");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "显示工作簿信息失败");
                    MessageHelper.ShowError($"显示工作簿信息失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 数据清理
        /// </summary>
        private void CleanData()
        {
            using (_logger.MeasurePerformance("数据清理"))
            {
                try
                {
                    _logger.Info("开始数据清理");
                    
                    var selection = ExcelApp?.Selection as Excel.Range;
                    if (selection == null || selection.Cells.Count == 0)
                    {
                        MessageHelper.ShowWarning("请选择要清理的数据区域");
                        return;
                    }
                    
                    var options = new[]
                    {
                        "1. 去除首尾空格",
                        "2. 去除所有空格",
                        "3. 转换为大写",
                        "4. 转换为小写",
                        "5. 首字母大写",
                        "6. 删除空单元格",
                        "7. 统一日期格式",
                        "8. 删除非数字字符"
                    };
                    
                    var choices = Microsoft.VisualBasic.Interaction.InputBox(
                        $"选择要执行的清理操作（可多选，用逗号分隔）:\n\n{string.Join("\n", options)}",
                        "数据清理",
                        "1,6");
                    
                    if (string.IsNullOrEmpty(choices))
                    {
                        _logger.Info("用户取消了数据清理");
                        return;
                    }
                    
                    var selectedOptions = choices.Split(',');
                    _logger.Debug("选择的清理选项: {0}", choices);
                    
                    var cleanedCount = 0;
                    
                    foreach (Excel.Range cell in selection.Cells)
                    {
                        if (cell.Value == null)
                            continue;
                        
                        var originalValue = cell.Value.ToString();
                        var newValue = originalValue;
                        var changed = false;
                        
                        foreach (var option in selectedOptions)
                        {
                            switch (option.Trim())
                            {
                                case "1": // 去除首尾空格
                                    newValue = newValue.Trim();
                                    break;
                                case "2": // 去除所有空格
                                    newValue = newValue.Replace(" ", "");
                                    break;
                                case "3": // 转换为大写
                                    newValue = newValue.ToUpper();
                                    break;
                                case "4": // 转换为小写
                                    newValue = newValue.ToLower();
                                    break;
                                case "5": // 首字母大写
                                    newValue = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(newValue.ToLower());
                                    break;
                                case "6": // 删除空单元格
                                    if (string.IsNullOrWhiteSpace(newValue))
                                    {
                                        cell.Clear();
                                        changed = true;
                                        continue;
                                    }
                                    break;
                                case "7": // 统一日期格式
                                    if (DateTime.TryParse(newValue, out DateTime date))
                                    {
                                        cell.Value = date;
                                        cell.NumberFormat = "yyyy-mm-dd";
                                        changed = true;
                                        continue;
                                    }
                                    break;
                                case "8": // 删除非数字字符
                                    newValue = System.Text.RegularExpressions.Regex.Replace(newValue, @"[^\d.-]", "");
                                    break;
                            }
                        }
                        
                        if (newValue != originalValue || changed)
                        {
                            if (!changed)
                                cell.Value = newValue;
                            cleanedCount++;
                        }
                    }
                    
                    MessageHelper.ShowInfo($"数据清理完成！\n\n处理单元格数: {selection.Cells.Count}\n清理单元格数: {cleanedCount}", "清理完成");
                    _logger.Info("数据清理完成，清理了 {0} 个单元格", cleanedCount);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "数据清理失败");
                    MessageHelper.ShowError($"数据清理失败: {ex.Message}");
                }
            }
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 获取文件大小
        /// </summary>
        private string GetFileSize(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    var fileInfo = new FileInfo(filePath);
                    var size = fileInfo.Length;
                    
                    if (size < 1024)
                        return $"{size} B";
                    else if (size < 1024 * 1024)
                        return $"{size / 1024.0:F2} KB";
                    else if (size < 1024 * 1024 * 1024)
                        return $"{size / (1024.0 * 1024.0):F2} MB";
                    else
                        return $"{size / (1024.0 * 1024.0 * 1024.0):F2} GB";
                }
            }
            catch { }
            return "未知";
        }

        /// <summary>
        /// 获取文件创建时间
        /// </summary>
        private string GetFileCreationTime(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    return File.GetCreationTime(filePath).ToString("yyyy-MM-dd HH:mm:ss");
                }
            }
            catch { }
            return "未知";
        }

        /// <summary>
        /// 获取文件修改时间
        /// </summary>
        private string GetFileModifiedTime(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    return File.GetLastWriteTime(filePath).ToString("yyyy-MM-dd HH:mm:ss");
                }
            }
            catch { }
            return "未知";
        }

        /// <summary>
        /// 生成随机姓名
        /// </summary>
        private string GenerateRandomName(Random random)
        {
            var surnames = new[] { "张", "王", "李", "赵", "刘", "陈", "杨", "黄", "周", "吴" };
            var names = new[] { "伟", "芳", "娜", "敏", "静", "强", "磊", "洋", "艳", "杰" };
            return surnames[random.Next(surnames.Length)] + names[random.Next(names.Length)];
        }

        /// <summary>
        /// 生成随机邮箱
        /// </summary>
        private string GenerateRandomEmail(Random random)
        {
            var names = new[] { "john", "mary", "david", "sarah", "michael", "emma", "robert", "lisa", "james", "anna" };
            var domains = new[] { "gmail.com", "outlook.com", "163.com", "qq.com", "hotmail.com" };
            return $"{names[random.Next(names.Length)]}{random.Next(100, 999)}@{domains[random.Next(domains.Length)]}";
        }

        /// <summary>
        /// 生成随机手机号
        /// </summary>
        private string GenerateRandomPhone(Random random)
        {
            var prefixes = new[] { "130", "131", "132", "155", "156", "185", "186", "150", "151" };
            return $"{prefixes[random.Next(prefixes.Length)]}{random.Next(10000000, 99999999)}";
        }

        /// <summary>
        /// 生成随机产品名称
        /// </summary>
        private string GenerateRandomProduct(Random random)
        {
            var adjectives = new[] { "高级", "专业", "智能", "便携", "多功能", "豪华", "经典", "时尚", "创新", "精品" };
            var products = new[] { "笔记本", "手机", "平板", "耳机", "音箱", "键盘", "鼠标", "显示器", "相机", "手表" };
            return $"{adjectives[random.Next(adjectives.Length)]}{products[random.Next(products.Length)]}";
        }

        #endregion
    }
} 