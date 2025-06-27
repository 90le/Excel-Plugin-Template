using System;
using System.Collections.Generic;
using BasePlugin.Core;
using BasePlugin.Models;
using DTI_Tool.AddIn.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasePlugin.Features
{
    /// <summary>
    /// 基础功能类 - 提供基本的Excel操作示例
    /// </summary>
    public class BasicFeatures : IFeatureProvider
    {
        #region 私有字段

        private readonly PluginLogger _logger;

        #endregion

        #region 构造函数

        public BasicFeatures(PluginLogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            
            _logger.Debug("BasicFeatures 已初始化");
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
        /// 获取基础功能列表
        /// </summary>
        public List<PluginFeature> GetFeatures()
        {
            return new List<PluginFeature>
            {
                new PluginFeature
                {
                    Id = "hello_world",
                    Name = "Hello World",
                    Description = "显示一个简单的欢迎消息",
                    Category = "基础功能",
                    Tags = new List<string> { "示例", "基础", "入门" },
                    ImageMso = "HappyFace",
                    Action = HelloWorld
                },
                new PluginFeature
                {
                    Id = "get_selection_info",
                    Name = "获取选择信息",
                    Description = "显示当前选中区域的详细信息",
                    Category = "基础功能",
                    Tags = new List<string> { "选择", "信息", "区域" },
                    ImageMso = "TableExcelSelect",
                    Action = GetSelectionInfo
                },
                new PluginFeature
                {
                    Id = "insert_current_time",
                    Name = "插入当前时间",
                    Description = "在活动单元格插入当前日期和时间",
                    Category = "基础功能",
                    Tags = new List<string> { "时间", "日期", "插入" },
                    ImageMso = "DateAndTimePicker",
                    Action = InsertCurrentTime
                },
                new PluginFeature
                {
                    Id = "cell_value_operations",
                    Name = "单元格值操作",
                    Description = "演示如何读取和修改单元格的值",
                    Category = "基础功能",
                    Tags = new List<string> { "单元格", "值", "读写" },
                    ImageMso = "TableCellProperties",
                    Action = CellValueOperations
                },
                new PluginFeature
                {
                    Id = "range_navigation",
                    Name = "区域导航",
                    Description = "演示如何在工作表中导航和选择区域",
                    Category = "基础功能",
                    Tags = new List<string> { "导航", "区域", "选择" },
                    ImageMso = "Navigation",
                    Action = RangeNavigation
                }
            };
        }

        public void Dispose()
        {
            _logger.Debug("BasicFeatures 已释放");
        }

        #endregion

        #region 功能实现

        /// <summary>
        /// Hello World 示例
        /// </summary>
        private void HelloWorld()
        {
            using (_logger.MeasurePerformance("Hello World"))
            {
                try
                {
                    _logger.Info("执行 Hello World 功能");
                    
                    var message = "欢迎使用 BasePlugin！\n\n" +
                                 "这是一个功能完整的Excel插件开发模板。\n" +
                                 "您可以基于此模板快速开发自己的插件。\n\n" +
                                 "主要特点：\n" +
                                 "• 清晰的项目结构\n" +
                                 "• 完整的功能示例\n" +
                                 "• 详细的日志记录\n" +
                                 "• WPF界面支持\n" +
                                 "• 任务窗格支持";
                    
                    MessageHelper.ShowInfo(message, "欢迎使用 BasePlugin");
                    _logger.Info("Hello World 功能执行完成");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Hello World 功能执行失败");
                    MessageHelper.ShowError($"Hello World 功能执行失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 获取选择信息
        /// </summary>
        private void GetSelectionInfo()
        {
            using (_logger.MeasurePerformance("获取选择信息"))
            {
                try
                {
                    _logger.Info("开始获取选择信息");
                    
                    var selection = ExcelApp?.Selection as Excel.Range;
                    if (selection == null)
                    {
                        _logger.Warning("用户未选择任何区域");
                        MessageHelper.ShowWarning("请先选择一个区域");
                        return;
                    }
                    
                    _logger.Debug("选中区域地址: {0}", selection.Address);
                    
                    var info = $"选中区域详细信息:\n\n" +
                              $"地址: {selection.Address}\n" +
                              $"工作表: {selection.Worksheet.Name}\n" +
                              $"行数: {selection.Rows.Count}\n" +
                              $"列数: {selection.Columns.Count}\n" +
                              $"单元格数: {selection.Cells.Count}\n" +
                              $"起始行: {selection.Row}\n" +
                              $"起始列: {selection.Column}\n" +
                              $"结束行: {selection.Row + selection.Rows.Count - 1}\n" +
                              $"结束列: {selection.Column + selection.Columns.Count - 1}";
                    
                    // 如果是单个单元格，显示其值
                    if (selection.Cells.Count == 1)
                    {
                        var value = selection.Value?.ToString() ?? "(空)";
                        info += $"\n\n单元格值: {value}";
                    }
                    
                    MessageHelper.ShowInfo(info, "选择信息");
                    _logger.Info("成功显示选择信息，区域: {0}", selection.Address);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "获取选择信息失败");
                    MessageHelper.ShowError($"获取选择信息失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 插入当前时间
        /// </summary>
        private void InsertCurrentTime()
        {
            using (_logger.MeasurePerformance("插入当前时间"))
            {
                try
                {
                    _logger.Info("开始插入当前时间");
                    
                    var activeCell = ExcelApp?.ActiveCell;
                    if (activeCell == null)
                    {
                        _logger.Warning("没有活动单元格");
                        MessageHelper.ShowWarning("请先选择一个单元格");
                        return;
                    }
                    
                    var currentTime = DateTime.Now;
                    var formattedTime = currentTime.ToString("yyyy-MM-dd HH:mm:ss");
                    
                    _logger.Debug("活动单元格地址: {0}, 插入时间: {1}", activeCell.Address, formattedTime);
                    
                    // 保存原值用于撤销
                    var originalValue = activeCell.Value;
                    
                    // 插入时间
                    activeCell.Value = currentTime;
                    activeCell.NumberFormat = "yyyy-mm-dd hh:mm:ss";
                    
                    MessageHelper.ShowInfo($"已在单元格 {activeCell.Address} 插入当前时间:\n{formattedTime}", 
                                         "插入时间成功");
                    
                    _logger.Info("成功在单元格 {0} 插入时间: {1}", activeCell.Address, formattedTime);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "插入时间失败");
                    MessageHelper.ShowError($"插入时间失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 单元格值操作示例
        /// </summary>
        private void CellValueOperations()
        {
            using (_logger.MeasurePerformance("单元格值操作"))
            {
                try
                {
                    _logger.Info("开始演示单元格值操作");
                    
                    var worksheet = ExcelApp?.ActiveSheet as Excel.Worksheet;
                    if (worksheet == null)
                    {
                        _logger.Warning("没有活动工作表");
                        MessageHelper.ShowWarning("请先打开一个工作表");
                        return;
                    }
                    
                    // 读取A1单元格的值
                    var cellA1 = worksheet.Range["A1"];
                    var originalValue = cellA1.Value?.ToString() ?? "(空)";
                    _logger.Debug("A1单元格原值: {0}", originalValue);
                    
                    // 写入新值
                    cellA1.Value = "Hello from BasePlugin!";
                    _logger.Debug("已向A1单元格写入新值");
                    
                    // 读取并格式化B1单元格
                    var cellB1 = worksheet.Range["B1"];
                    cellB1.Value = DateTime.Now;
                    cellB1.NumberFormat = "yyyy-mm-dd";
                    cellB1.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);
                    _logger.Debug("已向B1单元格写入日期并设置格式");
                    
                    // 在C1单元格写入公式
                    var cellC1 = worksheet.Range["C1"];
                    cellC1.Formula = "=NOW()";
                    cellC1.NumberFormat = "hh:mm:ss";
                    _logger.Debug("已向C1单元格写入公式");
                    
                    // 批量操作示例
                    var range = worksheet.Range["A3:C5"];
                    range.Value = "批量填充";
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    _logger.Debug("已对A3:C5区域进行批量操作");
                    
                    MessageHelper.ShowInfo(
                        "单元格操作完成！\n\n" +
                        $"• A1: 写入文本\n" +
                        $"• B1: 写入日期并格式化\n" +
                        $"• C1: 写入公式\n" +
                        $"• A3:C5: 批量操作\n\n" +
                        $"A1原值: {originalValue}",
                        "操作成功");
                    
                    _logger.Info("单元格值操作演示完成");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "单元格值操作失败");
                    MessageHelper.ShowError($"单元格值操作失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 区域导航示例
        /// </summary>
        private void RangeNavigation()
        {
            using (_logger.MeasurePerformance("区域导航"))
            {
                try
                {
                    _logger.Info("开始演示区域导航");
                    
                    var worksheet = ExcelApp?.ActiveSheet as Excel.Worksheet;
                    if (worksheet == null)
                    {
                        _logger.Warning("没有活动工作表");
                        MessageHelper.ShowWarning("请先打开一个工作表");
                        return;
                    }
                    
                    // 1. 选择特定区域
                    var rangeA1C3 = worksheet.Range["A1:C3"];
                    rangeA1C3.Select();
                    _logger.Debug("已选择区域 A1:C3");
                    System.Threading.Thread.Sleep(500); // 短暂停留让用户看到
                    
                    // 2. 查找最后使用的单元格
                    var usedRange = worksheet.UsedRange;
                    var lastRow = usedRange.Row + usedRange.Rows.Count - 1;
                    var lastCol = usedRange.Column + usedRange.Columns.Count - 1;
                    _logger.Debug("工作表使用范围: 行1-{0}, 列1-{1}", lastRow, lastCol);
                    
                    // 3. 导航到特定单元格
                    var targetCell = worksheet.Cells[5, 5] as Excel.Range; // E5
                    if (targetCell != null)
                    {
                        targetCell.Select();
                        targetCell.Value = "导航目标";
                        targetCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    }
                    _logger.Debug("已导航到单元格 E5");
                    System.Threading.Thread.Sleep(500);
                    
                    // 4. 选择整行
                    var row7 = worksheet.Rows[7] as Excel.Range;
                    row7?.Select();
                    _logger.Debug("已选择第7行");
                    System.Threading.Thread.Sleep(500);
                    
                    // 5. 选择整列
                    var columnD = worksheet.Columns["D"] as Excel.Range;
                    columnD?.Select();
                    _logger.Debug("已选择D列");
                    System.Threading.Thread.Sleep(500);
                    
                    // 6. 回到A1
                    worksheet.Range["A1"].Select();
                    _logger.Debug("已返回A1");
                    
                    MessageHelper.ShowInfo(
                        "区域导航演示完成！\n\n" +
                        "已演示:\n" +
                        "• 选择特定区域 (A1:C3)\n" +
                        "• 查找使用范围\n" +
                        "• 导航到特定单元格 (E5)\n" +
                        "• 选择整行 (第7行)\n" +
                        "• 选择整列 (D列)\n" +
                        "• 返回起始位置 (A1)\n\n" +
                        $"工作表使用范围: {usedRange.Address}",
                        "导航完成");
                    
                    _logger.Info("区域导航演示完成");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "区域导航失败");
                    MessageHelper.ShowError($"区域导航失败: {ex.Message}");
                }
            }
        }

        #endregion
    }
} 