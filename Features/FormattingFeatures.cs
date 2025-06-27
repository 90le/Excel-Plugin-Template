using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using BasePlugin.Core;
using BasePlugin.Models;
using DTI_Tool.AddIn.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasePlugin.Features
{
    /// <summary>
    /// 格式化功能类 - 提供单元格格式化相关的示例功能
    /// </summary>
    public class FormattingFeatures : IFeatureProvider
    {
        #region 私有字段

        private readonly PluginLogger _logger;

        #endregion

        #region 构造函数

        public FormattingFeatures(PluginLogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            
            _logger.Debug("FormattingFeatures 已初始化");
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
        /// 获取格式化功能列表
        /// </summary>
        public List<PluginFeature> GetFeatures()
        {
            return new List<PluginFeature>
            {
                new PluginFeature
                {
                    Id = "apply_number_format",
                    Name = "数字格式化",
                    Description = "为选中单元格应用数字格式",
                    Category = "格式化",
                    Tags = new List<string> { "格式", "数字", "货币", "百分比" },
                    ImageMso = "NumberFormatCurrency",
                    Action = ApplyNumberFormat
                },
                new PluginFeature
                {
                    Id = "apply_conditional_format",
                    Name = "条件格式",
                    Description = "为选中区域添加条件格式",
                    Category = "格式化",
                    Tags = new List<string> { "条件", "格式", "颜色", "规则" },
                    ImageMso = "ConditionalFormattingMenu",
                    Action = ApplyConditionalFormat
                },
                new PluginFeature
                {
                    Id = "apply_cell_styles",
                    Name = "单元格样式",
                    Description = "快速应用预定义的单元格样式",
                    Category = "格式化",
                    Tags = new List<string> { "样式", "主题", "颜色" },
                    ImageMso = "CellStylesGallery",
                    Action = ApplyCellStyles
                },
                new PluginFeature
                {
                    Id = "format_as_table",
                    Name = "格式化为表格",
                    Description = "将数据区域格式化为表格",
                    Category = "格式化",
                    Tags = new List<string> { "表格", "样式", "筛选" },
                    ImageMso = "TableInsert",
                    Action = FormatAsTable
                },
                new PluginFeature
                {
                    Id = "apply_borders",
                    Name = "边框设置",
                    Description = "为选中区域设置边框",
                    Category = "格式化",
                    Tags = new List<string> { "边框", "线条", "样式" },
                    ImageMso = "BordersAll",
                    Action = ApplyBorders
                }
            };
        }

        public void Dispose()
        {
            _logger.Debug("FormattingFeatures 已释放");
        }

        #endregion

        #region 功能实现

        /// <summary>
        /// 应用数字格式
        /// </summary>
        private void ApplyNumberFormat()
        {
            using (_logger.MeasurePerformance("应用数字格式"))
            {
                try
                {
                    _logger.Info("开始应用数字格式");
                    
                    var selection = ExcelApp?.Selection as Excel.Range;
                    if (selection == null)
                    {
                        MessageHelper.ShowWarning("请选择要格式化的单元格");
                        return;
                    }
                    
                    var formats = new Dictionary<string, string>
                    {
                        { "1. 通用格式", "General" },
                        { "2. 数字格式 (2位小数)", "#,##0.00" },
                        { "3. 货币格式", "¥#,##0.00" },
                        { "4. 百分比格式", "0.00%" },
                        { "5. 日期格式", "yyyy-mm-dd" },
                        { "6. 时间格式", "hh:mm:ss" },
                        { "7. 科学计数法", "0.00E+00" },
                        { "8. 文本格式", "@" }
                    };
                    
                    var formatList = string.Join("\n", formats.Keys);
                    var input = Microsoft.VisualBasic.Interaction.InputBox(
                        $"请选择数字格式（输入数字1-8）:\n\n{formatList}",
                        "选择格式",
                        "1");
                    
                    if (string.IsNullOrEmpty(input))
                    {
                        _logger.Info("用户取消了格式选择");
                        return;
                    }
                    
                    var key = formats.Keys.FirstOrDefault(k => k.StartsWith(input + "."));
                    if (key != null && formats.TryGetValue(key, out var format))
                    {
                        selection.NumberFormat = format;
                        MessageHelper.ShowInfo($"已应用格式: {key}", "格式化成功");
                        _logger.Info("已应用数字格式: {0}", format);
                    }
                    else
                    {
                        MessageHelper.ShowWarning("无效的选择");
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "应用数字格式失败");
                    MessageHelper.ShowError($"应用数字格式失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 应用条件格式
        /// </summary>
        private void ApplyConditionalFormat()
        {
            using (_logger.MeasurePerformance("应用条件格式"))
            {
                try
                {
                    _logger.Info("开始应用条件格式");
                    
                    var selection = ExcelApp?.Selection as Excel.Range;
                    if (selection == null)
                    {
                        MessageHelper.ShowWarning("请选择要应用条件格式的区域");
                        return;
                    }
                    
                    var options = "请选择条件格式类型:\n\n" +
                                "1. 突出显示大于某值的单元格\n" +
                                "2. 数据条\n" +
                                "3. 色阶（红-黄-绿）\n" +
                                "4. 图标集";
                    
                    var choice = Microsoft.VisualBasic.Interaction.InputBox(options, "条件格式", "1");
                    
                    if (string.IsNullOrEmpty(choice))
                    {
                        _logger.Info("用户取消了条件格式选择");
                        return;
                    }
                    
                    // 清除现有条件格式
                    selection.FormatConditions.Delete();
                    
                    switch (choice)
                    {
                        case "1":
                            ApplyHighlightCondition(selection);
                            break;
                        case "2":
                            ApplyDataBar(selection);
                            break;
                        case "3":
                            ApplyColorScale(selection);
                            break;
                        case "4":
                            ApplyIconSet(selection);
                            break;
                        default:
                            MessageHelper.ShowWarning("无效的选择");
                            return;
                    }
                    
                    _logger.Info("条件格式应用成功");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "应用条件格式失败");
                    MessageHelper.ShowError($"应用条件格式失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 应用单元格样式
        /// </summary>
        private void ApplyCellStyles()
        {
            using (_logger.MeasurePerformance("应用单元格样式"))
            {
                try
                {
                    _logger.Info("开始应用单元格样式");
                    
                    var selection = ExcelApp?.Selection as Excel.Range;
                    if (selection == null)
                    {
                        MessageHelper.ShowWarning("请选择要应用样式的单元格");
                        return;
                    }
                    
                    // 创建自定义样式
                    var styles = new[]
                    {
                        new { Name = "标题样式", Bold = true, Size = 16, Color = Color.DarkBlue, BgColor = Color.LightGray },
                        new { Name = "强调样式", Bold = true, Size = 11, Color = Color.White, BgColor = Color.DarkBlue },
                        new { Name = "警告样式", Bold = false, Size = 11, Color = Color.DarkRed, BgColor = Color.LightYellow },
                        new { Name = "成功样式", Bold = false, Size = 11, Color = Color.DarkGreen, BgColor = Color.LightGreen }
                    };
                    
                    var styleNames = string.Join("\n", styles.Select((s, i) => $"{i + 1}. {s.Name}"));
                    var choice = Microsoft.VisualBasic.Interaction.InputBox(
                        $"选择要应用的样式:\n\n{styleNames}",
                        "选择样式",
                        "1");
                    
                    if (string.IsNullOrEmpty(choice) || !int.TryParse(choice, out int index) || index < 1 || index > styles.Length)
                    {
                        _logger.Info("用户取消了样式选择或选择无效");
                        return;
                    }
                    
                    var selectedStyle = styles[index - 1];
                    
                    // 应用样式
                    selection.Font.Bold = selectedStyle.Bold;
                    selection.Font.Size = selectedStyle.Size;
                    selection.Font.Color = ColorTranslator.ToOle(selectedStyle.Color);
                    selection.Interior.Color = ColorTranslator.ToOle(selectedStyle.BgColor);
                    selection.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    selection.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    
                    MessageHelper.ShowInfo($"已应用 {selectedStyle.Name}", "样式应用成功");
                    _logger.Info("已应用单元格样式: {0}", selectedStyle.Name);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "应用单元格样式失败");
                    MessageHelper.ShowError($"应用单元格样式失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 格式化为表格
        /// </summary>
        private void FormatAsTable()
        {
            using (_logger.MeasurePerformance("格式化为表格"))
            {
                try
                {
                    _logger.Info("开始格式化为表格");
                    
                    var selection = ExcelApp?.Selection as Excel.Range;
                    if (selection == null || selection.Cells.Count < 2)
                    {
                        MessageHelper.ShowWarning("请选择要格式化为表格的数据区域（至少2个单元格）");
                        return;
                    }
                    
                    var worksheet = selection.Worksheet;
                    _logger.Debug("选中区域: {0}", selection.Address);
                    
                    // 确认是否包含标题行
                    var hasHeaders = MessageHelper.ShowConfirm("数据的第一行是否为标题行？", "确认标题");
                    
                    // 创建表格
                    var listObject = worksheet.ListObjects.Add(
                        Excel.XlListObjectSourceType.xlSrcRange,
                        selection,
                        Type.Missing,
                        hasHeaders ? Excel.XlYesNoGuess.xlYes : Excel.XlYesNoGuess.xlNo);
                    
                    // 设置表格样式
                    listObject.TableStyle = "TableStyleMedium2";
                    listObject.ShowTableStyleRowStripes = true;
                    
                    // 设置表格名称
                    var tableName = $"Table_{DateTime.Now:yyyyMMddHHmmss}";
                    listObject.Name = tableName;
                    
                    MessageHelper.ShowInfo($"已创建表格: {tableName}\n\n" +
                                        "表格功能:\n" +
                                        "• 自动筛选\n" +
                                        "• 排序功能\n" +
                                        "• 自动扩展\n" +
                                        "• 结构化引用",
                                        "表格创建成功");
                    
                    _logger.Info("表格创建成功: {0}", tableName);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "格式化为表格失败");
                    MessageHelper.ShowError($"格式化为表格失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 应用边框
        /// </summary>
        private void ApplyBorders()
        {
            using (_logger.MeasurePerformance("应用边框"))
            {
                try
                {
                    _logger.Info("开始应用边框");
                    
                    var selection = ExcelApp?.Selection as Excel.Range;
                    if (selection == null)
                    {
                        MessageHelper.ShowWarning("请选择要添加边框的区域");
                        return;
                    }
                    
                    var borderStyles = new[]
                    {
                        "1. 所有边框",
                        "2. 外边框",
                        "3. 内部边框", 
                        "4. 双线边框",
                        "5. 粗边框",
                        "6. 清除边框"
                    };
                    
                    var choice = Microsoft.VisualBasic.Interaction.InputBox(
                        $"选择边框样式:\n\n{string.Join("\n", borderStyles)}",
                        "边框样式",
                        "1");
                    
                    if (string.IsNullOrEmpty(choice))
                    {
                        _logger.Info("用户取消了边框选择");
                        return;
                    }
                    
                    _logger.Debug("应用边框到区域: {0}, 样式选择: {1}", selection.Address, choice);
                    
                    switch (choice)
                    {
                        case "1": // 所有边框
                            ApplyAllBorders(selection, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin);
                            break;
                        case "2": // 外边框
                            ApplyOutsideBorders(selection, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
                            break;
                        case "3": // 内部边框
                            ApplyInsideBorders(selection, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin);
                            break;
                        case "4": // 双线边框
                            ApplyAllBorders(selection, Excel.XlLineStyle.xlDouble, Excel.XlBorderWeight.xlThick);
                            break;
                        case "5": // 粗边框
                            ApplyAllBorders(selection, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
                            break;
                        case "6": // 清除边框
                            selection.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                            MessageHelper.ShowInfo("已清除所有边框", "操作成功");
                            _logger.Info("已清除边框");
                            return;
                        default:
                            MessageHelper.ShowWarning("无效的选择");
                            return;
                    }
                    
                    MessageHelper.ShowInfo("边框应用成功", "操作完成");
                    _logger.Info("边框应用成功");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "应用边框失败");
                    MessageHelper.ShowError($"应用边框失败: {ex.Message}");
                }
            }
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 应用高亮条件
        /// </summary>
        private void ApplyHighlightCondition(Excel.Range range)
        {
            var value = Microsoft.VisualBasic.Interaction.InputBox("输入阈值（大于此值的单元格将被突出显示）:", "设置阈值", "50");
            if (string.IsNullOrEmpty(value) || !double.TryParse(value, out double threshold))
            {
                MessageHelper.ShowWarning("无效的数值");
                return;
            }
            
            var condition = range.FormatConditions.Add(
                Excel.XlFormatConditionType.xlCellValue,
                Excel.XlFormatConditionOperator.xlGreater,
                threshold) as Excel.FormatCondition;
            
            if (condition != null)
            {
                condition.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                condition.Font.Bold = true;
            }
            
            MessageHelper.ShowInfo($"已设置条件格式：大于 {threshold} 的值将被突出显示", "条件格式");
            _logger.Debug("应用高亮条件，阈值: {0}", threshold);
        }

        /// <summary>
        /// 应用数据条
        /// </summary>
        private void ApplyDataBar(Excel.Range range)
        {
            var dataBar = range.FormatConditions.AddDatabar() as Excel.Databar;
            if (dataBar != null)
            {
                try
                {
                    // 简化颜色设置，避免复杂的COM类型转换
                    dataBar.BarFillType = Excel.XlDataBarFillType.xlDataBarFillGradient;
                    // 注意：颜色设置可能需要特定的COM接口，这里先跳过以避免编译错误
                    _logger.Debug("数据条已添加，使用默认颜色");
                }
                catch (Exception ex)
                {
                    _logger.Warning("设置数据条属性时发生错误: {0}", ex.Message);
                }
            }
            
            MessageHelper.ShowInfo("已添加数据条", "条件格式");
            _logger.Debug("应用数据条");
        }

        /// <summary>
        /// 应用色阶
        /// </summary>
        private void ApplyColorScale(Excel.Range range)
        {
            var colorScale = range.FormatConditions.AddColorScale(3) as Excel.ColorScale;
            
            if (colorScale != null)
            {
                // 设置最小值颜色（红色）
                colorScale.ColorScaleCriteria[1].Type = Excel.XlConditionValueTypes.xlConditionValueLowestValue;
                colorScale.ColorScaleCriteria[1].FormatColor.Color = ColorTranslator.ToOle(Color.Red);
                
                // 设置中间值颜色（黄色）
                colorScale.ColorScaleCriteria[2].Type = Excel.XlConditionValueTypes.xlConditionValuePercentile;
                colorScale.ColorScaleCriteria[2].Value = 50;
                colorScale.ColorScaleCriteria[2].FormatColor.Color = ColorTranslator.ToOle(Color.Yellow);
                
                // 设置最大值颜色（绿色）
                colorScale.ColorScaleCriteria[3].Type = Excel.XlConditionValueTypes.xlConditionValueHighestValue;
                colorScale.ColorScaleCriteria[3].FormatColor.Color = ColorTranslator.ToOle(Color.Green);
            }
            
            MessageHelper.ShowInfo("已添加色阶（红-黄-绿）", "条件格式");
            _logger.Debug("应用色阶");
        }

        /// <summary>
        /// 应用图标集
        /// </summary>
        private void ApplyIconSet(Excel.Range range)
        {
            var iconSet = range.FormatConditions.AddIconSetCondition() as Excel.IconSetCondition;
            if (iconSet != null)
            {
                iconSet.IconSet = (Excel.IconSets)ExcelApp.ActiveWorkbook.IconSets[Excel.XlIconSet.xl3TrafficLights1];
            }
            
            MessageHelper.ShowInfo("已添加图标集（交通灯）", "条件格式");
            _logger.Debug("应用图标集");
        }

        /// <summary>
        /// 应用所有边框
        /// </summary>
        private void ApplyAllBorders(Excel.Range range, Excel.XlLineStyle lineStyle, Excel.XlBorderWeight weight)
        {
            var borders = new[]
            {
                Excel.XlBordersIndex.xlEdgeLeft,
                Excel.XlBordersIndex.xlEdgeTop,
                Excel.XlBordersIndex.xlEdgeBottom,
                Excel.XlBordersIndex.xlEdgeRight,
                Excel.XlBordersIndex.xlInsideVertical,
                Excel.XlBordersIndex.xlInsideHorizontal
            };
            
            foreach (var border in borders)
            {
                try
                {
                    range.Borders[border].LineStyle = lineStyle;
                    range.Borders[border].Weight = weight;
                }
                catch
                {
                    // 忽略不适用的边框（如单个单元格的内部边框）
                }
            }
        }

        /// <summary>
        /// 应用外边框
        /// </summary>
        private void ApplyOutsideBorders(Excel.Range range, Excel.XlLineStyle lineStyle, Excel.XlBorderWeight weight)
        {
            range.BorderAround(lineStyle, weight);
        }

        /// <summary>
        /// 应用内部边框
        /// </summary>
        private void ApplyInsideBorders(Excel.Range range, Excel.XlLineStyle lineStyle, Excel.XlBorderWeight weight)
        {
            try
            {
                range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = lineStyle;
                range.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = weight;
                range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = lineStyle;
                range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = weight;
            }
            catch
            {
                MessageHelper.ShowWarning("选中区域没有内部边框可设置");
            }
        }

        #endregion
    }
} 