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
    /// 数据处理功能类 - 提供数据处理相关的示例功能
    /// </summary>
    public class DataProcessingFeatures : IFeatureProvider
    {
        #region 私有字段

        private readonly PluginLogger _logger;

        #endregion

        #region 构造函数

        public DataProcessingFeatures(PluginLogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            
            _logger.Debug("DataProcessingFeatures 已初始化");
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
        /// 获取数据处理功能列表
        /// </summary>
        public List<PluginFeature> GetFeatures()
        {
            return new List<PluginFeature>
            {
                new PluginFeature
                {
                    Id = "data_sort",
                    Name = "数据排序",
                    Description = "对选中区域的数据进行排序",
                    Category = "数据处理",
                    Tags = new List<string> { "排序", "数据", "升序", "降序" },
                    ImageMso = "SortUp",
                    Action = DataSort
                },
                new PluginFeature
                {
                    Id = "data_filter",
                    Name = "数据筛选",
                    Description = "为选中区域添加筛选器",
                    Category = "数据处理",
                    Tags = new List<string> { "筛选", "过滤", "数据" },
                    ImageMso = "FilterBySelection",
                    Action = DataFilter
                },
                new PluginFeature
                {
                    Id = "remove_duplicates",
                    Name = "删除重复项",
                    Description = "删除选中区域的重复数据",
                    Category = "数据处理",
                    Tags = new List<string> { "去重", "重复", "清理" },
                    ImageMso = "RemoveDuplicates",
                    Action = RemoveDuplicates
                },
                new PluginFeature
                {
                    Id = "data_statistics",
                    Name = "数据统计",
                    Description = "计算选中区域的基本统计信息",
                    Category = "数据处理",
                    Tags = new List<string> { "统计", "求和", "平均值", "计数" },
                    ImageMso = "FunctionWizard",
                    Action = DataStatistics
                },
                new PluginFeature
                {
                    Id = "fill_series",
                    Name = "填充序列",
                    Description = "自动填充数字或日期序列",
                    Category = "数据处理",
                    Tags = new List<string> { "填充", "序列", "自动" },
                    ImageMso = "FillSeriesDates",
                    Action = FillSeries
                }
            };
        }

        public void Dispose()
        {
            _logger.Debug("DataProcessingFeatures 已释放");
        }

        #endregion

        #region 功能实现

        /// <summary>
        /// 数据排序
        /// </summary>
        private void DataSort()
        {
            using (_logger.MeasurePerformance("数据排序"))
            {
                try
                {
                    _logger.Info("开始执行数据排序");
                    
                    var selection = ExcelApp?.Selection as Excel.Range;
                    if (selection == null)
                    {
                        MessageHelper.ShowWarning("请先选择要排序的数据区域");
                        return;
                    }
                    
                    if (selection.Cells.Count < 2)
                    {
                        MessageHelper.ShowWarning("请选择要排序的数据区域（至少2个单元格）");
                        return;
                    }
                    
                    // 检查是否为合并单元格
                    if (selection.MergeCells is true)
                    {
                        MessageHelper.ShowWarning("选择的区域包含合并单元格，无法排序。请先取消合并单元格。");
                        return;
                    }
                    
                    // 询问排序方式
                    var result = MessageHelper.ShowYesNoCancel(
                        "选择排序方式:\n\n" +
                        "是 - 升序排序\n" +
                        "否 - 降序排序\n" +
                        "取消 - 取消操作", 
                        "数据排序");
                    
                    if (result == System.Windows.Forms.DialogResult.Cancel)
                    {
                        _logger.Info("用户取消了排序操作");
                        return;
                    }
                    
                    var sortOrder = result == System.Windows.Forms.DialogResult.Yes 
                        ? Excel.XlSortOrder.xlAscending 
                        : Excel.XlSortOrder.xlDescending;
                    
                    _logger.Debug("排序区域: {0}, 排序方式: {1}", selection.Address, sortOrder);
                    
                    // 检查选择区域的有效性
                    if (selection.Columns.Count == 0)
                    {
                        MessageHelper.ShowWarning("选择的区域无效，无法进行排序");
                        return;
                    }
                    
                    // 执行排序
                    var sortAddress = selection.Address;
                    var worksheet = selection.Worksheet;
                    
                    try
                    {
                        // 尝试使用工作表排序对象
                        try
                        {
                            var sortFields = worksheet.Sort.SortFields;
                            sortFields.Clear();
                            
                            // 添加排序字段（使用第一列）
                            var firstColumn = selection.Columns[1] as Excel.Range;
                            sortFields.Add(firstColumn, Excel.XlSortOn.xlSortOnValues, sortOrder);
                            
                            // 应用排序
                            worksheet.Sort.SetRange(selection);
                            worksheet.Sort.Header = Excel.XlYesNoGuess.xlGuess;
                            worksheet.Sort.MatchCase = false;
                            worksheet.Sort.Orientation = Excel.XlSortOrientation.xlSortColumns;
                            worksheet.Sort.SortMethod = Excel.XlSortMethod.xlPinYin;
                            worksheet.Sort.Apply();
                        }
                        catch
                        {
                            _logger.Info("工作表排序方法失败，尝试使用备用排序方法");
                            
                            // 备用排序方法：使用简单的Range.Sort方法
                            var firstColumn = selection.Columns[1] as Excel.Range;
                            selection.Sort(firstColumn, sortOrder);
                        }
                        
                        var sortTypeName = sortOrder == Excel.XlSortOrder.xlAscending ? "升序" : "降序";
                        MessageHelper.ShowInfo($"已对区域 {sortAddress} 完成{sortTypeName}排序", "排序成功");
                        _logger.Info("数据排序完成，区域: {0}，排序方式: {1}", sortAddress, sortTypeName);
                    }
                    catch (System.Runtime.InteropServices.COMException comEx)
                    {
                        _logger.Error(comEx, "Excel COM 排序操作失败");
                        MessageHelper.ShowError($"排序操作失败: {comEx.Message}\n\n请检查数据格式是否正确。");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "数据排序失败");
                    MessageHelper.ShowError($"数据排序失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 数据筛选
        /// </summary>
        private void DataFilter()
        {
            using (_logger.MeasurePerformance("数据筛选"))
            {
                try
                {
                    _logger.Info("开始设置数据筛选");
                    
                    var selection = ExcelApp?.Selection as Excel.Range;
                    if (selection == null)
                    {
                        MessageHelper.ShowWarning("请选择要添加筛选器的数据区域");
                        return;
                    }
                    
                    var worksheet = selection.Worksheet;
                    
                    // 检查是否已有筛选器
                    if (worksheet.AutoFilterMode)
                    {
                        // 如果已有筛选器，询问是否移除
                        if (MessageHelper.ShowConfirm("当前工作表已有筛选器，是否移除？", "筛选器"))
                        {
                            worksheet.AutoFilterMode = false;
                            MessageHelper.ShowInfo("已移除筛选器", "操作成功");
                            _logger.Info("已移除筛选器");
                        }
                    }
                    else
                    {
                        // 添加筛选器
                        selection.AutoFilter();
                        MessageHelper.ShowInfo($"已为区域 {selection.Address} 添加筛选器\n\n" +
                                            "点击列标题的下拉箭头可以进行筛选", "筛选器已添加");
                        _logger.Info("已添加筛选器到区域: {0}", selection.Address);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "设置数据筛选失败");
                    MessageHelper.ShowError($"设置数据筛选失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 删除重复项
        /// </summary>
        private void RemoveDuplicates()
        {
            using (_logger.MeasurePerformance("删除重复项"))
            {
                try
                {
                    _logger.Info("开始删除重复项");
                    
                    var selection = ExcelApp?.Selection as Excel.Range;
                    if (selection == null || selection.Cells.Count < 2)
                    {
                        MessageHelper.ShowWarning("请选择要删除重复项的数据区域（至少2个单元格）");
                        return;
                    }
                    
                    // 记录原始行数
                    var originalRows = selection.Rows.Count;
                    _logger.Debug("原始数据行数: {0}", originalRows);
                    
                    // 询问用户确认
                    if (!MessageHelper.ShowConfirm($"将在区域 {selection.Address} 中删除重复项。\n\n此操作不可撤销，是否继续？", "确认删除重复项"))
                    {
                        _logger.Info("用户取消了删除重复项操作");
                        return;
                    }
                    
                    // 创建列索引数组（所有列）
                    var columns = new object[selection.Columns.Count];
                    for (int i = 0; i < selection.Columns.Count; i++)
                    {
                        columns[i] = i + 1;
                    }
                    
                    // 删除重复项
                    selection.RemoveDuplicates(columns, Excel.XlYesNoGuess.xlNo);
                    
                    // 计算删除的行数
                    var remainingRows = selection.Rows.Count;
                    var deletedRows = originalRows - remainingRows;
                    
                    MessageHelper.ShowInfo($"删除重复项完成！\n\n" +
                                        $"原始行数: {originalRows}\n" +
                                        $"删除行数: {deletedRows}\n" +
                                        $"剩余行数: {remainingRows}", 
                                        "操作成功");
                    
                    _logger.Info("删除重复项完成，删除了 {0} 行", deletedRows);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "删除重复项失败");
                    MessageHelper.ShowError($"删除重复项失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 数据统计
        /// </summary>
        private void DataStatistics()
        {
            using (_logger.MeasurePerformance("数据统计"))
            {
                try
                {
                    _logger.Info("开始计算数据统计");
                    
                    var selection = ExcelApp?.Selection as Excel.Range;
                    if (selection == null)
                    {
                        MessageHelper.ShowWarning("请选择要统计的数据区域");
                        return;
                    }
                    
                    _logger.Debug("统计区域: {0}", selection.Address);
                    
                    // 统计数值数据
                    var numericCells = new List<double>();
                    foreach (Excel.Range cell in selection.Cells)
                    {
                        if (cell.Value != null && IsNumeric(cell.Value))
                        {
                            numericCells.Add(Convert.ToDouble(cell.Value));
                        }
                    }
                    
                    if (numericCells.Count == 0)
                    {
                        MessageHelper.ShowWarning("选中区域没有数值数据可供统计");
                        return;
                    }
                    
                    // 计算统计信息
                    var count = numericCells.Count;
                    var sum = numericCells.Sum();
                    var average = numericCells.Average();
                    var min = numericCells.Min();
                    var max = numericCells.Max();
                    var stdDev = CalculateStandardDeviation(numericCells);
                    
                    // 显示统计结果
                    var statistics = $"数据统计结果:\n\n" +
                                   $"统计区域: {selection.Address}\n" +
                                   $"总单元格数: {selection.Cells.Count}\n" +
                                   $"数值单元格数: {count}\n\n" +
                                   $"求和: {sum:N2}\n" +
                                   $"平均值: {average:N2}\n" +
                                   $"最小值: {min:N2}\n" +
                                   $"最大值: {max:N2}\n" +
                                   $"标准差: {stdDev:N2}";
                    
                    MessageHelper.ShowInfo(statistics, "统计结果");
                    _logger.Info("数据统计完成，数值单元格数: {0}", count);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "数据统计失败");
                    MessageHelper.ShowError($"数据统计失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 填充序列
        /// </summary>
        private void FillSeries()
        {
            using (_logger.MeasurePerformance("填充序列"))
            {
                try
                {
                    _logger.Info("开始填充序列");
                    
                    var selection = ExcelApp?.Selection as Excel.Range;
                    if (selection == null || selection.Cells.Count < 2)
                    {
                        MessageHelper.ShowWarning("请选择至少2个单元格来填充序列");
                        return;
                    }
                    
                    // 询问填充类型
                    var message = "选择填充类型:\n\n" +
                                "是 - 数字序列（1, 2, 3...）\n" +
                                "否 - 日期序列（按天递增）\n" +
                                "取消 - 取消操作";
                    
                    var result = MessageHelper.ShowYesNoCancel(message, "填充序列");
                    
                    if (result == System.Windows.Forms.DialogResult.Cancel)
                    {
                        _logger.Info("用户取消了填充序列操作");
                        return;
                    }
                    
                    _logger.Debug("填充区域: {0}", selection.Address);
                    
                    if (result == System.Windows.Forms.DialogResult.Yes)
                    {
                        // 填充数字序列
                        FillNumberSeries(selection);
                    }
                    else
                    {
                        // 填充日期序列
                        FillDateSeries(selection);
                    }
                    
                    _logger.Info("序列填充完成");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "填充序列失败");
                    MessageHelper.ShowError($"填充序列失败: {ex.Message}");
                }
            }
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 判断是否为数值
        /// </summary>
        private bool IsNumeric(object value)
        {
            return double.TryParse(value?.ToString(), out _);
        }

        /// <summary>
        /// 计算标准差
        /// </summary>
        private double CalculateStandardDeviation(List<double> values)
        {
            if (values.Count < 2) return 0;
            
            var average = values.Average();
            var sumOfSquares = values.Sum(v => Math.Pow(v - average, 2));
            return Math.Sqrt(sumOfSquares / (values.Count - 1));
        }

        /// <summary>
        /// 填充数字序列
        /// </summary>
        private void FillNumberSeries(Excel.Range selection)
        {
            try
            {
                var firstCell = selection.Cells[1, 1] as Excel.Range;
                
                // 获取起始值
                double startValue = 1;
                if (firstCell.Value != null && IsNumeric(firstCell.Value))
                {
                    startValue = Convert.ToDouble(firstCell.Value);
                }
                else
                {
                    firstCell.Value = startValue;
                }
                
                // 填充序列
                var step = 1;
                int index = 1;
                foreach (Excel.Range cell in selection.Cells)
                {
                    if (index > 1)
                    {
                        cell.Value = startValue + (index - 1) * step;
                    }
                    index++;
                }
                
                MessageHelper.ShowInfo($"已填充数字序列（起始值: {startValue}，步长: {step}）", "填充成功");
                _logger.Debug("填充数字序列完成，起始值: {0}", startValue);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("填充数字序列时发生错误", ex);
            }
        }

        /// <summary>
        /// 填充日期序列
        /// </summary>
        private void FillDateSeries(Excel.Range selection)
        {
            try
            {
                var firstCell = selection.Cells[1, 1] as Excel.Range;
                
                // 获取起始日期
                DateTime startDate = DateTime.Today;
                if (firstCell.Value != null && firstCell.Value is DateTime)
                {
                    startDate = (DateTime)firstCell.Value;
                }
                else
                {
                    firstCell.Value = startDate;
                }
                
                // 填充序列
                int index = 1;
                foreach (Excel.Range cell in selection.Cells)
                {
                    if (index > 1)
                    {
                        cell.Value = startDate.AddDays(index - 1);
                    }
                    cell.NumberFormat = "yyyy-mm-dd";
                    index++;
                }
                
                MessageHelper.ShowInfo($"已填充日期序列（起始日期: {startDate:yyyy-MM-dd}）", "填充成功");
                _logger.Debug("填充日期序列完成，起始日期: {0}", startDate);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("填充日期序列时发生错误", ex);
            }
        }

        #endregion
    }
} 