using System;
using System.Collections.Generic;
using BasePlugin.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasePlugin.Features
{
    /// <summary>
    /// 示例功能类 - 提供基础功能示例
    /// </summary>
    public class SampleFeatures
    {
        private Excel.Application _excelApp;
        
        public SampleFeatures(Excel.Application excelApp)
        {
            _excelApp = excelApp;
        }
        
        /// <summary>
        /// 获取所有示例功能
        /// </summary>
        public List<PluginFeature> GetFeatures()
        {
            return new List<PluginFeature>
            {
                new PluginFeature
                {
                    Id = "hello_world",
                    Name = "Hello World",
                    Description = "显示一个简单的问候消息",
                    Category = "示例功能",
                    Tags = new List<string> { "示例", "基础" },
                    ImageMso = "HappyFace",
                    Action = HelloWorld
                },
                new PluginFeature
                {
                    Id = "get_selection_info",
                    Name = "获取选择信息",
                    Description = "显示当前选中区域的基本信息",
                    Category = "示例功能",
                    Tags = new List<string> { "选择", "信息" },
                    ImageMso = "TableExcelSelect",
                    Action = GetSelectionInfo
                },
                new PluginFeature
                {
                    Id = "insert_current_time",
                    Name = "插入当前时间",
                    Description = "在活动单元格插入当前日期和时间",
                    Category = "示例功能",
                    Tags = new List<string> { "时间", "插入" },
                    ImageMso = "DateAndTimePicker",
                    Action = InsertCurrentTime
                }
            };
        }
        
        #region 示例功能实现
        
        /// <summary>
        /// Hello World 示例
        /// </summary>
        private void HelloWorld()
        {
            try
            {
                System.Windows.Forms.MessageBox.Show(
                    "Hello World! 这是一个基础插件示例。\n\n您可以基于这个模板开发自己的Excel插件。", 
                    "基础插件示例", 
                    System.Windows.Forms.MessageBoxButtons.OK, 
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ShowError($"Hello World 功能执行失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 获取选择信息
        /// </summary>
        private void GetSelectionInfo()
        {
            try
            {
                var selection = _excelApp?.Selection as Excel.Range;
                if (selection == null)
                {
                    ShowMessage("请先选择一个区域");
                    return;
                }
                
                var info = $"选中区域信息:\n" +
                          $"地址: {selection.Address}\n" +
                          $"行数: {selection.Rows.Count}\n" +
                          $"列数: {selection.Columns.Count}\n" +
                          $"单元格数: {selection.Cells.Count}";
                
                ShowMessage(info);
            }
            catch (Exception ex)
            {
                ShowError($"获取选择信息失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 插入当前时间
        /// </summary>
        private void InsertCurrentTime()
        {
            try
            {
                var activeCell = _excelApp?.ActiveCell;
                if (activeCell == null)
                {
                    ShowMessage("请先选择一个单元格");
                    return;
                }
                
                activeCell.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                ShowMessage("已在活动单元格插入当前时间");
            }
            catch (Exception ex)
            {
                ShowError($"插入时间失败: {ex.Message}");
            }
        }
        
        #endregion
        
        #region 辅助方法
        
        private void ShowMessage(string message)
        {
            System.Windows.Forms.MessageBox.Show(message, "基础插件示例", 
                System.Windows.Forms.MessageBoxButtons.OK, 
                System.Windows.Forms.MessageBoxIcon.Information);
        }
        
        private void ShowError(string message)
        {
            System.Windows.Forms.MessageBox.Show(message, "基础插件示例", 
                System.Windows.Forms.MessageBoxButtons.OK, 
                System.Windows.Forms.MessageBoxIcon.Error);
        }
        
        #endregion
    }
} 