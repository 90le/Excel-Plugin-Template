using System;
using System.Collections.Generic;
using DTI_Tool.AddIn.Common.Interfaces;
using DTI_Tool.AddIn.Common.Models;
using DTI_Tool.AddIn.Core;

namespace BasePlugin.Models
{
    /// <summary>
    /// 插件功能基础模型
    /// </summary>
    public class PluginFeature
    {
        /// <summary>
        /// 功能唯一标识符
        /// </summary>
        public string Id { get; set; }
        
        /// <summary>
        /// 功能名称
        /// </summary>
        public string Name { get; set; }
        
        /// <summary>
        /// 功能描述
        /// </summary>
        public string Description { get; set; }
        
        /// <summary>
        /// 功能类别
        /// </summary>
        public string Category { get; set; }
        
        /// <summary>
        /// 功能标签
        /// </summary>
        public List<string> Tags { get; set; } = new List<string>();
        
        /// <summary>
        /// 图标文字
        /// </summary>
        public string IconText { get; set; }
        
        /// <summary>
        /// Office内置图标名称
        /// </summary>
        public string ImageMso { get; set; }
        
        /// <summary>
        /// 是否启用
        /// </summary>
        public bool IsEnabled { get; set; } = true;
        
        /// <summary>
        /// 执行动作
        /// </summary>
        public Action Action { get; set; }
        
        /// <summary>
        /// 转换为RibbonButton
        /// </summary>
        public RibbonButton ToRibbonButton()
        {
            return new RibbonButton
            {
                Type = "button",
                Id = Id,
                Label = Name,
                ImageMso = ImageMso,
                IconText = IconText,
                Tooltip = Description,
                OnAction = Id,
                Enabled = IsEnabled,
                Image = null
            };
        }
        
        /// <summary>
        /// 转换为PluginCommand
        /// </summary>
        public PluginCommand ToPluginCommand()
        {
            return new PluginCommand
            {
                Id = Id,
                Name = Name,
                Description = Description,
                Category = Category,
                Tags = Tags,
                IconText = IconText,
                IsEnabled = IsEnabled
            };
        }
    }
} 