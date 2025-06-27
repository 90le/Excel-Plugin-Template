using System;
using System.Drawing;
using System.Windows.Forms;
using BasePlugin.Core;
using DTI_Tool.AddIn.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace BasePlugin.WinForms
{
    /// <summary>
    /// WinForms 任务窗格控件
    /// </summary>
    public partial class TaskPaneControl : UserControl
    {
        #region 私有字段

        private readonly PluginLogger _logger;
        private readonly FeatureManager _featureManager;

        // 控件声明
        private TableLayoutPanel mainTableLayout;
        private Panel titlePanel;
        private Label titleLabel;
        private Button closeButton;
        private Panel contentPanel;
        
        // 信息区域控件
        private GroupBox infoGroupBox;
        private TableLayoutPanel infoTableLayout;
        private Label workbookLabel;
        private Label workbookValueLabel;
        private Label worksheetLabel;
        private Label worksheetValueLabel;
        private Label selectionLabel;
        private Label selectionValueLabel;
        private Label cellCountLabel;
        private Label cellCountValueLabel;
        private Button refreshButton;

        // 操作区域控件
        private GroupBox quickActionsGroupBox;
        private GroupBox dataProcessingGroupBox;
        private GroupBox formattingGroupBox;
        private GroupBox worksheetGroupBox;
        private GroupBox utilityGroupBox;

        // 状态区域
        private Panel statusPanel;
        private Label statusLabel;

        #endregion

        #region 私有属性

        /// <summary>
        /// 获取Excel应用程序对象
        /// </summary>
        private Excel.Application ExcelApp => HostApplication.Instance?.ExcelApplication;

        #endregion

        #region 构造函数

        public TaskPaneControl()
        {
            // 初始化日志记录器
            _logger = PluginLog.ForPlugin("BasePlugin.WinFormsTaskPane");
            
            // 初始化功能管理器
            _featureManager = new FeatureManager(_logger);
            try
            {
                _featureManager.Initialize();
                _logger.Info("WinForms TaskPane 功能管理器初始化成功");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "WinForms TaskPane 功能管理器初始化失败");
            }

            InitializeComponent();
            
            // 添加窗口大小改变事件处理
            this.SizeChanged += TaskPaneControl_SizeChanged;
            
            RefreshInfo();
            UpdateStatus("WinForms 任务窗格已加载，准备就绪");
        }

        private void TaskPaneControl_SizeChanged(object sender, EventArgs e)
        {
            // 当任务窗格大小改变时，调整所有GroupBox的宽度
            try
            {
                if (contentPanel != null)
                {
                    foreach (Control control in contentPanel.Controls)
                    {
                        if (control is GroupBox groupBox)
                        {
                            groupBox.Width = this.Width - 24;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "调整控件大小时发生错误");
            }
        }

        #endregion

        #region 组件初始化

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // 设置主控件属性
            this.Name = "TaskPaneControl";
            this.MinimumSize = new Size(300, 600);
            this.Size = new Size(340, 800);
            this.BackColor = Color.FromArgb(248, 249, 250);
            this.Padding = new Padding(0);
            this.AutoScroll = true;

            // 创建主布局
            CreateMainLayout();

            // 创建标题栏
            CreateTitlePanel();

            // 创建内容面板
            CreateContentPanel();

            // 创建信息区域
            CreateInfoSection();

            // 创建操作区域
            CreateActionSections();

            // 创建状态栏
            CreateStatusPanel();

            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void CreateMainLayout()
        {
            mainTableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 3,
                ColumnCount = 1,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };

            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F)); // 标题栏
            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // 内容区域
            mainTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60F)); // 状态栏

            this.Controls.Add(mainTableLayout);
        }

        private void CreateTitlePanel()
        {
            titlePanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(0, 122, 204),
                Padding = new Padding(12, 8, 12, 8)
            };

            titleLabel = new Label
            {
                Text = "BasePlugin 控制台",
                Font = new Font("Microsoft YaHei UI", 12F, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = false,
                Dock = DockStyle.Left,
                TextAlign = ContentAlignment.MiddleLeft,
                Width = 200
            };

            closeButton = new Button
            {
                Text = "✕",
                Font = new Font("Microsoft YaHei UI", 12F, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.Transparent,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(30, 30),
                Dock = DockStyle.Right,
                Cursor = Cursors.Hand
            };
            closeButton.FlatAppearance.BorderSize = 0;
            closeButton.Click += CloseButton_Click;

            titlePanel.Controls.Add(titleLabel);
            titlePanel.Controls.Add(closeButton);

            mainTableLayout.Controls.Add(titlePanel, 0, 0);
        }

        private void CreateContentPanel()
        {
            contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(12, 8, 12, 8),
                BackColor = Color.FromArgb(248, 249, 250)
            };

            mainTableLayout.Controls.Add(contentPanel, 0, 1);
        }

        private void CreateInfoSection()
        {
            infoGroupBox = new GroupBox
            {
                Text = "📊 工作表信息",
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(73, 80, 87),
                Height = 150,
                Width = this.Width - 24,
                Location = new Point(8, 0),
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.White,
                Padding = new Padding(4, 18, 4, 8)
            };

            infoTableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 5,
                ColumnCount = 2,
                Padding = new Padding(8)
            };

            // 设置行样式
            for (int i = 0; i < 4; i++)
            {
                infoTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26F));
            }
            infoTableLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));

            // 设置列样式
            infoTableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 85F));
            infoTableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            // 创建标签
            workbookLabel = CreateInfoLabel("工作簿：", 0, 0);
            workbookValueLabel = CreateValueLabel("未加载", 1, 0);
            
            worksheetLabel = CreateInfoLabel("工作表：", 0, 1);
            worksheetValueLabel = CreateValueLabel("未加载", 1, 1);
            
            selectionLabel = CreateInfoLabel("选中区域：", 0, 2);
            selectionValueLabel = CreateValueLabel("无选择", 1, 2);
            
            cellCountLabel = CreateInfoLabel("单元格数：", 0, 3);
            cellCountValueLabel = CreateValueLabel("0", 1, 3);

            refreshButton = new Button
            {
                Text = "🔄 刷新信息",
                Font = new Font("Microsoft YaHei UI", 8.5F, FontStyle.Bold),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Size = new Size(200, 32),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 8, 0, 0),
                Anchor = AnchorStyles.None
            };
            refreshButton.FlatAppearance.BorderSize = 0;
            refreshButton.Click += RefreshButton_Click;
            
            // 添加悬停效果
            refreshButton.MouseEnter += (s, e) => refreshButton.BackColor = Color.FromArgb(90, 98, 104);
            refreshButton.MouseLeave += (s, e) => refreshButton.BackColor = Color.FromArgb(108, 117, 125);

            infoTableLayout.Controls.Add(refreshButton, 0, 4);
            infoTableLayout.SetColumnSpan(refreshButton, 2);

            infoGroupBox.Controls.Add(infoTableLayout);
            contentPanel.Controls.Add(infoGroupBox);
        }

        private Label CreateInfoLabel(string text, int col, int row)
        {
            var label = new Label
            {
                Text = text,
                Font = new Font("Microsoft YaHei UI", 8.5F),
                ForeColor = Color.FromArgb(73, 80, 87),
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(0, 2, 0, 2)
            };
            infoTableLayout.Controls.Add(label, col, row);
            return label;
        }

        private Label CreateValueLabel(string text, int col, int row)
        {
            var label = new Label
            {
                Text = text,
                Font = new Font("Microsoft YaHei UI", 8.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(33, 37, 41),
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(2, 2, 0, 2),
                BackColor = Color.FromArgb(248, 249, 250)
            };
            infoTableLayout.Controls.Add(label, col, row);
            return label;
        }

        private void CreateActionSections()
        {
            int currentY = 160; // 信息区域后的起始位置

            // 快速操作
            quickActionsGroupBox = CreateGroupBox("⚡ 快速操作", currentY);
            var quickActionsButtons = new[]
            {
                CreateActionButton("📅 插入当前时间", "insert_current_time"),
                CreateActionButton("📝 获取选择信息", "get_selection_info"),
                CreateActionButton("🎨 应用单元格样式", "apply_cell_styles"),
                CreateActionButton("📊 数据统计", "data_statistics", Color.FromArgb(40, 167, 69))
            };
            AddButtonsToGroupBox(quickActionsGroupBox, quickActionsButtons);
            contentPanel.Controls.Add(quickActionsGroupBox);
            currentY += CalculateGroupBoxHeight(quickActionsButtons.Length) + 10;

            // 数据处理
            dataProcessingGroupBox = CreateGroupBox("🔧 数据处理", currentY);
            var dataProcessingButtons = new[]
            {
                CreateActionButton("🔤 数据排序", "data_sort"),
                CreateActionButton("🔍 数据筛选", "data_filter"),
                CreateActionButton("🗑️ 删除重复项", "remove_duplicates", Color.FromArgb(255, 193, 7)),
                CreateActionButton("🎲 生成随机数据", "generate_random_data", Color.FromArgb(108, 117, 125)),
                CreateActionButton("📋 填充序列", "fill_series")
            };
            AddButtonsToGroupBox(dataProcessingGroupBox, dataProcessingButtons);
            contentPanel.Controls.Add(dataProcessingGroupBox);
            currentY += CalculateGroupBoxHeight(dataProcessingButtons.Length) + 10;

            // 格式化
            formattingGroupBox = CreateGroupBox("🎨 格式化", currentY);
            var formattingButtons = new[]
            {
                CreateActionButton("🔢 应用数字格式", "apply_number_format"),
                CreateActionButton("🌈 条件格式化", "apply_conditional_format"),
                CreateActionButton("📋 应用边框", "apply_borders"),
                CreateActionButton("🎯 突出显示重要数据", "highlight_important", Color.FromArgb(40, 167, 69))
            };
            AddButtonsToGroupBox(formattingGroupBox, formattingButtons);
            contentPanel.Controls.Add(formattingGroupBox);
            currentY += CalculateGroupBoxHeight(formattingButtons.Length) + 10;

            // 工作表管理
            worksheetGroupBox = CreateGroupBox("📋 工作表管理", currentY);
            var worksheetButtons = new[]
            {
                CreateActionButton("➕ 创建新工作表", "create_worksheet", Color.FromArgb(40, 167, 69)),
                CreateActionButton("✏️ 重命名工作表", "rename_worksheet"),
                CreateActionButton("🔒 保护工作表", "protect_worksheet", Color.FromArgb(255, 193, 7)),
                CreateActionButton("🗂️ 工作表列表", "list_worksheets", Color.FromArgb(108, 117, 125))
            };
            AddButtonsToGroupBox(worksheetGroupBox, worksheetButtons);
            contentPanel.Controls.Add(worksheetGroupBox);
            currentY += CalculateGroupBoxHeight(worksheetButtons.Length) + 10;

            // 实用工具
            utilityGroupBox = CreateGroupBox("🛠️ 实用工具", currentY);
            var utilityButtons = new[]
            {
                CreateActionButton("💾 导出为CSV", "export_csv"),
                CreateActionButton("🔍 高级查找替换", "find_replace_advanced"),
                CreateActionButton("ℹ️ 工作簿信息", "workbook_info", Color.FromArgb(108, 117, 125)),
                CreateActionButton("🧹 数据清理", "clean_data", Color.FromArgb(255, 193, 7))
            };
            AddButtonsToGroupBox(utilityGroupBox, utilityButtons);
            contentPanel.Controls.Add(utilityGroupBox);
        }

        private int CalculateGroupBoxHeight(int buttonCount)
        {
            // 计算GroupBox高度：标题栏(20) + 顶部边距(12) + 按钮高度*数量(32*count) + 按钮间距*(count-1)(6*(count-1)) + 底部边距(12)
            return 20 + 12 + (buttonCount * 32) + ((buttonCount - 1) * 6) + 12;
        }

        private void AddButtonsToGroupBox(GroupBox groupBox, Button[] buttons)
        {
            var containerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = false,
                Padding = new Padding(10, 8, 10, 8)
            };

            int yOffset = 0;
            foreach (var button in buttons)
            {
                button.Location = new Point(0, yOffset);
                button.Size = new Size(containerPanel.Width - 20, 32); // 容器宽度减去左右边距
                button.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
                containerPanel.Controls.Add(button);
                yOffset += 38; // 按钮高度(32) + 间距(6)
            }

            // 当容器大小改变时，更新按钮宽度
            containerPanel.SizeChanged += (s, e) =>
            {
                foreach (Button btn in containerPanel.Controls)
                {
                    btn.Width = containerPanel.Width - 20;
                }
            };

            groupBox.Controls.Add(containerPanel);
            
            // 设置GroupBox高度
            groupBox.Height = CalculateGroupBoxHeight(buttons.Length);
        }

        private GroupBox CreateGroupBox(string text, int y)
        {
            return new GroupBox
            {
                Text = text,
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(73, 80, 87),
                Location = new Point(8, y),
                Width = this.Width - 24, // 自适应宽度，留出左右边距
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.White,
                Padding = new Padding(6)
            };
        }

        private Button CreateActionButton(string text, string tag, Color? backColor = null)
        {
            var button = new Button
            {
                Text = text,
                Tag = tag,
                Font = new Font("Microsoft YaHei UI", 8.5F, FontStyle.Regular),
                BackColor = backColor ?? Color.FromArgb(0, 122, 204),
                ForeColor = backColor == Color.FromArgb(255, 193, 7) ? Color.FromArgb(33, 37, 41) : Color.White,
                FlatStyle = FlatStyle.Flat,
                Height = 32,
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 10, 0),
                Margin = new Padding(0),
                AutoEllipsis = true // 文本过长时显示省略号
            };
            
            button.FlatAppearance.BorderSize = 0;
            button.Click += ActionButton_Click;
            
            // 添加鼠标悬停效果
            var originalColor = button.BackColor;
            var hoverColor = GetHoverColor(originalColor);
            
            button.MouseEnter += (s, e) => button.BackColor = hoverColor;
            button.MouseLeave += (s, e) => button.BackColor = originalColor;
            
            return button;
        }

        private Color GetHoverColor(Color originalColor)
        {
            // 根据原始颜色计算悬停颜色
            if (originalColor == Color.FromArgb(255, 193, 7)) // 警告色
                return Color.FromArgb(230, 173, 6);
            else if (originalColor == Color.FromArgb(108, 117, 125)) // 次要色
                return Color.FromArgb(90, 98, 104);
            else if (originalColor == Color.FromArgb(40, 167, 69)) // 成功色
                return Color.FromArgb(34, 142, 58);
            else // 默认主色
                return Color.FromArgb(0, 86, 153);
        }

        private void CreateStatusPanel()
        {
            statusPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(227, 242, 253),
                Padding = new Padding(12, 8, 12, 8)
            };

            statusLabel = new Label
            {
                Text = "准备就绪 - 选择工作表区域后使用相关功能",
                Font = new Font("Microsoft YaHei UI", 8.5F),
                ForeColor = Color.FromArgb(66, 66, 66),
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };

            statusPanel.Controls.Add(statusLabel);
            mainTableLayout.Controls.Add(statusPanel, 0, 2);
        }

        #endregion

        #region 事件处理

        private void CloseButton_Click(object sender, EventArgs e)
        {
            try
            {
                var hostApp = HostApplication.Instance;
                if (hostApp != null)
                {
                    hostApp.CloseTaskPane("BasePluginDemo_WinFormsTaskPane");
                    _logger.Info("用户手动关闭 WinForms 任务窗格");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "关闭 WinForms 任务窗格时发生错误");
            }
        }

        private void RefreshButton_Click(object sender, EventArgs e)
        {
            RefreshInfo();
            UpdateStatus("信息已刷新");
        }

        private void ActionButton_Click(object sender, EventArgs e)
        {
            var button = sender as Button;
            var commandId = button?.Tag as string;

            if (string.IsNullOrEmpty(commandId))
            {
                UpdateStatus("错误：未找到命令标识", true);
                return;
            }

            ExecuteCommand(commandId);
        }

        #endregion

        #region 私有方法

        private void RefreshInfo()
        {
            try
            {
                if (ExcelApp == null)
                {
                    workbookValueLabel.Text = "Excel 未连接";
                    worksheetValueLabel.Text = "无";
                    selectionValueLabel.Text = "无";
                    cellCountValueLabel.Text = "0";
                    return;
                }

                var workbook = ExcelApp.ActiveWorkbook;
                var worksheet = ExcelApp.ActiveSheet as Excel.Worksheet;
                var selection = ExcelApp.Selection as Excel.Range;

                workbookValueLabel.Text = workbook?.Name ?? "无工作簿";
                worksheetValueLabel.Text = worksheet?.Name ?? "无工作表";

                if (selection != null)
                {
                    selectionValueLabel.Text = selection.Address;
                    cellCountValueLabel.Text = selection.Cells.Count.ToString();
                }
                else
                {
                    selectionValueLabel.Text = "无选择";
                    cellCountValueLabel.Text = "0";
                }

                _logger.Debug("WinForms 任务窗格信息已刷新");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "刷新 WinForms 任务窗格信息时发生错误");
                workbookValueLabel.Text = "错误";
                worksheetValueLabel.Text = "错误";
                selectionValueLabel.Text = "错误";
                cellCountValueLabel.Text = "0";
            }
        }

        private void ExecuteCommand(string commandId)
        {
            try
            {
                if (_featureManager == null)
                {
                    UpdateStatus("错误：功能管理器未初始化", true);
                    return;
                }

                _logger.Info("WinForms 任务窗格执行命令: {0}", commandId);
                UpdateStatus($"正在执行操作...");

                _featureManager.ExecuteCommand(commandId);

                RefreshInfo();
                UpdateStatus($"操作已完成");
                _logger.Info("WinForms 任务窗格命令执行成功: {0}", commandId);
            }
            catch (ArgumentException)
            {
                var errorMsg = $"未找到命令: {commandId}";
                UpdateStatus(errorMsg, true);
                _logger.Warning("WinForms 任务窗格命令未找到: {0}", commandId);

                MessageBox.Show($"功能 '{commandId}' 暂未实现或不可用。\n\n请检查插件配置。",
                    "功能不可用", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                var errorMsg = $"操作失败: {ex.Message}";
                UpdateStatus(errorMsg, true);
                _logger.Error(ex, "WinForms 任务窗格执行命令失败: {0}", commandId);

                MessageBox.Show($"执行操作时发生错误：\n\n{ex.Message}\n\n请查看日志文件了解详细信息。",
                    "操作失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateStatus(string message, bool isError = false)
        {
            try
            {
                if (statusLabel != null)
                {
                    statusLabel.Text = $"{DateTime.Now:HH:mm:ss} - {message}";
                    statusLabel.ForeColor = isError ? Color.Red : Color.FromArgb(66, 66, 66);
                }
            }
            catch (Exception ex)
            {
                _logger?.Error(ex, "更新 WinForms 任务窗格状态信息时发生错误");
            }
        }

        #endregion

        #region 清理资源

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try
                {
                    _featureManager?.Dispose();
                    _logger?.Info("WinForms TaskPaneControl 资源已清理");
                }
                catch (Exception ex)
                {
                    _logger?.Error(ex, "清理 WinForms TaskPaneControl 资源时发生错误");
                }
            }
            base.Dispose(disposing);
        }

        #endregion
    }
} 