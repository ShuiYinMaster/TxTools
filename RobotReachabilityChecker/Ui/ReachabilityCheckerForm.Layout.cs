// ============================================================================
// ReachabilityCheckerForm.Layout.cs
//
// UI 构建：5 张并列卡片 + ToolStrip + LogPanel + StatusStrip。
// 拆出来便于后续调整布局而不影响业务逻辑。
//
// 布局结构：
//   ┌─ TxToolStrip（仅检查时显示进度条，默认隐藏）
//   ├─ Panel(_cardsPanel)（5 张并列卡片：检查目标 / TCP / 轴角 / 干涉 / 功能区）
//   ├─ TxFlexGrid（结果表，Fill）
//   └─ StatusStrip + 日志面板（底部，可折叠）
// ============================================================================
using System;
using System.Drawing;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;
using Tecnomatix.Engineering.Ui.WPF;
using TxTools.Common;
using static TxTools.RobotReachabilityChecker.Ui.Theme;

namespace TxTools.RobotReachabilityChecker.Ui
{
    public partial class ReachabilityCheckerForm
    {
        // =====================================================================
        // UI 总入口
        // =====================================================================
        private void InitializeComponent()
        {
            SuspendLayout();
            BuildToolStrip();
            BuildCardsPanel();
            BuildLogPanel();
            BuildStatusStrip();
            BuildGrid();   // 在 Grid 文件中实现

            // DockStyle 优先级：后 Add 的 Top 显示在上；Fill 最后
            _toolStrip.SendToBack();
            _cardsPanel.SendToBack();
            _statusStrip.SendToBack();
            _logPanel.SendToBack();
            _grid.BringToFront();
            ResumeLayout(false);
        }

        // =====================================================================
        // ToolStrip（精简：仅保留进度条，检查时显示）
        // =====================================================================
        private void BuildToolStrip()
        {
            _toolStrip = new TxToolStrip
            {
                Dock = DockStyle.Top,
                GripStyle = ToolStripGripStyle.Hidden,
                Padding = new Padding(4, 2, 4, 2),
                Visible = false
            };

            _tsProgress = new ToolStripProgressBar { Width = 200, Visible = false };
            _toolStrip.Items.Add(_tsProgress);
            Controls.Add(_toolStrip);
        }

        // =====================================================================
        // 5 张卡片 — 水平等分顶部，自适应高度
        // =====================================================================
        private void BuildCardsPanel()
        {
            _cardsPanel = new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = SystemColors.Control,
                Padding = new Padding(3, 3, 3, 3)
            };

            var table = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top,
                ColumnCount = 5,
                RowCount = 1,
                BackColor = Color.Transparent,
                Padding = new Padding(0)
            };
            for (int i = 0; i < 5; i++)
                table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20f));
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            table.Controls.Add(BuildCard1_Target(),       0, 0);
            table.Controls.Add(BuildCard2_TcpMargin(),    1, 0);
            table.Controls.Add(BuildCard3_JointMargin(),  2, 0);
            table.Controls.Add(BuildCard4_Interference(), 3, 0);
            table.Controls.Add(BuildCard5_Functions(),    4, 0);

            _cardsPanel.Controls.Add(table);
            Controls.Add(_cardsPanel);
        }

        // — 卡片容器（ColoredGroupBox 自绘彩色标题）
        private GroupBox MkCard(string title) => new FormUiKit.ColoredGroupBox
        {
            Text = title,
            Dock = DockStyle.Fill,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold),
            HeaderColor = ClrAccent,
            ForeColor = ClrAccent,
            Margin = new Padding(2, 2, 2, 2),
            Padding = new Padding(8, 6, 8, 4)
        };

        // 卡片内容 Panel —— Dock=Top 用 DisplayRectangle 自动避开标题区域
        private FlowLayoutPanel MkCardContent() => new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            BackColor = Color.Transparent,
            Font = SystemFonts.DefaultFont,
            Padding = new Padding(0, 2, 0, 0)
        };

        // ---------------------------------------------------------------------
        // Card1: 检查目标及筛选
        // ---------------------------------------------------------------------
        private GroupBox BuildCard1_Target()
        {
            var card = MkCard("1. 检查目标及筛选");
            var flow = MkCardContent();

            // 行1：OP节点 拾取器
            var row1 = MkInlineRow();
            row1.Controls.Add(MkLabel("OP节点"));
            _txtOpNode = new TxObjEditBoxCtrl
            {
                Width = 120, Height = 22, Font = SystemFonts.DefaultFont,
                Margin = new Padding(2, 1, 0, 0),
                PickOnly = true, ListenToPick = true
            };
            _txtOpNode.Picked += OnOpNodePicked;
            row1.Controls.Add(_txtOpNode);
            flow.Controls.Add(row1);

            // 行2：点筛选
            var row2 = MkInlineRow();
            row2.Controls.Add(MkLabel("点筛选"));
            _cbPointTypeFilter = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            _cbPointTypeFilter.Items.AddRange(new object[] { "所有类型", "仅焊点", "仅Via" });
            _cbPointTypeFilter.SelectedIndex = 0;
            AutoFitComboBoxWidth(_cbPointTypeFilter);
            _cbPointTypeFilter.SelectedIndexChanged += (s, e) => ApplyFilterNow();
            row2.Controls.Add(_cbPointTypeFilter);
            flow.Controls.Add(row2);

            // 行3：品牌
            var row3 = MkInlineRow();
            row3.Controls.Add(MkLabel("品牌"));
            _cbBrand = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            _cbBrand.Items.AddRange(new object[] { "自动", "KUKA", "ABB", "FANUC", "其他" });
            _cbBrand.SelectedIndex = 0;
            AutoFitComboBoxWidth(_cbBrand);
            row3.Controls.Add(_cbBrand);
            flow.Controls.Add(row3);

            // 行4：隐藏正常结果项
            _chkHideNormal = new CheckBox
            {
                Text = "隐藏正常结果项",
                AutoSize = true, Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 0, 0, 0)
            };
            _chkHideNormal.CheckedChanged += (s, e) => ApplyFilterNow();
            flow.Controls.Add(_chkHideNormal);

            card.Controls.Add(flow);
            return card;
        }

        // ---------------------------------------------------------------------
        // Card2: TCP余量
        // ---------------------------------------------------------------------
        private GroupBox BuildCard2_TcpMargin()
        {
            var card = MkCard("2. TCP余量检查");
            var flow = MkCardContent();

            _chkTcpXyz = new CheckBox
            {
                Text = "点位XYZ余量检查",
                Checked = true, AutoSize = true,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 2, 0, 4)
            };
            flow.Controls.Add(_chkTcpXyz);

            var row = MkInlineRow();
            row.Controls.Add(MkLabel("余量(mm):"));
            _nudTcpMargin = new NumericUpDown
            {
                Minimum = 0, Maximum = 9999, Value = 200, DecimalPlaces = 0,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            AutoFitNumericWidth(_nudTcpMargin);
            row.Controls.Add(_nudTcpMargin);
            flow.Controls.Add(row);

            card.Controls.Add(flow);
            return card;
        }

        // ---------------------------------------------------------------------
        // Card3: 轴角度余量
        // ---------------------------------------------------------------------
        private GroupBox BuildCard3_JointMargin()
        {
            var card = MkCard("3. 轴角度余量检查");
            var flow = MkCardContent();

            _chkJointMargin = new CheckBox
            {
                Text = "各轴软限位余量检查",
                Checked = true, AutoSize = true,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 2, 0, 4)
            };
            flow.Controls.Add(_chkJointMargin);

            var row = MkInlineRow();
            row.Controls.Add(MkLabel("余量(度):"));
            _nudJointMarginDeg = new NumericUpDown
            {
                Minimum = 0, Maximum = 180, Value = 10, DecimalPlaces = 0,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            AutoFitNumericWidth(_nudJointMarginDeg);
            row.Controls.Add(_nudJointMarginDeg);
            flow.Controls.Add(row);

            card.Controls.Add(flow);
            return card;
        }

        // ---------------------------------------------------------------------
        // Card4: 干涉检查（仅展示）
        // ---------------------------------------------------------------------
        private GroupBox BuildCard4_Interference()
        {
            var card = MkCard("4. 点位干涉检查");
            var flow = MkCardContent();

            _chkStaticInterference = new CheckBox
            {
                Text = "启用静态干涉检查",
                AutoSize = true, Enabled = true,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 2, 0, 4)
            };
            _chkDynamicInterference = new CheckBox
            {
                Text = "启用动态干涉检查",
                AutoSize = true, Enabled = false,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 0, 0, 0)
            };
            flow.Controls.AddRange(new Control[] { _chkStaticInterference, _chkDynamicInterference });

            card.Controls.Add(flow);
            return card;
        }

        // ---------------------------------------------------------------------
        // Card5: 功能区
        // ---------------------------------------------------------------------
        private GroupBox BuildCard5_Functions()
        {
            var card = MkCard("5. 功能区");
            var outerFlow = MkCardContent();

            // 第一行：开始检查 + 检查所有路径
            var row1 = MkButtonRow();
            var btnCheck = MkFuncButton("开始检查", TxClrBtnCheck.Color);
            btnCheck.Click += BtnCheck_Click;
            var btnAll = MkFuncButton("检查所有路径", TxClrBtnAll.Color);
            btnAll.Click += BtnCheckAll_Click;
            row1.Controls.AddRange(new Control[] { btnCheck, btnAll });

            // 第二行：导出 + 重置 + OLP诊断 + 关闭
            var row2 = MkButtonRow();
            var btnExport = MkFuncButton("结果导出", TxClrBtnExport.Color);
            btnExport.Click += BtnExport_Click;
            var btnReset = MkFuncButton("重置窗口", TxClrBtnReset.Color);
            btnReset.Click += BtnReset_Click;
            var btnClose = MkFuncButton("关闭", TxClrBtnClose.Color);
            btnClose.Click += (s, e) => Close();
            row2.Controls.AddRange(new Control[] { btnExport, btnReset, btnClose });

            outerFlow.Controls.Add(row1);
            outerFlow.Controls.Add(row2);
            card.Controls.Add(outerFlow);
            return card;
        }

        private FlowLayoutPanel MkInlineRow() => new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            BackColor = Color.Transparent,
            Margin = new Padding(0, 2, 0, 2)
        };

        private FlowLayoutPanel MkButtonRow() => new FlowLayoutPanel
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            BackColor = Color.Transparent,
            Margin = new Padding(0, 0, 0, 2)
        };

        // =====================================================================
        // 日志面板
        // =====================================================================
        private void BuildLogPanel()
        {
            _logPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 140,
                BackColor = SystemColors.Control,
                BorderStyle = BorderStyle.FixedSingle,
                Visible = false
            };
            var hdr = new Panel { Dock = DockStyle.Top, Height = 20, BackColor = ClrAccent };
            var lblHdr = new Label
            {
                Text = "运行日志",
                Dock = DockStyle.Fill,
                ForeColor = TxColor.TxColorWhite.Color,
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(6, 0, 0, 0)
            };
            var btnClear = new FormUiKit.FlatColorButton
            {
                Text = "清空", Dock = DockStyle.Right, Width = 44, Height = 26,
                FlatStyle = FlatStyle.Flat,
                BgColor = ClrAccent,
                ForeColor = Color.White,
                BorderColor = ClrAccent,
                Font = SystemFonts.DefaultFont
            };
            btnClear.Click += (s, e) => _logBox?.Clear();
            hdr.Controls.AddRange(new Control[] { btnClear, lblHdr });

            _logBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = TxClrLogBg.Color,
                ForeColor = TxClrLogText.Color,
                Font = new Font("Consolas", 8f),
                BorderStyle = BorderStyle.None,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                WordWrap = false
            };
            _logPanel.Controls.Add(_logBox);
            _logPanel.Controls.Add(hdr);
            Controls.Add(_logPanel);
        }

        private void ToggleLogPanel()
        {
            if (_logPanel == null) return;
            _logVisible = !_logVisible;
            _logPanel.Visible = _logVisible;
            if (_btnFooterLog != null) _btnFooterLog.Text = _logVisible ? "▼ 隐藏日志" : "▲ 显示日志";
        }

        // =====================================================================
        // StatusStrip
        // =====================================================================
        private void BuildStatusStrip()
        {
            _statusStrip = new StatusStrip { SizingGrip = true };
            _lblStatus = new ToolStripStatusLabel("就绪")
            {
                TextAlign = ContentAlignment.MiddleLeft,
                Spring = true
            };

            var tsLogBtn = new ToolStripButton("▲ 显示日志")
            {
                ToolTipText = "显示/隐藏底部日志面板",
                Alignment = ToolStripItemAlignment.Right,
                DisplayStyle = ToolStripItemDisplayStyle.Text,
                AutoSize = true
            };
            tsLogBtn.Click += (s, e) => ToggleLogPanel();

            _statusStrip.Items.Add(_lblStatus);
            _statusStrip.Items.Add(tsLogBtn);

            // 维护 Button 类型的引用以便 ToggleLogPanel 更新文字
            _btnFooterLog = new Button();
            _btnFooterLog.TextChanged += (s, e) => { try { tsLogBtn.Text = _btnFooterLog.Text; } catch { } };

            Controls.Add(_statusStrip);
        }
    }
}
