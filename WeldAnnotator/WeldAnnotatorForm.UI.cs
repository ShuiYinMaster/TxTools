using System;
using System.Drawing;
using System.Windows.Forms;
using C1.Win.C1FlexGrid;
using Tecnomatix.Engineering.Ui;
using TxTools.Common;

namespace TxTools.WeldAnnotator
{
    public partial class WeldAnnotatorForm
    {
        // ════════════════════════════════════════════════════════════════
        //  UI 构建方法
        // ════════════════════════════════════════════════════════════════

        private void BuildCardPanel()
        {
            Panel rail = new Panel
            {
                Dock = DockStyle.Left,
                Width = 270,
                AutoScroll = true,
                BackColor = SystemColors.Control,
                Padding = new Padding(6, 4, 6, 4)
            };

            FlowLayoutPanel stack = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                BackColor = Color.Transparent
            };

            // 卡片①：操作节点（OP）
            GroupBox opCard = MakeRailCard("操作节点 (OP)");
            _opGrid = new TxObjGridCtrl
            {
                Dock = DockStyle.Top,
                Height = 110,
                ListenToPick = true,
                EnableMultipleSelection = false,
                EnableRecurringObjects = false
            };
            _opGrid.ObjectInserted += OnOpInserted;
            _opGrid.RowDeleted      += OnOpDeleted;
            _lblOpHint = new Label
            {
                Text = "在 PS 中选 OP 或视口选中",
                Dock = DockStyle.Bottom,
                Height = 18,
                ForeColor = Color.Gray,
                Font = new Font(_font.Name, 8f),
                TextAlign = ContentAlignment.MiddleCenter
            };
            opCard.Controls.Add(_lblOpHint);
            opCard.Controls.Add(_opGrid);

            // 卡片②：显示控制
            GroupBox snapCard = MakeRailCard("显示控制");
            _btnSnap     = MkRailBtn("拍摄快照",   BtnSnap_Click,    Color.FromArgb(0, 112, 192));
            _btnRestore  = MkRailBtn("恢复快照",   BtnRestore_Click, Color.FromArgb(84, 130, 53));
            _btnShowOnly = MkRailBtn("仅显示外观", BtnShowOnly_Click,Color.FromArgb(197, 90, 17));
            _btnShowAll  = MkRailBtn("显示全部",   BtnShowAll_Click, Color.FromArgb(100, 100, 100));
            _btnRestore.Enabled = false;
            _lblSnapStatus = new Label
            {
                Text = "", AutoSize = false, Height = 16, Width = 240,
                ForeColor = Color.Gray, Font = new Font(_font.Name, 8f),
                Margin = new Padding(0, 4, 0, 0), TextAlign = ContentAlignment.MiddleLeft
            };
            FillRailCardGrid(snapCard, 2,
                new Control[] { _btnSnap, _btnRestore, _btnShowOnly, _btnShowAll },
                new Control[] { _lblSnapStatus });

            // 卡片③：导出设置
            GroupBox cfgCard = MakeRailCard("导出设置");
            _chkNewSheet = new CheckBox
            {
                Text = "写入新 Sheet", Checked = false, AutoSize = true, Margin = new Padding(0, 2, 0, 2)
            };
            _chkWriteList = new CheckBox
            {
                Text = "附焊点数据表", Checked = false, AutoSize = true, Margin = new Padding(0, 2, 0, 2)
            };
            new ToolTip().SetToolTip(_chkWriteList,
                "勾选后，在截图下方附加焊点数据表（序号/名称/操作/XYZ/视口状态/类型）。");
            _chkAutoThick = new CheckBox
            {
                Text = "自动检测板厚", Checked = false, AutoSize = true, Margin = new Padding(0, 2, 0, 2)
            };
            new ToolTip().SetToolTip(_chkAutoThick,
                "按焊点绑定的不重复零件数自动设置类型：\n  2 个零件 → 二层板点焊\n  3+ 个零件 → 三层板及以上点焊\n已手动修改过类型的焊点不会被覆盖。");
            _chkAutoThick.CheckedChanged += (s, e) => ApplyAutoThicknessIfEnabled();

            Label lblMode = new Label
            {
                Text = "标注命名：", AutoSize = true, Margin = new Padding(0, 8, 0, 2),
                Font = new Font(_font.Name, 8.5f, FontStyle.Bold)
            };
            _cmbLabelMode = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList, Width = 200, Margin = new Padding(0, 0, 0, 2)
            };
            _cmbLabelMode.Items.AddRange(new object[] { "按序号", "按焊点名", "前缀+序号", "序号+后缀", "序号+焊点名" });
            _cmbLabelMode.SelectedIndex = 0;
            _cmbLabelMode.SelectedIndexChanged += (s, e) => UpdateLabelCardEnabled();

            FlowLayoutPanel pfxRow = new FlowLayoutPanel
            {
                AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight, WrapContents = false,
                Margin = new Padding(0, 0, 0, 2)
            };
            _lblPrefix = new Label { Text = "前缀", AutoSize = true, Margin = new Padding(0, 7, 2, 0) };
            _txtPrefix = new TextBox { Width = 60, Margin = new Padding(0, 4, 8, 0) };
            _lblSuffix = new Label { Text = "后缀", AutoSize = true, Margin = new Padding(0, 7, 2, 0) };
            _txtSuffix = new TextBox { Width = 60, Margin = new Padding(0, 4, 0, 0) };
            pfxRow.Controls.AddRange(new Control[] { _lblPrefix, _txtPrefix, _lblSuffix, _txtSuffix });

            FlowLayoutPanel cfgInner = new FlowLayoutPanel
            {
                Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(2, 2, 2, 2)
            };
            cfgInner.Controls.Add(_chkNewSheet);
            cfgInner.Controls.Add(_chkWriteList);
            cfgInner.Controls.Add(_chkAutoThick);
            cfgInner.Controls.Add(lblMode);
            cfgInner.Controls.Add(_cmbLabelMode);
            cfgInner.Controls.Add(pfxRow);
            cfgCard.Controls.Add(cfgInner);
            UpdateLabelCardEnabled();

            // 卡片④：操作按钮 + 状态
            GroupBox actCard = MakeRailCard("操作");
            _btnExport = MkRailBtnWide("导出到 Excel", BtnExport_Click, Color.FromArgb(0, 120, 215));
            _btnExport.Font = new Font(_font.Name, 10f, FontStyle.Bold);
            _btnExport.Height = 32;

            _btnClear = MkRailBtn("清空列表", BtnClear_Click, Color.FromArgb(183, 28, 28));
            _btnStyle = MkRailBtn("标注样式", BtnStyle_Click, Color.FromArgb(66, 66, 66));
            Button btnMinimize = MkRailBtn("最小化",
                (s, e) => WindowState = FormWindowState.Minimized, Color.FromArgb(120, 120, 120));
            _sharedToolTip.SetToolTip(btnMinimize, "最小化本窗口，以便在 PS 中调整视角、选择/隐藏对象。");

            _lblCount = new Label
            {
                Text = "焊点：0  视口内：0", AutoSize = true, Dock = DockStyle.Top,
                Margin = new Padding(0, 4, 0, 4),
                Font = new Font(_font.Name, 9f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 70, 127), TextAlign = ContentAlignment.MiddleLeft
            };

            FillRailCardGrid(actCard, 1,
                new Control[] { _btnExport },
                new Control[] { });

            TableLayoutPanel subGrid = new TableLayoutPanel
            {
                Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2, RowCount = 2, Margin = new Padding(0, 4, 0, 0)
            };
            subGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            subGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            subGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            subGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            subGrid.Controls.Add(_btnClear, 0, 0);
            subGrid.Controls.Add(_btnStyle, 1, 0);
            subGrid.Controls.Add(btnMinimize, 0, 1);

            actCard.Controls.Add(subGrid);
            actCard.Controls.Add(_lblCount);

            stack.Controls.Add(opCard);
            stack.Controls.Add(snapCard);
            stack.Controls.Add(cfgCard);
            stack.Controls.Add(actCard);

            rail.Controls.Add(stack);
            Controls.Add(rail);
        }

        // ── 卡片 / 按钮工厂 ──────────────────────────────────────────────

        private GroupBox MakeRailCard(string title)
        {
            return new FormUiKit.ColoredGroupBox
            {
                Text = title, Dock = DockStyle.Top, Width = 254,
                AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Font = new Font(_font.Name, 9f, FontStyle.Bold),
                HeaderColor = Color.FromArgb(0, 70, 127),
                ForeColor = Color.FromArgb(0, 70, 127),
                Padding = new Padding(6, 12, 6, 4),
                Margin = new Padding(0, 0, 0, 4),
                MinimumSize = new Size(254, 0)
            };
        }

        private Button MkRailBtn(string text, EventHandler h, Color fg)
        {
            var b = new FormUiKit.FlatColorButton
            {
                Text = text, Width = 116, Height = 26,
                BgColor = fg, ForeColor = Color.White, BorderColor = fg,
                Font = new Font(_font.Name, 9f, FontStyle.Regular),
                Margin = new Padding(0, 2, 4, 2), Padding = new Padding(4, 2, 4, 2),
                FlatStyle = FlatStyle.Flat
            };
            b.Click += h;
            return b;
        }

        private Button MkRailBtnWide(string text, EventHandler h, Color fg)
        {
            var b = new FormUiKit.FlatColorButton
            {
                Text = text, Width = 240, Height = 28,
                Dock = DockStyle.Top,
                BgColor = fg, ForeColor = Color.White, BorderColor = fg,
                Font = new Font(_font.Name, 10f, FontStyle.Bold),
                Padding = new Padding(4), FlatStyle = FlatStyle.Flat
            };
            b.Click += h;
            return b;
        }

        private void FillRailCardGrid(GroupBox card, int cols, Control[] buttons, Control[] extras)
        {
            var tlp = new TableLayoutPanel
            {
                Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = cols, RowCount = (buttons.Length + cols - 1) / cols
            };
            for (int c = 0; c < cols; c++)
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / cols));
            for (int i = 0; i < buttons.Length; i++)
                tlp.Controls.Add(buttons[i], i % cols, i / cols);
            card.Controls.Add(tlp);
            if (extras != null)
                foreach (var c in extras) card.Controls.Add(c);
        }

        private void UpdateLabelCardEnabled()
        {
            if (_cmbLabelMode == null) return;
            var m = (LabelNamingMode)_cmbLabelMode.SelectedIndex;
            bool pfx = m == LabelNamingMode.Prefix;
            bool sfx = m == LabelNamingMode.Suffix;
            _lblPrefix.Enabled = _txtPrefix.Enabled = pfx;
            _lblSuffix.Enabled = _txtSuffix.Enabled = sfx;
        }

        private void BuildMainArea()
        {
            Panel main = new Panel { Dock = DockStyle.Fill };

            _grid = new TxFlexGrid
            {
                Dock = DockStyle.Fill, AllowEditing = true, SelectionMode = SelectionModeEnum.Row
            };
            _grid.Rows.Fixed = 1;
            _grid.Cols.Count = 8;
            _grid.Rows.Count = 1;

            string[] hdrs  = { "#", "焊点名称", "操作名", "X(mm)", "Y(mm)", "Z(mm)", "视口", "类型" };
            int[]  widths = { 36, 140, 110, 78, 78, 78, 52, 92 };
            for (int c = 0; c < hdrs.Length; c++)
            {
                _grid[0, c] = hdrs[c];
                _grid.Cols[c].Width = widths[c];
                _grid.Cols[c].AllowSorting = false;
                _grid.Cols[c].AllowEditing = (c == C_TYPE);
            }
            _grid.Cols[C_TYPE].ComboList = string.Join("|", CATEGORY_OPTIONS);
            _grid.Cols[C_TYPE].TextAlign = TextAlignEnum.CenterCenter;

            _grid.Styles[CellStyleEnum.Fixed].BackColor    = Color.FromArgb(68, 114, 196);
            _grid.Styles[CellStyleEnum.Fixed].ForeColor    = Color.White;
            _grid.Styles[CellStyleEnum.Fixed].Font         = new Font(_font.Name, 9f, FontStyle.Bold);
            _grid.Styles[CellStyleEnum.Alternate].BackColor = Color.FromArgb(242, 242, 242);

            _grid.AfterEdit += Grid_AfterEdit;

            main.Controls.Add(_grid);
            Controls.Add(main);
        }

        private void BuildStatusBar()
        {
            _status = new StatusStrip();
            _lblStatus = new ToolStripStatusLabel("就绪")
            {
                Spring = true, TextAlign = ContentAlignment.MiddleLeft
            };
            _progress = new ToolStripProgressBar
            {
                Width = 150, Visible = false, Alignment = ToolStripItemAlignment.Right
            };
            var tsLog = new ToolStripButton("日志 ▲");
            tsLog.Click += (s, e) => ToggleLog(tsLog);
            _status.Items.AddRange(new ToolStripItem[] { _lblStatus, _progress, tsLog });
            Controls.Add(_status);
        }

        private void BuildLogPanel()
        {
            _logPanel = new Panel { Dock = DockStyle.Bottom, Height = 110, Visible = false };
            _logBox = new RichTextBox
            {
                Dock = DockStyle.Fill, BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.LightGray, Font = new Font("Consolas", 8f), ReadOnly = true
            };
            _logPanel.Controls.Add(_logBox);
            Controls.Add(_logPanel);
        }

        private void ToggleLog(ToolStripButton btn)
        {
            _logVisible = !_logVisible;
            _logPanel.Visible = _logVisible;
            btn.Text = _logVisible ? "日志 ▼" : "日志 ▲";
        }
    }
}
