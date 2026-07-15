using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;
using TxTools.Common;

namespace TxTools.LineToSolid
{
    /// <summary>
    /// 线生成实体 — 从 3D 曲线特征生成矩形截面或圆柱截面的几何体。
    ///
    /// GUI v3（本次重构）：
    ///   · 左列：特征列表卡（TxObjGridCtrl 原生拾取）+ 共享采样参数（曲线弦高）+ 段统计行。
    ///     弦高本质是"曲线离散采样精度"，与截面形状无关，原来矩形/圆柱卡各放一份
    ///     且实际只有圆柱那份被读取（矩形那份是死参数）——合并为一处，消除歧义。
    ///   · 右列：单张"截面参数"卡 + TabControl（矩形 / 圆柱），不再两卡常驻堆叠滚动；
    ///     底部一个生成按钮，文案随选项卡切换，点击按当前选项卡分发。
    ///   · 参数行全部基于 TableLayoutPanel（标签列 AutoSize + 内容列 100%），
    ///     替换原 FlowLayoutPanel + 固定宽 Label 的写法（AutoSize 与 Width 冲突，
    ///     高 DPI 下标签列对不齐）。
    ///   · 状态联动：拐弯半径仅在勾选"拐角过渡"时可编辑；间距/排列仅在根数 > 1 时可编辑。
    ///   · 底部日志卡保持可折叠。
    /// </summary>
    public class LineToSolidForm : TxForm
    {
        // ── 工具条 ──────────────────────────────────────────────────────
        private TxToolStrip _toolStrip;
        private ToolStripButton _btnRefresh;
        private ToolStripButton _btnToggleLog;

        // ── 左列：基线特征列表 + 共享采样参数 ───────────────────────────
        private TxObjGridCtrl _featureGrid;
        private Label _lblSegInfo;
        private NumericUpDown _numSagitta;     // 共享曲线弦高（采样精度）

        // ── 右列：截面参数选项卡 ────────────────────────────────────────
        private TabControl _tabs;
        private TabPage _tabRect;
        private TabPage _tabCyl;
        private FormUiKit.FlatColorButton _btnBuild;

        // 矩形参数
        private NumericUpDown _numWidth;
        private NumericUpDown _numHeight;
        private NumericUpDown _numRectOffsetX;
        private NumericUpDown _numRectOffsetY;
        private CheckBox _chkRectOnGround;
        private TextBox _txtRectPrefix;

        // 圆柱参数
        private NumericUpDown _numDiameter;
        private NumericUpDown _numPipeCount;
        private NumericUpDown _numPipeSpacing;
        private ComboBox _cmbPipeLayout;
        private NumericUpDown _numCylOffsetX;
        private NumericUpDown _numCylOffsetY;
        private CheckBox _chkCylOnGround;
        private CheckBox _chkRoundCorner;
        private NumericUpDown _numBendRadius;
        private TextBox _txtCylPrefix;

        // ── 日志 ────────────────────────────────────────────────────────
        private Panel _logPanel;
        private RichTextBox _logBox;
        private bool _logCollapsed;

        // ── 数据 ────────────────────────────────────────────────────────
        private List<LineSegment> _currentSegments = new List<LineSegment>();

        // ── DPI ─────────────────────────────────────────────────────────
        private static readonly Size _designSize = new Size(720, 660);
        private bool _dpiApplied;

        // =================================================================
        // 构造 + 生命周期
        // =================================================================
        public LineToSolidForm()
        {
            SemiModal = false;
            FormUiKit.InitStandardForm(this, "线生成实体 - TxTools.LineToSolid",
                _designSize, new Size(620, 500));
            this.Padding = new Padding(4);

            BuildToolStrip();
            BuildBody();
            BuildLogPanel();




            Log("插件已加载。在场景中选中曲线特征（Polyline/Line/Curve）或资源/零件容器，");
            Log("然后在网格中用 PS 标准拾取方式添加，选择右侧截面选项卡后点 [生成几何体]。");
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            FormUiKit.ApplyDpiScaling(this, ref _dpiApplied, _designSize);

            // TxObjGridCtrl 首次拾取需要先获得焦点（ExportGunForm 同款）
            BeginInvoke(new Action(() =>
            {
                try
                {
                    if (_featureGrid != null && _featureGrid.Visible)
                    {
                        _featureGrid.Focus();
                        try { _featureGrid.SetCurrentCell(0, 0); } catch { }
                    }
                }
                catch { }
            }));
        }

        // =================================================================
        // UI — 工具条
        // =================================================================
        private void BuildToolStrip()
        {
            _toolStrip = new TxToolStrip
            {
                Dock = DockStyle.Top,
                GripStyle = ToolStripGripStyle.Hidden
            };

            _btnRefresh = new ToolStripButton("刷新段信息");
            _btnRefresh.Click += (s, e) => RefreshSegmentInfo();

            _btnToggleLog = new ToolStripButton("折叠日志");
            _btnToggleLog.Click += (s, e) => ToggleLog();

            _toolStrip.Items.Add(_btnRefresh);
            _toolStrip.Items.Add(new ToolStripSeparator());
            _toolStrip.Items.Add(_btnToggleLog);
            this.Controls.Add(_toolStrip);
        }

        // =================================================================
        // UI — 主体：左右分栏
        // =================================================================
        private void BuildBody()
        {
            var body = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0, 2, 0, 2),
                Padding = Padding.Empty
            };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 340F));
            body.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            body.Controls.Add(BuildLeftColumn(), 0, 0);
            body.Controls.Add(BuildRightColumn(), 1, 0);
            this.Controls.Add(body);
        }

        // ── 左列：特征列表 + 共享采样参数 + 段信息 ────────────────────
        private Control BuildLeftColumn()
        {
            var col = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 0, 4, 0) };

            var card = new FormUiKit.ColoredGroupBox
            {
                Text = "基线特征列表（PS 原生拾取）",
                Dock = DockStyle.Fill,
                HeaderColor = Color.FromArgb(60, 90, 150),
                ForeColor = Color.FromArgb(60, 90, 150),
                Padding = new Padding(8, 18, 8, 8)
            };

            _featureGrid = new TxObjGridCtrl { Dock = DockStyle.Fill };
            try { _featureGrid.ListenToPick = true; } catch { }
            try { _featureGrid.EnableMultipleSelection = true; } catch { }
            try { _featureGrid.EnableRecurringObjects = false; } catch { }
            try { _featureGrid.ObjectInserted += (s, e) => RefreshSegmentInfo(); } catch { }
            try { _featureGrid.RowDeleted += (s, e) => RefreshSegmentInfo(); } catch { }

            // 底部：曲线弦高（共享采样参数）+ 段统计
            var bottom = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 3,
                Padding = new Padding(0, 4, 0, 0)
            };
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            var lblSag = new Label
            {
                Text = "曲线弦高（mm）",
                AutoSize = true,
                Font = FormUiKit.BaseFont,
                ForeColor = SystemColors.GrayText,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 4, 6, 0)
            };
            _numSagitta = MkNud(0.5m, 0.001m, 100m, 3);
            _numSagitta.Width = 90;
            _numSagitta.Anchor = AnchorStyles.Left;
            _numSagitta.ValueChanged += (s, e) => RefreshSegmentInfo();

            var lblSagHint = new Label
            {
                Text = "（曲线离散采样精度，越小越精细，对矩形/圆柱均生效）",
                AutoSize = true,
                Font = FormUiKit.BaseFont,
                ForeColor = SystemColors.GrayText,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(6, 4, 0, 0)
            };

            _lblSegInfo = new Label
            {
                Text = "特征数：0    段数：-    总长度：-",
                AutoSize = true,
                Font = FormUiKit.BaseFont,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 6, 0, 0)
            };

            bottom.RowCount = 2;
            bottom.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            bottom.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            bottom.Controls.Add(lblSag, 0, 0);
            bottom.Controls.Add(_numSagitta, 1, 0);
            bottom.Controls.Add(lblSagHint, 2, 0);
            bottom.Controls.Add(_lblSegInfo, 0, 1);
            bottom.SetColumnSpan(_lblSegInfo, 3);

            card.Controls.Add(_featureGrid);
            card.Controls.Add(bottom);
            col.Controls.Add(card);
            return col;
        }

        // ── 右列：截面参数卡（TabControl + 生成按钮）──────────────────
        private Control BuildRightColumn()
        {
            var card = new FormUiKit.ColoredGroupBox
            {
                Text = "截面参数",
                Dock = DockStyle.Fill,
                HeaderColor = Color.FromArgb(60, 90, 150),
                ForeColor = Color.FromArgb(60, 90, 150),
                Padding = new Padding(8, 18, 8, 8)
            };

            _tabs = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = FormUiKit.BaseFont
            };

            _tabRect = new TabPage("矩形截面")
            {
                Padding = new Padding(6),
                UseVisualStyleBackColor = true
            };
            _tabRect.Controls.Add(BuildRectTable());

            _tabCyl = new TabPage("圆柱截面")
            {
                Padding = new Padding(6),
                UseVisualStyleBackColor = true
            };
            _tabCyl.Controls.Add(BuildCylTable());

            _tabs.TabPages.Add(_tabRect);
            _tabs.TabPages.Add(_tabCyl);
            _tabs.SelectedIndexChanged += (s, e) => UpdateBuildButtonText();

            // 底部生成按钮（随选项卡切换文案，点击按选项卡分发）
            var btnHost = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 42,
                Padding = new Padding(0, 8, 0, 0)
            };
            _btnBuild = new FormUiKit.FlatColorButton
            {
                Dock = DockStyle.Fill,
                FlatStyle = FlatStyle.Flat,
                BgColor = Color.FromArgb(0, 100, 167),
                BorderColor = Color.FromArgb(0, 100, 167),
                ForeColor = Color.White,
                Font = FormUiKit.BaseFont
            };
            _btnBuild.Click += (s, e) => OnBuild();
            btnHost.Controls.Add(_btnBuild);

            card.Controls.Add(_tabs);
            card.Controls.Add(btnHost);
            UpdateBuildButtonText();
            return card;
        }

        private void UpdateBuildButtonText()
        {
            if (_btnBuild == null || _tabs == null) return;
            _btnBuild.Text = (_tabs.SelectedTab == _tabCyl)
                ? "生成圆柱几何体"
                : "生成矩形几何体";
        }

        // =================================================================
        // 参数行工厂（TableLayoutPanel：标签列 AutoSize + 内容列 100%）
        // =================================================================
        private static TableLayoutPanel MkParamTable()
        {
            var t = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Padding = new Padding(0, 4, 0, 0),
                BackColor = Color.Transparent
            };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            return t;
        }

        private static void AddRow(TableLayoutPanel t, string label, Control ctrl)
        {
            int r = t.RowCount;
            t.RowCount = r + 1;
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var lbl = new Label
            {
                Text = label,
                AutoSize = true,
                Font = FormUiKit.BaseFont,
                ForeColor = SystemColors.GrayText,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 6, 8, 2)
            };
            ctrl.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            ctrl.Margin = new Padding(0, 2, 0, 2);

            t.Controls.Add(lbl, 0, r);
            t.Controls.Add(ctrl, 1, r);
        }

        /// <summary>整行控件（如 CheckBox），跨两列。</summary>
        private static void AddSpanRow(TableLayoutPanel t, Control ctrl)
        {
            int r = t.RowCount;
            t.RowCount = r + 1;
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            ctrl.Anchor = AnchorStyles.Left;
            ctrl.Margin = new Padding(0, 4, 0, 2);
            t.Controls.Add(ctrl, 0, r);
            t.SetColumnSpan(ctrl, 2);
        }

        private NumericUpDown MkNud(decimal val, decimal min, decimal max, int dec)
        {
            return new NumericUpDown
            {
                Minimum = min,
                Maximum = max,
                DecimalPlaces = dec,
                Value = val,
                Font = FormUiKit.BaseFont
            };
        }

        // =================================================================
        // 矩形截面选项卡
        // =================================================================
        private Control BuildRectTable()
        {
            var t = MkParamTable();

            _numWidth = MkNud(50, 0.01m, 100000m, 2);
            AddRow(t, "宽（mm）", _numWidth);

            _numHeight = MkNud(50, 0.01m, 100000m, 2);
            _numHeight.ValueChanged += (s, e) =>
            {
                if (_chkRectOnGround.Checked) _numRectOffsetY.Value = _numHeight.Value / 2m;
            };
            AddRow(t, "高（mm）", _numHeight);

            _numRectOffsetX = MkNud(0, -100000m, 100000m, 2);
            AddRow(t, "局部X偏移", _numRectOffsetX);

            _numRectOffsetY = MkNud(0, -100000m, 100000m, 2);
            AddRow(t, "局部Y偏移", _numRectOffsetY);

            _chkRectOnGround = new CheckBox
            {
                Text = "底面贴地（自动按高/2 设 Y 偏移）",
                AutoSize = true,
                Font = FormUiKit.BaseFont
            };
            _chkRectOnGround.CheckedChanged += (s, e) =>
            {
                if (_chkRectOnGround.Checked) _numRectOffsetY.Value = _numHeight.Value / 2m;
            };
            AddSpanRow(t, _chkRectOnGround);

            _txtRectPrefix = new TextBox { Text = "LTS_Rect", Font = FormUiKit.BaseFont };
            AddRow(t, "Resource 前缀", _txtRectPrefix);

            return t;
        }

        // =================================================================
        // 圆柱截面选项卡
        // =================================================================
        private Control BuildCylTable()
        {
            var t = MkParamTable();

            _numDiameter = MkNud(30, 0.01m, 100000m, 2);
            _numDiameter.ValueChanged += (s, e) =>
            {
                if (_chkCylOnGround.Checked) _numCylOffsetY.Value = _numDiameter.Value / 2m;
            };
            AddRow(t, "直径（mm）", _numDiameter);

            _numPipeCount = MkNud(1, 1m, 1000m, 0);
            _numPipeCount.ValueChanged += (s, e) => UpdateCylEnabledState();
            AddRow(t, "圆柱根数", _numPipeCount);

            _numPipeSpacing = MkNud(50, 0m, 100000m, 2);
            AddRow(t, "圆柱间距（mm）", _numPipeSpacing);

            _cmbPipeLayout = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = FormUiKit.BaseFont
            };
            _cmbPipeLayout.Items.Add("对称中线");
            _cmbPipeLayout.Items.Add("起点偏移");
            _cmbPipeLayout.SelectedIndex = 0;
            AddRow(t, "排列方式", _cmbPipeLayout);

            _numCylOffsetX = MkNud(0, -100000m, 100000m, 2);
            AddRow(t, "局部X偏移", _numCylOffsetX);

            _numCylOffsetY = MkNud(0, -100000m, 100000m, 2);
            AddRow(t, "局部Y偏移", _numCylOffsetY);

            _chkCylOnGround = new CheckBox
            {
                Text = "底面贴地（自动按直径/2 设 Y 偏移）",
                AutoSize = true,
                Font = FormUiKit.BaseFont
            };
            _chkCylOnGround.CheckedChanged += (s, e) =>
            {
                if (_chkCylOnGround.Checked) _numCylOffsetY.Value = _numDiameter.Value / 2m;
            };
            AddSpanRow(t, _chkCylOnGround);

            _chkRoundCorner = new CheckBox
            {
                Text = "拐角过渡（单根球体 / 多根圆环）",
                AutoSize = true,
                Font = FormUiKit.BaseFont
            };
            _chkRoundCorner.CheckedChanged += (s, e) => UpdateCylEnabledState();
            AddSpanRow(t, _chkRoundCorner);

            _numBendRadius = MkNud(0, 0m, 100000m, 2);
            AddRow(t, "拐弯半径（mm）", _numBendRadius);

            _txtCylPrefix = new TextBox { Text = "LTS_Cyl", Font = FormUiKit.BaseFont };
            AddRow(t, "Resource 前缀", _txtCylPrefix);

            UpdateCylEnabledState();
            return t;
        }

        /// <summary>圆柱选项卡内的状态联动：根数 > 1 才有间距/排列；勾了拐角过渡才有拐弯半径。</summary>
        private void UpdateCylEnabledState()
        {
            bool multi = _numPipeCount != null && _numPipeCount.Value > 1;
            if (_numPipeSpacing != null) _numPipeSpacing.Enabled = multi;
            if (_cmbPipeLayout != null) _cmbPipeLayout.Enabled = multi;
            if (_numBendRadius != null && _chkRoundCorner != null)
                _numBendRadius.Enabled = _chkRoundCorner.Checked;
        }

        // =================================================================
        // 日志面板（可折叠）
        // =================================================================
        private void BuildLogPanel()
        {
            _logPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 120,
                Padding = new Padding(0, 2, 0, 0)
            };

            _logBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 8f),
                BackColor = Color.FromArgb(30, 30, 30),
                ForeColor = Color.LightGray,
                BorderStyle = BorderStyle.None,
                WordWrap = false,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };
            _logPanel.Controls.Add(_logBox);
            this.Controls.Add(_logPanel);
        }

        private void ToggleLog()
        {
            _logCollapsed = !_logCollapsed;
            _logPanel.Visible = !_logCollapsed;
            _btnToggleLog.Text = _logCollapsed ? "展开日志" : "折叠日志";
        }

        // =================================================================
        // 段信息刷新
        // =================================================================
        private List<ITxObject> GetGridObjects()
        {
            var list = new List<ITxObject>();
            try
            {
                int count = 0;
                try { count = _featureGrid.Count; } catch { }
                for (int i = 0; i < count; i++)
                {
                    try
                    {
                        var obj = _featureGrid.GetObject(i);
                        if (obj != null) list.Add(obj);
                    }
                    catch { }
                }
            }
            catch { }
            return list;
        }

        private void RefreshSegmentInfo()
        {
            _currentSegments.Clear();
            var features = GetGridObjects();

            if (features.Count == 0)
            {
                _lblSegInfo.Text = "特征数：0    段数：-    总长度：-";
                return;
            }

            double sagitta = 0.5;
            try { sagitta = (double)_numSagitta.Value; } catch { }

            var opts = new PolylineReader.ReadOptions { MaxSagitta = sagitta };
            var curveFeats = ExpandContainersToCurves(features);
            var res = PolylineReader.ExtractAll(curveFeats, opts);
            foreach (var m in res.Diagnostics) Log(m);

            _currentSegments = res.Segments;
            double total = 0;
            foreach (var s in res.Segments) total += s.Length;
            _lblSegInfo.Text = string.Format("特征数：{0}    段数：{1}    总长度：{2:F3} mm",
                res.FeatureCount, res.Segments.Count, total);
        }

        private List<ITxObject> ExpandContainersToCurves(List<ITxObject> input)
        {
            var result = new List<ITxObject>();
            foreach (var obj in input)
            {
                if (obj == null) continue;
                var kind = PolylineReader.ClassifyFeature(obj);
                if (kind != FeatureKind.Unknown)
                {
                    if (!result.Contains(obj)) result.Add(obj);
                }
                else
                {
                    PolylineReader.CollectCurvesRecursive(obj, result);
                }
            }
            return result;
        }

        // =================================================================
        // 生成
        // =================================================================
        private GeometryParams CollectRectParams()
        {
            return new GeometryParams
            {
                Section = CrossSectionType.Rectangle,
                Width = (double)_numWidth.Value,
                Height = (double)_numHeight.Value,
                OffsetX = (double)_numRectOffsetX.Value,
                OffsetY = (double)_numRectOffsetY.Value,
                PartNamePrefix = string.IsNullOrEmpty(_txtRectPrefix.Text)
                    ? "LTS_Rect" : _txtRectPrefix.Text.Trim()
            };
        }

        private GeometryParams CollectCylParams()
        {
            return new GeometryParams
            {
                Section = CrossSectionType.Circle,
                Diameter = (double)_numDiameter.Value,
                PipeCount = (int)_numPipeCount.Value,
                PipeSpacing = (double)_numPipeSpacing.Value,
                PipeLayout = (_cmbPipeLayout.SelectedIndex == 0)
                    ? MultiPipeLayout.SymmetricCentered
                    : MultiPipeLayout.OffsetFromStart,
                OffsetX = (double)_numCylOffsetX.Value,
                OffsetY = (double)_numCylOffsetY.Value,
                RoundCornerForCylinder = _chkRoundCorner.Checked,
                BendRadius = (double)_numBendRadius.Value,
                PartNamePrefix = string.IsNullOrEmpty(_txtCylPrefix.Text)
                    ? "LTS_Cyl" : _txtCylPrefix.Text.Trim()
            };
        }

        /// <summary>按当前选项卡分发生成。</summary>
        private void OnBuild()
        {
            if (_tabs.SelectedTab == _tabCyl) OnBuildCyl();
            else OnBuildRect();
        }

        private void OnBuildRect()
        {
            RefreshSegmentInfo();
            if (_currentSegments == null || _currentSegments.Count == 0)
            {
                Log("[警告] 没有可用段数据。");
                return;
            }
            var p = CollectRectParams();
            Log(string.Format("[开始] 矩形 {0}x{1}，共 {2} 段",
                p.Width, p.Height, _currentSegments.Count));
            var result = GeometryBuilder.BuildForSegments(_currentSegments, p);
            ReportResult(result);
        }

        private void OnBuildCyl()
        {
            RefreshSegmentInfo();
            if (_currentSegments == null || _currentSegments.Count == 0)
            {
                Log("[警告] 没有可用段数据。");
                return;
            }
            var p = CollectCylParams();
            string desc = string.Format("圆柱 D={0}", p.Diameter);
            if (p.PipeCount > 1) desc += string.Format("，{0}根@{1}", p.PipeCount, p.PipeSpacing);
            if (p.RoundCornerForCylinder) desc += "，拐角过渡";
            Log(string.Format("[开始] {0}，共 {1} 段", desc, _currentSegments.Count));
            var result = GeometryBuilder.BuildForSegments(_currentSegments, p);
            ReportResult(result);
        }

        private void ReportResult(GeometryBuilder.BuildResult result)
        {
            Log(string.Format("[完成] 成功 {0}，失败 {1}", result.SuccessCount, result.FailCount));
            foreach (var msg in result.Messages) Log("  · " + msg);
        }

        // =================================================================
        // 工具
        // =================================================================
        private void Log(string msg)
        {
            try
            {
                _logBox.AppendText(string.Format("[{0:HH:mm:ss}] {1}{2}",
                    DateTime.Now, msg, Environment.NewLine));
            }
            catch { }
        }
    }
}
