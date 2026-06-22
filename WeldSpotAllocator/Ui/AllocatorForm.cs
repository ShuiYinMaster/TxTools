// AllocatorForm.cs  —  C# 7.3
// 布局：外层 AutoScroll Panel + 垂直 stack(TableLayoutPanel Dock=Top AutoSize)，内容超出滚动不裁切。
// 卡片 MkCard(GroupBox Dock=Top AutoSize)；网格/预览/日志用固定高卡。
// 操作选择：强类型 TxObjGridCtrl(OpGridHost)，ObjectInserted 实时剔除非焊接操作。

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;

using Button = System.Windows.Forms.Button;
using Label = System.Windows.Forms.Label;
using ComboBox = System.Windows.Forms.ComboBox;
using CheckBox = System.Windows.Forms.CheckBox;
using TextBox = System.Windows.Forms.TextBox;
using RadioButton = System.Windows.Forms.RadioButton;
using Panel = System.Windows.Forms.Panel;

namespace MyPlugin.WeldSpotAllocator
{
    public class AllocatorForm : TxForm
    {
        private static AllocatorForm _inst;
        private SynchronizationContext _ctx;
        private bool _inited;

        private List<OpData> _refOps = new List<OpData>();
        private List<OpData> _targetOps = new List<OpData>();
        private AllocPlan _plan;

        private readonly OpGridHost _refHost = new OpGridHost();
        private readonly OpGridHost _tgtHost = new OpGridHost();
        private Panel _refHostPanel, _tgtHostPanel;

        private RadioButton _rbA, _rbB, _rbC;
        private ComboBox _cboFlip;
        private Button _btnPreview, _btnApply;
        private Label _lblRef, _lblTarget, _lblHint;
        private CheckBox _chkByDist, _chkByName, _chkRot, _chkVias, _chkParams, _chkConsume, _chkCountStrict, _chkPosOnly, _chkDiffFrame;
        private Button _btnExport;
        private NumericUpDown _numTol;
        private TextBox _txtPrefix;
        private DataGridView _grid;
        private TextBox _log;

        public static void ShowSingleton(SynchronizationContext ctx)
        {
            if (_inst == null || _inst.IsDisposed)
            {
                _inst = new AllocatorForm { _ctx = ctx };
                try { _inst.SemiModal = false; } catch { }
                _inst.Name = typeof(AllocatorForm).FullName;
                _inst.Show();
            }
            else
            {
                if (_inst.WindowState == FormWindowState.Minimized) _inst.WindowState = FormWindowState.Normal;
                _inst.Activate();
            }
        }

        public AllocatorForm()
        {
            try { SemiModal = false; } catch { }
            try { FlatStyleEnabled = false; } catch { }
            Text = "焊点分配 / 更新";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.None;
            ClientSize = new Size(840, 820);
            MinimumSize = new Size(700, 500);
            StartPosition = FormStartPosition.CenterParent;
            BuildUi();
            FormClosed += (s, e) => _inst = null;
        }

        public override void OnInitTxForm() { try { base.OnInitTxForm(); } catch { } EnsureInit(); }
        protected override void OnLoad(EventArgs e) { base.OnLoad(e); EnsureInit(); }
        private void EnsureInit()
        {
            if (_inited) return; _inited = true;
            try
            {
                _refHost.Init(_refHostPanel, IsWeldOp, DropLog, () => { _lblRef.Text = RefText(); }, Log);
                _tgtHost.Init(_tgtHostPanel, IsWeldOp, DropLog, () => { _lblTarget.Text = TgtText(); }, Log);
            }
            catch (Exception ex) { Log("网格初始化异常：" + ex.Message); }
        }

        // ── 整体：可滚动垂直 stack ─────────────────────────────────────────
        private void BuildUi()
        {
            var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };
            var stack = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, RowCount = 7 };
            stack.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 7; i++) stack.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            stack.Controls.Add(BuildModeCard(), 0, 0);
            stack.Controls.Add(BuildHint(), 0, 1);
            stack.Controls.Add(BuildDataCard(), 0, 2);
            stack.Controls.Add(BuildOptionCard(), 0, 3);
            stack.Controls.Add(BuildActionRow(), 0, 4);
            stack.Controls.Add(BuildPreviewCard(), 0, 5);
            stack.Controls.Add(BuildLogCard(), 0, 6);

            scroll.Controls.Add(stack);
            Controls.Add(scroll);
            SyncModeUi();
        }

        // ── 卡片工厂 ──────────────────────────────────────────────────────
        private static GroupBox MkCard(string title, Control inner)
        {
            inner.Dock = DockStyle.Top;
            var g = new GroupBox { Text = title, Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(10, 6, 10, 8), Margin = new Padding(0, 0, 0, 6) };
            g.Controls.Add(inner);
            return g;
        }
        private static GroupBox MkFixedCard(string title, int height, Control inner)
        {
            inner.Dock = DockStyle.Fill;
            var g = new GroupBox { Text = title, Dock = DockStyle.Top, AutoSize = false, Height = height, Padding = new Padding(8, 4, 8, 8), Margin = new Padding(0, 0, 0, 6) };
            g.Controls.Add(inner);
            return g;
        }

        private Control BuildModeCard()
        {
            var flow = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = true };
            _rbA = NewRb("A 位置更新"); _rbA.Checked = true;
            _rbB = NewRb("B 新焊点分配");
            _rbC = NewRb("C 对称分配");
            _rbA.CheckedChanged += (s, e) => { if (_rbA.Checked) SyncModeUi(); };
            _rbB.CheckedChanged += (s, e) => { if (_rbB.Checked) SyncModeUi(); };
            _rbC.CheckedChanged += (s, e) => { if (_rbC.Checked) SyncModeUi(); };
            flow.Controls.Add(_rbA); flow.Controls.Add(_rbB); flow.Controls.Add(_rbC);
            return MkCard("模式（单选）", flow);
        }

        private Control BuildHint()
        {
            _lblHint = new Label { Dock = DockStyle.Top, AutoSize = true, MaximumSize = new Size(800, 0), ForeColor = Color.DimGray, Margin = new Padding(3, 0, 3, 6) };
            return _lblHint;
        }

        private Control BuildDataCard()
        {
            var two = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 2, RowCount = 1 };
            two.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            two.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            _refHostPanel = new Panel { Dock = DockStyle.Fill, BackColor = SystemColors.Window, BorderStyle = BorderStyle.FixedSingle };
            _tgtHostPanel = new Panel { Dock = DockStyle.Fill, BackColor = SystemColors.Window, BorderStyle = BorderStyle.FixedSingle };
            _lblRef = new Label { Text = RefText(), AutoSize = true, Margin = new Padding(3, 4, 3, 2) };
            _lblTarget = new Label { Text = TgtText(), AutoSize = true, Margin = new Padding(3, 4, 3, 2) };

            two.Controls.Add(BuildSide("参考集", _refHostPanel, _lblRef), 0, 0);
            two.Controls.Add(BuildSide("目标集", _tgtHostPanel, _lblTarget), 1, 0);

            return MkCard("操作选择（仅焊接操作：点网格内的拾取行，再去 PS 树/3D 点选对象）", two);
        }

        private Control BuildSide(string title, Panel hostPanel, Label info)
        {
            var t = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Margin = new Padding(2) };
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.RowStyles.Add(new RowStyle(SizeType.Absolute, 200));   // 网格固定高
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            t.Controls.Add(new Label { Text = title, AutoSize = true, Font = new Font(Font, FontStyle.Bold), Margin = new Padding(3, 2, 3, 2) }, 0, 0);
            t.Controls.Add(hostPanel, 0, 1);
            t.Controls.Add(info, 0, 2);
            return t;
        }

        private Control BuildOptionCard()
        {
            var flow = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = true };

            _chkByDist = NewChk("焊点距离匹配"); _chkByDist.Checked = true;
            _chkByName = NewChk("焊点名匹配");
            _numTol = new NumericUpDown { Minimum = 0, Maximum = 100000, Value = 5, Width = 70, Margin = new Padding(3, 4, 3, 4) };
            _txtPrefix = new TextBox { Width = 110, Margin = new Padding(3, 4, 3, 4) };
            _chkDiffFrame = NewChk("不同车件参考系");
            flow.Controls.Add(_chkByDist);
            flow.Controls.Add(LabeledPair("距离阈值 mm", _numTol));
            flow.Controls.Add(_chkByName);
            flow.Controls.Add(LabeledPair("剥前缀(名)", _txtPrefix));
            flow.Controls.Add(_chkDiffFrame);
            AddBreak(flow);

            _chkRot = NewChk("复制旋转姿态");
            _cboFlip = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 70, Margin = new Padding(3, 4, 3, 4) };
            _cboFlip.Items.AddRange(new object[] { "None", "X", "Y", "Z" }); _cboFlip.SelectedIndex = 0;
            _chkVias = NewChk("镜像过渡点(C)");
            _chkParams = NewChk("复制轨迹参数");
            _chkConsume = NewChk("挪动焊点(移入新轨迹,不复制)"); _chkConsume.Checked = true;
            _chkCountStrict = NewChk("操作数量严格(A)"); _chkCountStrict.Checked = true;
            _chkPosOnly = NewChk("仅更新位置(A)");
            flow.Controls.Add(_chkRot);
            flow.Controls.Add(LabeledPair("镜像翻轴(X前后/Y左右/Z上下)", _cboFlip));
            flow.Controls.Add(_chkVias);
            flow.Controls.Add(_chkParams);
            flow.Controls.Add(_chkConsume);
            AddBreak(flow);
            flow.Controls.Add(_chkCountStrict);
            flow.Controls.Add(_chkPosOnly);
            return MkCard("匹配条件 + 分配后处理", flow);
        }

        private Control BuildActionRow()
        {
            var flow = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Top, Margin = new Padding(0, 0, 0, 4) };
            _btnPreview = NewBtn("预览匹配"); _btnPreview.Click += (s, e) => DoPreview();
            _btnApply = NewBtn("执行写入"); _btnApply.Enabled = false; _btnApply.Click += (s, e) => DoApply();
            _btnExport = NewBtn("导出Excel"); _btnExport.Click += (s, e) => DoExport();
            var bSwap = NewBtn("交换参考/目标"); bSwap.Click += (s, e) => DoSwap();
            var bClearAll = NewBtn("清空列表");
            bClearAll.Click += (s, e) => { _refHost.Clear(); _tgtHost.Clear(); _lblRef.Text = RefText(); _lblTarget.Text = TgtText(); _grid.Rows.Clear(); _btnApply.Enabled = false; };
            flow.Controls.Add(_btnPreview); flow.Controls.Add(_btnApply); flow.Controls.Add(_btnExport); flow.Controls.Add(bSwap); flow.Controls.Add(bClearAll);
            return flow;
        }

        private Control BuildPreviewCard()
        {
            _grid = new DataGridView
            {
                Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            };
            _grid.Columns.Add("op", "新轨迹");
            _grid.Columns.Add("refn", "参考焊点");
            _grid.Columns.Add("tgt", "待分配焊点");
            _grid.Columns.Add("by", "依据");
            _grid.Columns.Add("dist", "距离mm");
            return MkFixedCard("预览结果", 260, _grid);
        }

        private Control BuildLogCard()
        {
            _log = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true };
            return MkFixedCard("日志", 130, _log);
        }

        // 工厂
        private static RadioButton NewRb(string t) => new RadioButton { Text = t, AutoSize = true, Margin = new Padding(3, 4, 18, 4) };
        private static Button NewBtn(string t) => new Button { Text = t, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 2, 8, 2), Margin = new Padding(3), Anchor = AnchorStyles.Left };
        private static CheckBox NewChk(string t) => new CheckBox { Text = t, AutoSize = true, Margin = new Padding(3, 6, 12, 6) };
        private static void AddBreak(FlowLayoutPanel f) { var sep = new Label { AutoSize = false, Width = 1, Height = 1, Margin = new Padding(0) }; f.Controls.Add(sep); f.SetFlowBreak(sep, true); }
        private static Control LabeledPair(string text, Control c)
        {
            var f = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(3, 0, 12, 0), WrapContents = false };
            f.Controls.Add(new Label { Text = text, AutoSize = true, Margin = new Padding(0, 7, 4, 0) });
            f.Controls.Add(c);
            return f;
        }

        private AllocMode Mode => _rbB.Checked ? AllocMode.NewSpot : (_rbC.Checked ? AllocMode.Symmetric : AllocMode.UpdatePosition);
        private string RefText() => $"参考操作 {_refHost.Count} 条";
        private string TgtText() => $"目标操作 {_tgtHost.Count} 条";

        private void SyncModeUi()
        {
            bool assign = Mode != AllocMode.UpdatePosition;
            bool sym = Mode == AllocMode.Symmetric;
            switch (Mode)
            {
                case AllocMode.UpdatePosition:
                    _lblHint.Text = "①旧版焊接操作 ②新版焊接操作。把旧版焊点更新到新版坐标，分配不变。"; break;
                case AllocMode.NewSpot:
                    _lblHint.Text = "①已分配参考操作 ②待分配焊点清单(操作)。参考在工位、新点在导入原点对不上时，勾「不同车件参考系」(目标车身须在世界原点)。"; break;
                case AllocMode.Symmetric:
                    _lblHint.Text = "①左侧参考操作 ②右侧待分配焊点清单(操作)。参考镜像后匹配，via 镜像复制到 _Mapped。"; break;
            }
            _chkRot.Enabled = assign; _chkParams.Enabled = assign; _chkConsume.Enabled = assign; _chkDiffFrame.Enabled = assign;
            _chkVias.Enabled = sym; _cboFlip.Enabled = sym;
            _chkCountStrict.Enabled = !assign; _chkPosOnly.Enabled = !assign;
            _btnApply.Enabled = false;
        }

        // ── 数据 ──────────────────────────────────────────────────────────
        private static bool IsWeldOp(ITxObject o)
        {
            string tn = o.GetType().Name;
            return tn == "TxWeldOperation" || (tn.Contains("Weld") && tn.Contains("Operation") && !tn.Contains("Location"));
        }
        private void DropLog(ITxObject o) => Log($"⚠ 已剔除非焊接操作：{SafeName(o)}");

        private List<OpData> ReadHost(OpGridHost h)
        {
            var res = new List<OpData>();
            foreach (var o in h.GetObjects())
            {
                var od = SpotReader.ReadOneWeldOp(o, Log);
                if (od != null && (od.Spots.Count > 0 || od.Vias.Count > 0)) res.Add(od);
                else Log($"⚠ 跳过非焊接/空操作：{SafeName(o)}");
            }
            return res;
        }

        private void DoPreview()
        {
            try
            {
                _lblRef.Text = RefText(); _lblTarget.Text = TgtText();
                _refOps = ReadHost(_refHost);
                _targetOps = ReadHost(_tgtHost);

                var cfg = new MatchSettings
                {
                    ByDistance = _chkByDist.Checked,
                    ByName = _chkByName.Checked,
                    MaxDistMm = (double)_numTol.Value,
                    CountStrict = _chkCountStrict.Checked,
                    DiffFrame = _chkDiffFrame.Checked,
                    Names = new NameOptions { StripPrefix = string.IsNullOrWhiteSpace(_txtPrefix.Text) ? null : _txtPrefix.Text.Trim() }
                };
                if (!cfg.ByDistance && !cfg.ByName) { Log("⚠ 未勾任何匹配条件，按仅距离处理"); cfg.ByDistance = true; }

                _plan = SpotMatchEngine.Build(Mode, _refOps, _targetOps, cfg, (FlipAxis)_cboFlip.SelectedIndex,
                    _chkVias.Checked, _chkRot.Checked, _chkParams.Checked, _chkConsume.Checked, Log);

                FillGrid(_plan);
                foreach (var w in _plan.Warnings) Log("⚠ " + w);
                _btnApply.Enabled = _plan.TotalMatches > 0;
                Log($"预览：匹配 {_plan.TotalMatches} 点，警告 {_plan.Warnings.Count} 条");
            }
            catch (Exception ex) { Log("预览异常：" + ex.Message); }
        }

        private void FillGrid(AllocPlan plan)
        {
            _grid.Rows.Clear();
            foreach (var om in plan.OpMatches)
            {
                string opName = plan.Mode == AllocMode.UpdatePosition ? (om.TargetOp?.Name ?? om.RefOp?.Name) : (om.RefOp?.Name + "_Mapped");
                foreach (var mt in om.Matches)
                    _grid.Rows.Add(opName, mt.Ref?.Name, mt.Target?.Name, mt.By + (mt.Mirrored ? "·变换" : ""), mt.Dist.ToString("F1"));
            }
        }

        private void DoApply()
        {
            if (_plan == null) { Log("请先预览"); return; }
            var r = MessageBox.Show($"将写入 {_plan.TotalMatches} 处更改（{Mode}）。\nB/C 会 Paste 出 _Mapped 新轨迹；可用 PS 撤销回退。继续？",
                "确认执行", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (r != DialogResult.OK) return;
            try
            {
                var rep = SpotWriter.ApplyPlan(_plan, _chkPosOnly.Checked, Log);
                foreach (var e in rep.Errors) Log("✗ " + e);
                Log("完成：" + rep);
                MessageBox.Show(rep.ToString(), "执行完成", MessageBoxButtons.OK, rep.Failed > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
            catch (Exception ex) { Log("执行异常：" + ex.Message); }
        }

        private void DoSwap()
        {
            var a = _refHost.GetObjects();
            var b = _tgtHost.GetObjects();
            _refHost.Clear(); _tgtHost.Clear();
            foreach (var o in b) _refHost.Append(o);
            foreach (var o in a) _tgtHost.Append(o);
            _lblRef.Text = RefText(); _lblTarget.Text = TgtText();
            _grid.Rows.Clear(); _btnApply.Enabled = false;
            Log($"已交换参考/目标：参考 {_refHost.Count} 条 ↔ 目标 {_tgtHost.Count} 条");
        }

        // 启动 Excel（late-bound COM，免引用 Interop），新建空白工作簿，注入预览表数据
        private void DoExport()
        {
            if (_grid.Rows.Count == 0) { Log("无匹配结果，请先「预览匹配」"); return; }
            object excel = null;
            try
            {
                Type t = Type.GetTypeFromProgID("Excel.Application");
                if (t == null) { Log("未检测到 Excel（ProgID Excel.Application 缺失）"); return; }
                excel = Activator.CreateInstance(t);
                dynamic app = excel;
                app.Visible = true;
                app.DisplayAlerts = false;

                dynamic wb = app.Workbooks.Add();
                dynamic ws = wb.ActiveSheet;

                int cols = _grid.Columns.Count;
                for (int c = 0; c < cols; c++) ws.Cells[1, c + 1] = _grid.Columns[c].HeaderText;

                int r = 2;
                foreach (DataGridViewRow gr in _grid.Rows)
                {
                    if (gr.IsNewRow) continue;
                    for (int c = 0; c < cols; c++)
                        ws.Cells[r, c + 1] = gr.Cells[c].Value == null ? "" : gr.Cells[c].Value.ToString();
                    r++;
                }

                try { ws.Range[ws.Cells[1, 1], ws.Cells[1, cols]].Font.Bold = true; ws.Columns.AutoFit(); } catch { }
                Log($"已导出 {r - 2} 行到 Excel（新建工作簿）");
            }
            catch (Exception ex)
            {
                Log("导出 Excel 异常：" + ex.Message);
                try { if (excel != null) ((dynamic)excel).Visible = true; } catch { }
            }
        }

        private static string SafeName(ITxObject o) { try { dynamic d = o; return (d.Name as string) ?? o.GetType().Name; } catch { return o.GetType().Name; } }
        private void Log(string s)
        {
            if (_log.InvokeRequired) { _log.BeginInvoke((Action)(() => Log(s))); return; }
            _log.AppendText(DateTime.Now.ToString("HH:mm:ss ") + s + Environment.NewLine);
        }
    }
}
