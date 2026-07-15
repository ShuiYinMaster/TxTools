// GrouperForm.cs — C# 7.3
// 焊点自动分组 GUI。TxForm + TxObjGridCtrl 原生拾取。
//
// 关键约定（均为验证过的 PS 2402 写法）：
//   · TxForm：SemiModal=false；Name=GetType().FullName（独立持久化键）；form.Show() 在主线程；
//     静态字段持引用防 GC。
//   · TxObjGridCtrl 直接当字段用（不继承）：ListenToPick=true；Count/GetObject(i) 读对象；
//     无 Clear() → DeleteRow 循环；OnInitTxForm 是「窗体」的 public override，不调 grid.Init，
//     BeginInvoke 里 Focus()+SetCurrentCell(0,0) 激活拾取。
//   · 高 DPI：AutoScaleMode.None；OnLoad 里 sc=DpiX/96，sc>1.01 时 Scale + 显式 Size×sc。

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;

namespace TxTools.WeldSpotGrouper
{
    public class GrouperForm : TxForm
    {
        private TxObjGridCtrl _scopeGrid;
        private TextBox _prefix;
        private CheckBox _ignoreCase, _skipUnbound;
        private Button _btnScan, _btnApply, _btnClear, _btnLog;
        private Label _lblScope, _lblCount;
        private ListView _lv;
        private TextBox _txtLog;

        private List<SpotGroup> _groups;
        private readonly List<string> _log = new List<string>();
        private bool _scaled;

        private readonly Size _design = new Size(560, 660);

        public GrouperForm()
        {
            try { SemiModal = false; } catch { }
            try
            {
                var flatStyleProp = this.GetType().GetProperty("FlatStyleEnabled");
                if (flatStyleProp != null && flatStyleProp.CanWrite)
                {
                    flatStyleProp.SetValue(this, false, null);
                }
            }
            catch
            {
                // 反射失败时静默忽略，确保插件继续运行
            }
            BuildUI();
            try { Name = GetType().FullName; } catch { }
        }

        // ── 窗体 OnInitTxForm：public override，不调 grid.Init ──
        public override void OnInitTxForm()
        {
            base.OnInitTxForm();
            try { SemiModal = false; } catch { }
            BeginInvoke(new Action(delegate ()
            {
                try
                {
                    if (_scopeGrid != null) { _scopeGrid.Focus(); try { _scopeGrid.SetCurrentCell(0, 0); } catch { } }
                }
                catch { }
            }));
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            if (_scaled) return;
            _scaled = true;
            try
            {
                float sc = CreateGraphics().DpiX / 96f;
                if (sc > 1.01f)
                {
                    Scale(new SizeF(sc, sc));
                    ClientSize = new Size((int)(_design.Width * sc), (int)(_design.Height * sc));
                }
            }
            catch { }
        }

        // ════════════════════════════════════════════════════════════
        private void BuildUI()
        {
            Text = "焊点自动分组";
            AutoScaleMode = AutoScaleMode.None;
            Font = new Font("Microsoft YaHei UI", 9F);
            ClientSize = _design;
            MinimumSize = new Size(_design.Width, 420);

            var title = new Label
            {
                Text = "按绑定零件自动分组焊点",
                Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold),
                Location = new Point(12, 10),
                Size = new Size(_design.Width - 24, 24),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            Controls.Add(title);

            // 1. 范围拾取
            Controls.Add(new Label { Text = "① 拾取范围节点（焊接/复合操作或资源），可拾取多个", Location = new Point(12, 42), AutoSize = true });
            _scopeGrid = new TxObjGridCtrl
            {
                Location = new Point(12, 64),
                Size = new Size(_design.Width - 24, 88),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            try { _scopeGrid.ListenToPick = true; } catch { }
            try { _scopeGrid.EnableMultipleSelection = true; } catch { }
            try { _scopeGrid.EnableRecurringObjects = false; } catch { }
            try { _scopeGrid.ObjectInserted += (s, a) => RefreshScopeLabel(); } catch { }
            Controls.Add(_scopeGrid);

            _lblScope = new Label { Text = "已选范围：0", Location = new Point(12, 156), AutoSize = true, ForeColor = Color.DimGray };
            Controls.Add(_lblScope);

            // 2. 选项
            Controls.Add(new Label { Text = "② 选项", Location = new Point(12, 182), AutoSize = true, Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold) });
            Controls.Add(new Label { Text = "命名前缀", Location = new Point(12, 208), AutoSize = true });
            _prefix = new TextBox { Text = "焊点分组_", Location = new Point(80, 205), Size = new Size(180, 24) };
            Controls.Add(_prefix);
            _ignoreCase = new CheckBox { Text = "零件名忽略大小写", Location = new Point(280, 206), AutoSize = true, Checked = false };
            Controls.Add(_ignoreCase);
            _skipUnbound = new CheckBox { Text = "跳过无绑定零件的点", Location = new Point(12, 236), AutoSize = true, Checked = true };
            Controls.Add(_skipUnbound);

            // 3. 按钮
            _btnScan = new Button { Text = "扫描预览", Location = new Point(12, 266), Size = new Size(96, 30) };
            _btnScan.Click += (s, e) => DoScan();
            Controls.Add(_btnScan);
            _btnApply = new Button { Text = "执行分组", Location = new Point(116, 266), Size = new Size(96, 30), Enabled = false };
            _btnApply.Click += (s, e) => DoApply();
            Controls.Add(_btnApply);
            _btnClear = new Button { Text = "清空", Location = new Point(220, 266), Size = new Size(70, 30) };
            _btnClear.Click += (s, e) => DoClear();
            Controls.Add(_btnClear);
            _btnLog = new Button { Text = "日志", Location = new Point(_design.Width - 90, 266), Size = new Size(70, 30), Anchor = AnchorStyles.Top | AnchorStyles.Right };
            _btnLog.Click += (s, e) => ToggleLog();
            Controls.Add(_btnLog);

            // 4. 预览表
            _lblCount = new Label { Text = "预览：0 组", Location = new Point(12, 302), AutoSize = true };
            Controls.Add(_lblCount);
            _lv = new ListView
            {
                Location = new Point(12, 324),
                Size = new Size(_design.Width - 24, 232),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            _lv.Columns.Add("组", 40);
            _lv.Columns.Add("焊点数", 60);
            _lv.Columns.Add("绑定零件（指纹）", _design.Width - 24 - 40 - 60 - 24);
            Controls.Add(_lv);

            // 5. 日志（默认隐藏）
            _txtLog = new TextBox
            {
                Location = new Point(12, 564),
                Size = new Size(_design.Width - 24, 84),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 8.5F),
                Visible = false,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            Controls.Add(_txtLog);
        }

        // ════════════════════════════════════════════════════════════
        private void Log(string s) { _log.Add(DateTime.Now.ToString("HH:mm:ss.fff") + "  " + s); }

        private void RefreshScopeLabel()
        {
            int n = 0;
            try { n = _scopeGrid.Count; } catch { }
            if (_lblScope != null) _lblScope.Text = "已选范围：" + n;
        }

        private List<ITxObject> ReadScope()
        {
            var list = new List<ITxObject>();
            try
            {
                int n = _scopeGrid.Count;
                for (int i = 0; i < n; i++)
                {
                    ITxObject o = null;
                    try { o = _scopeGrid.GetObject(i); } catch { }
                    if (o != null) list.Add(o);
                }
            }
            catch { }
            return list;
        }

        private GroupOptions ReadOptions()
        {
            return new GroupOptions
            {
                NamePrefix = string.IsNullOrEmpty(_prefix.Text) ? "焊点分组_" : _prefix.Text,
                IgnoreCase = _ignoreCase.Checked,
                SkipUnbound = _skipUnbound.Checked
            };
        }

        // ── 扫描预览 ──
        private void DoScan()
        {
            _log.Clear();
            var scope = ReadScope();
            if (scope.Count == 0) { Info("请先在树里拾取至少一个范围节点（绿色拾取行）。"); return; }

            var opt = ReadOptions();
            var rep = new GroupReport();
            var all = new Dictionary<string, SpotGroup>(StringComparer.Ordinal);

            try
            {
                // 多范围节点：各自扫描后按指纹合并
                foreach (var node in scope)
                {
                    var subRep = new GroupReport();
                    var gs = SpotGrouper.ScanAndGroup(node, opt, subRep, Log);
                    rep.ScannedSpots += subRep.ScannedSpots;
                    rep.BoundSpots += subRep.BoundSpots;
                    rep.SkippedUnbound += subRep.SkippedUnbound;
                    foreach (var g in gs)
                    {
                        SpotGroup m;
                        if (!all.TryGetValue(g.Signature, out m)) { m = new SpotGroup { Signature = g.Signature }; m.SamplePartNames.AddRange(g.SamplePartNames); all[g.Signature] = m; }
                        m.Spots.AddRange(g.Spots);
                    }
                }
            }
            catch (Exception ex) { Log("[扫描] 异常：" + ex); Info("扫描异常：" + ex.Message); FlushLog(); return; }

            _groups = new List<SpotGroup>(all.Values);
            _groups.Sort((a, b) => b.Spots.Count.CompareTo(a.Spots.Count));

            // 填表
            _lv.BeginUpdate();
            _lv.Items.Clear();
            int idx = 1;
            foreach (var g in _groups)
            {
                var it = new ListViewItem(idx.ToString());
                it.SubItems.Add(g.Spots.Count.ToString());
                it.SubItems.Add(g.PartLabel);
                _lv.Items.Add(it);
                idx++;
            }
            _lv.EndUpdate();

            _lblCount.Text = string.Format("预览：{0} 组（{1}）", _groups.Count, rep);
            _btnApply.Enabled = _groups.Count > 0;
            FlushLog();
            if (_groups.Count == 0) Info("未找到可分组的焊点（有绑定零件的焊点 0 个）。");
        }

        // ── 执行 ──
        private void DoApply()
        {
            if (_groups == null || _groups.Count == 0) { Info("请先扫描预览。"); return; }
            var ans = MessageBox.Show(
                string.Format("将新建 {0} 个空白焊接操作（挂 OperationRoot 根）并移动焊点进去。\n确定执行吗？", _groups.Count),
                "焊点自动分组", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (ans != DialogResult.OK) return;

            var opt = ReadOptions();
            var rep = new GroupReport();
            try { GroupWriter.Apply(_groups, opt, rep, Log); }
            catch (Exception ex) { Log("[执行] 异常：" + ex); }
            FlushLog();

            _btnApply.Enabled = false;
            string msg = rep.ToString();
            if (rep.Errors.Count > 0)
            {
                msg += "\n\n失败/警告（前 8 条）：";
                for (int i = 0; i < rep.Errors.Count && i < 8; i++) msg += "\n· " + rep.Errors[i];
            }
            MessageBox.Show(msg, "分组完成", MessageBoxButtons.OK,
                rep.Failed > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
        }

        private void DoClear()
        {
            try { while (_scopeGrid.Count > 0) _scopeGrid.DeleteRow(0); } catch { }
            _lv.Items.Clear();
            _groups = null;
            _btnApply.Enabled = false;
            _lblCount.Text = "预览：0 组";
            RefreshScopeLabel();
        }

        private void ToggleLog()
        {
            _txtLog.Visible = !_txtLog.Visible;
            // 给日志腾空间：可见时压缩预览表
            int bottom = _txtLog.Visible ? _txtLog.Top - 6 : ClientSize.Height - 12;
            _lv.Height = bottom - _lv.Top;
        }

        private void FlushLog()
        {
            try { _txtLog.Text = string.Join(Environment.NewLine, _log); _txtLog.SelectionStart = _txtLog.TextLength; _txtLog.ScrollToCaret(); } catch { }
            try
            {
                string p = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "WeldSpotGrouper.log");
                System.IO.File.WriteAllText(p, string.Join(Environment.NewLine, _log), System.Text.Encoding.UTF8);
            }
            catch { }
        }

        private void Info(string m) { MessageBox.Show(m, "焊点自动分组", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }
}
