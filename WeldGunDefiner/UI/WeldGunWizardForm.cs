// WeldGunWizardForm.cs — C# 7.3
// 布局：AutoScroll Panel + 垂直 stack(TableLayoutPanel Dock=Top AutoSize)
// 卡片 MkCard(GroupBox AutoSize) / MkFixedCard(固定高)
// 参考 AllocatorForm.cs 风格

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;
using TxTools.WeldGunDefiner.Core;
using TxTools.WeldGunDefiner.Math;
using TxTools.Common;

using Button = System.Windows.Forms.Button;
using Label = System.Windows.Forms.Label;
using Panel = System.Windows.Forms.Panel;
using TextBox = System.Windows.Forms.TextBox;
using CheckBox = System.Windows.Forms.CheckBox;

namespace TxTools.WeldGunDefiner.UI
{
    public class WeldGunWizardForm : TxForm
    {
        // ── 单例 ──────────────────────────────────────────────────────────
        private static WeldGunWizardForm _inst;
        public static void ShowSingleton()
        {
            if (_inst == null || _inst.IsDisposed) { _inst = new WeldGunWizardForm(); _inst.Show(); }
            else { if (_inst.WindowState == FormWindowState.Minimized) _inst.WindowState = FormWindowState.Normal; _inst.Activate(); }
        }

        // ── 数据模型 & 服务 ──────────────────────────────────────────────
        private readonly WeldGunModel _model = new WeldGunModel();
        private WeldGunService _service;

        // ── Step0：焊枪本体 + 建模检测 ──────────────────────────────────
        private Label _lblModelingStatus;
        private Button _btnEnableModeling;

        // ── Step1：铰点 ──────────────────────────────────────────────────
        private TxObjGridCtrl _gridO, _gridA, _gridC;
        private Label _lblErrO, _lblErrA, _lblErrC;
        private TxObjGridCtrl _gridStaticTip, _gridMovingTip;
        private TxFrameEditBoxCtrl _tcpEditBox;
        private TxTransformation _tcpPickedLocation;   // TCP点选位置(若非已有Frame,生成时建TCPF用)

        // ── Step1：Link 几何体 ───────────────────────────────────────────
        private TxObjGridCtrl _gridFixedLink, _gridInputLink, _gridCouplerLink, _gridOutputLink, _gridLnk2;
        private bool _gunFilterReliable = false;  // 需求2：祖先链判断是否可靠(自检结果)，不可靠则不剔除
        // 需求2/3：每个Link grid 关联(grid, 几何体list, 显示名, 高亮色)，用于交叉校验+刷新
        private System.Collections.Generic.List<System.Tuple<TxObjGridCtrl, System.Collections.Generic.List<ITxObject>, string, TxColor>> _linkGrids
            = new System.Collections.Generic.List<System.Tuple<TxObjGridCtrl, System.Collections.Generic.List<ITxObject>, string, TxColor>>();
        private bool _crossCheckActive = false;  // 校验弹窗进行中，防重入

        // ── Step2：参数 ──────────────────────────────────────────────────
        private TextBox _txtPistonStroke, _txtOpenGap, _txtWear;
        private bool _suppressOpenGapEvent = false;  // 程序填开口量时抑制TextChanged
        private bool _openGapUserEdited = false;     // 用户是否手动改过开口量
        private TextBox _txtRAuto, _txtDAuto, _txtDAO, _txtr, _txtL, _txtAlpha;
        private RichTextBox _txtCoordCard;

        // ── Step3：生成 ──────────────────────────────────────────────────
        private TxObjGridCtrl _gridTargetDevice;
        private TextBox _txtJ1, _txtJ2, _txtInputJ1;
        private CheckBox _chkOpenAdapt;   // 需求4：OPEN状态适配勾选
        private RichTextBox _txtFormula2, _txtFormulaInput;
        private RichTextBox _txtResult;

        // ── DPI（FormUiKit 标准方案）────────────────────────────────────────
        private bool _dpiApplied;
        private static readonly Size _designSize = new Size(560, 680);

        // ── 步骤导航 ─────────────────────────────────────────────────────
        private Panel[] _stepPanels;
        private Button _btnBack, _btnNext, _btnGenerate, _btnFinish;
        private Label _lblStepTitle;
        private Panel _stepIndicator;
        private int _currentStep = 0;

        // ── 构造 ─────────────────────────────────────────────────────────
        public WeldGunWizardForm()
        {
            try { SemiModal = false; } catch { }
            FormUiKit.InitStandardForm(this,
                "X 型焊枪快速定义向导",
                _designSize,
                new Size(480, 520),
                sizable: true);
            StartPosition = FormStartPosition.CenterParent;
            FormClosed += (s, e) =>
            {
                // 取消所有高亮，避免颜色残留在场景里
                try
                {
                    PsSdkHelper.UnhighlightAll(_model.FixedLinkBodies);
                    PsSdkHelper.UnhighlightAll(_model.InputLinkBodies);
                    PsSdkHelper.UnhighlightAll(_model.CouplerLinkBodies);
                    PsSdkHelper.UnhighlightAll(_model.OutputLinkBodies);
                    PsSdkHelper.UnhighlightAll(_model.Lnk2Bodies);
                    PsSdkHelper.Unhighlight(_model.ObjStaticTip);
                    PsSdkHelper.Unhighlight(_model.ObjMovingTip);
                }
                catch { }
                _inst = null;
            };
            _service = new WeldGunService(_model);
            BuildUi();
        }

        public override void OnInitTxForm() { try { base.OnInitTxForm(); } catch { } try { SemiModal = false; } catch { } }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // FormUiKit 标准 DPI 缩放（AutoScaleMode.None + 手动 Scale，ExportGun 同款）
            FormUiKit.ApplyDpiScaling(this, ref _dpiApplied, _designSize);
        }

        // ═════════════════════════════════════════════════════════════════
        // 整体布局
        // ═════════════════════════════════════════════════════════════════
        private void BuildUi()
        {
            // 外框：步骤指示 + 标题 + 内容区 + 按钮栏
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(6)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));  // 步骤指示
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));  // 标题
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));  // 内容
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));  // 按钮
            Controls.Add(root);

            root.Controls.Add(BuildStepIndicator(), 0, 0);

            _lblStepTitle = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 100, 180),
                Padding = new Padding(2, 0, 0, 0)
            };
            root.Controls.Add(_lblStepTitle, 0, 1);

            // 内容区（各 Step 用 Panel 切换）
            var host = new Panel { Dock = DockStyle.Fill };
            root.Controls.Add(host, 0, 2);

            _stepPanels = new Panel[]
            {
                BuildStep0(),
                BuildStep1(),
                BuildStep2(),
                BuildStep3(),
            };
            foreach (var p in _stepPanels) { p.Dock = DockStyle.Fill; p.Visible = false; host.Controls.Add(p); }

            root.Controls.Add(BuildButtonBar(), 0, 3);
            UpdateStepUi();
        }

        // ─────────────────────────────────────────────────────────────────
        // 步骤指示条
        // ─────────────────────────────────────────────────────────────────
        private Panel BuildStepIndicator()
        {
            _stepIndicator = new Panel { Dock = DockStyle.Fill };
            string[] labels = { "1. 焊枪本体", "2. 铰点/Link", "3. 参数", "4. 生成" };
            int w = 120;
            for (int i = 0; i < labels.Length; i++)
            {
                int idx = i;
                _stepIndicator.Controls.Add(new Label
                {
                    Text = labels[i],
                    Tag = idx,
                    Location = new Point(i * w, 2),
                    Size = new Size(w - 2, 28),
                    TextAlign = ContentAlignment.MiddleCenter,
                    BorderStyle = BorderStyle.FixedSingle
                });
            }
            return _stepIndicator;
        }

        private Panel BuildButtonBar()
        {
            var bar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(4, 6, 4, 0)
            };
            _btnGenerate = NewBtn("生成"); _btnGenerate.Click += (s, e) => DoGenerate(); _btnGenerate.Visible = false;
            _btnNext = NewBtn("下一步 >"); _btnNext.Click += (s, e) => GoNext();
            _btnBack = NewBtn("< 上一步"); _btnBack.Click += (s, e) => GoBack(); _btnBack.Enabled = false;
            // 7. 完成按钮（生成成功后才可见）
            _btnFinish = NewBtn("完成 ✓"); _btnFinish.Click += (s, e) => Close(); _btnFinish.Visible = false;
            bar.Controls.Add(_btnFinish);
            bar.Controls.Add(_btnGenerate);
            bar.Controls.Add(_btnNext);
            bar.Controls.Add(_btnBack);
            return bar;
        }

        // ═════════════════════════════════════════════════════════════════
        // Step 0：选取焊枪本体 + 建模状态检测
        // ═════════════════════════════════════════════════════════════════
        private Panel BuildStep0()
        {
            var scroll = MkScroll();
            var stack = MkStack(3);

            // 说明
            var hint = new Label
            {
                Text = "第一步：选取焊枪设备根节点（Component 层级）。\n选取后将自动检测建模状态，未开启则可一键启用。",
                Dock = DockStyle.Top,
                AutoSize = true,
                ForeColor = Color.DimGray,
                Margin = new Padding(2, 2, 2, 6)
            };
            stack.Controls.Add(MkCard("操作说明", hint), 0, 0);

            // 设备节点选取（Component 级）
            var devPanel = new Panel { Dock = DockStyle.Fill, Height = 80, BackColor = SystemColors.Window, BorderStyle = BorderStyle.FixedSingle };
            _gridTargetDevice = MkGrid(); _gridTargetDevice.PickLevel = TxPickLevel.Component;
            _gridTargetDevice.Dock = DockStyle.Fill;
            _gridTargetDevice.ObjectInserted += (s, e) =>
            {
                var a = e as TxObjGridCtrl_ObjectInsertedEventArgs;
                var obj = a?.Obj ?? _gridTargetDevice.GetObject(0) as ITxObject;
                if (obj == null) return;
                _model.TargetComponent = obj as ITxComponent;
                _model.TargetKinematics = obj as ITxKinematicsModellable;
                _model.TargetDevice = obj as ITxDevice;
                // 需求2：自检祖先链判断是否可靠——取焊钳一个已知后代测试，
                // 若连自己的后代都识别不了，说明祖先链API在此环境不通，
                // 则禁用外部剔除(放行所有选取)，避免误删本体几何。
                _gunFilterReliable = TestAncestorReliable(_model.TargetComponent);
                // 延迟更新状态，避免在拾取回调里同步改布局干扰焦点
                try { BeginInvoke((Action)UpdateModelingStatus); } catch { UpdateModelingStatus(); }
            };
            _gridTargetDevice.RowDeleted += (s, e) =>
            {
                _model.TargetComponent = null; _model.TargetKinematics = null; _model.TargetDevice = null;
                try { BeginInvoke((Action)UpdateModelingStatus); } catch { UpdateModelingStatus(); }
            };
            devPanel.Controls.Add(_gridTargetDevice);
            stack.Controls.Add(MkFixedCard("焊枪设备根节点", 120, devPanel), 0, 1);

            // 建模状态 + 启用按钮
            var statusPanel = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2 };
            statusPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            statusPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            _lblModelingStatus = new Label
            {
                Text = "请先选取焊枪设备",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.Gray
            };
            _btnEnableModeling = NewBtn("启用建模状态");
            _btnEnableModeling.Click += (s, e) => EnableModeling();
            _btnEnableModeling.Enabled = false;
            statusPanel.Controls.Add(_lblModelingStatus, 0, 0);
            statusPanel.Controls.Add(_btnEnableModeling, 1, 0);
            stack.Controls.Add(MkCard("建模状态检测", statusPanel), 0, 2);

            scroll.Controls.Add(stack);
            var p = new Panel { Dock = DockStyle.Fill };
            p.Controls.Add(scroll);
            return p;
        }

        // 更新建模状态显示
        private void UpdateModelingStatus()
        {
            if (_lblModelingStatus == null) return;
            if (_model.TargetKinematics == null)
            {
                _lblModelingStatus.Text = "请先选取焊枪设备";
                _lblModelingStatus.ForeColor = Color.Gray;
                _btnEnableModeling.Enabled = false;
                return;
            }
            bool open = false;
            try { open = _model.TargetKinematics.IsOpenForKinematicsModeling; } catch { }
            if (open)
            {
                _lblModelingStatus.Text = "[OK] 已开启运动学建模状态";
                _lblModelingStatus.ForeColor = Color.DarkGreen;
                _btnEnableModeling.Enabled = false;
            }
            else
            {
                _lblModelingStatus.Text = "[!] 未开启建模状态，请点击右侧按钮启用";
                _lblModelingStatus.ForeColor = Color.DarkOrange;
                _btnEnableModeling.Enabled = true;
            }
        }

        private void EnableModeling()
        {
            try
            {
                (_model.TargetComponent as ITxComponent)?.SetModelingScope();
                UpdateModelingStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show("启用建模状态失败: " + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ═════════════════════════════════════════════════════════════════
        // Step 1：铰点 + Link 几何体
        // ═════════════════════════════════════════════════════════════════
        private Panel BuildStep1()
        {
            var scroll = MkScroll();
            var stack = MkStack(3);

            // ── 铰点卡：3列（标签100 + Grid300 + 误差80），行高130（约6行）──
            var hingeLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 3 };
            hingeLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            hingeLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300));
            hingeLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
            int hrow = 0;

            void AddHingeRow(string lbl, ref TxObjGridCtrl grid, ref Label errLbl, string tag)
            {
                string t = tag;
                hingeLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 130));
                hingeLayout.Controls.Add(new Label { Text = lbl, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Padding = new Padding(0, 0, 6, 0) }, 0, hrow);
                // PickLevel.Entity：允许选子零件/Frame（修复子项无法选择的问题）
                var g = new TxObjGridCtrl { Dock = DockStyle.Fill, EnableMultipleSelection = false, PickLevel = TxPickLevel.Entity };
                WireGridPick(g);
                grid = g;
                BindHingeGrid(g,
                    pos => { if (t == "O") { _model.WorldO = pos; _model.HasO = true; } else if (t == "A") { _model.WorldA = pos; _model.HasA = true; } else { _model.WorldC = pos; _model.HasC = true; } TryUpdateErrors(); },
                    () => { if (t == "O") _model.HasO = false; else if (t == "A") _model.HasA = false; else _model.HasC = false; ClearErr(t == "O" ? _lblErrO : t == "A" ? _lblErrA : _lblErrC); },
                    // A点专用：拾取时确定参考平面法向 N
                    t == "A" ? (Action<Vec3>)(n => {
                        _model.PlaneNormal = n;
                        TryUpdateErrors();
                    }) : null);
                hingeLayout.Controls.Add(g, 1, hrow);
                var el = new Label { Text = "--", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.Gray };
                errLbl = el;
                hingeLayout.Controls.Add(el, 2, hrow);
                hrow++;
            }

            AddHingeRow("O 动臂转轴:", ref _gridO, ref _lblErrO, "O");
            AddHingeRow("A 活塞铰点:", ref _gridA, ref _lblErrA, "A");
            AddHingeRow("C 连杆铰点:", ref _gridC, ref _lblErrC, "C");

            var hintHinge = new Label { Text = "A点确定参考平面法向（Face法向或Frame的Z轴）；三点投影此平面强制共面", Dock = DockStyle.Top, AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(2, 4, 2, 2) };
            var hingeWrap = new Panel { Dock = DockStyle.Top, AutoSize = true };
            hingeWrap.Controls.Add(hintHinge);
            hingeWrap.Controls.Add(hingeLayout);
            stack.Controls.Add(MkCard("三铰点（连杆机构）", hingeWrap), 0, 0);

            // ── 活塞杆 & 电极帽：2列（标签100 + Grid300），行高130 ──
            var miscLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2 };
            miscLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            miscLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300));
            int mrow = 0;

            void AddMiscRow(string lbl, ref TxObjGridCtrl grid, Action<ITxObject> onInsert, Action onDelete)
            {
                miscLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 130));
                miscLayout.Controls.Add(new Label { Text = lbl, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Padding = new Padding(0, 0, 6, 0) }, 0, mrow);
                var g = new TxObjGridCtrl { Dock = DockStyle.Fill, EnableMultipleSelection = false, PickLevel = TxPickLevel.Entity };
                WireGridPick(g);
                grid = g;
                g.ObjectInserted += (s, e) => { var a = e as TxObjGridCtrl_ObjectInsertedEventArgs; var obj = a?.Obj ?? g.GetObject(0) as ITxObject; PsSdkHelper.Highlight(obj, PsSdkHelper.ColorTip); onInsert(obj); };
                g.RowDeleted += (s, e) => onDelete();
                miscLayout.Controls.Add(g, 1, mrow);
                mrow++;
            }

            AddMiscRow("静电极帽:", ref _gridStaticTip, o => _model.ObjStaticTip = o, () => _model.ObjStaticTip = null);
            AddMiscRow("动电极帽:", ref _gridMovingTip, o => _model.ObjMovingTip = o, () => _model.ObjMovingTip = null);

            // 焊钳TCP：用 PDPS 坐标选择器 TxFrameEditBoxCtrl
            miscLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            miscLayout.Controls.Add(new Label { Text = "焊钳TCP:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Padding = new Padding(0, 0, 6, 0) }, 0, mrow);
            _tcpEditBox = new TxFrameEditBoxCtrl { Dock = DockStyle.Fill };
            try { _tcpEditBox.ListenToPick = true; } catch { }
            try { _tcpEditBox.PickLevel = TxPickLevel.Entity; } catch { }
            _tcpEditBox.ValidFrameSet += (s, e) =>
            {
                try
                {
                    var loc = _tcpEditBox.GetLocation();
                    if (loc != null)
                    {
                        _model.WorldGunTcp = new Vec3(loc[0, 3], loc[1, 3], loc[2, 3]);
                        _tcpPickedLocation = loc;   // 保存完整变换，生成时若需建TCPF用
                    }
                }
                catch { }
            };
            // Picked：若用户点选的是已有 TxFrame 对象，直接作为 TCP Frame
            try
            {
                _tcpEditBox.Picked += (s, e) =>
                {
                    try
                    {
                        var pe = e as TxFrameEditBoxCtrl_PickedEventArgs;
                        _model.TcpFrame = pe?.Object as TxFrame;   // 是Frame就直接用，否则null(生成时建TCPF)
                        if (pe != null && pe.Location != null)
                        {
                            _tcpPickedLocation = pe.Location;
                            _model.WorldGunTcp = new Vec3(pe.Location[0, 3], pe.Location[1, 3], pe.Location[2, 3]);
                        }
                    }
                    catch { }
                };
            }
            catch { }
            miscLayout.Controls.Add(_tcpEditBox, 1, mrow);
            mrow++;
            stack.Controls.Add(MkCard("电极帽 & 焊钳TCP（坐标选择器）", miscLayout), 0, 1);

            // ── Link 几何体绑定：2×2 网格（避免单列过长）──
            // Component 级别可多选
            var linkGrid = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, RowCount = 2 };
            linkGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            linkGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            linkGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            linkGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            void AddLinkCell(string title, ref TxObjGridCtrl grid, List<ITxObject> list, int col, int row, TxColor color)
            {
                var lst = list;
                var col2 = color;
                var ttl = title;
                var g = new TxObjGridCtrl { Dock = DockStyle.Fill, EnableMultipleSelection = true, PickLevel = TxPickLevel.Entity };
                WireGridPick(g);
                grid = g;
                Action refresh = () => {
                    PsSdkHelper.UnhighlightAll(lst);            // 先取消旧高亮
                    lst.Clear();
                    for (int i = 0; i < 200; i++) { var o = g.GetObject(i) as ITxObject; if (o == null) break; lst.Add(o); }
                    PsSdkHelper.HighlightAll(lst, col2);       // 用本Link专属颜色高亮
                };
                g.ObjectInserted += (s, e) => {
                    // 需求1/2：静默剔除不属于当前焊钳的对象（无弹窗）
                    // 仅在祖先链判断可靠时启用；不可靠则放行所有(防误删本体)
                    var a = e as TxObjGridCtrl_ObjectInsertedEventArgs;
                    var obj = a?.Obj;
                    if (obj != null && _model.TargetComponent != null && _gunFilterReliable)
                    {
                        if (!PsSdkHelper.IsDescendantOf(obj, _model.TargetComponent))
                        {
                            try { RemoveObjectFromGrid(g, obj); } catch { }
                            refresh();
                            return;
                        }
                    }
                    refresh();
                };
                g.RowDeleted += (s, e) => refresh();
                // 焦点离开时做交叉校验（需求2）
                try { g.Leave += (s, e) => OnLinkGridLeave(g); } catch { }
                _linkGrids.Add(System.Tuple.Create(g, lst, ttl, col2));
                var cell = MkFixedCard(title, 150, g);
                cell.Margin = new Padding(2);
                linkGrid.Controls.Add(cell, col, row);
            }

            // 改为 2×3 网格（加 lnk2）
            linkGrid.ColumnStyles.Clear();
            linkGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            linkGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            linkGrid.RowCount = 3;
            linkGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            AddLinkCell("Fixed Link 枪体固定 (黄)", ref _gridFixedLink, _model.FixedLinkBodies, 0, 0, PsSdkHelper.ColorFixed);
            AddLinkCell("Input Link 活塞缸体 (橙)", ref _gridInputLink, _model.InputLinkBodies, 1, 0, PsSdkHelper.ColorInput);
            AddLinkCell("Coupler Link 活塞杆/连杆 (青)", ref _gridCouplerLink, _model.CouplerLinkBodies, 0, 1, PsSdkHelper.ColorCoupler);
            AddLinkCell("Output Link 动臂 (品红)", ref _gridOutputLink, _model.OutputLinkBodies, 1, 1, PsSdkHelper.ColorOutput);
            AddLinkCell("lnk2 磨量补偿 (粉)", ref _gridLnk2, _model.Lnk2Bodies, 0, 2, PsSdkHelper.ColorLnk2);

            stack.Controls.Add(MkCard("Link 几何体绑定（可多选，Component 级）", linkGrid), 0, 2);

            scroll.Controls.Add(stack);
            var p = new Panel { Dock = DockStyle.Fill };
            p.Controls.Add(scroll);
            return p;
        }
        // ═════════════════════════════════════════════════════════════════
        private Panel BuildStep2()
        {
            var scroll = MkScroll();
            var stack = MkStack(3);

            // 自动推算（只读）
            var autoLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 3 };
            autoLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            autoLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            autoLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 36));
            int ar = 0;
            TextBox AddAutoRow(string lbl, string unit)
            {
                autoLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
                autoLayout.Controls.Add(new Label { Text = lbl, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Padding = new Padding(0, 0, 4, 0) }, 0, ar);
                var tb = new TextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = Color.FromArgb(242, 242, 242) };
                autoLayout.Controls.Add(tb, 1, ar);
                autoLayout.Controls.Add(new Label { Text = unit, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Color.Gray }, 2, ar);
                ar++; autoLayout.RowCount = ar; return tb;
            }
            _txtRAuto = AddAutoRow("R  臂长(动TCP→O):", "mm");
            _txtDAuto = AddAutoRow("最大张开距离:", "mm");
            _txtDAO = AddAutoRow("d_AO  活塞铰->轴:", "mm");
            _txtr = AddAutoRow("r  连杆铰->轴:", "mm");
            _txtL = AddAutoRow("L  初始伸出量:", "mm");
            _txtAlpha = AddAutoRow("Alpha 结构角:", "deg");
            stack.Controls.Add(MkCard("自动推算参数（只读）", autoLayout), 0, 0);

            // 用户输入
            var userLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 3 };
            userLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            userLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            userLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 36));
            int ur = 0;
            TextBox AddUserRow(string lbl, string defVal, string unit)
            {
                userLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
                userLayout.Controls.Add(new Label { Text = lbl, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Padding = new Padding(0, 0, 4, 0) }, 0, ur);
                var tb = new TextBox { Dock = DockStyle.Fill, Text = defVal };
                userLayout.Controls.Add(tb, 1, ur);
                userLayout.Controls.Add(new Label { Text = unit, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Color.Gray }, 2, ur);
                ur++; userLayout.RowCount = ur; return tb;
            }
            // s0(活塞闭合长) = |AC| - 行程，自动算，无需用户输入
            _txtPistonStroke = AddUserRow("活塞行程:", "100", "mm");
            _txtOpenGap = AddUserRow("焊钳开口量:", "185", "mm");
            _txtWear = AddUserRow("磨损补偿量:", "10", "mm");
            _txtPistonStroke.TextChanged += (s, e) => { OnParamChanged(); SyncOpenGapFromCalc(); };
            _txtOpenGap.TextChanged += (s, e) => { if (_suppressOpenGapEvent) return; _openGapUserEdited = true; OnParamChanged(); };
            _txtWear.TextChanged += (s, e) => OnParamChanged();
            stack.Controls.Add(MkCard("焊钳参数（用户填写）", userLayout), 0, 1);

            // RPRR 坐标卡
            _txtCoordCard = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.FromArgb(245, 250, 255),
                Font = new Font("Consolas", 8),
                Text = "（请先完成铰点选取，参数将自动填入）"
            };
            stack.Controls.Add(MkFixedCard("RPRR 向导坐标参考", 120, _txtCoordCard), 0, 2);

            scroll.Controls.Add(stack);
            var p = new Panel { Dock = DockStyle.Fill };
            p.Controls.Add(scroll);
            return p;
        }

        // ═════════════════════════════════════════════════════════════════
        // Step 3：生成
        // ═════════════════════════════════════════════════════════════════
        private Panel BuildStep3()
        {
            var scroll = MkScroll();
            var stack = MkStack(3);

            // 关节命名（设备已在Step1选取）
            var nameLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2 };
            nameLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
            nameLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            int nr = 0;
            TextBox AddNameRow(string lbl, string def)
            {
                nameLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
                nameLayout.Controls.Add(new Label { Text = lbl, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Padding = new Padding(0, 0, 4, 0) }, 0, nr);
                var tb = new TextBox { Dock = DockStyle.Fill, Text = def };
                nameLayout.Controls.Add(tb, 1, nr);
                nr++; nameLayout.RowCount = nr; return tb;
            }
            _txtJ1 = AddNameRow("J1 主动关节:", "j1");
            _txtJ2 = AddNameRow("J2 磨量补偿:", "j2");
            _txtInputJ1 = AddNameRow("input_j1 动臂:", "input_j1");
            _txtJ1.TextChanged += (s, e) => { _model.J1_Name = _txtJ1.Text; RefreshFormulaPreview(); };
            _txtJ2.TextChanged += (s, e) => { _model.J2_Name = _txtJ2.Text; };
            _txtInputJ1.TextChanged += (s, e) => { _model.InputJ1_Name = _txtInputJ1.Text; RefreshFormulaPreview(); };

            // 需求4：OPEN状态适配勾选
            _chkOpenAdapt = new CheckBox
            {
                Text = "OPEN状态适配（焊钳模型为开口最大状态时勾选）",
                Dock = DockStyle.Top,
                AutoSize = true,
                Margin = new Padding(4, 6, 4, 2)
            };
            _chkOpenAdapt.CheckedChanged += (s, e) => { _model.OpenStateAdapt = _chkOpenAdapt.Checked; };
            var hintOpen = new Label
            {
                Text = "勾选后：先建主链→正向驱动活塞回闭合→定为0位→重写限位→再接驱动链。\n适用于厂家给的OPEN状态(滑块在最大值)模型。默认(不勾)适用CLOSE状态模型。",
                Dock = DockStyle.Top,
                AutoSize = true,
                ForeColor = Color.DimGray,
                Margin = new Padding(20, 0, 4, 4)
            };
            var nameWrap = new Panel { Dock = DockStyle.Top, AutoSize = true };
            nameWrap.Controls.Add(hintOpen);
            nameWrap.Controls.Add(_chkOpenAdapt);
            nameWrap.Controls.Add(nameLayout);
            stack.Controls.Add(MkCard("关节命名 & 状态适配", nameWrap), 0, 0);

            // 公式预览
            _txtFormula2 = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = Color.FromArgb(248, 248, 248), Font = new Font("Consolas", 8) };
            _txtFormulaInput = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = Color.FromArgb(248, 248, 248), Font = new Font("Consolas", 8) };
            var fmlLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2 };
            fmlLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            fmlLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            fmlLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            fmlLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            fmlLayout.Controls.Add(new Label { Text = "J2 公式:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight }, 0, 0);
            fmlLayout.Controls.Add(_txtFormula2, 1, 0);
            fmlLayout.Controls.Add(new Label { Text = "input_j1:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight }, 0, 1);
            fmlLayout.Controls.Add(_txtFormulaInput, 1, 1);
            stack.Controls.Add(MkFixedCard("Joint Dependency 公式预览（可复制）", 110, fmlLayout), 0, 1);

            // 结果
            _txtResult = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = Color.FromArgb(250, 252, 255) };
            stack.Controls.Add(MkFixedCard("生成结果", 200, _txtResult), 0, 2);

            scroll.Controls.Add(stack);
            var p = new Panel { Dock = DockStyle.Fill };
            p.Controls.Add(scroll);
            return p;
        }

        // ═════════════════════════════════════════════════════════════════
        // 辅助：卡片工厂（参考 AllocatorForm）
        // ═════════════════════════════════════════════════════════════════
        private static GroupBox MkCard(string title, Control inner)
        {
            inner.Dock = DockStyle.Top;
            var g = new GroupBox
            {
                Text = title,
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(8, 4, 8, 8),
                Margin = new Padding(0, 0, 0, 6)
            };
            g.Controls.Add(inner);
            return g;
        }
        private static GroupBox MkFixedCard(string title, int height, Control inner)
        {
            inner.Dock = DockStyle.Fill;
            var g = new GroupBox
            {
                Text = title,
                Dock = DockStyle.Top,
                AutoSize = false,
                Height = height,
                Padding = new Padding(8, 4, 8, 8),
                Margin = new Padding(0, 0, 0, 6)
            };
            g.Controls.Add(inner);
            return g;
        }
        private static Panel MkScroll()
            => new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(6) };
        private static TableLayoutPanel MkStack(int rows)
        {
            var t = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, RowCount = rows };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < rows; i++) t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            return t;
        }
        private static TxObjGridCtrl MkGrid()
        {
            var g = new TxObjGridCtrl { Dock = DockStyle.Fill, EnableMultipleSelection = false };
            WireGridPick(g);
            return g;
        }

        /// <summary>从 grid 移除指定对象（用 DeleteRow，按行匹配）</summary>
        private void RemoveObjectFromGrid(TxObjGridCtrl g, ITxObject target)
        {
            if (g == null || target == null) return;
            try
            {
                for (int i = 0; i < g.Count; i++)
                {
                    var o = g.GetObject(i) as ITxObject;
                    if (o == target) { g.DeleteRow(i); return; }
                }
            }
            catch { }
        }

        /// <summary>用对象列表重置 grid 内容（清空后 Objects 整体赋值）</summary>
        private void ResetGridObjects(TxObjGridCtrl g, System.Collections.Generic.List<ITxObject> objs)
        {
            if (g == null) return;
            try
            {
                var list = new TxObjectList();
                if (objs != null) foreach (var o in objs) if (o != null) list.Add(o);
                g.Objects = list;
            }
            catch { }
        }

        /// <summary>
        /// 自检：祖先链判断在当前环境/此焊钳是否可靠。
        /// 取焊钳几个已知后代，若 IsDescendantOf 能正确识别它们属于焊钳，则可靠；
        /// 否则(API在此环境走不通)返回false，调用方据此禁用外部剔除，避免误删本体。
        /// </summary>
        private bool TestAncestorReliable(ITxComponent comp)
        {
            if (comp == null) return false;
            try
            {
                var coll = comp as Tecnomatix.Engineering.ITxObjectCollection;
                if (coll == null) return false;
                var filter = new Tecnomatix.Engineering.TxTypeFilter(typeof(ITxObject));
                var desc = coll.GetAllDescendants(filter);
                if (desc == null) return false;
                int tested = 0, passed = 0;
                foreach (ITxObject o in desc)
                {
                    if (o == null) continue;
                    tested++;
                    if (PsSdkHelper.IsDescendantOf(o, comp)) passed++;
                    if (tested >= 5) break;   // 抽样几个即可
                }
                // 抽样后代里至少能识别一个，才认为祖先链可用
                return tested > 0 && passed > 0;
            }
            catch { return false; }
        }


        /// 若离开的 grid 与其他 grid 有共同几何体，弹窗三选项：
        ///   (1) 保留到此 grid，从其他 grid 移除
        ///   (2) 不保留交叉项（从此 grid 移除，其他保留）
        ///   (3) 取消，焦点返回此 grid
        /// </summary>
        private void OnLinkGridLeave(TxObjGridCtrl leftGrid)
        {
            if (_crossCheckActive) return;
            if (leftGrid == null) return;

            // 找到离开grid对应的list和名字
            System.Collections.Generic.List<ITxObject> leftList = null;
            string leftName = null;
            foreach (var t in _linkGrids)
                if (t.Item1 == leftGrid) { leftList = t.Item2; leftName = t.Item3; break; }
            if (leftList == null || leftList.Count == 0) return;

            // 收集与其他grid的交叉项
            var crossing = new System.Collections.Generic.List<ITxObject>();
            var otherInfo = new System.Collections.Generic.List<string>();
            foreach (var t in _linkGrids)
            {
                if (t.Item1 == leftGrid) continue;
                foreach (var o in t.Item2)
                {
                    if (o != null && leftList.Contains(o) && !crossing.Contains(o))
                    {
                        crossing.Add(o);
                        if (!otherInfo.Contains(t.Item3)) otherInfo.Add(t.Item3);
                    }
                }
            }
            if (crossing.Count == 0) return;

            _crossCheckActive = true;
            try
            {
                string msg = $"检测到 {crossing.Count} 个几何体同时存在于：\n" +
                             $"「{leftName}」与「{string.Join("、", otherInfo)}」\n\n" +
                             "请选择处理方式：\n" +
                             "  [是]  保留到「" + leftName + "」，从其他Link移除\n" +
                             "  [否]  不保留交叉项（从「" + leftName + "」移除）\n" +
                             "  [取消]  什么都不做，返回继续编辑";
                var r = MessageBox.Show(msg, "Link几何体交叉",
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (r == DialogResult.Yes)
                {
                    // 保留到本grid，从其他grid移除交叉项
                    foreach (var t in _linkGrids)
                    {
                        if (t.Item1 == leftGrid) continue;
                        bool changed = false;
                        foreach (var o in crossing)
                            if (t.Item2.Contains(o)) { RemoveObjectFromGrid(t.Item1, o); changed = true; }
                        if (changed)
                        {
                            t.Item2.RemoveAll(o => crossing.Contains(o));
                        }
                    }
                    RefreshAllLinkHighlights();   // 问题3：刷新高亮颜色
                }
                else if (r == DialogResult.No)
                {
                    // 从本grid移除交叉项（其他保留）
                    foreach (var o in crossing) RemoveObjectFromGrid(leftGrid, o);
                    leftList.RemoveAll(o => crossing.Contains(o));
                    RefreshAllLinkHighlights();   // 问题3：刷新高亮颜色
                }
                else
                {
                    // 取消：焦点返回本grid
                    try { BeginInvoke((Action)(() => { leftGrid.Focus(); ActivateGridPick(leftGrid); })); } catch { }
                }
            }
            finally { _crossCheckActive = false; }
        }

        /// <summary>问题3：交叉处理后重新着色所有Link几何体(先全清再按各自颜色上色)</summary>
        private void RefreshAllLinkHighlights()
        {
            try
            {
                foreach (var t in _linkGrids) PsSdkHelper.UnhighlightAll(t.Item2);
                foreach (var t in _linkGrids) PsSdkHelper.HighlightAll(t.Item2, t.Item4);
            }
            catch { }
        }

        /// <summary>
        /// 统一处理 grid 拾取激活，解决"点击选择框后鼠标不变十字"的焦点问题。
        /// 根因：TxObjGridCtrl 首次显示时 PS 拾取提供者未注册，需在控件获得焦点/
        /// 点击时主动 Focus() + 重设 ListenToPick 来触发 PS 重新注册拾取监听。
        /// </summary>
        private static void WireGridPick(TxObjGridCtrl g)
        {
            try { g.ListenToPick = true; } catch { }
            // 点击/进入控件时强制激活拾取
            g.Enter += (s, e) => ActivateGridPick(g);
            g.MouseDown += (s, e) => ActivateGridPick(g);
            g.Click += (s, e) => ActivateGridPick(g);
        }

        private static void ActivateGridPick(TxObjGridCtrl g)
        {
            try
            {
                g.Focus();              // 抢回输入焦点
                g.ListenToPick = false; // 切换一次触发 PS 重新注册拾取提供者
                g.ListenToPick = true;
            }
            catch { }
        }
        private static Button NewBtn(string t) => new Button { Text = t, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 2, 8, 2), Margin = new Padding(4) };
        private static Label MakeLabel(string t) => new Label { Text = t, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight, Padding = new Padding(0, 0, 6, 0) };

        // ═════════════════════════════════════════════════════════════════
        // 铰点 Grid：双路坐标（Frame原点 / Picked点击位置）
        // ═════════════════════════════════════════════════════════════════
        private void BindHingeGrid(TxObjGridCtrl grid, Action<Vec3> onSet, Action onClear,
                                   Action<Vec3> onNormal = null)
        {
            grid.Picked += (s, e) =>
            {
                var args = e as TxObjGridCtrl_PickedEventArgs;
                if (args == null) return;
                var loc = args.Location;
                onSet(new Vec3(loc[0, 3], loc[1, 3], loc[2, 3]));
                // A点专用：Face Pick 时提取面法向（Location Z轴）作为参考平面法向 N
                if (onNormal != null)
                {
                    var n = new Vec3(loc[0, 2], loc[1, 2], loc[2, 2]);
                    if (n.Length > 0.5) onNormal(n.Normalized());
                }
            };
            grid.ObjectInserted += (s, e) =>
            {
                var args = e as TxObjGridCtrl_ObjectInsertedEventArgs;
                var obj = args?.Obj ?? grid.GetObject(0) as ITxObject;
                if (obj == null) return;
                if (obj is TxFrame frame)
                {
                    var tx = frame.AbsoluteLocation;
                    onSet(new Vec3(tx[0, 3], tx[1, 3], tx[2, 3]));
                    // A点专用：Frame 的 Z 轴作为法向 N（以 Frame 的 XY 为平面）
                    if (onNormal != null)
                    {
                        var n = new Vec3(tx[0, 2], tx[1, 2], tx[2, 2]);
                        if (n.Length > 0.5) onNormal(n.Normalized());
                    }
                    return;
                }
                // 非Frame：等 Picked 事件取点击位置
            };
            grid.RowDeleted += (s, e) => onClear();
        }

        private void TryUpdateErrors()
        {
            if (!_model.AllHingesReady) return;
            string err;
            if (_service.UpdateHingePoints(out err))
            {
                UpdateErrLabel(_lblErrO, _model.ErrO);
                UpdateErrLabel(_lblErrA, _model.ErrA);
                UpdateErrLabel(_lblErrC, _model.ErrC);
            }
        }

        private void UpdateErrLabel(Label lbl, double err)
        {
            if (lbl == null) return;
            if (err < 0) { lbl.Text = "自动 [OK]"; lbl.ForeColor = Color.DarkGreen; }
            else if (err < 0.5) { lbl.Text = $"{err:F2}mm [OK]"; lbl.ForeColor = Color.DarkGreen; }
            else if (err < 1.0) { lbl.Text = $"{err:F2}mm [!]"; lbl.ForeColor = Color.DarkOrange; }
            else { lbl.Text = $"{err:F2}mm [X]"; lbl.ForeColor = Color.Red; }
        }
        private void ClearErr(Label lbl) { if (lbl == null) return; lbl.Text = "--"; lbl.ForeColor = Color.Gray; }

        // ═════════════════════════════════════════════════════════════════
        // 参数变更
        // ═════════════════════════════════════════════════════════════════
        private void OnParamChanged()
        {
            double v;
            if (double.TryParse(_txtPistonStroke.Text, out v)) _model.PistonStroke = v;
            if (double.TryParse(_txtOpenGap.Text, out v)) _model.OpenGap = v;
            if (double.TryParse(_txtWear.Text, out v)) _model.WearAllowance = v;
            RefreshParameterDisplay();
        }

        /// <summary>
        /// 活塞行程变化时，把计算出的最大开口量自动填入"焊钳开口量"框。
        /// 用户一旦手动改过开口量，则不再自动覆盖（尊重用户输入）。
        /// </summary>
        private void SyncOpenGapFromCalc()
        {
            if (_openGapUserEdited) return;       // 用户改过就不自动填
            if (_model.MaxOpenGap <= 0) return;
            _suppressOpenGapEvent = true;
            try
            {
                _txtOpenGap.Text = $"{_model.MaxOpenGap:F1}";
                _model.OpenGap = _model.MaxOpenGap;
            }
            finally { _suppressOpenGapEvent = false; }
        }

        private void RefreshParameterDisplay()
        {
            if (_model.Mechanism == null) return;
            var m = _model.Mechanism;
            m.S0 = _model.S0; m.PistonStroke = _model.PistonStroke;
            m.OpenGap = _model.OpenGap; m.WearAllowance = _model.WearAllowance;

            // ── 臂长 R = TCP(静电极尖端)→ O 投影距离 ──
            // 注意：电极帽自身坐标在【尾端】不可靠，TCP固定在静电极【尖端】
            // X型焊钳闭合时两电极尖端接触在TCP位置，动臂张开时动电极尖绕O摆动，
            // 所以摆动有效半径 = |TCP → O|，这也是臂长
            double armLen = 0;
            bool hasTcp = _model.WorldGunTcp.Length > 1e-6;
            if (_model.Plane != null && hasTcp)
            {
                Vec3 projTcp = _model.Plane.Project(_model.WorldGunTcp);
                Vec3 projO = _model.Plane.Project(_model.WorldO);
                armLen = (projTcp - projO).Length;
                _model.ArmLength = armLen;
            }
            else if (_model.Plane != null && _model.ObjMovingTip != null)
            {
                // 兜底：TCP未选时用动电极帽（精度较低）
                Vec3 wMov;
                if (PsSdkHelper.TryGetWorldPosition(_model.ObjMovingTip, out wMov))
                {
                    Vec3 projMov = _model.Plane.Project(wMov);
                    Vec3 projO = _model.Plane.Project(_model.WorldO);
                    armLen = (projMov - projO).Length;
                    _model.ArmLength = armLen;
                }
            }

            // ── 最大张开距离：动臂摆角 × 有效半径(TCP→O) ──
            // output_j1 = -(coup_output_j1 + fixed_Input_j1)，动臂相对基座转角
            // 开口 = 2 × |TCP→O| × sin(摆角/2)，TCP=静电极尖端=闭合时电极接触点
            double maxOpen = 0;
            if (armLen > 0 && m.d_AO > 0 && m.r > 0)
            {
                // s0 = 活塞闭合长 = |AC| - 行程（m.L 即 |AC|）
                double a = m.d_AO, b = m.r;
                double s0 = m.L - _model.PistonStroke;
                if (s0 <= 1e-6) s0 = m.L;
                if (s0 > 1e-6)
                {
                    double innerConst = b * b - a * a;
                    // fixed_Input_j1 角（余弦定理）
                    System.Func<double, double> fixedInput = d =>
                    {
                        double cv = (a * a + d * d - b * b) / (2 * a * d);
                        return System.Math.Acos(System.Math.Max(-1, System.Math.Min(1, cv)));
                    };
                    double offsetA = fixedInput(s0);
                    double offsetC = System.Math.Atan2(a * System.Math.Sin(offsetA),
                                                       (s0 * s0 + innerConst) / (2 * s0));
                    // coup_output_j1 角
                    System.Func<double, double> coupOutput = d =>
                    {
                        double y = a * System.Math.Sin(fixedInput(d));
                        double x = (d * d + innerConst) / (2 * d);
                        return System.Math.Atan2(y, x) - offsetC;
                    };
                    // output_j1 角 = -(coup + (fixed-offsetA))
                    System.Func<double, double> outputAngle = d =>
                        -(coupOutput(d) + (fixedInput(d) - offsetA));

                    double out0 = outputAngle(s0);
                    double outMax = outputAngle(s0 + _model.PistonStroke);
                    double swing = System.Math.Abs(outMax - out0);
                    maxOpen = 2 * armLen * System.Math.Sin(swing / 2);
                    _model.MaxOpenGap = maxOpen;
                }
            }

            _txtRAuto.Text = armLen > 0 ? $"{armLen:F2}" : "--";
            _txtDAuto.Text = maxOpen > 0 ? $"{maxOpen:F2}" : "--";
            _txtDAO.Text = $"{m.d_AO:F2}";
            _txtr.Text = $"{m.r:F2}";
            _txtL.Text = $"{m.L:F2}";
            _txtAlpha.Text = $"{m.Alpha * 180.0 / System.Math.PI:F2}";
            _txtCoordCard.Text = m.FormatHingePointsForWizard(_model.WorldO, _model.WorldA, _model.WorldC);
        }

        private void RefreshFormulaPreview()
        {
            if (_model.Mechanism == null) return;
            _model.Mechanism.PistonStroke = _model.PistonStroke;
            _model.Mechanism.OpenGap = _model.OpenGap;
            _model.Mechanism.WearAllowance = _model.WearAllowance;
            var formulas = _model.Mechanism.BuildFormulas(_model.J1_Name, _model.J2_Name, _model.InputJ1_Name);
            _txtFormula2.Text = formulas.J2_Formula;
            _txtFormulaInput.Text = formulas.InputJ1_Formula;
            _model.Formulas = formulas;
        }

        // ═════════════════════════════════════════════════════════════════
        // 生成
        // ═════════════════════════════════════════════════════════════════
        private void DoGenerate()
        {
            if (_model.TargetKinematics == null)
            {
                MessageBox.Show("请在「目标焊枪设备节点」中选取焊枪根节点。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 6. 建模状态检测
            if (!_model.TargetKinematics.IsOpenForKinematicsModeling)
            {
                var dr = MessageBox.Show(
                    "焊枪设备尚未开启运动学建模状态。\n\n是否立即启用？",
                    "需要建模状态", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr != DialogResult.Yes) return;
                try { (_model.TargetComponent as ITxComponent)?.SetModelingScope(); }
                catch (Exception ex)
                {
                    MessageBox.Show("启用建模状态失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            _model.J1_Name = _txtJ1?.Text ?? "j1";
            _model.J2_Name = _txtJ2?.Text ?? "j2";
            _model.InputJ1_Name = _txtInputJ1?.Text ?? "input_j1";

            // 需求3：TCP Frame 准备
            // 若用户点选的是已有Frame(_model.TcpFrame已设)，直接用；
            // 否则在点选位置创建名为 TCPF 的 Frame
            if (_model.TcpFrame == null && _tcpPickedLocation != null && _model.TargetComponent != null)
            {
                string fErr;
                var tcpf = PsSdkHelper.CreateFrame(_model.TargetComponent, "TCPF", _tcpPickedLocation, out fErr);
                if (tcpf != null) _model.TcpFrame = tcpf;
                else MessageBox.Show("创建TCPF失败: " + fErr + "\n焊钳定义的TCP将为空。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            var result = _service.GenerateAll();
            _txtResult.Text = result.Summary();
            _txtResult.ForeColor = result.Success ? Color.DarkGreen : Color.DarkRed;

            // 7. 生成成功后显示完成按钮
            if (result.Success)
            {
                _btnFinish.Visible = true;
                // 问题5：OPEN适配成功后，几何已被驱动并置零为CLOSE状态。
                // 自动取消勾选，避免再次点生成时重复驱动+置零导致过冲。
                if (_model.OpenStateAdapt)
                {
                    _model.OpenStateAdapt = false;
                    try { if (_chkOpenAdapt != null) _chkOpenAdapt.Checked = false; } catch { }
                }
            }
        }

        // ═════════════════════════════════════════════════════════════════
        // 步骤导航
        // ═════════════════════════════════════════════════════════════════
        private void GoNext()
        {
            if (!ValidateStep(out string err))
            { MessageBox.Show(err, "请完成当前步骤", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            if (_currentStep == 1)
            {
                _model.PistonStroke = 100; _model.OpenGap = 185; _model.WearAllowance = 10;
                double v;
                if (double.TryParse(_txtPistonStroke?.Text, out v)) _model.PistonStroke = v;
                if (double.TryParse(_txtOpenGap?.Text, out v)) _model.OpenGap = v;
                if (double.TryParse(_txtWear?.Text, out v)) _model.WearAllowance = v;
                string calcErr;
                if (!_service.ComputeMechanismFromPoints(out calcErr))
                { MessageBox.Show("机构参数计算失败:\n" + calcErr, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                RefreshParameterDisplay();
                // 首次进入参数页：把计算的最大开口量自动填入开口量框
                _openGapUserEdited = false;
                SyncOpenGapFromCalc();
            }
            if (_currentStep == 2) RefreshFormulaPreview();

            _currentStep = System.Math.Min(_currentStep + 1, _stepPanels.Length - 1);
            UpdateStepUi();
        }

        private void GoBack() { _currentStep = System.Math.Max(_currentStep - 1, 0); UpdateStepUi(); }

        private bool ValidateStep(out string err)
        {
            err = null;
            if (_currentStep == 0)
            {
                // Step0：必须选取焊枪本体且开启建模状态
                if (_model.TargetKinematics == null) { err = "请先选取焊枪设备根节点"; return false; }
                bool open = false;
                try { open = _model.TargetKinematics.IsOpenForKinematicsModeling; } catch { }
                if (!open) { err = "焊枪未开启运动学建模状态，请点击「启用建模状态」按钮"; return false; }
                return true;
            }
            if (_currentStep == 1)
            {
                if (!_model.HasO) { err = "请设置 O 点坐标"; return false; }
                if (!_model.HasA) { err = "请设置 A 点坐标（A点同时确定参考平面法向）"; return false; }
                if (!_model.HasC) { err = "请设置 C 点坐标"; return false; }
                if (_model.ObjStaticTip == null) { err = "请选取静电极帽"; return false; }
                if (_model.ObjMovingTip == null) { err = "请选取动电极帽"; return false; }
                if (_model.WorldGunTcp.Length < 1e-6) { err = "请选取焊钳TCP（静电极尖端坐标，用于臂长和开口计算）"; return false; }
                // 投影偏移已确定必须执行，不再提示
            }
            return true;
        }

        private void UpdateStepUi()
        {
            string[] titles = {
                "Step 1/4  --  选取焊枪本体 & 建模检测",
                "Step 2/4  --  铰点 & 电极 & Link 几何体",
                "Step 3/4  --  确认机构参数",
                "Step 4/4  --  生成运动学定义",
            };
            _lblStepTitle.Text = titles[_currentStep];
            for (int i = 0; i < _stepPanels.Length; i++) _stepPanels[i].Visible = i == _currentStep;
            _btnBack.Enabled = _currentStep > 0;
            _btnNext.Visible = _currentStep < _stepPanels.Length - 1;
            _btnGenerate.Visible = _currentStep == _stepPanels.Length - 1;
            // 切换步骤时隐藏完成按钮（只在生成成功后显示）
            if (_currentStep < _stepPanels.Length - 1) _btnFinish.Visible = false;

            foreach (Control c in _stepIndicator.Controls)
            {
                if (c is Label lbl && lbl.Tag is int idx)
                {
                    if (idx == _currentStep) { lbl.BackColor = Color.FromArgb(0, 114, 188); lbl.ForeColor = Color.White; }
                    else if (idx < _currentStep) { lbl.BackColor = Color.FromArgb(198, 230, 198); lbl.ForeColor = Color.DarkGreen; }
                    else { lbl.BackColor = SystemColors.Control; lbl.ForeColor = SystemColors.ControlText; }
                }
            }
        }
    }
}