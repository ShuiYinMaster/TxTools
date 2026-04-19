// ExportGunForm.cs  —  C# 7.3
//
// 重构说明（基于确认的 API 文档）：
//
// [重构1] 窗体基类：Form → TxForm
// [重构2] 操作拾取：TxObjEditBoxCtrl
//         确认 API：.Object 属性, .ListenToPick 属性, Picked 事件
//         事件参数：TxObjEditBoxCtrl_PickedEventArgs.Object / .IsValidObjectPicked
// [重构3] 参考坐标：TxFrameComboBoxCtrl
//         确认 API：.GetLocation(), .Clear(), .SelectFrame(), .ListenToPick
//         事件：ValidFrameSet / InvalidFrameSet / Picked
//         事件参数：TxFrameComboBoxCtrl_ValidFrameSetEventArgs.Location / .Object
// [重构4] 移除 P/Invoke 置顶，移除 Timer 轮询
// [重构5] 卡片风格重构 → GroupBox 卡片，与 RobotReachabilityChecker 风格统一

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;

namespace MyPlugin.ExportGun
{
    public partial class ExportGunForm : TxForm
    {
        // ── 状态 ──────────────────────────────────────────────────────
        private ExportService _svc;
        private List<OperationInfo> _ops = new List<OperationInfo>();
        private string _refFrameName = "世界坐标系";
        private double[] _refFrameMatrix;

        // ── [重构2] PS 原生对象选择 ──────────────────────────────────
        // 优先使用 TxObjGridCtrl（原生多对象列表），失败时回落到 TxObjEditBoxCtrl + ListView
        private TxObjGridCtrl _objGrid;                     // PS 原生对象列表（主方案）
        private TxObjEditBoxCtrl _objEditOp;                // 回落：单对象选择
        private ListView lvOps;                              // 回落：列表展示
        private bool _useGrid;                               // 是否启用 Grid 主方案

        // ── [重构3] PS 原生坐标选择 ──────────────────────────────────
        private TxFrameComboBoxCtrl _frameCombo;
        private Label lblRefCoordStatus;

        // ── 控件 ──────────────────────────────────────────────────────
        private Label lblListHint;
        private ComboBox cmbPointType;
        private CheckBox chkUseMfgName;
        private Label lblPointCount;
        private CheckBox chkExportTCP;
        private CheckBox chkGunOriginTCP;
        private CheckBox chkCustomGun;
        private TextBox txtGunModel;
        private Button btnBrowseGun;
        private TextBox txtGunProductName;                  // [需求3] 插枪：自定义产品名
        private Button btnExportGun;
        private ComboBox cmbBallTarget;
        private ComboBox cmbBallOption;
        private NumericUpDown nudDiameter;
        private TextBox txtGeomSet;
        private TextBox txtNamePrefix;
        private TextBox txtBallPartName;                    // [需求3] 点球：自定义零件名
        private Button btnExportBall;
        private RichTextBox rtbLog;
        private ProgressBar progressBar;
        private Label lblProgress;
        private Button btnReset;
        private Button btnClose;
        private Button btnHelp;
        private Button btnExportExcel;
        private Label lblGunInfo;
        private Button btnPickFromSel;

        // ── TxColor 配色（与 RobotReachabilityChecker 风格统一）─────
        private static readonly TxColor TxClrAccent = new TxColor(0, 70, 127);
        private static readonly TxColor TxClrCol1 = new TxColor(0, 100, 140);
        private static readonly TxColor TxClrGun = new TxColor(155, 120, 0);
        private static readonly TxColor TxClrBall = new TxColor(150, 70, 90);
        private static readonly TxColor TxClrLog = new TxColor(50, 120, 60);
        // 功能按钮色
        private static readonly TxColor TxClrBtnPrimary = new TxColor(0, 100, 167);
        private static readonly TxColor TxClrBtnSecondary = new TxColor(80, 120, 140);
        private static readonly TxColor TxClrBtnMuted = new TxColor(120, 124, 135);
        private static readonly TxColor TxClrBtnDanger = new TxColor(130, 50, 50);
        private static readonly TxColor TxClrBtnExport = new TxColor(80, 80, 130);
        // 日志面板
        private static readonly TxColor TxClrLogBg = new TxColor(20, 22, 27);
        private static readonly TxColor TxClrLogText = new TxColor(178, 200, 178);
        private static readonly TxColor TxClrLogOk = new TxColor(90, 210, 110);
        private static readonly TxColor TxClrLogErr = new TxColor(228, 88, 88);
        private static readonly TxColor TxClrLogWarn = new TxColor(228, 180, 70);
        private static readonly TxColor TxClrLogPs = new TxColor(110, 180, 228);
        // 状态色
        private static readonly Color ClrStatusOk = Color.FromArgb(25, 110, 25);
        private static readonly Color ClrStatusBgOk = Color.FromArgb(210, 252, 210);
        private static readonly Color ClrStatusRef = Color.FromArgb(25, 60, 130);
        private static readonly Color ClrStatusBgRef = Color.FromArgb(180, 220, 255);

        // ════════════════════════════════════════════════════════════
        //  构造
        // ════════════════════════════════════════════════════════════
        public ExportGunForm(SynchronizationContext psCtx)
        {
            InitializeComponent();
            BuildUI();

            if (System.ComponentModel.LicenseManager.UsageMode ==
                System.ComponentModel.LicenseUsageMode.Designtime) return;
            _svc = new ExportService(psCtx);
        }

        // ════════════════════════════════════════════════════════════
        //  OnInitTxForm — PS 宿主框架回调
        //
        //  API Remarks (TxObjGridCtrl.Init):
        //    "If this method is not called, the grid is initialized with
        //     one text column having an 'Object' header."
        //  
        //  结论：不手动调 Init，让 PS 框架走默认初始化（单列 Object）。
        //        手动调会导致列叠加（默认一列 + 手动一列 = 两列）。
        // ════════════════════════════════════════════════════════════
        public override void OnInitTxForm()
        {
            base.OnInitTxForm();
            // 不调 _objGrid.Init(...)——让 PS 框架完成默认单列初始化
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            if (_svc == null) return;
            BindFrameComboEvents();
            // [修复] 启动欢迎日志
            Log("[系统] 插件已启动");
            Log("[系统] 在左侧操作列表中拾取 PS 对象即可开始");
            if (_useGrid)
            {
                Log("[系统] 点击列表内高亮行，在 PS 中选中对象即可加入");
                // [修复] TxObjGridCtrl 启动后首次拾取需要先获得焦点
                //        延迟设焦点确保控件已完成渲染
                BeginInvoke(new Action(delegate ()
                {
                    try
                    {
                        if (_objGrid != null && _objGrid.Visible)
                        {
                            _objGrid.Focus();
                            // 尝试选中第 0 行激活拾取模式
                            try { _objGrid.SetCurrentCell(0, 0); } catch { }
                        }
                    }
                    catch { }
                }));
            }
        }

        // ════════════════════════════════════════════════════════════
        //  [重构2] TxObjEditBoxCtrl 事件
        // ════════════════════════════════════════════════════════════

        private void OnObjEditPicked(object sender, TxObjEditBoxCtrl_PickedEventArgs e)
        {
            if (e.Object == null) return;
            ITxObject pickedObj = e.Object as ITxObject;
            if (pickedObj == null) return;

            ThreadPool.QueueUserWorkItem(delegate (object s)
            {
                List<OperationInfo> ops = null;
                try
                {
                    _svc.InvokeOnPs(delegate ()
                    { ops = PsReader.ParsePickedObjectToOperations(pickedObj); });
                }
                catch (Exception ex)
                { UI(delegate () { Log("[错误] " + ex.Message); }); }

                if (ops != null && ops.Count > 0)
                {
                    UI(delegate ()
                    {
                        foreach (var op in ops)
                            if (!_ops.Exists(o => o.Name == op.Name)) _ops.Add(op);
                        RefreshOpList(); UpdatePointCount();
                        Log("[PS] 通过原生选择器添加 " + ops.Count + " 个操作");
                    });
                }
            });
        }

        // ════════════════════════════════════════════════════════════
        //  [重构3] TxFrameComboBoxCtrl 事件
        // ════════════════════════════════════════════════════════════

        private void BindFrameComboEvents()
        {
            if (_frameCombo == null) return;
            try
            {
                _frameCombo.ValidFrameSet += new TxFrameComboBoxCtrl_ValidFrameSetEventHandler(OnFrameValidSet);
                _frameCombo.InvalidFrameSet += new TxFrameComboBoxCtrl_InvalidFrameSetEventHandler(OnFrameInvalidSet);
                _frameCombo.Picked += new TxFrameComboBoxCtrl_PickedEventHandler(OnFramePicked);
            }
            catch (Exception ex) { Log("[警告] TxFrameComboBoxCtrl 事件绑定失败: " + ex.Message); }
        }

        private void OnFrameValidSet(object sender, TxFrameComboBoxCtrl_ValidFrameSetEventArgs e)
        {
            try
            {
                TxTransformation tx = e.Location as TxTransformation;
                if (tx != null)
                {
                    double[] arr = PsReader.TxToArr(tx);
                    if (!PsReader.IsIdentity(arr))
                    {
                        string name = "用户选定坐标系";
                        if (e.Object != null)
                        {
                            ITxObject txObj = e.Object as ITxObject;
                            if (txObj != null) try { name = txObj.Name; } catch { name = txObj.GetType().Name; }
                        }
                        _refFrameName = name; _refFrameMatrix = arr;
                        UpdateRefCoordStatus(name, false);
                        Log("[坐标] 参考坐标已设置：" + name);
                        return;
                    }
                }
                string fallbackName = "用户选定坐标系";
                if (e.Object != null)
                {
                    ITxObject txObj = e.Object as ITxObject;
                    if (txObj != null) try { fallbackName = txObj.Name; } catch { }
                }
                _refFrameName = fallbackName; _refFrameMatrix = null;
                UpdateRefCoordStatus(fallbackName + " (世界坐标系)", true);
                Log("[坐标] 坐标系: " + fallbackName + " (与世界坐标系等同)");
            }
            catch (Exception ex) { Log("[坐标] 异常: " + ex.Message); }
        }

        private void OnFrameInvalidSet(object sender, TxFrameComboBoxCtrl_InvalidFrameSetEventArgs e)
        { Log("[坐标] 所选对象不是有效坐标系"); }

        private void OnFramePicked(object sender, TxFrameComboBoxCtrl_PickedEventArgs e) { }

        private void UpdateRefCoordStatus(string name, bool isWorld)
        {
            if (lblRefCoordStatus == null) return;
            lblRefCoordStatus.Text = "当前：" + name;
            lblRefCoordStatus.BackColor = isWorld ? ClrStatusBgOk : ClrStatusBgRef;
            lblRefCoordStatus.ForeColor = isWorld ? ClrStatusOk : ClrStatusRef;
        }

        // ════════════════════════════════════════════════════════════
        //  界面搭建 — GroupBox 卡片风格（与 RobotReachabilityChecker 统一）
        // ════════════════════════════════════════════════════════════
        private void BuildUI()
        {
            SuspendLayout();
            Text = "导出插枪 / 点云到 CATIA";
            Size = new Size(960, 685);
            MinimumSize = new Size(960, 685);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = false;
            BackColor = SystemColors.Control;
            Font = SystemFonts.DefaultFont;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Margin = Padding.Empty,
                Padding = new Padding(6, 4, 6, 4)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 46F));
            root.Controls.Add(BuildHeaderRow(), 0, 0);
            root.Controls.Add(BuildBody(), 0, 1);
            root.Controls.Add(BuildBottom(), 0, 2);
            Controls.Add(root);
            ResumeLayout(false); PerformLayout();
        }

        private Control BuildHeaderRow()
        {
            var row = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                Margin = new Padding(0, 0, 0, 2),
                Padding = Padding.Empty
            };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3F));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3F));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3F));
            row.Controls.Add(MkHeaderLabel("通用信息", TxClrCol1), 0, 0);
            row.Controls.Add(MkHeaderLabel("导插枪 / 点球", TxClrGun), 1, 0);
            row.Controls.Add(MkHeaderLabel("日志信息", TxClrLog), 2, 0);
            return row;
        }

        private Control BuildBody()
        {
            var body = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                Margin = new Padding(0, 2, 0, 2),
                Padding = Padding.Empty
            };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3F));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3F));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3F));
            body.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            body.Controls.Add(BuildCol1(), 0, 0);
            body.Controls.Add(BuildCol2(), 1, 0);
            body.Controls.Add(BuildCol3(), 2, 0);
            return body;
        }

        private Control BuildBottom()
        {
            var bottom = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0, 2, 0, 0),
                Padding = Padding.Empty
            };
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60F));

            var lf = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(2, 6, 2, 2)
            };
            progressBar = new ProgressBar
            {
                Height = 18,
                Width = 260,
                Style = ProgressBarStyle.Continuous,
                Margin = new Padding(0, 2, 8, 2),
                Visible = false                         // [修复] 默认隐藏，导出时才显示
            };
            lblProgress = new Label
            {
                AutoSize = true,
                Text = "就绪",                          // [修复] 默认文案替代空白
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 4, 0, 2)
            };
            lf.Controls.Add(progressBar); lf.Controls.Add(lblProgress);

            var rf = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 4, 0, 0),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnReset = MkFuncButton("复位", TxClrBtnMuted.Color);
            btnExportExcel = MkFuncButton("导出Excel", TxClrBtnExport.Color);
            btnClose = MkFuncButton("关闭", TxClrBtnDanger.Color);
            btnHelp = MkFuncButton("?", TxClrAccent.Color);
            btnReset.Click += new EventHandler(OnReset);
            btnExportExcel.Click += new EventHandler(OnExportExcel);
            btnClose.Click += delegate { Close(); };
            btnHelp.Click += delegate { ShowHelp(); };
            // [需求8] 移除"输出目录"按钮（连同 3dxml/CATProduct 选择一并取消）
            rf.Controls.AddRange(new Control[] { btnReset, btnExportExcel, btnClose, btnHelp });

            bottom.Controls.Add(lf, 0, 0); bottom.Controls.Add(rf, 1, 0);
            return bottom;
        }

        // ════════════════════════════════════════════════════════════
        //  第1列：通用信息 — 3 张 GroupBox 卡片
        //  [需求7] 展示顺序：顶部=操作选择，中间=参考坐标，底部=导出设置
        // ════════════════════════════════════════════════════════════
        private Control BuildCol1()
        {
            var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(2) };
            var stack = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                Padding = Padding.Empty
            };
            stack.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            // ── 卡片1（顶）：操作选择 ── [需求6] 使用 PS 原生 TxObjGridCtrl，合并选择器与列表 ──
            var cardOps = MkCard("操作选择", TxClrCol1);
            var flowOps = MkCardContent();

            var gridPanel = new Panel
            {
                AutoSize = false,
                Height = 220,
                Dock = DockStyle.Top,
                BackColor = SystemColors.Window,
                Padding = new Padding(1),
                Margin = new Padding(0, 0, 0, 4),
                BorderStyle = BorderStyle.FixedSingle
            };
            // 主方案：TxObjGridCtrl（API 已确认）
            _useGrid = false;
            try
            {
                _objGrid = new TxObjGridCtrl
                {
                    Dock = DockStyle.Fill,
                    ListenToPick = true,                  // API 确认：监听拾取
                    EnableMultipleSelection = true,       // API 确认：允许多选
                    EnableRecurringObjects = false        // API 确认：禁止重复对象
                };
                // Init(TxGridDescription) 必须在 OnInitTxForm 中调用（API Remarks 明确要求），
                // 这里只做控件创建和事件绑定，Init 在下方 OnInitTxForm 重写中执行。
                _objGrid.ObjectInserted += new TxObjGridCtrl_ObjectInsertedEventHandler(OnGridObjectInserted);
                _objGrid.RowDeleted += new TxObjGridCtrl_RowDeletedEventHandler(OnGridRowDeleted);
                gridPanel.Controls.Add(_objGrid);
                _useGrid = true;
            }
            catch (Exception ex)
            {
                // 回落：TxObjEditBoxCtrl + ListView
                _objGrid = null;
                try
                {
                    _objEditOp = new TxObjEditBoxCtrl { Dock = DockStyle.Top, Height = 26, ListenToPick = true, Width = gridPanel.Width - 2 };
                    _objEditOp.Picked += new TxObjEditBoxCtrl_PickedEventHandler(OnObjEditPicked);
                    gridPanel.Controls.Add(_objEditOp);
                }
                catch { }
                lvOps = new ListView
                {
                    Dock = DockStyle.Fill,
                    View = View.Details,
                    FullRowSelect = true,
                    CheckBoxes = true,
                    GridLines = false,
                    BorderStyle = BorderStyle.None,
                    HeaderStyle = ColumnHeaderStyle.Nonclickable,
                    BackColor = SystemColors.Window,
                    Font = SystemFonts.DefaultFont
                };
                lvOps.Columns.Add("操作名称", -2);
                lvOps.ItemChecked += delegate { UpdatePointCount(); };
                gridPanel.Controls.Add(lvOps);
                Log("[警告] TxObjGridCtrl 不可用，已回落到 TxObjEditBoxCtrl + ListView：" + ex.Message);
            }

            lblListHint = new Label
            {
                Text = "列表为空",
                Dock = DockStyle.Bottom,
                Height = 18,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = SystemColors.GrayText,
                BackColor = SystemColors.ControlLight,
                Font = SystemFonts.DefaultFont
            };
            gridPanel.Controls.Add(lblListHint);
            flowOps.Controls.Add(gridPanel);

            var pickRow = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 0, 0, 2)
            };
            btnPickFromSel = MkFuncButton("拾取自PS", TxClrBtnSecondary.Color);
            btnPickFromSel.Click += new EventHandler(OnPickSel);
            var btnClearOps = MkFuncButton("清空列表", TxClrBtnMuted.Color);
            btnClearOps.Click += new EventHandler(OnClearOps);
            pickRow.Controls.AddRange(new Control[] { btnPickFromSel, btnClearOps });
            flowOps.Controls.Add(pickRow);

            flowOps.Controls.Add(new Label
            {
                Text = _useGrid
                    ? "提示：亮绿色行为拾取入口，点击后去 PS 中选中对象即可加入"
                    : "在上方选择框拾取单个操作，自动加入列表；也可在PS中选中后点[拾取自PS]",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 0, 0, 2)
            });

            cardOps.Controls.Add(flowOps);
            stack.Controls.Add(cardOps, 0, 0);

            // ── 卡片2（中）：参考坐标 ──
            var cardCoord = MkCard("参考坐标", TxClrCol1);
            var flowCoord = MkCardContent();

            var frameCtrlPanel = new Panel
            { AutoSize = false, Height = 28, Dock = DockStyle.Top, Margin = new Padding(0, 2, 0, 4) };
            try
            {
                _frameCombo = new TxFrameComboBoxCtrl { Dock = DockStyle.Fill, ListenToPick = true };
                frameCtrlPanel.Controls.Add(_frameCombo);
            }
            catch (Exception ex)
            {
                frameCtrlPanel.Controls.Add(new Label
                {
                    Text = "坐标控件不可用: " + ex.Message,
                    Dock = DockStyle.Fill,
                    ForeColor = TxClrBtnDanger.Color,
                    Font = SystemFonts.DefaultFont
                });
            }
            flowCoord.Controls.Add(frameCtrlPanel);

            lblRefCoordStatus = new Label
            {
                Text = "当前：世界坐标系",
                AutoSize = true,
                ForeColor = ClrStatusOk,
                BackColor = ClrStatusBgOk,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 0, 0, 4),
                Padding = new Padding(6, 3, 6, 3)
            };
            flowCoord.Controls.Add(lblRefCoordStatus);

            var btnClrCoord = MkFuncButton("重置为世界坐标系", TxClrBtnMuted.Color);
            btnClrCoord.Margin = new Padding(0, 0, 0, 2);
            btnClrCoord.Click += delegate
            {
                _refFrameName = "世界坐标系"; _refFrameMatrix = null;
                if (_frameCombo != null) try { _frameCombo.Clear(); } catch { }
                UpdateRefCoordStatus("世界坐标系", true);
                Log("[坐标] 已重置为世界坐标系");
            };
            flowCoord.Controls.Add(btnClrCoord);

            flowCoord.Controls.Add(new Label
            {
                Text = "在PS中选择Component/Frame即可自动获取",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 0, 0, 2)
            });

            cardCoord.Controls.Add(flowCoord);
            stack.Controls.Add(cardCoord, 0, 1);

            // ── 卡片3（底）：导出设置 ──
            var cardSettings = MkCard("导出设置", TxClrCol1);
            var flowSettings = MkCardContent();

            var rowType = MkRowFlow();
            rowType.Controls.Add(MkLabel("点类型"));
            cmbPointType = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            cmbPointType.Items.AddRange(new object[] { "焊点", "路径点", "连续点", "全部类型" });
            cmbPointType.SelectedIndex = 3;
            cmbPointType.SelectedIndexChanged += delegate { UpdatePointCount(); };
            AutoFitComboBoxWidth(cmbPointType);
            rowType.Controls.Add(cmbPointType);
            flowSettings.Controls.Add(rowType);

            chkUseMfgName = new CheckBox
            {
                Text = "采用MFG名称",
                AutoSize = true,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 2, 0, 2)
            };
            chkUseMfgName.CheckedChanged += delegate { UpdatePointCount(); };
            flowSettings.Controls.Add(chkUseMfgName);

            lblPointCount = new Label
            {
                Text = "将导出点数量：0",
                AutoSize = true,
                ForeColor = ClrStatusOk,
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold),
                Margin = new Padding(0, 2, 0, 4)
            };
            flowSettings.Controls.Add(lblPointCount);
            cardSettings.Controls.Add(flowSettings);
            stack.Controls.Add(cardSettings, 0, 2);

            scroll.Controls.Add(stack);
            return scroll;
        }

        // ════════════════════════════════════════════════════════════
        //  TxObjGridCtrl 事件（[需求6] 原生列表模式 — API 已确认）
        //
        //  设计：不依赖 EventArgs 的具体成员（API 文档未列出 ObjectInsertedEventArgs /
        //        RowDeletedEventArgs 的属性细节），统一用 SyncOpsFromGrid() 重算 _ops。
        //  读取 grid 内容只用确认的 API：Count（属性）+ GetObject(int)（方法）。
        // ════════════════════════════════════════════════════════════
        private void OnGridObjectInserted(object sender, TxObjGridCtrl_ObjectInsertedEventArgs e)
        { SyncOpsFromGrid(); }

        private void OnGridRowDeleted(object sender, TxObjGridCtrl_RowDeletedEventArgs e)
        { SyncOpsFromGrid(); }

        /// <summary>
        /// 从 Grid 当前内容重新构建 _ops。
        /// 读取使用 API 确认的 Count 属性 + GetObject(int) 方法。
        /// 解析在后台线程跑，最终在 UI 线程更新 _ops。
        /// </summary>
        private void SyncOpsFromGrid()
        {
            if (!_useGrid || _objGrid == null) return;

            // 1. 收集 grid 中当前所有对象（确认 API：Count + GetObject(int)）
            var gridObjs = new List<ITxObject>();
            int n = 0;
            try { n = _objGrid.Count; }
            catch (Exception ex) { Log("[警告] 读取 _objGrid.Count 失败：" + ex.Message); return; }

            for (int i = 0; i < n; i++)
            {
                ITxObject txo = null;
                try { txo = _objGrid.GetObject(i); }   // API 确认：返回 ITxObject
                catch { /* 单行读取失败跳过 */ }
                if (txo != null) gridObjs.Add(txo);
            }

            // 2. 异步在 PS 主线程逐个解析为 OperationInfo
            ThreadPool.QueueUserWorkItem(delegate (object s)
            {
                var newOps = new List<OperationInfo>();
                try
                {
                    _svc.InvokeOnPs(delegate ()
                    {
                        foreach (var obj in gridObjs)
                        {
                            List<OperationInfo> parsed = null;
                            try { parsed = PsReader.ParsePickedObjectToOperations(obj); } catch { }
                            if (parsed == null) continue;
                            foreach (var op in parsed)
                                if (!newOps.Exists(o => o.Name == op.Name))
                                    newOps.Add(op);
                        }
                    });
                }
                catch (Exception ex) { UI(delegate () { Log("[错误] 同步操作列表异常：" + ex.Message); }); return; }

                UI(delegate ()
                {
                    bool firstFill = (_ops.Count == 0 && newOps.Count > 0);
                    _ops = newOps;
                    RefreshHintFromOps();
                    UpdatePointCount();
                    if (firstFill || newOps.Count == 0)
                        Log("[PS] 列表已同步：" + newOps.Count + " 个操作");
                    else
                        Log("[PS] 列表更新：" + newOps.Count + " 个操作");

                    // 首次有 op 时，自动加载工具/坐标信息
                    if (firstFill)
                    {
                        LogOperationTools();
                        if (_refFrameName == "世界坐标系" && _refFrameMatrix == null)
                            AutoLoadRefFrameFromOps();
                    }
                    UpdateGunInfo();
                });
            });


        }

        private void RefreshHintFromOps()
        {
            if (lblListHint == null) return;
            lblListHint.Text = _ops.Count > 0 ? "共 " + _ops.Count + " 个操作" : "列表为空";
            lblListHint.ForeColor = _ops.Count > 0 ? ClrStatusOk : SystemColors.GrayText;
        }

        // ════════════════════════════════════════════════════════════
        //  第2列：导插枪 + 导点球 — 2 张 GroupBox 卡片
        // ════════════════════════════════════════════════════════════
        private Control BuildCol2()
        {
            var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(2) };
            var stack = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                Padding = Padding.Empty
            };
            stack.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            // ── 卡片：导插枪 ──
            var cardGun = MkCard("导插枪", TxClrGun);
            var flowGun = MkCardContent();

            // [需求8] 移除 3dxml / CATProduct 单选（二者等效），默认走原先 3dxml 流程

            // [需求3] 自定义产品名称
            var rowProd = MkRowFlow();
            rowProd.Controls.Add(MkLabel("产品名称"));
            txtGunProductName = new TextBox
            {
                Width = 140,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            rowProd.Controls.Add(txtGunProductName);
            flowGun.Controls.Add(rowProd);
            flowGun.Controls.Add(new Label
            {
                Text = "留空则使用活动文档内产品名",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 0, 0, 2)
            });

            var rowCustom = MkRowFlow();
            chkCustomGun = new CheckBox
            {
                Text = "自定义焊枪数模",
                AutoSize = true,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 4, 6, 2)
            };
            var seletGun = MkRowFlow();
            txtGunModel = new TextBox
            {
                Width = 120,
                ReadOnly = true,
                Enabled = false,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 4, 6, 2)
            };
            btnBrowseGun = MkFuncButton("选择", TxClrBtnMuted.Color);
            btnBrowseGun.Enabled = false;
            chkCustomGun.CheckedChanged += delegate
            { txtGunModel.Enabled = chkCustomGun.Checked; btnBrowseGun.Enabled = chkCustomGun.Checked; };
            btnBrowseGun.Click += new EventHandler(OnBrowseGun);
            rowCustom.Controls.AddRange(new Control[] { chkCustomGun });
            seletGun.Controls.AddRange(new Control[] { txtGunModel, btnBrowseGun });
            flowGun.Controls.Add(rowCustom);
            flowGun.Controls.Add(seletGun);

            lblGunInfo = new Label
            {
                Text = "焊钳：（待选取）",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 2, 0, 4)
            };
            flowGun.Controls.Add(lblGunInfo);

            chkGunOriginTCP = new CheckBox
            {
                Text = "焊枪以TCP为原点",
                AutoSize = true,
                Checked = true,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 2, 0, 2)
            };
            flowGun.Controls.Add(chkGunOriginTCP);

            chkExportTCP = new CheckBox
            {
                Text = "导出TCP坐标",
                AutoSize = true,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 2, 0, 4)
            };
            flowGun.Controls.Add(chkExportTCP);

            btnExportGun = MkFuncButton("导出插枪", Color.Orange);
            btnExportGun.Margin = new Padding(0, 6, 0, 4);
            btnExportGun.Click += new EventHandler(OnExportGun);
            flowGun.Controls.Add(btnExportGun);

            cardGun.Controls.Add(flowGun);
            stack.Controls.Add(cardGun, 0, 0);

            // ── 卡片：导点球 ──
            var cardBall = MkCard("导点球", TxClrBall);
            var flowBall = MkCardContent();

            var rowTgt = MkRowFlow();
            rowTgt.Controls.Add(MkLabel("导出到"));
            cmbBallTarget = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            cmbBallTarget.Items.AddRange(new object[] { "当前Part文档", "新建Part文档" });
            cmbBallTarget.SelectedIndex = 0;
            AutoFitComboBoxWidth(cmbBallTarget);
            rowTgt.Controls.Add(cmbBallTarget);
            flowBall.Controls.Add(rowTgt);

            var rowOpt = MkRowFlow();
            rowOpt.Controls.Add(MkLabel("导出选项"));
            cmbBallOption = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            cmbBallOption.Items.AddRange(new object[] { "轨迹点 + 点球", "仅轨迹点", "仅点球" });
            cmbBallOption.SelectedIndex = 0;
            AutoFitComboBoxWidth(cmbBallOption);
            rowOpt.Controls.Add(cmbBallOption);
            flowBall.Controls.Add(rowOpt);

            var rowDiam = MkRowFlow();
            rowDiam.Controls.Add(MkLabel("球直径(mm)"));
            nudDiameter = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 500,
                Value = 10,
                DecimalPlaces = 0,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            AutoFitNumericWidth(nudDiameter);
            rowDiam.Controls.Add(nudDiameter);
            flowBall.Controls.Add(rowDiam);

            var rowGeom = MkRowFlow();
            rowGeom.Controls.Add(MkLabel("几何集"));
            txtGeomSet = new TextBox
            {
                Text = "Geometry_Spheres",
                Width = 140,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            rowGeom.Controls.Add(txtGeomSet);
            flowBall.Controls.Add(rowGeom);

            var rowPrefix = MkRowFlow();
            rowPrefix.Controls.Add(MkLabel("名称前缀"));
            txtNamePrefix = new TextBox
            {
                Text = "SPHERE",
                Width = 140,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            rowPrefix.Controls.Add(txtNamePrefix);
            flowBall.Controls.Add(rowPrefix);

            // [需求3] 自定义零件名称（新建 Part 时生效）
            var rowPart = MkRowFlow();
            rowPart.Controls.Add(MkLabel("零件名称"));
            txtBallPartName = new TextBox
            {
                Width = 140,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            rowPart.Controls.Add(txtBallPartName);
            flowBall.Controls.Add(rowPart);
            flowBall.Controls.Add(new Label
            {
                Text = "新建Part时生效，留空用默认名",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 0, 0, 2)
            });

            btnExportBall = MkFuncButton("导出点球", TxClrBall.Color);
            btnExportBall.Margin = new Padding(0, 6, 0, 4);
            btnExportBall.Click += new EventHandler(OnExportBall);
            flowBall.Controls.Add(btnExportBall);

            cardBall.Controls.Add(flowBall);
            stack.Controls.Add(cardBall, 0, 1);

            scroll.Controls.Add(stack);
            return scroll;
        }

        // ════════════════════════════════════════════════════════════
        //  第3列：日志面板
        // ════════════════════════════════════════════════════════════
        private Control BuildCol3()
        {
            var card = MkCard("运行日志", TxClrLog);
            var innerPanel = new Panel { Dock = DockStyle.Fill };

            rtbLog = new RichTextBox
            {
                Dock = DockStyle.Fill,
                BackColor = TxClrLogBg.Color,
                ForeColor = TxClrLogText.Color,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                Font = new Font("Consolas", 8f),
                ScrollBars = RichTextBoxScrollBars.Vertical,
                WordWrap = false
            };

            var btnClear = new FlatColorButton
            {
                Text = "清除日志",
                Dock = DockStyle.Top,
                Height = 22,
                BgColor = TxClrAccent.Color,
                ForeColor = TxColor.TxColorWhite.Color,
                Font = SystemFonts.DefaultFont
            };
            btnClear.Click += delegate { rtbLog.Clear(); };

            innerPanel.Controls.Add(rtbLog); innerPanel.Controls.Add(btnClear);
            card.Controls.Add(innerPanel);
            return card;
        }
        private void UpdateGunInfo()
        {
            if (lblGunInfo == null) return;
            if (_ops == null || _ops.Count == 0)
            {
                lblGunInfo.Text = "焊钳：（待选取）";
                lblGunInfo.ForeColor = SystemColors.GrayText;
                return;
            }

            var names = new System.Collections.Generic.List<string>();
            foreach (var op in _ops)
            {
                Name = PsReader.GetToolNameFromOperation(op);
                if (!string.IsNullOrEmpty(Name))
                    names.Add(Name);
                Log("[调试] 操作 " + op.Name + " 使用焊钳：" + (string.IsNullOrEmpty(Name) ? "(无)" : Name));
            }
            if (names.Count == 0)
            {
                lblGunInfo.Text = "焊钳：（无可用信息）";
                lblGunInfo.ForeColor = SystemColors.GrayText;
                Log("[调试] 所有操作均无焊钳信息");
            }
            else
            {
                lblGunInfo.Text = "焊钳：" + string.Join(", ", names);
                lblGunInfo.ForeColor = SystemColors.GrayText;
                Log("[调试] 共 " + names.Count + " 个不同焊钳信息：" + string.Join(", ", names));

            }
        }

        // ════════════════════════════════════════════════════════════
        //  控件工厂 — 与 RobotReachabilityChecker 风格统一
        // ════════════════════════════════════════════════════════════

        /// <summary>标题栏标签（顶部三列分区标识）— 自绘确保背景色在 PS 宿主下生效</summary>
        private static Label MkHeaderLabel(string text, TxColor bg)
        {
            return new FlatColorLabel
            {
                Text = text,
                Dock = DockStyle.Fill,
                BgColor = bg.Color,
                ForeColor = TxColor.TxColorWhite.Color,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold),
                Margin = new Padding(1, 0, 1, 0)
            };
        }

        /// <summary>卡片容器：自绘 GroupBox，解决 ForeColor 被系统主题忽略的问题</summary>
        private static GroupBox MkCard(string title, TxColor titleColor)
        {
            return new ColoredGroupBox
            {
                Text = title,
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold),
                TitleColor = titleColor.Color,
                Margin = new Padding(2, 2, 2, 4),
                Padding = new Padding(8, 6, 8, 4)
            };
        }

        /// <summary>卡片内部内容面板（与 Checker.MkCardContent 同源）</summary>
        private static FlowLayoutPanel MkCardContent()
        {
            return new FlowLayoutPanel
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
        }

        /// <summary>行内标签（与 Checker.MkLabel 同源）</summary>
        private static Label MkLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Font = SystemFonts.DefaultFont,
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(0, 7, 4, 0)
            };
        }

        /// <summary>水平行流式面板（Label + 控件横排）</summary>
        private static FlowLayoutPanel MkRowFlow()
        {
            return new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 0, 0, 2)
            };
        }
        /// <summary>功能按钮：自绘按钮，绕过 PS 宿主对 Button.BackColor 的劫持</summary>
        private static Button MkFuncButton(string text, Color bgColor)
        {
            return new FlatColorButton
            {
                Text = text,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Height = 26,
                Font = SystemFonts.DefaultFont,
                BgColor = bgColor,
                Margin = new Padding(0, 2, 4, 2),
                Padding = new Padding(8, 2, 8, 2)
            };
        }

        /// <summary>根据 ComboBox 最长项文本自动设定宽度（与 Checker 同源）</summary>
        private void AutoFitComboBoxWidth(ComboBox cb)
        {
            if (cb == null || cb.Items.Count == 0) return;
            using (var g = CreateGraphics())
            {
                float maxW = 0;
                foreach (var item in cb.Items)
                {
                    float w = g.MeasureString(item.ToString(), cb.Font).Width;
                    if (w > maxW) maxW = w;
                }
                cb.Width = (int)maxW + 28;
            }
        }

        /// <summary>根据 NumericUpDown 的 Maximum 位数自动设定宽度（与 Checker 同源）</summary>
        private void AutoFitNumericWidth(NumericUpDown nud)
        {
            if (nud == null) return;
            string maxText = nud.Maximum.ToString("F" + nud.DecimalPlaces);
            using (var g = CreateGraphics())
            {
                float w = g.MeasureString(maxText, nud.Font).Width;
                nud.Width = (int)w + 26;
            }
        }

        // ════════════════════════════════════════════════════════════
        //  PS 数据加载 / 列表管理
        // ════════════════════════════════════════════════════════════
        private void OnPickSel(object sender, EventArgs e)
        {
            Log("[PS] 从PS当前选中拾取...");
            btnPickFromSel.Enabled = false;
            ThreadPool.QueueUserWorkItem(delegate
            {
                List<OperationInfo> ops = null;
                try { ops = _svc.LoadFromSelection(new Action<string>(Log)); }
                catch (Exception ex) { UI(delegate () { Log("[错误] " + ex.Message); }); }
                var final = ops ?? new List<OperationInfo>();

                UI(delegate ()
                {
                    if (_useGrid && _objGrid != null)
                    {
                        // [优化] 把拾取出的每个 op 的 PsObject 推入 grid（API 确认：AppendObject）
                        // _ops 由 ObjectInserted → SyncOpsFromGrid 自动同步
                        int added = 0, skipped = 0;
                        foreach (var op in final)
                        {
                            if (op.PsObject == null) { skipped++; continue; }
                            try { _objGrid.AppendObject(op.PsObject); added++; }
                            catch (Exception ex) { Log("[警告] AppendObject 失败：" + ex.Message); skipped++; }
                        }
                        Log("[PS] 已推入 " + added + " 个对象到列表" + (skipped > 0 ? "（跳过 " + skipped + "）" : ""));
                        // 兜底：若 AppendObject 未触发 ObjectInserted，手动同步
                        if (added > 0) SyncOpsFromGrid();
                    }
                    else
                    {
                        // 回落模式
                        _ops = final; RefreshOpList(); UpdatePointCount();
                    }
                    btnPickFromSel.Enabled = true;
                });
            });
        }

        private void OnClearOps(object sender, EventArgs e)
        {
            if (lvOps != null) lvOps.Items.Clear();
            _ops.Clear();
            if (_objEditOp != null) _objEditOp.Object = null;

            // [需求6] 清空 Grid：API 确认 Objects 属性可写，赋空集合即清空
            if (_objGrid != null)
            {
                try { _objGrid.Objects = new TxObjectList(); }
                catch (Exception ex) { Log("[警告] 清空 grid 失败：" + ex.Message); }
            }

            if (lblListHint != null)
            {
                lblListHint.Text = "列表为空";
                lblListHint.ForeColor = SystemColors.GrayText;
            }
            UpdatePointCount();
            UpdateGunInfo();
        }

        private void RefreshOpList()
        {
            // [需求6] Grid 模式下 lvOps 为 null；回落模式才更新 ListView
            if (lvOps != null)
            {
                lvOps.Items.Clear(); lvOps.Columns[0].Width = -2;
                foreach (OperationInfo op in _ops)
                {
                    ListViewItem item = new ListViewItem(op.Name);
                    item.Checked = true; item.Tag = op;
                    lvOps.Items.Add(item);
                }
            }
            if (lblListHint != null)
            {
                lblListHint.Text = _ops.Count > 0 ? "共 " + _ops.Count + " 个操作" : "列表为空";
                lblListHint.ForeColor = _ops.Count > 0 ? ClrStatusOk : SystemColors.GrayText;
            }
            if (_ops.Count > 0)
            {
                Log("[PS] 加载 " + _ops.Count + " 个操作");
                LogOperationTools();
                if (_refFrameName == "世界坐标系" && _refFrameMatrix == null)
                    AutoLoadRefFrameFromOps();
            }
        }

        private void LogOperationTools()
        {
            var ops = new List<OperationInfo>(_ops);
            ThreadPool.QueueUserWorkItem(delegate
            {
                foreach (OperationInfo op in ops)
                {
                    string toolName = null;
                    try { _svc.InvokeOnPs(delegate () { toolName = PsReader.GetToolNameFromOperation(op); }); } catch { }
                    string on = op.Name; string tn = toolName;
                    UI(delegate ()
                    { Log("[工具] " + on + " → " + (string.IsNullOrEmpty(tn) ? "未绑定" : tn)); });
                }
            });
        }

        private void AutoLoadRefFrameFromOps()
        {
            var ops = new List<OperationInfo>(_ops);
            var ptType = GetPtType();
            var useMfg = chkUseMfgName.Checked;

            ThreadPool.QueueUserWorkItem(delegate
            {
                PsReader.RefFrameResult rf = null;
                var logBuf = new List<string>();
                Action<string> collect = delegate (string s) { lock (logBuf) logBuf.Add(s); };

                try
                {
                    _svc.InvokeOnPs(delegate ()
                    {
                        foreach (var op in ops)
                        {
                            if (op.Points == null || op.Points.Count == 0)
                                PsReader.FillPoints(op, ptType, useMfg, collect);

                            rf = PsReader.ResolveOperationRefFrame(op, fallbackToWorld: false, log: collect);
                            if (rf != null && rf.Source != PsReader.RefFrameSource.None
                                           && rf.Source != PsReader.RefFrameSource.World)
                                break;
                        }
                    });
                }
                catch (Exception ex) { collect("[坐标] PS 线程异常：" + ex.Message); }

                var rfLocal = rf;
                var bufLocal = logBuf;
                UI(delegate ()
                {
                    foreach (var line in bufLocal) Log(line);

                    if (rfLocal == null || rfLocal.Source == PsReader.RefFrameSource.None
                                         || rfLocal.Source == PsReader.RefFrameSource.World)
                    {
                        Log("[坐标] 未找到焊点的零件/外观绑定，保持世界坐标系（可在 PS 中选择 Component/Frame 手动覆盖）");
                        return;
                    }

                    if (_refFrameName != "世界坐标系" || _refFrameMatrix != null) return;

                    // 一致性检查：若需要用户选择，弹窗
                    if (rfLocal.NeedsUserChoice)
                    {
                        var picked = AskUserPickRefFrame(rfLocal);
                        if (picked == null)
                        {
                            Log("[坐标] 用户取消，保持世界坐标系");
                            return;
                        }
                        _refFrameName = picked.Item1;
                        _refFrameMatrix = picked.Item2;
                        UpdateRefCoordStatus(picked.Item1, false);
                        Log("[坐标] 用户选择参考坐标：" + picked.Item1);
                        return;
                    }

                    _refFrameName = rfLocal.Name;
                    _refFrameMatrix = rfLocal.Matrix;
                    UpdateRefCoordStatus(rfLocal.Name, false);
                    string srcLabel;
                    switch (rfLocal.Source)
                    {
                        case PsReader.RefFrameSource.Appearance: srcLabel = "外观"; break;
                        case PsReader.RefFrameSource.Part: srcLabel = "零件"; break;
                        case PsReader.RefFrameSource.Fixture: srcLabel = "夹具"; break;
                        default: srcLabel = rfLocal.Source.ToString(); break;
                    }
                    string pointHint = string.IsNullOrEmpty(rfLocal.PointName) ? "" : "，基于焊点 " + rfLocal.PointName;
                    Log("[坐标] 自动获取参考坐标：" + rfLocal.Name + "（来源=" + srcLabel + pointHint + "）");
                });
            });
        }

        /// <summary>
        /// 当零件与夹具坐标不一致时弹窗让用户选择。
        /// 返回 (显示名, 4x4矩阵)；取消返回 null。
        /// </summary>
        private Tuple<string, double[]> AskUserPickRefFrame(PsReader.RefFrameResult rf)
        {
            var options = new List<Tuple<string, double[], string>>(); // (显示名, 矩阵, 原始名)
            if (rf.PartCandidates != null)
                foreach (var a in rf.PartCandidates)
                    options.Add(Tuple.Create(a.Name + "（零件）", a.Matrix, a.Name));
            if (rf.FixtureCandidates != null)
                foreach (var a in rf.FixtureCandidates)
                    options.Add(Tuple.Create(a.Name + "（夹具）", a.Matrix, a.Name));

            if (options.Count == 0) return null;
            if (options.Count == 1)
                return Tuple.Create(options[0].Item1, options[0].Item2);

            using (var dlg = new Form())
            {
                dlg.Text = "选择参考坐标";
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                dlg.MinimizeBox = false; dlg.MaximizeBox = false;
                dlg.Width = 520; dlg.Height = 340;

                var lbl = new Label
                {
                    Text = "检测到零件与夹具坐标不一致：\r\n" + (rf.ConflictReason ?? "") + "\r\n\r\n请选择参考坐标：",
                    Dock = DockStyle.Top,
                    Height = 90,
                    Padding = new Padding(10)
                };
                var lb = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false };
                foreach (var opt in options) lb.Items.Add(opt.Item1);
                lb.SelectedIndex = 0;

                var pnlBtn = new Panel { Dock = DockStyle.Bottom, Height = 44 };
                var btnOk = new Button { Text = "确定", DialogResult = DialogResult.OK, Width = 80, Top = 8, Left = pnlBtn.Width - 180, Anchor = AnchorStyles.Top | AnchorStyles.Right };
                var btnCancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Width = 80, Top = 8, Left = pnlBtn.Width - 90, Anchor = AnchorStyles.Top | AnchorStyles.Right };
                pnlBtn.Controls.Add(btnOk); pnlBtn.Controls.Add(btnCancel);
                dlg.AcceptButton = btnOk; dlg.CancelButton = btnCancel;

                dlg.Controls.Add(lb);
                dlg.Controls.Add(pnlBtn);
                dlg.Controls.Add(lbl);

                if (dlg.ShowDialog(this) != DialogResult.OK) return null;
                int idx = lb.SelectedIndex;
                if (idx < 0 || idx >= options.Count) return null;
                return Tuple.Create(options[idx].Item1, options[idx].Item2);
            }
        }

        private List<OperationInfo> GetCheckedOps()
        {
            // [需求6] Grid 模式：无勾选框，直接取全部 _ops（等同全选）
            if (_useGrid || lvOps == null) return new List<OperationInfo>(_ops);

            // 回落模式：从 ListView 取勾选项
            var list = new List<OperationInfo>();
            foreach (ListViewItem item in lvOps.Items)
                if (item.Checked && item.Tag is OperationInfo)
                    list.Add((OperationInfo)item.Tag);
            return list;
        }

        private void UpdatePointCount()
        {
            try
            {
                var ops = GetCheckedOps();
                if (ops.Count == 0)
                { lblPointCount.Text = "将导出点数量：0（列表为空）"; lblPointCount.ForeColor = TxClrBtnDanger.Color; return; }
                int n = _svc.PreviewPointCount(ops, GetPtType(), chkUseMfgName.Checked, delegate { });
                lblPointCount.Text = "将导出点数量：" + n;
                lblPointCount.ForeColor = n > 0 ? ClrStatusOk : TxClrBtnDanger.Color;
            }
            catch { }
        }

        private PointType GetPtType()
        {
            switch (cmbPointType.SelectedIndex)
            {
                case 0: return PointType.WeldPoint;
                case 1: return PointType.PathPoint;
                case 2: return PointType.ContinuousPoint;
                default: return PointType.All;
            }
        }

        // ════════════════════════════════════════════════════════════
        //  事件处理（导出等）
        // ════════════════════════════════════════════════════════════
        private void OnBrowseGun(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "选择焊枪数模";
                dlg.Filter = "CATIA文件|*.CATProduct;*.CATPart;*.cgr|所有文件|*.*";
                if (dlg.ShowDialog() == DialogResult.OK)
                { txtGunModel.Tag = dlg.FileName; txtGunModel.Text = Path.GetFileName(dlg.FileName); }
            }
        }

        private void OnExportExcel(object sender, EventArgs e)
        {
            var ops = GetCheckedOps();
            if (ops.Count == 0) { MessageBox.Show("请先添加操作到列表。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            SetBusy(true); Log("[Excel] 开始导出，参考坐标：" + _refFrameName);
            _svc.ExportExcelAsync(ops, GetPtType(), chkUseMfgName.Checked,
                DefaultOut(), _refFrameMatrix, _refFrameName,
                new Action<string>(msg => UI(delegate () { Log(msg); })),
                new Action<bool, string>((ok, result) => UI(delegate ()
                {
                    SetBusy(false);
                    if (ok) { Log("✓ Excel：" + result); if (MessageBox.Show("成功：\n" + result + "\n\n打开？", "完成", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes) try { System.Diagnostics.Process.Start(result); } catch { } }
                    else { Log("✗ 失败：" + result); MessageBox.Show("失败：\n" + result, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                })));
        }

        private void OnExportGun(object sender, EventArgs e)
        {
            var ops = GetCheckedOps();
            if (ops.Count == 0) { MessageBox.Show("请先添加操作到列表。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            string globalCustomPath = chkCustomGun.Checked
                ? (txtGunModel.Tag as string ?? txtGunModel.Text)
                : null;
            bool hasGlobalCustom = !string.IsNullOrEmpty(globalCustomPath) && File.Exists(globalCustomPath);

            // [需求1 & 2] CGR 前置校验：逐个操作在同级目录下做模糊匹配
            const double HIGH_CONFIDENCE = 0.75;
            var perOpPaths = new Dictionary<string, string>();

            foreach (var op in ops)
            {
                if (hasGlobalCustom)
                {
                    perOpPaths[op.Name] = globalCustomPath;
                    continue;
                }

                PsReader.CgrLookupResult lookup = null;
                try
                {
                    _svc.InvokeOnPs(delegate ()
                    {
                        lookup = PsReader.LookupCgrForOperation(op, msg => UI(delegate () { Log(msg); }));
                    });
                }
                catch (Exception ex) { Log("[错误] CGR 查找异常：" + ex.Message); }

                // [需求2] 未找到同级目录 CGR → 弹窗告知，不进入导出
                if (lookup == null || lookup.Candidates == null || lookup.Candidates.Count == 0)
                {
                    string reason = lookup?.Reason ?? "未知原因";
                    string dir = lookup?.SameDir;
                    string toolName = lookup?.ToolName ?? op.Name;
                    string msg = "操作 [" + op.Name + "] 的工具 [" + toolName + "] 未能找到 CGR。\n"
                        + (!string.IsNullOrEmpty(dir) ? ("同级目录：" + dir + "\n") : "")
                        + "原因：" + reason + "\n\n"
                        + "请检查目录或勾选 [自定义焊枪数模] 手动选择后重试。";
                    MessageBox.Show(msg, "未找到 CGR", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Log("[CGR] 已中止：" + reason);
                    return;
                }

                var best = lookup.Candidates[0];

                // [需求1] 高相似度自动采用；低相似度让用户确认
                if (best.Similarity >= HIGH_CONFIDENCE)
                {
                    perOpPaths[op.Name] = best.Path;
                    Log("[CGR] " + op.Name + " → " + Path.GetFileName(best.Path)
                        + "（相似度 " + (best.Similarity * 100).ToString("F0") + "%）");
                }
                else
                {
                    string chosen = PromptCgrSelection(op.Name, lookup.ToolName, lookup.SameDir, lookup.Candidates);
                    if (string.IsNullOrEmpty(chosen))
                    {
                        Log("[CGR] 用户取消，已中止导出");
                        return;
                    }
                    perOpPaths[op.Name] = chosen;
                    Log("[CGR] " + op.Name + " → " + Path.GetFileName(chosen) + "（用户确认）");
                }
            }

            SetBusy(true);
            var p = new GunExportParams
            {
                Operations = ops,
                ExportTCP = chkExportTCP.Checked,
                GunOriginAtTCP = chkGunOriginTCP.Checked,
                CustomModelPath = hasGlobalCustom ? globalCustomPath : null,
                PerOpModelPaths = perOpPaths,
                CustomProductName = string.IsNullOrWhiteSpace(txtGunProductName.Text) ? null : txtGunProductName.Text.Trim(),
                Format = ExportFormat.Xml3d,                    // [需求8] 固定 3dxml
                OutputPath = DefaultOut(),                       // [需求8] 无输出目录按钮，用默认路径
                RefMatrix = _refFrameMatrix,
                RefName = _refFrameName
            };
            _svc.ExportGunsAsync(p, new Action<string>(msg => UI(delegate () { Log(msg); })),
                new Action<ExportProgress>(pg => UI(delegate () { SetProgress(pg); })),
                new Action<bool, string>((ok, msg) => UI(delegate ()
                { SetBusy(false); Log(ok ? "✓ " + msg : "✗ " + msg); ShowResult(ok, msg); })));
        }

        // ════════════════════════════════════════════════════════════
        //  CGR 候选选择对话框（[需求1] 低相似度时弹出）
        // ════════════════════════════════════════════════════════════
        private string PromptCgrSelection(string opName, string toolName, string sameDir,
            List<PsReader.CgrMatch> candidates)
        {
            using (var dlg = new Form())
            {
                dlg.Text = "请确认 CGR 文件 — " + opName;
                dlg.Size = new Size(560, 380);
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                dlg.MaximizeBox = false; dlg.MinimizeBox = false;
                dlg.BackColor = SystemColors.Control;

                var header = new Label
                {
                    Text = "工具名称：" + (toolName ?? "(未知)")
                        + "\n同级目录：" + (sameDir ?? "(未解析)")
                        + "\n未找到高相似度的 CGR 文件，请从以下候选中选择，或手动浏览：",
                    Dock = DockStyle.Top,
                    Height = 60,
                    Padding = new Padding(10, 8, 10, 0),
                    Font = SystemFonts.DefaultFont,
                    ForeColor = SystemColors.ControlText
                };

                var lb = new ListBox
                {
                    Dock = DockStyle.Fill,
                    IntegralHeight = false,
                    Font = SystemFonts.DefaultFont,
                    BorderStyle = BorderStyle.FixedSingle
                };
                foreach (var c in candidates)
                {
                    string sim = (c.Similarity * 100).ToString("F0");
                    lb.Items.Add(c.CandidateName + ".cgr   (相似度 " + sim + "%)");
                }
                if (lb.Items.Count > 0) lb.SelectedIndex = 0;

                var pnlBtns = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    Height = 44,
                    FlowDirection = FlowDirection.RightToLeft,
                    Padding = new Padding(10, 6, 10, 6),
                    BackColor = SystemColors.Control
                };

                bool browseFallback = false;
                var btnOk = MkFuncButton("使用选中", TxClrBtnPrimary.Color);
                var btnBrowse = MkFuncButton("手动浏览...", TxClrBtnSecondary.Color);
                var btnCancel = MkFuncButton("取消", TxClrBtnDanger.Color);
                btnOk.DialogResult = DialogResult.OK;
                btnCancel.DialogResult = DialogResult.Cancel;
                btnBrowse.Click += delegate
                {
                    browseFallback = true;
                    dlg.DialogResult = DialogResult.OK;
                    dlg.Close();
                };

                pnlBtns.Controls.Add(btnCancel);
                pnlBtns.Controls.Add(btnBrowse);
                pnlBtns.Controls.Add(btnOk);

                dlg.Controls.Add(lb);
                dlg.Controls.Add(pnlBtns);
                dlg.Controls.Add(header);
                dlg.AcceptButton = btnOk;
                dlg.CancelButton = btnCancel;

                if (dlg.ShowDialog(this) != DialogResult.OK) return null;

                if (browseFallback)
                {
                    using (var ofd = new OpenFileDialog())
                    {
                        ofd.Title = "为 [" + opName + "] 选择 CGR";
                        ofd.Filter = "CATIA 文件|*.CATProduct;*.CATPart;*.cgr|所有文件|*.*";
                        if (!string.IsNullOrEmpty(sameDir) && Directory.Exists(sameDir))
                            ofd.InitialDirectory = sameDir;
                        if (ofd.ShowDialog(this) != DialogResult.OK) return null;
                        return ofd.FileName;
                    }
                }

                int idx = lb.SelectedIndex;
                if (idx < 0 || idx >= candidates.Count) return null;
                return candidates[idx].Path;
            }
        }

        private void OnExportBall(object sender, EventArgs e)
        {
            var ops = GetCheckedOps();
            if (ops.Count == 0) { MessageBox.Show("请先添加操作到列表。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            SetBusy(true);
            var optMap = new[] { BallExportOption.TrajectoryAndBall, BallExportOption.TrajectoryOnly, BallExportOption.BallOnly };
            var p = new BallExportParams
            {
                Operations = ops,
                ExportToCurrentDoc = cmbBallTarget.SelectedIndex == 0,
                Option = optMap[cmbBallOption.SelectedIndex],
                BallDiameter = (double)nudDiameter.Value,
                OutputPath = DefaultOut(),                       // [需求8] 无输出目录按钮，用默认路径
                PointFilter = GetPtType(),
                UseMfgName = chkUseMfgName.Checked,
                GeomSetName = string.IsNullOrWhiteSpace(txtGeomSet.Text) ? "Geometry_Spheres" : txtGeomSet.Text.Trim(),
                NamePrefix = string.IsNullOrWhiteSpace(txtNamePrefix.Text) ? "SPHERE" : txtNamePrefix.Text.Trim(),
                CustomPartName = string.IsNullOrWhiteSpace(txtBallPartName.Text) ? null : txtBallPartName.Text.Trim(),  // [需求3]
                RefMatrix = _refFrameMatrix,
                RefName = _refFrameName
            };
            _svc.ExportBallsAsync(p, new Action<string>(msg => UI(delegate () { Log(msg); })),
                new Action<ExportProgress>(pg => UI(delegate () { SetProgress(pg); })),
                new Action<bool, string>((ok, msg) => UI(delegate ()
                { SetBusy(false); Log(ok ? "✓ " + msg : "✗ " + msg); ShowResult(ok, msg); })));
        }

        private void OnReset(object sender, EventArgs e)
        {
            chkExportTCP.Checked = false; chkGunOriginTCP.Checked = true;
            chkCustomGun.Checked = false; txtGunModel.Text = ""; txtGunModel.Tag = null;
            if (txtGunProductName != null) txtGunProductName.Text = "";   // [需求3] 清空自定义产品名
            cmbPointType.SelectedIndex = 3;
            cmbBallTarget.SelectedIndex = 0; cmbBallOption.SelectedIndex = 0;
            nudDiameter.Value = 10; txtGeomSet.Text = "Geometry_Spheres";
            txtNamePrefix.Text = "SPHERE"; chkUseMfgName.Checked = false;
            if (txtBallPartName != null) txtBallPartName.Text = "";       // [需求3] 清空自定义零件名
            _refFrameName = "世界坐标系"; _refFrameMatrix = null;
            UpdateRefCoordStatus("世界坐标系", true);
            if (_frameCombo != null) try { _frameCombo.Clear(); } catch { }
            if (_objEditOp != null) _objEditOp.Object = null;
            // [需求6] 清空 Grid：赋空 TxObjectList
            if (_objGrid != null)
            {
                try { _objGrid.Objects = new TxObjectList(); }
                catch (Exception ex) { Log("[警告] 清空 grid 失败：" + ex.Message); }
            }
            progressBar.Value = 0; lblProgress.Text = "";
            if (lvOps != null) lvOps.Items.Clear();
            _ops.Clear();
            if (lblListHint != null) { lblListHint.Text = "列表为空"; lblListHint.ForeColor = SystemColors.GrayText; }
            Log("[操作] 已复位所有设置");
        }

        // ════════════════════════════════════════════════════════════
        //  UI 辅助
        // ════════════════════════════════════════════════════════════
        private void SetBusy(bool busy)
        {
            btnExportGun.Enabled = !busy; btnExportBall.Enabled = !busy;
            btnPickFromSel.Enabled = !busy; btnExportExcel.Enabled = !busy;
            // [修复] 进度条按需显示
            progressBar.Visible = busy;
            if (busy) { progressBar.Value = 0; lblProgress.Text = "运行中..."; }
            else { lblProgress.Text = "就绪"; }
        }

        private void SetProgress(ExportProgress p)
        {
            if (p.Total <= 0) return;
            progressBar.Value = Math.Min((int)((double)p.Current / p.Total * 100), 100);
            lblProgress.Text = p.Current + "/" + p.Total + "  " + p.CurrentItem;
        }

        private void Log(string msg)
        {
            if (rtbLog == null || IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action<string>(Log), msg); return; }
            Color c;
            if (msg.StartsWith("✓") || msg.Contains("完成")) c = TxClrLogOk.Color;
            else if (msg.StartsWith("✗") || msg.Contains("失败") || msg.Contains("错误")) c = TxClrLogErr.Color;
            else if (msg.Contains("⚠")) c = TxClrLogWarn.Color;
            else if (msg.StartsWith("[PS]")) c = TxClrLogPs.Color;
            else if (msg.StartsWith("[坐标]")) c = Color.FromArgb(160, 200, 255);
            else if (msg.StartsWith("[Excel]")) c = Color.FromArgb(180, 228, 160);
            else c = TxClrLogText.Color;
            rtbLog.SelectionStart = rtbLog.TextLength; rtbLog.SelectionLength = 0;
            rtbLog.SelectionColor = c;
            rtbLog.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + msg + Environment.NewLine);
            rtbLog.ScrollToCaret();
        }

        private void UI(Action act) { if (IsDisposed) return; if (InvokeRequired) BeginInvoke(act); else act(); }
        private void ShowResult(bool ok, string msg) { MessageBox.Show(msg, ok ? "完成" : "失败", MessageBoxButtons.OK, ok ? MessageBoxIcon.Information : MessageBoxIcon.Error); }

        private void ShowHelp()
        {
            MessageBox.Show(
                "导出插枪 / 点云到 CATIA\n\n" +
                "【参考坐标】使用PS原生坐标选择控件\n" +
                "【拾取操作】方式1: 在上方原生列表内拾取，自动加入\n" +
                "          方式2: 在PS选中后点击[拾取自PS]\n" +
                "【导出】列表内的操作即为导出对象，点击底栏按钮",
                "帮助", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ════════════════════════════════════════════════════════════
        //  窗体关闭
        // ════════════════════════════════════════════════════════════
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_frameCombo != null)
            {
                try { _frameCombo.ValidFrameSet -= new TxFrameComboBoxCtrl_ValidFrameSetEventHandler(OnFrameValidSet); } catch { }
                try { _frameCombo.InvalidFrameSet -= new TxFrameComboBoxCtrl_InvalidFrameSetEventHandler(OnFrameInvalidSet); } catch { }
                try { _frameCombo.Picked -= new TxFrameComboBoxCtrl_PickedEventHandler(OnFramePicked); } catch { }
            }
            if (_objEditOp != null)
                try { _objEditOp.Picked -= new TxObjEditBoxCtrl_PickedEventHandler(OnObjEditPicked); } catch { }
            // [需求6] 解绑 _objGrid 事件
            if (_objGrid != null)
            {
                try { _objGrid.ObjectInserted -= new TxObjGridCtrl_ObjectInsertedEventHandler(OnGridObjectInserted); } catch { }
                try { _objGrid.RowDeleted -= new TxObjGridCtrl_RowDeletedEventHandler(OnGridRowDeleted); } catch { }
            }
            if (_svc != null) _svc.Dispose();
            base.OnFormClosing(e);
        }

        private static string DefaultOut() => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "CatiaExport");

        private void InitializeComponent()
        { SuspendLayout(); ClientSize = new Size(1200, 850); Name = "ExportGunForm"; ResumeLayout(false); }

        // ════════════════════════════════════════════════════════════
        //  自绘控件 — 绕开 PS 宿主主题对 BackColor/ForeColor 的劫持
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 自绘按钮，绕过 WinForms 主题覆盖 BackColor。
        /// 支持 Hover/Pressed/Disabled 三态，AutoSize 沿用 Button 基类行为。
        /// </summary>
        private class FlatColorButton : Button
        {
            public Color BgColor { get; set; } = Color.FromArgb(0, 100, 167);
            private bool _hover;
            private bool _pressed;

            public FlatColorButton()
            {
                SetStyle(ControlStyles.UserPaint
                       | ControlStyles.AllPaintingInWmPaint
                       | ControlStyles.OptimizedDoubleBuffer
                       | ControlStyles.ResizeRedraw
                       | ControlStyles.SupportsTransparentBackColor, true);
                FlatStyle = FlatStyle.Flat;
                FlatAppearance.BorderSize = 0;
                ForeColor = Color.White;
                Cursor = Cursors.Hand;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                Color fill;
                if (!Enabled)
                    fill = Color.FromArgb(
                        (BgColor.R + 255) / 2,
                        (BgColor.G + 255) / 2,
                        (BgColor.B + 255) / 2);   // 淡化表示禁用
                else if (_pressed)
                    fill = ControlPaint.Dark(BgColor, 0.15f);
                else if (_hover)
                    fill = ControlPaint.Light(BgColor, 0.25f);
                else
                    fill = BgColor;

                e.Graphics.Clear(fill);
                var textColor = Enabled ? ForeColor : Color.FromArgb(230, 230, 230);
                TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, textColor,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter
                    | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis);
            }

            protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
            protected override void OnMouseLeave(EventArgs e) { _hover = false; _pressed = false; Invalidate(); base.OnMouseLeave(e); }
            protected override void OnMouseDown(MouseEventArgs e) { _pressed = true; Invalidate(); base.OnMouseDown(e); }
            protected override void OnMouseUp(MouseEventArgs e) { _pressed = false; Invalidate(); base.OnMouseUp(e); }
            protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }
        }

        /// <summary>
        /// 自绘标签，绕过 WinForms 主题覆盖 BackColor。
        /// 用于顶部分区标题条（通用信息 / 导插枪 / 日志）。
        /// </summary>
        private class FlatColorLabel : Label
        {
            public Color BgColor { get; set; } = Color.FromArgb(0, 100, 140);

            public FlatColorLabel()
            {
                SetStyle(ControlStyles.UserPaint
                       | ControlStyles.AllPaintingInWmPaint
                       | ControlStyles.OptimizedDoubleBuffer
                       | ControlStyles.ResizeRedraw, true);
                ForeColor = Color.White;
                TextAlign = ContentAlignment.MiddleCenter;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.Clear(BgColor);
                TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, ForeColor,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
            }
        }

        /// <summary>
        /// 自绘 GroupBox，绕过 WinForms 视觉主题对 GroupBox.ForeColor 的忽略。
        /// 保持 AutoSize 行为（尺寸仍由 GroupBox 基类计算）。
        /// </summary>
        private class ColoredGroupBox : GroupBox
        {
            public Color TitleColor { get; set; } = Color.Black;
            public Color BorderColor { get; set; } = Color.FromArgb(200, 200, 200);

            public ColoredGroupBox()
            {
                SetStyle(ControlStyles.UserPaint
                       | ControlStyles.AllPaintingInWmPaint
                       | ControlStyles.OptimizedDoubleBuffer
                       | ControlStyles.ResizeRedraw, true);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics;
                string title = Text ?? "";
                Size textSize = TextRenderer.MeasureText(g, title, Font, Size.Empty, TextFormatFlags.NoPadding);
                int halfH = textSize.Height / 2;

                // 背景（传承父级背景，避免黑块）
                using (var bgBrush = new SolidBrush(BackColor))
                    g.FillRectangle(bgBrush, ClientRectangle);

                // 边框（留出顶部标题一半高度）
                var borderRect = new Rectangle(0, halfH, Width - 1, Height - halfH - 1);
                using (var borderPen = new Pen(BorderColor))
                    g.DrawRectangle(borderPen, borderRect);

                if (!string.IsNullOrEmpty(title))
                {
                    // 标题覆盖边框线：先涂底色，再写字
                    var titleRect = new Rectangle(8, 0, textSize.Width + 6, textSize.Height);
                    using (var bgBrush = new SolidBrush(BackColor))
                        g.FillRectangle(bgBrush, titleRect);
                    TextRenderer.DrawText(g, title, Font,
                        new Point(titleRect.X + 3, titleRect.Y), TitleColor,
                        TextFormatFlags.NoPadding);
                }
            }
        }
    }
}