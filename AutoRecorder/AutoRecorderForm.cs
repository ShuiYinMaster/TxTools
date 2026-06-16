using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;
using TxTools.Common;

// System.Threading.Timer 和 System.Windows.Forms.Timer 命名冲突，显式选 Forms 那个
using WinFormsTimer = System.Windows.Forms.Timer;

namespace TxTools.AutoRecorder
{
    /// <summary>
    /// 操作录像主窗口。
    ///
    /// GUI v2（本次重构）：
    ///   · 原版整窗绝对坐标手排（Left/Top/y+=），折叠展开靠 RelayoutBelowKeyframePanel
    ///     遍历平移所有下方控件 + 硬编码 798/928 高度常量，高 DPI 下脆弱且难维护。
    ///   · 现改为：单列根 TableLayoutPanel（AutoSize 行），关键帧卡 / 日志卡折叠展开
    ///     只改自身 Height，下方内容自动回流；窗体高度由 UpdateFormHeight 按内容实测
    ///     （常量法兜底），RelayoutBelowKeyframePanel 整个删除。
    ///   · 录像参数卡内部改为 TableLayoutPanel 参数行（标签列 AutoSize + 内容列 100%），
    ///     替换逐控件像素定位，200% DPI 下对齐稳定。
    ///   · 保持 600 宽窄竖条 + 右上角定位（不遮挡 3D 主视口拾取的设计意图不变）。
    ///   · 录制服务、关键帧编辑、批量任务等业务逻辑零改动。
    /// </summary>
    public class AutoRecorderForm : TxForm
    {
        // ===== 根布局 =====
        private TableLayoutPanel _root;

        // ===== 操作选择 =====
        private TxObjGridCtrl _grid;
        private Label _lblOpInfo;
        private Button _btnSaveView;        // "📷 记录当前视角"（起始视角）
        private CheckBox _chkAutoApplyView;   // "拾取时自动定位视角"

        // op → 该 op 的视角脚本（起始 + 关键帧）
        private readonly System.Collections.Generic.Dictionary<ITxObject, OperationViewSchedule> _opSchedules
            = new System.Collections.Generic.Dictionary<ITxObject, OperationViewSchedule>();
        private ITxObject _lastPickedOp;

        // ===== 关键帧面板（折叠） =====
        private Panel _grpKf;            // 折叠卡片（用 Panel 不用 GroupBox，避免 Title 区占高）
        private Button _btnKfToggle;      // 标题栏点击展开/收起
        private bool _kfExpanded;       // 当前是否展开
        private ComboBox _cmbKfOpSelector;  // 当前编辑哪个 op
        private DataGridView _gridKf;           // location 表
        private Button _btnKfSet;         // 用当前视角设定/更新
        private Button _btnKfPreview;     // 预览
        private Button _btnKfClear;       // 清除
        private Label _lblKfStatus;      // 状态行
        private List<object> _kfLocations;    // 当前编辑 op 的 location 缓存（PM items 或 ITxObject）
        private ITxObject _kfEditingOp;      // 当前正在编辑的 op（独立于主拾取）

        // ===== 录像参数 =====
        private Label _lblPath;          // "输出文件：" / "输出目录："
        private TextBox _txtPath;
        private bool _isPathInDirMode;  // true=目录模式(多操作), false=单文件
        private Button _btnBrowse;
        private ComboBox _cmbCodec;
        private NumericUpDown _numFps;
        private NumericUpDown _numCompression;
        private ComboBox _cmbTimeSource;
        private CheckBox _chkCustomRes;
        private NumericUpDown _numResW;
        private NumericUpDown _numResH;
        private CheckBox _chkFocus;
        private ComboBox _cmbSpeedup;
        private Label _lblSpeedupHint;

        // ===== 主按钮 =====
        private Button _btnStart;
        private Button _btnCancel;
        private Button _btnClose;

        // ===== 日志 =====
        private Button _btnToggleLog;
        private Panel _pnlLogContainer;
        private RichTextBox _rtbLog;
        private bool _logExpanded;

        // ===== 状态栏 =====
        private StatusStrip _statusStrip;
        private ToolStripStatusLabel _lblStatus;
        private ToolStripStatusLabel _lblTimer;

        // ===== 服务 + 计时 =====
        private RecordingService _svc;
        private WinFormsTimer _uiTimer;
        private double _tickElapsed;

        // PS 主线程的 SynchronizationContext（用于把 player 事件回调切回 UI 线程）
        private readonly SynchronizationContext _psCtx;

        // 布局常量（仅作初始值 / 兜底；实际高度由 UpdateFormHeight 按内容实测）
        private const int PAD = 8;
        private const int FORM_W = 520;
        private const int CARD1_H = 258;
        private const int LOG_H = 130;
        private const int FORM_H_FALLBACK = 798;   // 全折叠时的兜底高度（沿用原版实测值）
        private static readonly Size _designSize = new Size(FORM_W, FORM_H_FALLBACK);
        private bool _dpiApplied;

        public AutoRecorderForm() : this(null) { }

        public AutoRecorderForm(SynchronizationContext psCtx)
        {
            SemiModal = false;
            _psCtx = psCtx
                     ?? SynchronizationContext.Current
                     ?? new SynchronizationContext();

            // 统一窗体规范 + DPI（套件唯一一处缩放设置，详见 FormUiKit）。
            // 串扰修复：InitStandardForm 内部已按 ExportGun 方案设唯一 Name，
            // 各窗体几何独立持久化、互不干扰。本窗体固定尺寸、折叠展开由 ClientSize 驱动。
            FormUiKit.InitStandardForm(this, "TxTools 操作录像",
                _designSize, Size.Empty, sizable: false);
            this.StartPosition = FormStartPosition.Manual;   // 自行 PositionTopRight，不挡主视口
            this.ClientSize = new Size(FORM_W, FORM_H_FALLBACK);

            _svc = new RecordingService(_psCtx);
            _svc.Log += OnSvcLog;
            _svc.StateChanged += OnSvcState;
            _svc.JobProgress += OnSvcJobProgress;





            _uiTimer = new WinFormsTimer { Interval = 200 };
            _uiTimer.Tick += OnUiTick;

            BuildUi();
            BindEvents();
        }

        public override void OnInitTxForm()
        {
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            FormUiKit.ApplyDpiScaling(this, ref _dpiApplied, _designSize);
            LoadCodecs();
            LoadTimeSources();
            UpdateFormHeight();      // DPI 缩放后按内容实测重设窗体高度
            PositionTopRight();      // 不挡主视口
            ApplyViewerSizeDefault();// 把视口尺寸填到分辨率输入框作为默认
            RefreshButtonState();

            // 让焦点初始就在 grid 上，用户打开窗口就能直接到主视口拾取
            try { this.ActiveControl = _grid; _grid.Focus(); } catch { }
        }

        /// <summary>窗口移到当前屏幕右上角，避免遮挡 3D 主视口的拾取。</summary>
        private void PositionTopRight()
        {
            try
            {
                var screen = Screen.FromControl(this) ?? Screen.PrimaryScreen;
                var wa = screen.WorkingArea;
                this.Location = new Point(
                    wa.Right - this.Width - 20,
                    wa.Top + 60);
            }
            catch { }
        }

        /// <summary>读取主视口当前像素尺寸，填到 ResW/ResH 输入框作为默认。</summary>
        private void ApplyViewerSizeDefault()
        {
            try
            {
                var v = PsReader.GetGraphicViewer();
                int w, h;
                if (v == null || !PsReader.TryGetViewerSize(v, out w, out h)) return;
                if (w <= 0 || h <= 0) return;

                int w4 = Math.Max(4, (w / 4) * 4);
                int h4 = Math.Max(4, (h / 4) * 4);
                if (w4 >= _numResW.Minimum && w4 <= _numResW.Maximum)
                    _numResW.Value = w4;
                if (h4 >= _numResH.Minimum && h4 <= _numResH.Maximum)
                    _numResH.Value = h4;
            }
            catch { }
        }

        // ============================================================
        // UI 构建 —— 单列根 TableLayoutPanel，行序：
        //   [0] 操作选择卡（固定高）
        //   [1] 视角关键帧卡（AutoSize，Height 切换 28/288）
        //   [2] 录像参数卡（AutoSize，内部参数表自撑）
        //   [3] 按钮行
        //   [4] 日志折叠按钮
        //   [5] 日志容器（AutoSize，Height 切换 0/130）
        // StatusStrip 独立 Dock=Bottom。
        // ============================================================
        private void BuildUi()
        {
            _root = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                Padding = new Padding(PAD, PAD, PAD, 0),
                Margin = Padding.Empty
            };
            _root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            // ---------- 卡片 1：操作选择（固定高，grid Fill） ----------
            var grp1 = NewGroupBox("操作选择（多选批量；每个操作可独立设置视角）");
            grp1.Height = CARD1_H;
            grp1.Padding = new Padding(10, 20, 10, 8);

            _grid = new TxObjGridCtrl { Dock = DockStyle.Fill };
            try
            {
                _grid.ListenToPick = true;
                _grid.EnableMultipleSelection = true;
                _grid.EnableRecurringObjects = false;
            }
            catch { }

            // 底部：信息行 + 视角控制行（表格自排，替代绝对坐标）
            var grp1Bottom = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Padding = new Padding(0, 4, 0, 0)
            };
            grp1Bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            grp1Bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            grp1Bottom.RowCount = 2;
            grp1Bottom.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            grp1Bottom.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            _lblOpInfo = new Label
            {
                Text = "未选择操作（在主视口拾取操作，可多次拾取以批量录制）",
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.MessageBoxFont,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(2, 2, 0, 2)
            };
            grp1Bottom.Controls.Add(_lblOpInfo, 0, 0);
            grp1Bottom.SetColumnSpan(_lblOpInfo, 2);

            _chkAutoApplyView = new CheckBox
            {
                Text = "拾取时自动定位视角",
                Font = SystemFonts.MessageBoxFont,
                Checked = true,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(2, 4, 0, 0)
            };
            grp1Bottom.Controls.Add(_chkAutoApplyView, 0, 1);

            _btnSaveView = new FormUiKit.FlatColorButton
            {
                Text = "[记录当前视角为起始]",
                Height = 26,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = new Size(150, 26),
                FlatStyle = FlatStyle.Flat,
                Font = SystemFonts.MessageBoxFont,
                Enabled = false,
                BgColor = Color.FromArgb(80, 120, 140),
                ForeColor = Color.White,
                BorderColor = Color.FromArgb(80, 120, 140),
                Anchor = AnchorStyles.Right,
                Margin = new Padding(6, 2, 0, 0)
            };
            grp1Bottom.Controls.Add(_btnSaveView, 1, 1);

            grp1.Controls.Add(_grid);          // Fill（先 Add）
            grp1.Controls.Add(grp1Bottom);     // Bottom（后 Add → 占住底边）
            AddRootRow(grp1, SizeType.Absolute, CARD1_H, dockFill: true);

            // ---------- 卡片 1.5：视角关键帧（默认折叠） ----------
            BuildKeyframeCard();
            AddRootRow(_grpKf, SizeType.AutoSize, 0, dockFill: false);

            // ---------- 卡片 2：录像参数 ----------
            var grp2 = NewGroupBox("录像参数");
            grp2.AutoSize = true;
            grp2.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            grp2.Padding = new Padding(10, 20, 10, 10);
            grp2.Controls.Add(BuildParamsTable());
            AddRootRow(grp2, SizeType.AutoSize, 0, dockFill: false);

            // ---------- 按钮行 ----------
            _btnStart = NewButton("开始录制", 110, primary: true);
            _btnCancel = NewButton("取消录制", 110, primary: false);
            _btnClose = NewButton("关闭", 90, primary: false);

            var pnlBtns = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.RightToLeft,
                Margin = new Padding(0, 0, 0, PAD)
            };
            pnlBtns.Controls.Add(_btnClose);
            pnlBtns.Controls.Add(_btnCancel);
            pnlBtns.Controls.Add(_btnStart);
            AddRootRow(pnlBtns, SizeType.AutoSize, 0, dockFill: false);

            // ---------- 日志折叠按钮 ----------
            _btnToggleLog = new FormUiKit.FlatColorButton
            {
                Width = 140,
                Height = 26,
                Text = "▼ 展开日志",
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = SystemFonts.MessageBoxFont,
                BgColor = SystemColors.ControlLight,
                ForeColor = SystemColors.ControlText,
                BorderColor = SystemColors.ControlDark,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 0, 0, 4)
            };
            AddRootRow(_btnToggleLog, SizeType.AutoSize, 0, dockFill: false, stretch: false);

            // ---------- 日志容器（默认收起，Height 切换 0/LOG_H）----------
            _pnlLogContainer = new Panel
            {
                Height = 0,
                BorderStyle = BorderStyle.FixedSingle,
                Visible = false,
                Margin = new Padding(0, 0, 0, PAD)
            };
            _rtbLog = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font(FontFamily.GenericMonospace, 8.5F),
                BackColor = Color.FromArgb(250, 250, 250),
                BorderStyle = BorderStyle.None,
                DetectUrls = false,
                HideSelection = false,
            };
            _pnlLogContainer.Controls.Add(_rtbLog);
            AddRootRow(_pnlLogContainer, SizeType.AutoSize, 0, dockFill: false);

            this.Controls.Add(_root);

            // ---------- StatusStrip ----------
            _statusStrip = new StatusStrip { Font = SystemFonts.MessageBoxFont };
            _lblStatus = new ToolStripStatusLabel("状态：待机")
            {
                Spring = true,
                TextAlign = ContentAlignment.MiddleLeft,
            };
            _lblTimer = new ToolStripStatusLabel("0.0s")
            {
                AutoSize = false,
                Width = 90,
                TextAlign = ContentAlignment.MiddleRight,
            };
            _statusStrip.Items.Add(_lblStatus);
            _statusStrip.Items.Add(_lblTimer);
            this.Controls.Add(_statusStrip);
        }

        /// <summary>
        /// 往根表追加一行。
        /// dockFill=true 用于固定高行（控件 Fill 撑满格子）；
        /// 其余行 AutoSize，控件 Dock=Top（宽度占满、高度自带）。
        /// stretch=false 时控件保持自身宽度靠左（如日志折叠按钮）。
        /// </summary>
        private void AddRootRow(Control ctrl, SizeType sizeType, float height,
                                bool dockFill, bool stretch = true)
        {
            int r = _root.RowCount;
            _root.RowCount = r + 1;
            _root.RowStyles.Add(sizeType == SizeType.Absolute
                ? new RowStyle(SizeType.Absolute, height)
                : new RowStyle(SizeType.AutoSize));

            if (dockFill) ctrl.Dock = DockStyle.Fill;
            else if (stretch) ctrl.Dock = DockStyle.Top;
            // stretch=false：不 Dock，按控件自身 Width + Anchor 排

            if (ctrl.Margin == Padding.Empty)
                ctrl.Margin = new Padding(0, 0, 0, PAD);
            _root.Controls.Add(ctrl, 0, r);
        }

        // ============================================================
        // 录像参数表（标签列 AutoSize + 内容列 100%）
        // ============================================================
        private Control BuildParamsTable()
        {
            var t = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Padding = new Padding(2, 2, 2, 0)
            };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            // ---- Row 1: 输出路径 + 浏览 ----
            _lblPath = ParamLabel("输出路径：");
            _txtPath = new TextBox
            {
                Font = SystemFonts.MessageBoxFont,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Margin = new Padding(0, 3, 4, 3)
            };
            _btnBrowse = new FormUiKit.FlatColorButton
            {
                Text = "浏览...",
                Width = 80,
                Height = 26,
                FlatStyle = FlatStyle.Flat,
                Font = SystemFonts.MessageBoxFont,
                BgColor = Color.FromArgb(120, 124, 135),
                ForeColor = Color.White,
                BorderColor = Color.FromArgb(120, 124, 135),
                Anchor = AnchorStyles.Right,
                Margin = new Padding(0, 1, 0, 1)
            };
            var pathRow = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Margin = Padding.Empty
            };
            pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            pathRow.RowCount = 1;
            pathRow.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            pathRow.Controls.Add(_txtPath, 0, 0);
            pathRow.Controls.Add(_btnBrowse, 1, 0);
            AddParamRow(t, _lblPath, pathRow);

            // ---- Row 2: 视频编码（整行）----
            _cmbCodec = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = SystemFonts.MessageBoxFont,
                MaxDropDownItems = 12,
                DropDownHeight = 240,
                IntegralHeight = false,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Margin = new Padding(0, 3, 0, 3)
            };
            AddParamRow(t, ParamLabel("视频编码："), _cmbCodec);

            // ---- Row 3: 时间源（整行）----
            _cmbTimeSource = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = SystemFonts.MessageBoxFont,
                MaxDropDownItems = 12,
                DropDownHeight = 240,
                IntegralHeight = false,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Margin = new Padding(0, 3, 0, 3)
            };
            AddParamRow(t, ParamLabel("时间源："), _cmbTimeSource);

            // ---- Row 4: 帧率 + 压缩率（同行）----
            _numFps = new NumericUpDown
            {
                Width = 78,
                Minimum = 1,
                Maximum = 120,
                Value = 30,
                Increment = 1,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 2, 0, 2)
            };
            _numCompression = new NumericUpDown
            {
                Width = 78,
                Minimum = 0,
                Maximum = 100,
                Value = 70,
                Increment = 5,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 2, 0, 2)
            };
            var fpsFlow = MkInlineFlow();
            fpsFlow.Controls.Add(_numFps);
            fpsFlow.Controls.Add(InlineLabel("fps", gray: true));
            fpsFlow.Controls.Add(InlineLabel("        压缩率：", gray: false));
            fpsFlow.Controls.Add(_numCompression);
            fpsFlow.Controls.Add(InlineLabel("%", gray: true));
            AddParamRow(t, ParamLabel("帧率："), fpsFlow);

            // ---- Row 5: 自定义分辨率 + 宽×高 ----
            _chkCustomRes = new CheckBox
            {
                Text = "自定义分辨率",
                Font = SystemFonts.MessageBoxFont,
                Checked = false,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(2, 4, 8, 2)
            };
            _numResW = new NumericUpDown
            {
                Width = 80,
                Minimum = 320,
                Maximum = 7680,
                Value = 1920,
                Increment = 4,
                Font = SystemFonts.MessageBoxFont,
                Enabled = false,
                Margin = new Padding(0, 2, 0, 2)
            };
            _numResH = new NumericUpDown
            {
                Width = 80,
                Minimum = 240,
                Maximum = 4320,
                Value = 1080,
                Increment = 4,
                Font = SystemFonts.MessageBoxFont,
                Enabled = false,
                Margin = new Padding(0, 2, 0, 2)
            };
            var resFlow = MkInlineFlow();
            resFlow.Controls.Add(InlineLabel("宽：", gray: false));
            resFlow.Controls.Add(_numResW);
            resFlow.Controls.Add(InlineLabel(" × ", gray: true));
            resFlow.Controls.Add(InlineLabel("高：", gray: false));
            resFlow.Controls.Add(_numResH);
            AddParamRow(t, _chkCustomRes, resFlow);

            // ---- Row 6: 录后加速（ffmpeg）----
            _cmbSpeedup = new ComboBox
            {
                Width = 110,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = SystemFonts.MessageBoxFont,
                MaxDropDownItems = 10,
                DropDownHeight = 240,
                IntegralHeight = false,
                Margin = new Padding(0, 2, 0, 2)
            };
            _cmbSpeedup.Items.AddRange(new object[]
            {
                new SpeedupItem(1.0,  "1.0x（原速，不加速）"),
                new SpeedupItem(1.5,  "1.5x"),
                new SpeedupItem(2.0,  "2.0x"),
                new SpeedupItem(3.0,  "3.0x"),
                new SpeedupItem(5.0,  "5.0x"),
                new SpeedupItem(8.0,  "8.0x"),
                new SpeedupItem(10.0, "10x"),
                new SpeedupItem(20.0, "20x（极速预览）"),
            });
            _cmbSpeedup.SelectedIndex = 0;

            _lblSpeedupHint = new Label
            {
                AutoSize = true,
                Text = "需 ffmpeg.exe（放插件目录或 PATH）",
                Font = SystemFonts.MessageBoxFont,
                ForeColor = SystemColors.GrayText,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(6, 5, 0, 0)
            };
            var spdFlow = MkInlineFlow();
            spdFlow.Controls.Add(_cmbSpeedup);
            spdFlow.Controls.Add(_lblSpeedupHint);
            AddParamRow(t, ParamLabel("录后加速："), spdFlow);

            // ---- Row 7: 自动聚焦（跨两列）----
            _chkFocus = new CheckBox
            {
                Text = "录制前自动聚焦视角到操作（推荐）",
                Font = SystemFonts.MessageBoxFont,
                Checked = true,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(2, 6, 0, 2)
            };
            int rFocus = t.RowCount;
            t.RowCount = rFocus + 1;
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.Controls.Add(_chkFocus, 0, rFocus);
            t.SetColumnSpan(_chkFocus, 2);

            return t;
        }

        private static void AddParamRow(TableLayoutPanel t, Control labelCtrl, Control content)
        {
            int r = t.RowCount;
            t.RowCount = r + 1;
            t.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            t.Controls.Add(labelCtrl, 0, r);
            t.Controls.Add(content, 1, r);
        }

        private static Label ParamLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Font = SystemFonts.MessageBoxFont,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(2, 6, 8, 2)
            };
        }

        private static Label InlineLabel(string text, bool gray)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Font = SystemFonts.MessageBoxFont,
                ForeColor = gray ? SystemColors.GrayText : SystemColors.ControlText,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(2, 5, 2, 0)
            };
        }

        private static FlowLayoutPanel MkInlineFlow()
        {
            return new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Anchor = AnchorStyles.Left,
                Margin = Padding.Empty
            };
        }

        // ============================================================
        // 控件工厂
        // ============================================================
        private static GroupBox NewGroupBox(string text)
        {
            return new FormUiKit.ColoredGroupBox
            {
                Text = text,
                Font = SystemFonts.MessageBoxFont,
                HeaderColor = Color.FromArgb(60, 90, 150),
                ForeColor = Color.FromArgb(60, 90, 150)
            };
        }

        private static Button NewButton(string text, int width, bool primary)
        {
            return new FormUiKit.FlatColorButton
            {
                Text = text,
                Width = width,
                Height = 26,
                FlatStyle = FlatStyle.Flat,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(4, 0, 4, 0),
                BgColor = primary ? Color.FromArgb(0, 122, 204) : SystemColors.ControlLight,
                ForeColor = primary ? Color.White : SystemColors.ControlText,
                BorderColor = primary ? Color.FromArgb(0, 122, 204) : SystemColors.ControlDark
            };
        }

        // ============================================================
        // 事件绑定
        // ============================================================
        private void BindEvents()
        {
            // TxObjGridCtrl 拾取
            try { _grid.ObjectInserted += (s, e) => OnOperationPicked(); } catch { }
            try { _grid.RowDeleted += (s, e) => OnOperationCleared(); } catch { }

            _btnBrowse.Click += OnBrowse;
            _btnStart.Click += OnStart;
            _btnCancel.Click += OnCancel;
            _btnClose.Click += (s, e) => this.Close();
            _btnToggleLog.Click += OnToggleLog;
            _btnSaveView.Click += OnSaveCurrentView;

            _chkCustomRes.CheckedChanged += (s, e) =>
            {
                _numResW.Enabled = _chkCustomRes.Checked;
                _numResH.Enabled = _chkCustomRes.Checked;
            };

            _numResW.ValueChanged += (s, e) =>
            {
                var v = RoundDownTo4((int)_numResW.Value);
                if ((int)_numResW.Value != v) _numResW.Value = v;
            };
            _numResH.ValueChanged += (s, e) =>
            {
                var v = RoundDownTo4((int)_numResH.Value);
                if ((int)_numResH.Value != v) _numResH.Value = v;
            };

            _txtPath.TextChanged += (s, e) => RefreshButtonState();

            this.FormClosing += OnFormClosing;
        }

        // ============================================================
        // 数据加载
        // ============================================================
        private void LoadCodecs()
        {
            _cmbCodec.Items.Clear();
            foreach (var c in PsReader.GetSupportedCodecs())
                _cmbCodec.Items.Add(new CodecItem(c));

            // 默认选 H.264（最兼容）
            for (int i = 0; i < _cmbCodec.Items.Count; i++)
            {
                if (((CodecItem)_cmbCodec.Items[i]).Value == TxVideoCodec.MPEG4_AVC_H264)
                { _cmbCodec.SelectedIndex = i; return; }
            }
            if (_cmbCodec.Items.Count > 0) _cmbCodec.SelectedIndex = 0;
        }

        private void LoadTimeSources()
        {
            _cmbTimeSource.Items.Clear();
            _cmbTimeSource.Items.Add(new TimeSourceItem("仿真时间（推荐）", TxMovieTimeSource.SimulationTime));
            _cmbTimeSource.Items.Add(new TimeSourceItem("实时时间", TxMovieTimeSource.RealTime));
            _cmbTimeSource.SelectedIndex = 0;
        }

        // ============================================================
        // 操作选择回调
        // ============================================================
        private void OnOperationPicked()
        {
            // 【焦点修复】拾取后焦点要回到 grid，否则下次拾取/键盘操作无响应
            try { _grid.Focus(); } catch { }

            int count = GetGridCount();
            if (count == 0) { OnOperationCleared(); return; }

            // 找出最新拾取的 op —— 在 grid 末尾
            ITxObject latest = null;
            try { latest = _grid.GetObject(count - 1) as ITxObject; } catch { }
            if (latest != null && !ReferenceEquals(latest, _lastPickedOp))
            {
                _lastPickedOp = latest;
                // 为新 op 自动计算相机，并应用到视口
                HandleNewOpCamera(latest);
            }

            if (count == 1)
            {
                var op = _grid.GetObject(0) as ITxObject;
                var name = PsReader.GetObjectName(op);
                var dur = PsReader.GetOperationDuration(op);
                _lblOpInfo.ForeColor = SystemColors.ControlText;
                _lblOpInfo.Text = dur.HasValue
                    ? string.Format("已选 1 个操作：{0}    (估计时长 {1:F2}s)", name, dur.Value)
                    : "已选 1 个操作：" + name;

                _lblPath.Text = "输出文件：";
                if (string.IsNullOrWhiteSpace(_txtPath.Text) || _isPathInDirMode)
                {
                    var safe = MakeSafeFileName(name);
                    _txtPath.Text = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.MyVideos),
                        string.Format("PsRecord_{0}_{1:yyyyMMdd_HHmmss}.mp4",
                                      safe, DateTime.Now));
                }
                _isPathInDirMode = false;
            }
            else
            {
                _lblOpInfo.ForeColor = SystemColors.ControlText;
                _lblOpInfo.Text = string.Format(
                    "已选 {0} 个操作（将依次录制 {0} 个视频文件）", count);

                _lblPath.Text = "输出目录：";
                if (string.IsNullOrWhiteSpace(_txtPath.Text) || !_isPathInDirMode)
                {
                    _txtPath.Text = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.MyVideos),
                        string.Format("PsRecord_{0:yyyyMMdd_HHmmss}", DateTime.Now));
                }
                _isPathInDirMode = true;
            }

            _btnSaveView.Enabled = (_lastPickedOp != null);
            RefreshButtonState();
            RefreshKfOpSelector();
        }

        // ============================================================
        // 关键帧面板（折叠卡片）
        // 内部用 Dock 布局：顶部 [op 选择行] / 底部 [按钮行][状态行] / 中间 DGV 填满
        // 外层放在根表 AutoSize 行里，Height 切换后下方内容自动回流。
        // ============================================================
        private const int KF_COLLAPSED_H = 28;
        private const int KF_EXPANDED_H = 288;

        private void BuildKeyframeCard()
        {
            _grpKf = new Panel
            {
                Height = KF_COLLAPSED_H,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(0),
            };

            // 折叠/展开 toggle（用 + / - 避免某些字体不支持 ▶ ▼）
            _btnKfToggle = new FormUiKit.FlatColorButton
            {
                Dock = DockStyle.Top,
                Height = KF_COLLAPSED_H - 2,   // 占满折叠高度（减 Panel 边框）
                Text = "[+] 视角关键帧（点击展开） — 高级：操作内多视角切换",
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = SystemFonts.MessageBoxFont,
                BgColor = SystemColors.ControlLight,
                ForeColor = SystemColors.ControlText,
                BorderColor = SystemColors.ControlDark
            };
            _btnKfToggle.Click += OnToggleKeyframePanel;

            // === 折叠展开时显示的内容（用容器分区，Dock 自动布局） ===
            // 顺序：先 Add 的处于 Z 序底层；Dock=Bottom 优先底部
            //   1. 状态行 Dock=Bottom（最底）
            //   2. 按钮行 Dock=Bottom（次底）
            //   3. op 选择行 Dock=Top
            //   4. DGV Dock=Fill（占据剩余空间）

            _lblKfStatus = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 20,
                Text = "",
                Font = SystemFonts.MessageBoxFont,
                ForeColor = SystemColors.GrayText,
                TextAlign = ContentAlignment.MiddleLeft,
                Visible = false,
            };

            var pnlKfBtns = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 36,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = false,
                Padding = new Padding(0, 4, 0, 4),
                Margin = new Padding(0),
                Visible = false,
            };

            _btnKfSet = new FormUiKit.FlatColorButton
            {
                Text = "[设定为当前视角]",
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = new Size(140, 26),
                Height = 26,
                FlatStyle = FlatStyle.Flat,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 0, 6, 0),
                Enabled = false,
                BgColor = Color.FromArgb(0, 100, 167),
                ForeColor = Color.White,
                BorderColor = Color.FromArgb(0, 100, 167)
            };
            _btnKfSet.Click += OnKfSet;

            _btnKfPreview = new FormUiKit.FlatColorButton
            {
                Text = "[预览]",
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = new Size(70, 26),
                Height = 26,
                FlatStyle = FlatStyle.Flat,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 0, 6, 0),
                Enabled = false,
                BgColor = Color.FromArgb(80, 120, 140),
                ForeColor = Color.White,
                BorderColor = Color.FromArgb(80, 120, 140)
            };
            _btnKfPreview.Click += OnKfPreview;

            _btnKfClear = new FormUiKit.FlatColorButton
            {
                Text = "[清除]",
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = new Size(70, 26),
                Height = 26,
                FlatStyle = FlatStyle.Flat,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0),
                Enabled = false,
                BgColor = Color.FromArgb(120, 124, 135),
                ForeColor = Color.White,
                BorderColor = Color.FromArgb(120, 124, 135)
            };
            _btnKfClear.Click += OnKfClear;

            pnlKfBtns.Controls.Add(_btnKfSet);
            pnlKfBtns.Controls.Add(_btnKfPreview);
            pnlKfBtns.Controls.Add(_btnKfClear);

            // op 选择行（顶部）
            var pnlOpRow = new Panel
            {
                Dock = DockStyle.Top,
                Height = 28,
                Padding = new Padding(0),
                Margin = new Padding(0),
                Visible = false,
            };
            var lblOp = new Label
            {
                Left = 0,
                Top = 5,
                AutoSize = true,
                Text = "编辑操作:",
                Font = SystemFonts.MessageBoxFont,
                TextAlign = ContentAlignment.MiddleLeft,
            };
            _cmbKfOpSelector = new ComboBox
            {
                Top = 2,
                Height = 22,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = SystemFonts.MessageBoxFont,
                MaxDropDownItems = 12,
                DropDownHeight = 240,
                IntegralHeight = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            };
            _cmbKfOpSelector.SelectedIndexChanged += OnKfOpSelectorChanged;
            pnlOpRow.Controls.Add(lblOp);
            pnlOpRow.Controls.Add(_cmbKfOpSelector);
            // 加入容器后立即按 lblOp 实际宽度定位 ComboBox
            pnlOpRow.HandleCreated += (s, e) =>
            {
                _cmbKfOpSelector.Left = lblOp.Right + 8;
                _cmbKfOpSelector.Width = pnlOpRow.ClientSize.Width - _cmbKfOpSelector.Left;
            };

            // location 表（Fill，占据剩余空间）
            _gridKf = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AllowUserToResizeColumns = true,
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                RowHeadersVisible = false,
                ColumnHeadersVisible = true,
                ReadOnly = true,
                BackgroundColor = SystemColors.Window,
                BorderStyle = BorderStyle.FixedSingle,
                Font = SystemFonts.MessageBoxFont,
                Visible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 32,
                EnableHeadersVisualStyles = false,
                RowTemplate = { Height = 24 },
                ScrollBars = ScrollBars.Vertical,
            };
            _gridKf.ColumnHeadersDefaultCellStyle.Font =
                new Font(SystemFonts.MessageBoxFont, FontStyle.Bold);
            _gridKf.ColumnHeadersDefaultCellStyle.Padding = new Padding(2, 2, 2, 2);
            _gridKf.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(235, 240, 245);
            _gridKf.ColumnHeadersDefaultCellStyle.ForeColor = SystemColors.ControlText;
            _gridKf.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

            _gridKf.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "#",
                Width = 56,
                Name = "colIdx",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            });
            _gridKf.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "触发时机",
                Width = 140,
                Name = "colTrig",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            });
            _gridKf.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Location",
                Name = "colName",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            });
            _gridKf.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "视角",
                Width = 100,
                Name = "colCam",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            });

            _gridKf.SelectionChanged += (s, e) => RefreshKfButtonState();

            // 按 Dock 优先级 Add：WinForms 是"后 Add 的更贴边"，所以
            //   最贴底边的（状态行）最后 Add 在 Bottom 组里
            //   最贴顶边的（toggle 按钮）最后 Add 在 Top 组里
            _grpKf.Controls.Add(_gridKf);       // Fill
            _grpKf.Controls.Add(pnlKfBtns);     // Bottom（先 Add）
            _grpKf.Controls.Add(_lblKfStatus);  // Bottom（后 Add → 最底）
            _grpKf.Controls.Add(pnlOpRow);      // Top（先 Add）
            _grpKf.Controls.Add(_btnKfToggle);  // Top（后 Add → 最顶）

            // 保存引用用于折叠控制
            _pnlOpRow = pnlOpRow;
            _pnlKfBtns = pnlKfBtns;
        }

        // 保存对 op 选择行容器 / 按钮行容器的引用
        private Panel _pnlOpRow;
        private FlowLayoutPanel _pnlKfBtns;

        private void OnToggleKeyframePanel(object sender, EventArgs e)
        {
            SetKeyframePanelExpanded(!_kfExpanded);
        }

        private void SetKeyframePanelExpanded(bool expanded)
        {
            _kfExpanded = expanded;
            _grpKf.Height = expanded ? KF_EXPANDED_H : KF_COLLAPSED_H;
            if (_pnlOpRow != null) _pnlOpRow.Visible = expanded;
            if (_gridKf != null) _gridKf.Visible = expanded;
            if (_pnlKfBtns != null) _pnlKfBtns.Visible = expanded;
            if (_lblKfStatus != null) _lblKfStatus.Visible = expanded;

            _btnKfToggle.Text = expanded
                ? "[-] 视角关键帧（点击收起）"
                : "[+] 视角关键帧（点击展开） — 高级：操作内多视角切换";

            // 根表 AutoSize 行自动回流，只需重设窗体总高
            UpdateFormHeight();
            RefreshKfOpSelector();
            RefreshKfPanelContent();
        }

        /// <summary>
        /// 按根表实测内容高度重设窗体 ClientSize（替代原 RelayoutBelowKeyframePanel
        /// 的手动平移 + 硬编码高度常量；DPI 缩放后同样准确）。
        /// </summary>
        private void UpdateFormHeight()
        {
            int h;
            try
            {
                _root.PerformLayout();
                int statusH = (_statusStrip != null && _statusStrip.Height > 0)
                    ? _statusStrip.Height : 24;
                h = _root.Height + statusH;
                if (h < 300) throw new InvalidOperationException("layout not ready");
            }
            catch
            {
                // 兜底：常量法（与原版一致）
                int kfExtra = _kfExpanded ? (KF_EXPANDED_H - KF_COLLAPSED_H) : 0;
                int logExtra = _logExpanded ? (LOG_H + 2) : 0;
                h = FORM_H_FALLBACK + kfExtra + logExtra;
            }
            this.ClientSize = new Size(this.ClientSize.Width, h);
            if (_statusStrip != null) _statusStrip.BringToFront();
        }

        /// <summary>
        /// 刷新 op 下拉框 —— 列出当前所有已拾取的 op
        /// </summary>
        private void RefreshKfOpSelector()
        {
            if (_cmbKfOpSelector == null) return;
            var prevOp = _kfEditingOp;
            _cmbKfOpSelector.BeginUpdate();
            _cmbKfOpSelector.Items.Clear();

            int count = GetGridCount();
            var ops = new List<ITxObject>(count);
            for (int i = 0; i < count; i++)
            {
                ITxObject o = null;
                try { o = _grid.GetObject(i) as ITxObject; } catch { }
                if (o == null) continue;
                ops.Add(o);
                _cmbKfOpSelector.Items.Add(new KfOpItem(o));
            }

            // 优先保持之前选中的 op
            int sel = -1;
            if (prevOp != null)
            {
                for (int i = 0; i < ops.Count; i++)
                    if (ReferenceEquals(ops[i], prevOp)) { sel = i; break; }
            }
            // 否则跟随最近拾取的
            if (sel < 0 && _lastPickedOp != null)
            {
                for (int i = 0; i < ops.Count; i++)
                    if (ReferenceEquals(ops[i], _lastPickedOp)) { sel = i; break; }
            }
            if (sel < 0 && ops.Count > 0) sel = ops.Count - 1;

            _cmbKfOpSelector.SelectedIndex = sel;  // 会触发 OnKfOpSelectorChanged
            _kfEditingOp = (sel >= 0) ? ops[sel] : null;
            _cmbKfOpSelector.EndUpdate();
        }

        private void OnKfOpSelectorChanged(object sender, EventArgs e)
        {
            var item = _cmbKfOpSelector.SelectedItem as KfOpItem;
            _kfEditingOp = (item != null) ? item.Op : null;
            RefreshKfPanelContent();
        }

        /// <summary>下拉框条目 —— 提供友好显示名 + 持有 op 引用</summary>
        private class KfOpItem
        {
            public ITxObject Op;
            public KfOpItem(ITxObject op) { Op = op; }
            public override string ToString()
            {
                return Op == null ? "" : PsReader.GetObjectName(Op);
            }
        }

        /// <summary>加速倍数下拉条目</summary>
        private class SpeedupItem
        {
            public double Factor;
            public string Label;
            public SpeedupItem(double factor, string label) { Factor = factor; Label = label; }
            public override string ToString() { return Label; }
        }

        /// <summary>查找 ffmpeg.exe 是否就位（插件目录 / PATH）。逻辑跟 Service 同步</summary>
        private static bool HasFfmpegAvailable()
        {
            try
            {
                string asmDir = Path.GetDirectoryName(
                    System.Reflection.Assembly.GetExecutingAssembly().Location);
                if (!string.IsNullOrEmpty(asmDir)
                    && File.Exists(Path.Combine(asmDir, "ffmpeg.exe")))
                    return true;
            }
            catch { }
            try
            {
                string path = Environment.GetEnvironmentVariable("PATH") ?? "";
                foreach (var dir in path.Split(Path.PathSeparator))
                {
                    if (string.IsNullOrWhiteSpace(dir)) continue;
                    string candidate;
                    try { candidate = Path.Combine(dir.Trim(), "ffmpeg.exe"); }
                    catch { continue; }
                    if (File.Exists(candidate)) return true;
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 刷新关键帧 grid 内容（基于 _kfEditingOp）
        /// </summary>
        private void RefreshKfPanelContent()
        {
            if (!_kfExpanded || _gridKf == null) return;
            _gridKf.Rows.Clear();
            _kfLocations = null;

            if (_kfEditingOp == null)
            {
                _lblKfStatus.Text = "（拾取一个操作后开始编辑）";
                RefreshKfButtonState();
                return;
            }

            var sched = GetOrCreateSchedule(_kfEditingOp);

            // 行 0：起始视角
            int rowStart = _gridKf.Rows.Add("起",
                "op 起始时",
                "—",
                sched.InitialCamera != null ? "已设" : "—");
            _gridKf.Rows[rowStart].DefaultCellStyle.BackColor = Color.FromArgb(245, 247, 252);

            // 接下来一行行 = location 列表
            try
            {
                _kfLocations = PsReader.GetOperationLocations(_kfEditingOp);
            }
            catch { _kfLocations = new List<object>(); }
            if (_kfLocations == null) _kfLocations = new List<object>();

            for (int i = 0; i < _kfLocations.Count; i++)
            {
                int locIdx = i + 1;
                string locName = "?";
                try { locName = PsReader.GetLocationName(_kfLocations[i]); } catch { }

                bool hasKf = false;
                if (sched.Keyframes != null)
                {
                    for (int k = 0; k < sched.Keyframes.Count; k++)
                    {
                        if (sched.Keyframes[k] != null
                            && sched.Keyframes[k].LocationIndex == locIdx
                            && sched.Keyframes[k].Camera != null)
                        { hasKf = true; break; }
                    }
                }
                _gridKf.Rows.Add(locIdx.ToString(),
                                 "到达第 " + locIdx + " 点",
                                 locName,
                                 hasKf ? "已设" : "—");
            }

            int setCount = (sched.InitialCamera != null ? 1 : 0)
                         + (sched.Keyframes != null
                            ? sched.Keyframes.Count(k => k != null && k.Camera != null) : 0);
            int totalCount = 1 + _kfLocations.Count;

            if (_kfLocations.Count == 0)
            {
                _lblKfStatus.Text = string.Format(
                    "已设定 {0}/{1} 个视角 —— 该操作没有读到 location（仅有起始视角可编辑）",
                    setCount, totalCount);
                _lblKfStatus.ForeColor = Color.FromArgb(180, 100, 0);
            }
            else
            {
                _lblKfStatus.Text = string.Format(
                    "已设定 {0}/{1} 个视角（含起始视角），共 {2} 个 location",
                    setCount, totalCount, _kfLocations.Count);
                _lblKfStatus.ForeColor = SystemColors.GrayText;
            }

            if (_gridKf.Rows.Count > 0)
            {
                _gridKf.ClearSelection();
                _gridKf.Rows[0].Selected = true;
            }

            RefreshKfButtonState();
        }

        private void RefreshKfButtonState()
        {
            if (_btnKfSet == null) return;
            bool hasOp = _kfEditingOp != null;
            bool hasSel = _gridKf != null && _gridKf.SelectedRows.Count > 0;
            bool isRecording = _svc != null
                && (_svc.State == RecordingState.Recording
                 || _svc.State == RecordingState.Preparing);

            _btnKfSet.Enabled = hasOp && hasSel && !isRecording;
            _btnKfPreview.Enabled = hasOp && hasSel && HasCameraForSelectedRow() && !isRecording;
            _btnKfClear.Enabled = hasOp && hasSel && HasCameraForSelectedRow() && !isRecording;
        }

        private bool HasCameraForSelectedRow()
        {
            if (_kfEditingOp == null || _gridKf == null) return false;
            if (_gridKf.SelectedRows.Count == 0) return false;
            int rowIdx = _gridKf.SelectedRows[0].Index;
            var sched = GetOrCreateSchedule(_kfEditingOp);
            if (rowIdx == 0) return sched.InitialCamera != null;
            int locIdx = rowIdx;
            if (sched.Keyframes == null) return false;
            return sched.Keyframes.Any(k => k != null
                                         && k.LocationIndex == locIdx
                                         && k.Camera != null);
        }

        private void OnKfSet(object sender, EventArgs e)
        {
            if (_kfEditingOp == null) return;
            if (_gridKf.SelectedRows.Count == 0) return;

            var viewer = PsReader.GetGraphicViewer();
            if (viewer == null) { AppendLog("无法获取主视口", LogLevel.Error); return; }
            var cam = PsReader.GetCurrentCamera(viewer);
            if (cam == null) { AppendLog("无法读取当前视角", LogLevel.Warn); return; }

            int rowIdx = _gridKf.SelectedRows[0].Index;
            var sched = GetOrCreateSchedule(_kfEditingOp);

            if (rowIdx == 0)
            {
                sched.InitialCamera = cam;
                AppendLog("已设定 " + PsReader.GetObjectName(_kfEditingOp) + " 的起始视角",
                          LogLevel.Success);
            }
            else
            {
                int locIdx = rowIdx;
                if (sched.Keyframes == null) sched.Keyframes = new List<CameraKeyframe>();
                var existing = sched.Keyframes.FirstOrDefault(k => k != null && k.LocationIndex == locIdx);
                if (existing != null)
                {
                    existing.Camera = cam;
                }
                else
                {
                    sched.Keyframes.Add(new CameraKeyframe
                    {
                        LocationIndex = locIdx,
                        Camera = cam,
                    });
                    sched.Keyframes.Sort((a, b) => a.LocationIndex.CompareTo(b.LocationIndex));
                }
                AppendLog("已设定 " + PsReader.GetObjectName(_kfEditingOp)
                        + " 第 " + locIdx + " 点的视角", LogLevel.Success);
            }

            RefreshKfPanelContent();
            if (rowIdx < _gridKf.Rows.Count)
            {
                _gridKf.ClearSelection();
                _gridKf.Rows[rowIdx].Selected = true;
            }
        }

        private void OnKfPreview(object sender, EventArgs e)
        {
            if (_kfEditingOp == null || _gridKf.SelectedRows.Count == 0) return;
            int rowIdx = _gridKf.SelectedRows[0].Index;
            var sched = GetOrCreateSchedule(_kfEditingOp);

            TxCamera cam = null;
            if (rowIdx == 0)
            {
                cam = sched.InitialCamera;
            }
            else if (sched.Keyframes != null)
            {
                int locIdx = rowIdx;
                var kf = sched.Keyframes.FirstOrDefault(k => k != null && k.LocationIndex == locIdx);
                if (kf != null) cam = kf.Camera;
            }

            if (cam == null) { AppendLog("该位置尚未设定视角", LogLevel.Warn); return; }
            var viewer = PsReader.GetGraphicViewer();
            if (viewer != null && PsReader.SetCurrentCamera(viewer, cam))
                AppendLog("已预览第 " + rowIdx + " 行视角", LogLevel.Detail);
        }

        private void OnKfClear(object sender, EventArgs e)
        {
            if (_kfEditingOp == null || _gridKf.SelectedRows.Count == 0) return;
            int rowIdx = _gridKf.SelectedRows[0].Index;
            var sched = GetOrCreateSchedule(_kfEditingOp);

            if (rowIdx == 0)
            {
                sched.InitialCamera = null;
                AppendLog("已清除起始视角", LogLevel.Info);
            }
            else if (sched.Keyframes != null)
            {
                int locIdx = rowIdx;
                sched.Keyframes.RemoveAll(k => k != null && k.LocationIndex == locIdx);
                AppendLog("已清除第 " + locIdx + " 点视角切换", LogLevel.Info);
            }
            RefreshKfPanelContent();
            if (rowIdx < _gridKf.Rows.Count)
            {
                _gridKf.ClearSelection();
                _gridKf.Rows[rowIdx].Selected = true;
            }
        }

        // ============================================================
        // Schedule helper
        // ============================================================
        private OperationViewSchedule GetOrCreateSchedule(ITxObject op)
        {
            if (op == null) return null;
            OperationViewSchedule s;
            if (!_opSchedules.TryGetValue(op, out s))
            {
                s = new OperationViewSchedule();
                _opSchedules[op] = s;
            }
            return s;
        }

        /// <summary>
        /// 新拾取的 op 处理相机：自动计算 + 可选应用到视口 + 存储
        /// </summary>
        private void HandleNewOpCamera(ITxObject op)
        {
            try
            {
                var cam = PsReader.ComputeOptimalCamera(op);
                if (cam == null)
                {
                    AppendLog("无法自动计算 " + PsReader.GetObjectName(op)
                             + " 的视角（缺机器人或焊点位置），将使用 ZoomToSelection 兜底",
                             LogLevel.Warn);
                    return;
                }
                GetOrCreateSchedule(op).InitialCamera = cam;

                if (_chkAutoApplyView != null && _chkAutoApplyView.Checked)
                {
                    var viewer = PsReader.GetGraphicViewer();
                    if (viewer != null && PsReader.SetCurrentCamera(viewer, cam))
                    {
                        AppendLog("已自动定位 " + PsReader.GetObjectName(op) + " 的视角",
                                  LogLevel.Detail);
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLog("自动视角计算异常：" + ex.Message, LogLevel.Warn);
            }
        }

        /// <summary>
        /// "📷 记录当前视角" 按钮 —— 把当前视口相机覆盖存到最近拾取的 op 上
        /// </summary>
        private void OnSaveCurrentView(object sender, EventArgs e)
        {
            if (_lastPickedOp == null) { Beep("请先拾取一个操作"); return; }
            var viewer = PsReader.GetGraphicViewer();
            if (viewer == null) { AppendLog("无法获取主视口", LogLevel.Error); return; }
            var cam = PsReader.GetCurrentCamera(viewer);
            if (cam == null) { AppendLog("无法读取当前视角", LogLevel.Warn); return; }

            GetOrCreateSchedule(_lastPickedOp).InitialCamera = cam;
            AppendLog("✓ 已记录起始视角 → " + PsReader.GetObjectName(_lastPickedOp), LogLevel.Success);
        }

        private void OnOperationCleared()
        {
            int count = GetGridCount();
            if (count > 0) { OnOperationPicked(); return; }  // 还有别的项

            _lblOpInfo.Text = "未选择操作（在主视口拾取操作，可多次拾取以批量录制）";
            _lblOpInfo.ForeColor = SystemColors.GrayText;
            _lblPath.Text = "输出路径：";
            _opSchedules.Clear();
            _lastPickedOp = null;
            _kfEditingOp = null;
            _btnSaveView.Enabled = false;
            RefreshButtonState();
            RefreshKfOpSelector();
        }

        // ============================================================
        // 路径浏览 —— 自适应：单选 → 文件对话框；多选 → 目录对话框
        // ============================================================
        private void OnBrowse(object sender, EventArgs e)
        {
            if (_isPathInDirMode)
            {
                using (var dlg = new FolderBrowserDialog())
                {
                    dlg.Description = "选择视频输出目录（每个操作生成一个 .mp4 文件）";
                    dlg.ShowNewFolderButton = true;
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(_txtPath.Text)
                            && Directory.Exists(_txtPath.Text))
                            dlg.SelectedPath = _txtPath.Text;
                    }
                    catch { }
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        _txtPath.Text = dlg.SelectedPath;
                        RefreshButtonState();
                    }
                }
            }
            else
            {
                using (var dlg = new SaveFileDialog())
                {
                    dlg.Title = "选择输出视频文件";
                    dlg.Filter = "MP4 视频 (*.mp4)|*.mp4|"
                               + "WMV 视频 (*.wmv)|*.wmv|"
                               + "AVI 视频 (*.avi)|*.avi|"
                               + "MKV 视频 (*.mkv)|*.mkv";
                    dlg.DefaultExt = "mp4";

                    if (!string.IsNullOrWhiteSpace(_txtPath.Text))
                    {
                        try
                        {
                            dlg.InitialDirectory = Path.GetDirectoryName(_txtPath.Text);
                            dlg.FileName = Path.GetFileName(_txtPath.Text);
                        }
                        catch { }
                    }
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        _txtPath.Text = dlg.FileName;
                        RefreshButtonState();
                    }
                }
            }
        }

        // ============================================================
        // 开始录制 —— 自动适应单/多操作模式
        // ============================================================
        private void OnStart(object sender, EventArgs e)
        {
            var ops = GetSelectedOperations();
            if (ops.Count == 0) { Beep("请先选择至少一个操作"); return; }
            if (string.IsNullOrWhiteSpace(_txtPath.Text)) { Beep("请先指定输出路径"); return; }

            var path = _txtPath.Text.Trim();
            List<RecordingJob> jobs;

            try
            {
                if (_isPathInDirMode)
                {
                    // ----- 多操作：path 是目录 -----
                    if (!Directory.Exists(path))
                    {
                        try { Directory.CreateDirectory(path); }
                        catch (Exception ex)
                        {
                            AppendLog("创建目录失败：" + ex.Message, LogLevel.Error);
                            return;
                        }
                    }
                    jobs = BuildJobsForDirectory(ops, path);
                    if (!ConfirmBatchOverwrite(jobs)) return;
                }
                else
                {
                    // ----- 单操作：path 是完整文件路径 -----
                    var ext = (Path.GetExtension(path) ?? "").ToLowerInvariant();
                    if (ext != ".mp4" && ext != ".wmv" && ext != ".avi" && ext != ".mkv")
                    {
                        AppendLog("不支持的扩展名 " + ext + "，请使用 mp4 / wmv / avi / mkv",
                                  LogLevel.Error);
                        return;
                    }

                    var dir = Path.GetDirectoryName(path);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                        Directory.CreateDirectory(dir);

                    if (File.Exists(path))
                    {
                        var r = MessageBox.Show(this,
                            "文件已存在，是否覆盖？\n" + path,
                            "TxTools 操作录像",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);
                        if (r != DialogResult.Yes) return;
                        try { File.Delete(path); }
                        catch (Exception ex)
                        {
                            AppendLog("删除旧文件失败：" + ex.Message, LogLevel.Error);
                            return;
                        }
                    }

                    OperationViewSchedule schedSingle;
                    _opSchedules.TryGetValue(ops[0], out schedSingle);
                    jobs = new List<RecordingJob>
                    {
                        new RecordingJob
                        {
                            Operation     = ops[0],
                            FilePath      = path,
                            OperationName = PsReader.GetObjectName(ops[0]),
                            Schedule      = schedSingle,
                        }
                    };
                }
            }
            catch (Exception ex)
            {
                AppendLog("准备任务失败：" + ex.Message, LogLevel.Error);
                return;
            }

            // 组装 options
            var opts = new RecordingOptions
            {
                Codec = ((CodecItem)_cmbCodec.SelectedItem).Value,
                FrameRate = (int)_numFps.Value,
                Compression = (int)_numCompression.Value,
                TimeSource = ((TimeSourceItem)_cmbTimeSource.SelectedItem).Value,
                FocusOnOperation = _chkFocus.Checked,
            };
            if (_chkCustomRes.Checked)
            {
                opts.ResolutionWidth = (int)_numResW.Value;
                opts.ResolutionHeight = (int)_numResH.Value;
            }
            var spdItem = _cmbSpeedup.SelectedItem as SpeedupItem;
            if (spdItem != null) opts.SpeedupFactor = spdItem.Factor;

            // 加速倍数 > 1 时，提前提示一下 ffmpeg 是否就位
            if (opts.SpeedupFactor > 1.01)
            {
                if (!HasFfmpegAvailable())
                {
                    var r = MessageBox.Show(this,
                        "设定了录后加速 " + opts.SpeedupFactor.ToString("0.0") + "x，"
                      + "但未找到 ffmpeg.exe。\n\n"
                      + "查找位置：\n"
                      + "  1) 插件 DLL 所在目录\n"
                      + "  2) 系统 PATH\n\n"
                      + "不加速继续录制？（点否取消，回去先安置 ffmpeg）",
                        "TxTools 操作录像",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (r != DialogResult.Yes) return;
                    opts.SpeedupFactor = 1.0;  // 降级为原速
                }
            }

            _tickElapsed = 0;
            _uiTimer.Start();
            _svc.Start(jobs, opts);
        }

        /// <summary>多操作模式：在目录下为每个 op 生成唯一文件名。</summary>
        private List<RecordingJob> BuildJobsForDirectory(List<ITxObject> ops, string dir)
        {
            var jobs = new List<RecordingJob>(ops.Count);
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var op in ops)
            {
                var rawName = PsReader.GetObjectName(op);
                var safe = MakeSafeFileName(rawName);
                var path = Path.Combine(dir, safe + ".mp4");
                int n = 2;
                while (usedNames.Contains(path) || File.Exists(path))
                {
                    path = Path.Combine(dir, safe + "_" + n + ".mp4");
                    n++;
                }
                usedNames.Add(path);
                OperationViewSchedule sched;
                _opSchedules.TryGetValue(op, out sched);
                jobs.Add(new RecordingJob
                {
                    Operation = op,
                    FilePath = path,
                    OperationName = rawName,
                    Schedule = sched,
                });
            }
            return jobs;
        }

        /// <summary>批量模式开始前提示用户即将生成的文件清单。</summary>
        private bool ConfirmBatchOverwrite(List<RecordingJob> jobs)
        {
            var sb = new StringBuilder();
            sb.AppendLine("将依次录制 " + jobs.Count + " 个视频，文件如下：");
            sb.AppendLine();
            int previewCount = Math.Min(jobs.Count, 8);
            for (int i = 0; i < previewCount; i++)
            {
                sb.AppendLine(string.Format("  {0}. {1}",
                    i + 1, Path.GetFileName(jobs[i].FilePath)));
            }
            if (jobs.Count > previewCount)
                sb.AppendLine("  ... 等共 " + jobs.Count + " 个");
            sb.AppendLine();
            sb.AppendLine("是否开始批量录制？");

            var r = MessageBox.Show(this, sb.ToString(),
                "TxTools 操作录像 - 批量录制确认",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question);
            return r == DialogResult.OK;
        }

        private void OnCancel(object sender, EventArgs e)
        {
            _svc.Cancel();
            _uiTimer.Stop();
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            if (_svc.State == RecordingState.Recording ||
                _svc.State == RecordingState.Preparing)
            {
                var r = MessageBox.Show(this,
                    "录制正在进行中。关闭窗口将取消录制，确定？",
                    "TxTools 操作录像",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                if (r != DialogResult.Yes) { e.Cancel = true; return; }
                _svc.Cancel();
            }
            _uiTimer.Stop();
        }

        // ============================================================
        // 日志折叠
        // ============================================================
        private void OnToggleLog(object sender, EventArgs e)
        {
            _logExpanded = !_logExpanded;
            _pnlLogContainer.Visible = _logExpanded;
            _pnlLogContainer.Height = _logExpanded ? LOG_H : 0;
            _btnToggleLog.Text = _logExpanded ? "▲ 收起日志" : "▼ 展开日志";

            // 根表 AutoSize 行自动回流，只需重设窗体总高
            UpdateFormHeight();
        }

        // ============================================================
        // UI 计时
        // ============================================================
        private void OnUiTick(object sender, EventArgs e)
        {
            if (_svc.State == RecordingState.Recording)
            {
                _tickElapsed += _uiTimer.Interval / 1000.0;
                _lblTimer.Text = _tickElapsed.ToString("F1") + "s";
            }
        }

        // ============================================================
        // Service 事件（player 回调可能跨线程，用 PS 主线程 SyncContext 切回）
        // ============================================================
        private void OnSvcLog(string msg, LogLevel level)
        {
            if (this.IsDisposed) return;
            _psCtx.Post(_ =>
            {
                if (this.IsDisposed) return;
                AppendLog(msg, level);
                if ((level == LogLevel.Error || level == LogLevel.Warn) && !_logExpanded)
                    OnToggleLog(null, null);
            }, null);
        }

        private void OnSvcJobProgress(int current, int total, string opName)
        {
            if (this.IsDisposed) return;
            _psCtx.Post(_ =>
            {
                if (this.IsDisposed) return;
                _tickElapsed = 0;  // 每个任务的计时单独从 0 开始
                UpdateStatusLabel(
                    string.Format("● 录制中 [{0}/{1}] {2}", current, total, opName),
                    "0.0s");
                _lblStatus.ForeColor = Color.Firebrick;
            }, null);
        }

        private void OnSvcState(RecordingState s)
        {
            if (this.IsDisposed) return;
            _psCtx.Post(_ =>
            {
                if (this.IsDisposed) return;
                switch (s)
                {
                    case RecordingState.Idle:
                        UpdateStatusLabel("待机", "0.0s");
                        break;
                    case RecordingState.Preparing:
                        UpdateStatusLabel("准备中...", "0.0s");
                        break;
                    case RecordingState.Recording:
                        UpdateStatusLabel("● 录制中", "0.0s");
                        _lblStatus.ForeColor = Color.Firebrick;
                        break;
                    case RecordingState.Completed:
                        if (_svc.TotalJobs <= 1)
                        {
                            UpdateStatusLabel(string.Format("✓ 完成 ({0:F1}s)",
                                               _svc.LastDurationSeconds),
                                              _svc.LastDurationSeconds.ToString("F1") + "s");
                            _lblStatus.ForeColor = Color.SeaGreen;
                            _uiTimer.Stop();
                            OfferOpenFile(_svc.LastFilePath);
                        }
                        else
                        {
                            UpdateStatusLabel(string.Format("✓ 批量完成 {0}/{1}",
                                               _svc.SucceededCount, _svc.TotalJobs),
                                              "—");
                            _lblStatus.ForeColor = _svc.FailedCount > 0
                                                    ? Color.DarkOrange : Color.SeaGreen;
                            _uiTimer.Stop();
                            OfferOpenFolder(_svc.ProducedFiles);
                        }
                        break;
                    case RecordingState.Cancelled:
                        UpdateStatusLabel("已取消", "—");
                        _lblStatus.ForeColor = SystemColors.GrayText;
                        _uiTimer.Stop();
                        break;
                    case RecordingState.Failed:
                        UpdateStatusLabel("✗ 失败", "—");
                        _lblStatus.ForeColor = Color.Firebrick;
                        _uiTimer.Stop();
                        break;
                }
                RefreshButtonState();
            }, null);
        }

        // ============================================================
        // 工具方法
        // ============================================================
        private void RefreshButtonState()
        {
            bool isRunning = _svc.State == RecordingState.Recording
                          || _svc.State == RecordingState.Preparing;
            bool canStart = !isRunning
                          && GetGridCount() > 0
                          && !string.IsNullOrWhiteSpace(_txtPath.Text);

            _btnStart.Enabled = canStart;
            _btnCancel.Enabled = isRunning;

            // 参数编辑在录制中禁用
            _txtPath.Enabled = !isRunning;
            _btnBrowse.Enabled = !isRunning;
            _cmbCodec.Enabled = !isRunning;
            _numFps.Enabled = !isRunning;
            _numCompression.Enabled = !isRunning;
            _cmbTimeSource.Enabled = !isRunning;
            _chkCustomRes.Enabled = !isRunning;
            _numResW.Enabled = !isRunning && _chkCustomRes.Checked;
            _numResH.Enabled = !isRunning && _chkCustomRes.Checked;
            _chkFocus.Enabled = !isRunning;
            if (_cmbSpeedup != null) _cmbSpeedup.Enabled = !isRunning;
            try { _grid.Enabled = !isRunning; } catch { }
            if (_btnSaveView != null)
                _btnSaveView.Enabled = !isRunning && _lastPickedOp != null;
            if (_chkAutoApplyView != null)
                _chkAutoApplyView.Enabled = !isRunning;
            RefreshKfButtonState();
        }

        private int GetGridCount()
        {
            try { return _grid.Count; }
            catch { return 0; }
        }

        /// <summary>取出 grid 中所有 ITxObject。</summary>
        private List<ITxObject> GetSelectedOperations()
        {
            var list = new List<ITxObject>();
            try
            {
                int n = _grid.Count;
                for (int i = 0; i < n; i++)
                {
                    var t = _grid.GetObject(i) as ITxObject;
                    if (t != null) list.Add(t);
                }
            }
            catch { }
            return list;
        }

        /// <summary>取首项（兼容旧调用点用）。</summary>
        private ITxObject GetSelectedOperation()
        {
            try
            {
                if (_grid.Count == 0) return null;
                return _grid.GetObject(0) as ITxObject;
            }
            catch { return null; }
        }

        private static string MakeSafeFileName(string s)
        {
            if (string.IsNullOrEmpty(s)) return "operation";
            var invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(s.Length);
            foreach (var c in s)
                sb.Append(Array.IndexOf(invalid, c) >= 0 ? '_' : c);
            return sb.ToString();
        }

        private static int RoundDownTo4(int v) { return Math.Max(4, (v / 4) * 4); }

        private void AppendLog(string msg, LogLevel level)
        {
            if (_rtbLog == null || _rtbLog.IsDisposed) return;
            Color color; string prefix;
            switch (level)
            {
                case LogLevel.Error: color = Color.Firebrick; prefix = "[ERR ] "; break;
                case LogLevel.Warn: color = Color.DarkOrange; prefix = "[WARN] "; break;
                case LogLevel.Success: color = Color.SeaGreen; prefix = "[ OK ] "; break;
                case LogLevel.Info: color = Color.DarkSlateBlue; prefix = "[INFO] "; break;
                default: color = Color.Gray; prefix = "[ .. ] "; break;
            }
            string line = DateTime.Now.ToString("HH:mm:ss") + " " + prefix + msg + "\r\n";
            _rtbLog.SelectionStart = _rtbLog.TextLength;
            _rtbLog.SelectionLength = 0;
            _rtbLog.SelectionColor = color;
            _rtbLog.AppendText(line);
            _rtbLog.ScrollToCaret();
        }

        private void UpdateStatusLabel(string status, string timer)
        {
            _lblStatus.Text = "状态：" + status;
            _lblStatus.ForeColor = SystemColors.ControlText;
            _lblTimer.Text = timer;
        }

        private void Beep(string msg)
        {
            System.Media.SystemSounds.Beep.Play();
            AppendLog(msg, LogLevel.Warn);
            if (!_logExpanded) OnToggleLog(null, null);
        }

        private void OfferOpenFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;
            var r = MessageBox.Show(this,
                "录制完成：\n" + path + "\n\n是否打开视频文件？",
                "TxTools 操作录像",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information);
            if (r == DialogResult.Yes)
            {
                try { System.Diagnostics.Process.Start(path); }
                catch (Exception ex) { AppendLog("打开失败：" + ex.Message, LogLevel.Warn); }
            }
        }

        /// <summary>批量完成后弹窗：打开包含所有输出文件的目录。</summary>
        private void OfferOpenFolder(List<string> files)
        {
            if (files == null || files.Count == 0) return;
            string dir = null;
            try { dir = Path.GetDirectoryName(files[0]); } catch { }
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir)) return;

            var r = MessageBox.Show(this,
                string.Format("批量录制完成，共生成 {0} 个视频。\n\n输出目录：\n{1}\n\n是否打开目录？",
                              files.Count, dir),
                "TxTools 操作录像",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information);
            if (r == DialogResult.Yes)
            {
                try { System.Diagnostics.Process.Start("explorer.exe", "\"" + dir + "\""); }
                catch (Exception ex) { AppendLog("打开目录失败：" + ex.Message, LogLevel.Warn); }
            }
        }

        // ============================================================
        // ComboBox 包装项
        // ============================================================
        private class CodecItem
        {
            public TxVideoCodec Value { get; private set; }
            public CodecItem(TxVideoCodec v) { Value = v; }
            public override string ToString()
            {
                switch (Value)
                {
                    case TxVideoCodec.MPEG4: return "MPEG-4 (兼容性好)";
                    case TxVideoCodec.MPEG4_AVC_H264: return "H.264 / MP4 (推荐)";
                    case TxVideoCodec.MPEGH_HVEC_H265: return "H.265 / HEVC";
                    case TxVideoCodec.THEORA: return "Theora";
                    case TxVideoCodec.VP9: return "VP9";
                    default: return Value.ToString();
                }
            }
        }

        private class TimeSourceItem
        {
            public string Label { get; private set; }
            public TxMovieTimeSource Value { get; private set; }
            public TimeSourceItem(string label, TxMovieTimeSource v)
            { Label = label; Value = v; }
            public override string ToString() { return Label; }
        }

        // ============================================================
        // Dispose
        // ============================================================
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try { if (_uiTimer != null) { _uiTimer.Stop(); _uiTimer.Dispose(); } } catch { }
                try
                {
                    if (_svc != null)
                    {
                        _svc.Log -= OnSvcLog;
                        _svc.StateChanged -= OnSvcState;
                        _svc.JobProgress -= OnSvcJobProgress;
                    }
                }
                catch { }
            }
            base.Dispose(disposing);
        }
    }
}
