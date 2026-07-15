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
//
// 二次重构（修复 20 项问题）：
// - 修复 UpdateGunInfo 误用 Form.Name 的 Bug，改为后台线程跑工具名解析
// - 集中颜色到 Theme 静态类
// - 控件字段按 region 分组
// - 抽象 IOpListAdapter（GridAdapter / ListViewAdapter），消除 _useGrid 分支
// - BuildCol1 / BuildCol2 拆分为多个 BuildXxxCard
// - OnExportGun 拆分为 ResolveCgrPathsForOps + BuildGunExportParams
// - AutoFit* 改用 TextRenderer.MeasureText
// - Log 增加带 LogLevel 的重载
// - AskUserPickRefFrame 按钮改 FlowLayoutPanel
// - 首次填充副作用集中到 OnOpsFirstFilled
// - _ops 去重改 HashSet
// - EnsureOpsSelected 守卫
// - OnFormClosing 显式释放 PS 控件持有的对象引用
//
// 第三轮改动（CGR 检索体验优化）：
// - "未找到 CGR" 弹窗替换为 PromptCgrNotFound 对话框，新增"打开 JT 所在位置"
//   按钮（用资源管理器打开同级目录并选中 .jt/.cojt），同时允许直接手动浏览
//   选择一个 CGR/CATPart/CATProduct 继续导出，无需返回主界面勾选自定义焊枪。
// - PromptCgrSelection（候选选择对话框）增强：
//     · 增加"打开 JT 所在位置"按钮
//     · 列表项鼠标悬停 tooltip 显示完整路径
//     · 双击列表项 = 使用选中
//     · 底部状态条显示当前选中候选的完整路径，避免文件名截断歧义
//     · 对话框改为可调整大小，扩大默认尺寸
// - ResolveCgrPathsForOps 在高置信度分支中区分"精确命中"（相似度 ≥ 0.999）
//   与"高相似度"的日志输出，便于用户判断命中质量。
// - 新增 OpenJtLocation 辅助方法：用 explorer /select 选中目标文件，
//   未找到 .jt 时退化为打开目录本身；异常时给出 MessageBox 提示。

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;

namespace TxTools.ExportGun
{
    public partial class ExportGunForm : TxForm
    {
        // ════════════════════════════════════════════════════════════
        //  Theme：集中管理所有颜色常量
        // ════════════════════════════════════════════════════════════
        private static class Theme
        {
            // 区块/卡片标题色（TxColor 形式，传给 PS API 时用 .Color 取 System.Drawing.Color）
            public static readonly TxColor TxAccent = new TxColor(0, 70, 127);
            public static readonly TxColor TxCol1 = new TxColor(0, 100, 140);
            public static readonly TxColor TxGun = new TxColor(155, 120, 0);
            public static readonly TxColor TxBall = new TxColor(150, 70, 90);
            public static readonly TxColor TxLog = new TxColor(50, 120, 60);
            public static readonly TxColor TxPoints = new TxColor(95, 75, 140);

            // 功能按钮色
            public static readonly Color BtnPrimary = Color.FromArgb(0, 100, 167);
            public static readonly Color BtnSecondary = Color.FromArgb(80, 120, 140);
            public static readonly Color BtnMuted = Color.FromArgb(120, 124, 135);
            public static readonly Color BtnDanger = Color.FromArgb(130, 50, 50);
            public static readonly Color BtnExport = Color.FromArgb(80, 80, 130);
            public static readonly Color BtnGun = Color.Orange;

            // 日志面板
            public static readonly Color LogBg = Color.FromArgb(20, 22, 27);
            public static readonly Color LogText = Color.FromArgb(178, 200, 178);
            public static readonly Color LogOk = Color.FromArgb(90, 210, 110);
            public static readonly Color LogErr = Color.FromArgb(228, 88, 88);
            public static readonly Color LogWarn = Color.FromArgb(228, 180, 70);
            public static readonly Color LogPs = Color.FromArgb(110, 180, 228);
            public static readonly Color LogCoord = Color.FromArgb(160, 200, 255);
            public static readonly Color LogExcel = Color.FromArgb(180, 228, 160);

            // 状态色（参考坐标状态条）
            public static readonly Color StatusOkFg = Color.FromArgb(25, 110, 25);
            public static readonly Color StatusOkBg = Color.FromArgb(210, 252, 210);
            public static readonly Color StatusRefFg = Color.FromArgb(25, 60, 130);
            public static readonly Color StatusRefBg = Color.FromArgb(180, 220, 255);
        }

        // ════════════════════════════════════════════════════════════
        //  日志级别
        // ════════════════════════════════════════════════════════════
        private enum LogLevel { Info, Ok, Warn, Error, Ps, Coord, Excel, Debug }

        // ════════════════════════════════════════════════════════════
        //  状态字段
        // ════════════════════════════════════════════════════════════
        #region State
        private ExportService _svc;
        private List<OperationInfo> _ops = new List<OperationInfo>();
        private string _refFrameName = "世界坐标系";
        private double[] _refFrameMatrix;
        private IOpListAdapter _listAdapter;   // 列表适配器（Grid 或 ListView）

        // 操作点列表模型
        private List<PointRow> _pointRows = new List<PointRow>();
        private bool _pointsLoaded;            // 点是否已加载完成
        private int _pointsLoadToken;          // 加载令牌，丢弃过期的异步结果
        private bool _suppressPointCheck;      // 批量勾选/重建时抑制 ItemChecked
        private int _lastSelCount;             // 最近一次导出时的勾选点数（供完成弹窗显示）
        private bool _dpiApplied;              // OnLoad 手动 DPI 放大仅执行一次
        #endregion

        // ════════════════════════════════════════════════════════════
        //  控件字段（按区域分组）
        // ════════════════════════════════════════════════════════════
        #region UI Controls - 操作选择
        private TxObjGridCtrl _objGrid;        // PS 原生对象列表（主方案）
        private TxObjEditBoxCtrl _objEditOp;   // 回落：单对象选择
        private ListView _lvOps;               // 回落：列表展示
        private Label _lblListHint;
        private Button _btnPickFromSel;
        #endregion

        #region UI Controls - 操作点列表（新第二栏）
        private ListView _lvPoints;            // 操作内所有点的勾选列表
        private Label _lblPointSel;            // 点列表标题条计数（已选 X / 共 Y）
        #endregion

        #region UI Controls - 参考坐标
        private TxFrameComboBoxCtrl _frameCombo;
        private Label _lblRefCoordStatus;
        #endregion

        #region UI Controls - 导出设置
        private ComboBox _cmbPointType;
        private CheckBox _chkUseMfgName;
        private Label _lblPointCount;
        #endregion

        #region UI Controls - 导插枪
        private CheckBox _chkExportTCP;
        private CheckBox _chkGunOriginTCP;
        private CheckBox _chkCustomGun;
        private TextBox _txtGunModel;
        private Button _btnBrowseGun;
        private TextBox _txtGunProductName;
        private Button _btnExportGun;
        private Label _lblGunInfo;
        private ComboBox _cmbGunExportMode;   // 导出方式：共享几何 / 独立命名
        // TCP 选择
        private ComboBox _cmbTcp;                 // 多 TCP 下拉
        private Panel _tcpCustomPanel;            // 自定义坐标选择器容器
        private TxFrameComboBoxCtrl _tcpFrameCombo; // 自定义 TCP 坐标拾取（懒加载）
        private Label _lblTcpInfo;
        private string _tcpChoiceName;            // 选中的命名 TCP（默认项时为 null）
        private double[] _tcpCustomMatrix;        // 自定义 TCP 世界矩阵（非 null 优先）
        private int _tcpLoadToken;                // TCP 选项加载令牌
        #endregion

        private const string TcpCustomLabel = "自定义坐标…";

        #region UI Controls - 导点球
        private ComboBox _cmbBallTarget;
        private ComboBox _cmbBallOption;
        private NumericUpDown _nudDiameter;
        private TextBox _txtGeomSet;
        private TextBox _txtNamePrefix;
        private TextBox _txtBallPartName;
        private Button _btnExportBall;
        #endregion

        #region UI Controls - 日志/底栏
        private RichTextBox _rtbLog;
        private ProgressBar _progressBar;
        private Label _lblProgress;
        private Button _btnReset;
        private Button _btnClose;
        private Button _btnHelp;
        private Button _btnExportExcel;
        private TableLayoutPanel _root;        // 根布局（用于切换日志行高度）
        private Label _logToggle;              // 日志折叠标题条
        private bool _logExpanded;             // 日志当前是否展开
        #endregion

        // 日志折叠行高度
        private const int LOG_COLLAPSED_H = 26;
        private const int LOG_EXPANDED_H = 168;

        // ════════════════════════════════════════════════════════════
        //  构造
        // ════════════════════════════════════════════════════════════
        public ExportGunForm(SynchronizationContext psCtx)
        {

            SemiModal = false;

            InitializeComponent();
            BuildUI();
            // 关闭 Siemens flat style 皮肤以让自定义配色生效；
            // 低版本 PDPS 可能无此属性，故用反射设置，失败静默忽略。
            
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

            // DPI 放大（确定性，每次打开都执行一次）：
            // 实测 PS 2402 宿主不会触发 WinForms 的 AutoScale，控件停在 96-DPI 设计尺寸（偏小）。
            // 这里按设备 DPI 手动整体放大一次。先重置 Size 为设计尺寸（96-DPI 基准），
            // 再 Scale，避免 TxForm 持久化尺寸逐次叠加放大。AutoScaleMode 已设为 None。
            if (!_dpiApplied)
            {
                _dpiApplied = true;
                try
                {
                    // 重置为设计尺寸，防止 TxForm 持久化尺寸叠加放大
                    Size = new Size(960, 800);
                    float sc = DpiScale();
                    if (sc < 1f) sc = 1f;
                    if (sc > 1.01f) Scale(new SizeF(sc, sc));

                    // 两侧列固定宽（按内容贴合）、中间点列表吃掉剩余宽度。
                    // 用已缩放像素显式重设，不依赖 TableLayoutPanel 对 Absolute 列的缩放行为。
                    ApplySideColumnWidths(sc);
                }
                catch { }
            }

            if (_svc == null) return;
            BindFrameComboEvents();

            Log("[系统] 插件已启动");
            Log("[系统] 在左侧操作列表中拾取 PS 对象即可开始");

            if (_listAdapter is GridAdapter)
            {
                Log("[系统] 点击列表内高亮行，在 PS 中选中对象即可加入");
                // TxObjGridCtrl 启动后首次拾取需要先获得焦点
                BeginInvoke(new Action(delegate ()
                {
                    try
                    {
                        if (_objGrid != null && _objGrid.Visible)
                        {
                            _objGrid.Focus();
                            try { _objGrid.SetCurrentCell(0, 0); } catch { }
                        }
                    }
                    catch { }
                }));
            }
        }

        // ════════════════════════════════════════════════════════════
        //  TxObjEditBoxCtrl 事件（回落模式用）
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
                { UI(delegate () { Log("[错误] " + ex.Message, LogLevel.Error); }); }

                if (ops != null && ops.Count > 0)
                {
                    UI(delegate ()
                    {
                        MergeOpsByName(ops);
                        RefreshOpList();
                        ReloadPointsList();
                        UpdatePointCount();
                        Log("[PS] 通过原生选择器添加 " + ops.Count + " 个操作", LogLevel.Ps);
                    });
                }
            });
        }

        // ════════════════════════════════════════════════════════════
        //  TxFrameComboBoxCtrl 事件
        // ════════════════════════════════════════════════════════════
        private void BindFrameComboEvents()
        {
            if (_frameCombo != null)
            {
                try
                {
                    _frameCombo.ValidFrameSet += new TxFrameComboBoxCtrl_ValidFrameSetEventHandler(OnFrameValidSet);
                    _frameCombo.InvalidFrameSet += new TxFrameComboBoxCtrl_InvalidFrameSetEventHandler(OnFrameInvalidSet);
                    // 不绑 Picked：原实现为空方法，无意义
                }
                catch (Exception ex) { Log("[警告] TxFrameComboBoxCtrl 事件绑定失败: " + ex.Message, LogLevel.Warn); }
            }
            // _tcpFrameCombo 改为懒加载：仅在用户选「自定义坐标…」时创建并绑定（见 EnsureTcpFrameCombo）
        }

        /// <summary>懒加载自定义 TCP 坐标控件：首次需要时才创建，加快窗体打开速度。</summary>
        private void EnsureTcpFrameCombo()
        {
            if (_tcpFrameCombo != null || _tcpCustomPanel == null) return;
            try
            {
                _tcpFrameCombo = new TxFrameComboBoxCtrl { Dock = DockStyle.Fill, ListenToPick = true };
                _tcpCustomPanel.Controls.Add(_tcpFrameCombo);
                _tcpFrameCombo.ValidFrameSet += new TxFrameComboBoxCtrl_ValidFrameSetEventHandler(OnTcpFrameValidSet);
                _tcpFrameCombo.InvalidFrameSet += new TxFrameComboBoxCtrl_InvalidFrameSetEventHandler(OnTcpFrameInvalidSet);
            }
            catch (Exception ex)
            {
                _tcpCustomPanel.Controls.Add(new Label
                {
                    Text = "坐标控件不可用: " + ex.Message,
                    Dock = DockStyle.Fill,
                    ForeColor = Theme.BtnDanger,
                    Font = SystemFonts.MessageBoxFont
                });
                Log("[警告] 自定义TCP坐标控件创建失败: " + ex.Message, LogLevel.Warn);
            }
        }

        // ── 自定义 TCP 坐标选择器事件 ──
        private void OnTcpFrameValidSet(object sender, TxFrameComboBoxCtrl_ValidFrameSetEventArgs e)
        {
            try
            {
                TxTransformation tx = e.Location as TxTransformation;
                if (tx == null) { return; }
                double[] arr = PsReader.TxToArr(tx);
                _tcpCustomMatrix = arr;
                string name = "自定义坐标";
                if (e.Object != null)
                {
                    ITxObject txObj = e.Object as ITxObject;
                    if (txObj != null) try { name = txObj.Name; } catch { }
                }
                if (_lblTcpInfo != null) _lblTcpInfo.Text = "TCP：自定义（" + name + "）";
                Log("[TCP] 自定义 TCP 坐标已设置：" + name, LogLevel.Coord);
            }
            catch (Exception ex) { Log("[TCP] 自定义坐标异常：" + ex.Message, LogLevel.Error); }
        }

        private void OnTcpFrameInvalidSet(object sender, TxFrameComboBoxCtrl_InvalidFrameSetEventArgs e)
        { Log("[TCP] 所选对象不是有效坐标系", LogLevel.Coord); }

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
                        _refFrameName = name;
                        _refFrameMatrix = arr;
                        UpdateRefCoordStatus(name, false);
                        Log("[坐标] 参考坐标已设置：" + name, LogLevel.Coord);
                        return;
                    }
                }
                string fallbackName = "用户选定坐标系";
                if (e.Object != null)
                {
                    ITxObject txObj = e.Object as ITxObject;
                    if (txObj != null) try { fallbackName = txObj.Name; } catch { }
                }
                _refFrameName = fallbackName;
                _refFrameMatrix = null;
                UpdateRefCoordStatus(fallbackName + " (世界坐标系)", true);
                Log("[坐标] 坐标系: " + fallbackName + " (与世界坐标系等同)", LogLevel.Coord);
            }
            catch (Exception ex) { Log("[坐标] 异常: " + ex.Message, LogLevel.Error); }
        }

        private void OnFrameInvalidSet(object sender, TxFrameComboBoxCtrl_InvalidFrameSetEventArgs e)
        { Log("[坐标] 所选对象不是有效坐标系", LogLevel.Coord); }

        private void UpdateRefCoordStatus(string name, bool isWorld)
        {
            if (_lblRefCoordStatus == null) return;
            _lblRefCoordStatus.Text = "当前：" + name;
            _lblRefCoordStatus.BackColor = isWorld ? Theme.StatusOkBg : Theme.StatusRefBg;
            _lblRefCoordStatus.ForeColor = isWorld ? Theme.StatusOkFg : Theme.StatusRefFg;
        }

        // ════════════════════════════════════════════════════════════
        //  界面搭建
        // ════════════════════════════════════════════════════════════
        private void BuildUI()
        {
            SuspendLayout();
            // 高 DPI 一致性：以 96 DPI 为设计基准，按设备 DPI 自动缩放整套布局
            // DPI：关掉不可靠的 AutoScale（PS 2402 宿主不触发），改由 OnLoad 手动 Scale 放大。
            // AutoScaleDimensions 在 None 下被忽略，保留无害。
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.None;
            Text = "导出插枪 / 点云到 CATIA";
            Size = new Size(960, 800);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            BackColor = SystemColors.Control;
            Font = SystemFonts.MessageBoxFont;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Margin = Padding.Empty,
                Padding = new Padding(6, 4, 6, 4)
            };
            _root = root;
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));            // 0 标题
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));           // 1 主体
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, LOG_COLLAPSED_H)); // 2 日志（可折叠）
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 46F));           // 3 底栏
            root.Controls.Add(BuildHeaderRow(), 0, 0);
            root.Controls.Add(BuildBody(), 0, 1);
            root.Controls.Add(BuildLogPanel(), 0, 2);
            root.Controls.Add(BuildBottom(), 0, 3);
            Controls.Add(root);
            ResumeLayout(false);
            PerformLayout();
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
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));  // 左：通用信息（固定宽）
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 390F));   // 中：操作点列表（吃掉剩余宽度）
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));  // 右：导插枪/点球（固定宽，贴合内容）
            row.Controls.Add(MkHeaderLabel("通用信息", Theme.TxCol1), 0, 0);
            row.Controls.Add(MkHeaderLabel("操作点列表", Theme.TxPoints), 1, 0);
            row.Controls.Add(MkHeaderLabel("导插枪 / 点球", Theme.TxGun), 2, 0);
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
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));  // 左：通用信息（固定宽）
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 390F));   // 中：操作点列表（吃掉剩余宽度）
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));  // 右：导插枪/点球（固定宽，贴合内容）
            body.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            body.Controls.Add(BuildCol1(), 0, 0);          // 通用信息（不变）
            body.Controls.Add(BuildPointsColumn(), 1, 0);  // 操作点列表（满高勾选）
            body.Controls.Add(BuildCol2(), 2, 0);          // 导插枪 / 点球（原第二栏移来）
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
            _progressBar = new ProgressBar
            {
                Height = 18,
                Width = 260,
                Style = ProgressBarStyle.Continuous,
                Margin = new Padding(0, 2, 8, 2),
                Visible = false
            };
            _lblProgress = new Label
            {
                AutoSize = true,
                Text = "就绪",
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 4, 0, 2)
            };
            lf.Controls.Add(_progressBar);
            lf.Controls.Add(_lblProgress);

            var rf = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 4, 0, 0),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            _btnReset = MkFuncButton("复位", Theme.BtnMuted);
            _btnExportExcel = MkFuncButton("导出Excel", Theme.BtnExport);
            _btnClose = MkFuncButton("关闭", Theme.BtnDanger);
            _btnHelp = MkFuncButton("?", Theme.TxAccent.Color);
            _btnReset.Click += new EventHandler(OnReset);
            _btnExportExcel.Click += new EventHandler(OnExportExcel);
            _btnClose.Click += delegate { Close(); };
            _btnHelp.Click += delegate { ShowHelp(); };
            rf.Controls.AddRange(new Control[] { _btnReset, _btnExportExcel, _btnClose, _btnHelp });

            bottom.Controls.Add(lf, 0, 0);
            bottom.Controls.Add(rf, 1, 0);
            return bottom;
        }

        // ════════════════════════════════════════════════════════════
        //  第1列：通用信息
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

            stack.Controls.Add(BuildOpsCard(), 0, 0);
            stack.Controls.Add(BuildCoordCard(), 0, 1);
            stack.Controls.Add(BuildSettingsCard(), 0, 2);

            scroll.Controls.Add(stack);
            return scroll;
        }

        // ── 卡片：操作选择 ──
        private GroupBox BuildOpsCard()
        {
            var card = MkCard("操作选择", Theme.TxCol1);
            var flow = MkCardContent();

            var gridPanel = new Panel
            {
                AutoSize = false,
                Height = 220,
                BackColor = SystemColors.Window,
                Padding = new Padding(1),
                Margin = new Padding(0, 0, 0, 4),
                BorderStyle = BorderStyle.FixedSingle
            };

            // 主方案：TxObjGridCtrl
            bool usingGrid = false;
            _objGrid = new TxObjGridCtrl
            {
                Dock = DockStyle.Fill,
                ListenToPick = true,
                EnableMultipleSelection = true,
                EnableRecurringObjects = false,
            };
            _objGrid.ObjectInserted += new TxObjGridCtrl_ObjectInsertedEventHandler(OnGridObjectInserted);
            _objGrid.RowDeleted += new TxObjGridCtrl_RowDeletedEventHandler(OnGridRowDeleted);
            gridPanel.Controls.Add(_objGrid);
            usingGrid = true;

            // 创建对应的 adapter
            _listAdapter = usingGrid
                ? (IOpListAdapter)new GridAdapter(this)
                : new ListViewAdapter(this);

            _lblListHint = new Label
            {
                Text = "列表为空",
                AutoSize = true,
                Dock = DockStyle.Bottom,
                Padding = new Padding(0, 2, 0, 2),
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = SystemColors.GrayText,
                BackColor = SystemColors.ControlLight,
                Font = SystemFonts.MessageBoxFont
            };
            gridPanel.Controls.Add(_lblListHint);
            flow.Controls.Add(gridPanel);
            FillWidthInFlow(flow, gridPanel);   // 宽度填满卡片

            var pickRow = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 0, 0, 2)
            };
            _btnPickFromSel = MkFuncButton("拾取全部操作", Theme.BtnSecondary);
            _btnPickFromSel.Click += new EventHandler(OnPickSel);
            var btnClearOps = MkFuncButton("清空列表", Theme.BtnMuted);
            btnClearOps.Click += new EventHandler(OnClearOps);
            pickRow.Controls.AddRange(new Control[] { _btnPickFromSel, btnClearOps });
            flow.Controls.Add(pickRow);

            var opsHint = new Label
            {
                Text = usingGrid
                    ? "提示：亮绿色行为拾取入口，点击后去 PS 中选中对象即可加入"
                    : "在上方选择框拾取单个操作，自动加入列表；也可在PS中选中后点[拾取自PS]",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 0, 0, 6)
            };
            flow.Controls.Add(opsHint);
            WrapLabelInFlow(flow, opsHint);

            card.Controls.Add(flow);
            return card;
        }

        // ── 卡片：参考坐标 ──
        private GroupBox BuildCoordCard()
        {
            var card = MkCard("参考坐标", Theme.TxCol1);
            var flow = MkCardContent();

            var frameCtrlPanel = new Panel
            {
                AutoSize = false,
                Height = 28,
                Margin = new Padding(0, 2, 0, 4)
            };
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
                    ForeColor = Theme.BtnDanger,
                    Font = SystemFonts.MessageBoxFont
                });
            }
            flow.Controls.Add(frameCtrlPanel);
            FillWidthInFlow(flow, frameCtrlPanel);   // 坐标框宽度填满卡片

            _lblRefCoordStatus = new Label
            {
                Text = "当前：世界坐标系",
                AutoSize = true,
                ForeColor = Theme.StatusOkFg,
                BackColor = Theme.StatusOkBg,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 0, 0, 4),
                Padding = new Padding(6, 3, 6, 3)
            };
            flow.Controls.Add(_lblRefCoordStatus);

            var btnClrCoord = MkFuncButton("重置为世界坐标系", Theme.BtnMuted);
            btnClrCoord.Margin = new Padding(0, 0, 0, 2);
            btnClrCoord.Click += delegate
            {
                _refFrameName = "世界坐标系";
                _refFrameMatrix = null;
                if (_frameCombo != null) try { _frameCombo.Clear(); } catch { }
                UpdateRefCoordStatus("世界坐标系", true);
                Log("[坐标] 已重置为世界坐标系", LogLevel.Coord);
            };
            flow.Controls.Add(btnClrCoord);

            var coordHint = new Label
            {
                Text = "在PS中选择Component/Frame即可自动获取",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 0, 0, 2)
            };
            flow.Controls.Add(coordHint);
            WrapLabelInFlow(flow, coordHint);

            card.Controls.Add(flow);
            return card;
        }

        // ── 卡片：导出设置 ──
        private GroupBox BuildSettingsCard()
        {
            var card = MkCard("导出设置", Theme.TxCol1);
            var flow = MkCardContent();

            var rowType = MkRowFlow();
            rowType.Controls.Add(MkLabel("点类型"));
            _cmbPointType = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            _cmbPointType.Items.AddRange(new object[] { "焊点", "路径点", "连续点", "全部类型" });
            _cmbPointType.SelectedIndex = 3;
            _cmbPointType.SelectedIndexChanged += delegate { RebuildPointsListView(); UpdatePointCount(); };
            AutoFitComboBoxWidth(_cmbPointType);
            rowType.Controls.Add(_cmbPointType);
            flow.Controls.Add(rowType);

            _chkUseMfgName = new CheckBox
            {
                Text = "采用MFG名称",
                AutoSize = true,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 2, 0, 2)
            };
            _chkUseMfgName.CheckedChanged += delegate { ReloadPointsList(); };
            flow.Controls.Add(_chkUseMfgName);

            _lblPointCount = new Label
            {
                Text = "将导出点数量：0",
                AutoSize = true,
                ForeColor = Theme.StatusOkFg,
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold),
                Margin = new Padding(0, 2, 0, 4)
            };
            flow.Controls.Add(_lblPointCount);

            card.Controls.Add(flow);
            return card;
        }

        // ════════════════════════════════════════════════════════════
        //  第2列：导插枪 + 导点球
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

            stack.Controls.Add(BuildGunCard(), 0, 0);
            stack.Controls.Add(BuildBallCard(), 0, 1);

            scroll.Controls.Add(stack);
            return scroll;
        }

        // ── 卡片：导插枪 ──
        private GroupBox BuildGunCard()
        {
            var card = MkCard("导插枪", Theme.TxGun);
            var flow = MkCardContent();

            // 自定义产品名
            var rowProd = MkRowFlow();
            rowProd.Controls.Add(MkLabel("产品名称"));
            _txtGunProductName = new TextBox
            {
                Width = 140,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            rowProd.Controls.Add(_txtGunProductName);
            flow.Controls.Add(rowProd);
            var prodHint = new Label
            {
                Text = "留空则使用活动文档内产品名",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 0, 0, 2)
            };
            flow.Controls.Add(prodHint);
            WrapLabelInFlow(flow, prodHint);

            var rowCustom = MkRowFlow();
            _chkCustomGun = new CheckBox
            {
                Text = "自定义焊枪数模",
                AutoSize = true,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 4, 6, 2)
            };
            var seletGun = MkRowFlow();
            _txtGunModel = new TextBox
            {
                Width = 120,
                ReadOnly = true,
                Enabled = false,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 4, 6, 2)
            };
            _btnBrowseGun = MkFuncButton("选择", Theme.BtnMuted);
            _btnBrowseGun.Enabled = false;
            _chkCustomGun.CheckedChanged += delegate
            {
                _txtGunModel.Enabled = _chkCustomGun.Checked;
                _btnBrowseGun.Enabled = _chkCustomGun.Checked;
            };
            _btnBrowseGun.Click += new EventHandler(OnBrowseGun);
            rowCustom.Controls.AddRange(new Control[] { _chkCustomGun });
            seletGun.Controls.AddRange(new Control[] { _txtGunModel, _btnBrowseGun });
            flow.Controls.Add(rowCustom);
            flow.Controls.Add(seletGun);

            _lblGunInfo = new Label
            {
                Text = "焊钳：（待选取）",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 2, 0, 4)
            };
            flow.Controls.Add(_lblGunInfo);

            // ── 导出方式（共享几何 vs 独立命名） ──
            var rowMode = MkRowFlow();
            rowMode.Controls.Add(MkLabel("导出方式"));
            _cmbGunExportMode = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 220,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            _cmbGunExportMode.Items.Add("共享几何（体积最小，名称=CGR原名）");
            _cmbGunExportMode.Items.Add("独立命名（体积×焊点数，名称=焊点名）");
            _cmbGunExportMode.SelectedIndex = 0;
            rowMode.Controls.Add(_cmbGunExportMode);
            flow.Controls.Add(rowMode);

            var modeHint = new Label
            {
                Text = "共享几何：3DXML ≈ 1 份 CGR，实例名保持 CGR 原名\n独立命名：3DXML ≈ N 份 CGR，实例名可读为焊点名",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 0, 0, 4)
            };
            flow.Controls.Add(modeHint);
            WrapLabelInFlow(flow, modeHint);

            _chkGunOriginTCP = new CheckBox
            {
                Text = "焊枪以TCP为原点",
                AutoSize = true,
                Checked = true,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 2, 0, 2)
            };
            flow.Controls.Add(_chkGunOriginTCP);

            // ── TCP 选择（多 TCP 下拉 + 自定义坐标选择器）──
            var rowTcp = MkRowFlow();
            rowTcp.Controls.Add(MkLabel("TCP"));
            _cmbTcp = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 200,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            _cmbTcp.Items.Add(PsReader.DefaultTcpLabel);
            _cmbTcp.SelectedIndex = 0;
            _cmbTcp.SelectedIndexChanged += new EventHandler(OnTcpChoiceChanged);
            rowTcp.Controls.Add(_cmbTcp);
            flow.Controls.Add(rowTcp);

            // 自定义 TCP 坐标选择器容器（默认隐藏；内部坐标控件懒加载，加快打开速度）
            _tcpCustomPanel = new Panel
            {
                AutoSize = false,
                Height = 30,
                Width = 250,
                Visible = false,
                Margin = new Padding(0, 0, 0, 2)
            };
            flow.Controls.Add(_tcpCustomPanel);
            FillWidthInFlow(flow, _tcpCustomPanel);

            _lblTcpInfo = new Label
            {
                Text = "TCP：默认（机器人当前TCP）",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 0, 0, 4)
            };
            flow.Controls.Add(_lblTcpInfo);
            WrapLabelInFlow(flow, _lblTcpInfo);

            _chkExportTCP = new CheckBox
            {
                Text = "导出TCP坐标",
                AutoSize = true,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 2, 0, 4)
            };
            flow.Controls.Add(_chkExportTCP);

            _btnExportGun = MkFuncButton("导出插枪", Theme.BtnGun);
            _btnExportGun.Margin = new Padding(0, 6, 0, 4);
            _btnExportGun.Click += new EventHandler(OnExportGun);
            flow.Controls.Add(_btnExportGun);

            card.Controls.Add(flow);
            return card;
        }

        // ── 卡片：导点球 ──
        private GroupBox BuildBallCard()
        {
            var card = MkCard("导点球", Theme.TxBall);
            var flow = MkCardContent();

            var rowTgt = MkRowFlow();
            rowTgt.Controls.Add(MkLabel("导出到"));
            _cmbBallTarget = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            _cmbBallTarget.Items.AddRange(new object[] { "当前Part文档", "新建Part文档" });
            _cmbBallTarget.SelectedIndex = 0;
            AutoFitComboBoxWidth(_cmbBallTarget);
            rowTgt.Controls.Add(_cmbBallTarget);
            flow.Controls.Add(rowTgt);

            var rowOpt = MkRowFlow();
            rowOpt.Controls.Add(MkLabel("导出选项"));
            _cmbBallOption = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            _cmbBallOption.Items.AddRange(new object[] { "轨迹点 + 点球", "仅轨迹点", "仅点球" });
            _cmbBallOption.SelectedIndex = 0;
            AutoFitComboBoxWidth(_cmbBallOption);
            rowOpt.Controls.Add(_cmbBallOption);
            flow.Controls.Add(rowOpt);

            var rowDiam = MkRowFlow();
            rowDiam.Controls.Add(MkLabel("球直径(mm)"));
            _nudDiameter = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 500,
                Value = 10,
                DecimalPlaces = 0,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            AutoFitNumericWidth(_nudDiameter);
            rowDiam.Controls.Add(_nudDiameter);
            flow.Controls.Add(rowDiam);

            var rowGeom = MkRowFlow();
            rowGeom.Controls.Add(MkLabel("几何集"));
            _txtGeomSet = new TextBox
            {
                Text = "Geometry_Spheres",
                Width = 140,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            rowGeom.Controls.Add(_txtGeomSet);
            flow.Controls.Add(rowGeom);

            var rowPrefix = MkRowFlow();
            rowPrefix.Controls.Add(MkLabel("名称前缀"));
            _txtNamePrefix = new TextBox
            {
                Text = "SPHERE",
                Width = 140,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            rowPrefix.Controls.Add(_txtNamePrefix);
            flow.Controls.Add(rowPrefix);

            var rowPart = MkRowFlow();
            rowPart.Controls.Add(MkLabel("零件名称"));
            _txtBallPartName = new TextBox
            {
                Width = 140,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(2, 3, 0, 0)
            };
            rowPart.Controls.Add(_txtBallPartName);
            flow.Controls.Add(rowPart);
            var partHint = new Label
            {
                Text = "新建Part时生效，留空用默认名",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.MessageBoxFont,
                Margin = new Padding(0, 0, 0, 2)
            };
            flow.Controls.Add(partHint);
            WrapLabelInFlow(flow, partHint);

            _btnExportBall = MkFuncButton("导出点球", Theme.TxBall.Color);
            _btnExportBall.Margin = new Padding(0, 6, 0, 4);
            _btnExportBall.Click += new EventHandler(OnExportBall);
            flow.Controls.Add(_btnExportBall);

            card.Controls.Add(flow);
            return card;
        }

        // ════════════════════════════════════════════════════════════
        //  第2列（新）：操作点列表 —— 满高、按操作分组、勾选、默认全选
        // ════════════════════════════════════════════════════════════
        private Control BuildPointsColumn()
        {
            var card = new ColoredGroupBox
            {
                Text = "操作点列表",
                Dock = DockStyle.Fill,
                AutoSize = false,
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold),
                TitleColor = Theme.TxPoints.Color,
                Margin = new Padding(2, 2, 2, 4),
                Padding = new Padding(8, 6, 8, 6)
            };

            var container = new Panel { Dock = DockStyle.Fill, BackColor = Color.Transparent };

            // ── 列表（Fill，先加入）──
            _lvPoints = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                CheckBoxes = true,
                FullRowSelect = true,
                MultiSelect = true,
                HideSelection = false,
                ShowGroups = true,
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                BackColor = SystemColors.Window,
                Font = SystemFonts.MessageBoxFont
            };
            _lvPoints.Columns.Add("点名称", 170);
            _lvPoints.Columns.Add("类型", 70);
            _lvPoints.ItemChecked += new ItemCheckedEventHandler(OnPointItemChecked);
            _lvPoints.Resize += delegate { StretchPointsColumn(); };
            container.Controls.Add(_lvPoints);

            // ── 工具条（Top）──
            var bar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 0, 0, 2)
            };
            var btnAll = MkFuncButton("全选", Theme.BtnSecondary);
            var btnNone = MkFuncButton("全不选", Theme.BtnMuted);
            var btnInv = MkFuncButton("反选", Theme.BtnMuted);
            btnAll.Click += delegate { SetAllPointsChecked(true); };
            btnNone.Click += delegate { SetAllPointsChecked(false); };
            btnInv.Click += delegate { SetAllPointsChecked(null); };
            bar.Controls.AddRange(new Control[] { btnAll, btnNone, btnInv });
            container.Controls.Add(bar);

            // ── 计数条（Bottom）──
            _lblPointSel = new Label
            {
                Dock = DockStyle.Bottom,
                Text = "已选 0 / 共 0 点",
                AutoSize = false,
                Height = 22,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = SystemColors.GrayText,
                BackColor = SystemColors.ControlLight,
                Font = SystemFonts.MessageBoxFont,
                Padding = new Padding(6, 0, 0, 0)
            };
            container.Controls.Add(_lblPointSel);

            card.Controls.Add(container);
            return card;
        }

        // ════════════════════════════════════════════════════════════
        //  底部：可折叠日志面板（原日志栏改为折叠）
        // ════════════════════════════════════════════════════════════
        private Control BuildLogPanel()
        {
            var panel = new Panel { Dock = DockStyle.Fill, Margin = new Padding(0, 2, 0, 0) };

            // 日志正文（Fill，先加入；默认隐藏）
            _rtbLog = new RichTextBox
            {
                Dock = DockStyle.Fill,
                BackColor = Theme.LogBg,
                ForeColor = Theme.LogText,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                Font = new Font("Consolas", 8f),
                ScrollBars = RichTextBoxScrollBars.Vertical,
                WordWrap = false,
                Visible = false
            };
            panel.Controls.Add(_rtbLog);

            // 标题条（Top）：点击切换展开/收起 + 右侧清除按钮
            var bar = new Panel { Dock = DockStyle.Top, Height = 24, BackColor = Theme.TxLog.Color };

            _logToggle = new Label
            {
                Dock = DockStyle.Fill,
                Text = "▸  运行日志（点击展开）",
                ForeColor = TxColor.TxColorWhite.Color,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold),
                Padding = new Padding(8, 0, 0, 0),
                Cursor = Cursors.Hand
            };
            _logToggle.Click += delegate { ToggleLog(); };
            bar.Click += delegate { ToggleLog(); };
            bar.Controls.Add(_logToggle);

            var btnClear = new FlatColorButton
            {
                Text = "清除",
                Dock = DockStyle.Right,
                Width = 56,
                BgColor = Theme.TxAccent.Color,
                ForeColor = TxColor.TxColorWhite.Color,
                Font = SystemFonts.MessageBoxFont
            };
            btnClear.Click += delegate { if (_rtbLog != null) _rtbLog.Clear(); };
            bar.Controls.Add(btnClear);

            panel.Controls.Add(bar);

            _logExpanded = false;   // 默认折叠
            return panel;
        }

        /// <summary>
        /// 设置主体/标题行的三列：左右固定宽（按 DPI 缩放），中间点列表占满剩余宽度。
        /// 同时作用于标题行和主体行，保证三个彩色标题与下方三列对齐。
        /// </summary>
        private void ApplySideColumnWidths(float sc)
        {
            int leftW = (int)(330 * sc);
            int rightW = (int)(300 * sc);
            TableLayoutPanel[] tables =
            {
                _root != null ? _root.GetControlFromPosition(0, 0) as TableLayoutPanel : null, // 标题行
                _root != null ? _root.GetControlFromPosition(0, 1) as TableLayoutPanel : null  // 主体
            };
            foreach (TableLayoutPanel t in tables)
            {
                if (t == null || t.ColumnStyles.Count < 3) continue;
                t.ColumnStyles[0].SizeType = SizeType.Absolute; t.ColumnStyles[0].Width = leftW;
                t.ColumnStyles[1].SizeType = SizeType.Percent; t.ColumnStyles[1].Width = 100F;
                t.ColumnStyles[2].SizeType = SizeType.Absolute; t.ColumnStyles[2].Width = rightW;
            }
        }

        /// <summary>当前 DPI 相对 96 的缩放系数（PS 宿主有时不触发 AutoScale，此处兜底计算）。</summary>
        private float DpiScale()
        {
            try { using (var g = CreateGraphics()) return g.DpiX / 96f; }
            catch { return 1f; }
        }

        private void ToggleLog()
        {
            _logExpanded = !_logExpanded;

            float sc = DpiScale();
            int collapsedH = (int)Math.Round(LOG_COLLAPSED_H * sc);
            int expandedH = (int)Math.Round(LOG_EXPANDED_H * sc);
            int delta = expandedH - collapsedH;
            if (_root != null && _root.RowStyles.Count > 2)
                _root.RowStyles[2].Height = _logExpanded ? expandedH : collapsedH;

            // 日志向下展开：同步增/减窗体高度，避免压缩上方主体区域
            try
            {
                SuspendLayout();
                int target = Height + (_logExpanded ? delta : -delta);
                int maxH = Screen.FromControl(this).WorkingArea.Height;
                int minH = MinimumSize.Height > 0 ? MinimumSize.Height : 0;
                if (target > maxH) target = maxH;
                if (target < minH) target = minH;
                Height = target;
            }
            catch { }
            finally { ResumeLayout(true); }

            if (_rtbLog != null)
            {
                _rtbLog.Visible = _logExpanded;
                if (_logExpanded) try { _rtbLog.ScrollToCaret(); } catch { }
            }
            if (_logToggle != null)
                _logToggle.Text = _logExpanded ? "▾  运行日志（点击收起）" : "▸  运行日志（点击展开）";
        }

        // ════════════════════════════════════════════════════════════
        //  操作点列表 —— 加载 / 重建 / 勾选 / 计数
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 根据当前 _ops 重新加载所有点。
        /// 后台线程在 PS 上下文中对每个 op 重新 FillPoints(All)，
        /// 默认全部勾选，完成后回 UI 线程重建列表。
        /// </summary>
        private void ReloadPointsList()
        {
            if (_lvPoints == null) return;

            int token = ++_pointsLoadToken;
            var opsSnapshot = new List<OperationInfo>(_ops);

            if (opsSnapshot.Count == 0)
            {
                _pointRows = new List<PointRow>();
                _pointsLoaded = true;
                RebuildPointsListView();
                UpdatePointCount();
                ReloadTcpOptions();
                return;
            }

            _pointsLoaded = false;
            _suppressPointCheck = true;
            _lvPoints.BeginUpdate();
            _lvPoints.Items.Clear();
            _lvPoints.Groups.Clear();
            _lvPoints.EndUpdate();
            _suppressPointCheck = false;
            if (_lblPointSel != null) _lblPointSel.Text = "正在加载点…";

            bool useMfg = _chkUseMfgName.Checked;
            ThreadPool.QueueUserWorkItem(delegate (object s)
            {
                var rows = new List<PointRow>();
                var logBuf = new List<string>();
                Action<string> collect = delegate (string m) { lock (logBuf) logBuf.Add(m); };

                try
                {
                    _svc.InvokeOnPs(delegate ()
                    {
                        foreach (var op in opsSnapshot)
                        {
                            try
                            {
                                // 强制刷新：清空后按 All 重新填充，保证类型/名称与当前设置一致
                                if (op.Points == null) op.Points = new List<PointInfo>();
                                else op.Points.Clear();
                                PsReader.FillPoints(op, PointType.All, useMfg, collect);
                            }
                            catch (Exception ex)
                            { collect("[点列表] 填充失败 " + op.Name + "：" + ex.Message); }

                            if (op.Points != null)
                                foreach (var pt in op.Points)
                                    rows.Add(new PointRow { Op = op, Pt = pt, Checked = true });
                        }
                    });
                }
                catch (Exception ex) { collect("[点列表] PS 线程异常：" + ex.Message); }

                UI(delegate ()
                {
                    if (token != _pointsLoadToken) return;   // 已有更新的加载，丢弃本次结果
                    foreach (var line in logBuf) Log(line);
                    _pointRows = rows;
                    _pointsLoaded = true;
                    RebuildPointsListView();
                    UpdatePointCount();
                    Log("[点列表] 已加载 " + rows.Count + " 个点（默认全选）", LogLevel.Ps);
                    ReloadTcpOptions();
                });
            });
        }

        /// <summary>按当前点类型筛选，重建 ListView 显示（保留各行勾选状态）。</summary>
        private void RebuildPointsListView()
        {
            if (_lvPoints == null) return;
            PointType filter = GetPtType();

            _suppressPointCheck = true;
            _lvPoints.BeginUpdate();
            _lvPoints.Items.Clear();
            _lvPoints.Groups.Clear();

            var groups = new Dictionary<OperationInfo, ListViewGroup>();
            foreach (var row in _pointRows)
            {
                row.Item = null;
                if (filter != PointType.All && row.Pt.Type != filter) continue;

                ListViewGroup g;
                if (!groups.TryGetValue(row.Op, out g))
                {
                    g = new ListViewGroup(row.Op.Name == null ? "(操作)" : row.Op.Name)
                    { Header = row.Op.Name };
                    groups[row.Op] = g;
                    _lvPoints.Groups.Add(g);
                }

                var item = new ListViewItem(
                    string.IsNullOrEmpty(row.Pt.Name) ? "(未命名)" : row.Pt.Name, g)
                {
                    Tag = row,
                    Checked = row.Checked
                };
                item.SubItems.Add(PtTypeLabel(row.Pt.Type));
                row.Item = item;
                _lvPoints.Items.Add(item);
            }

            _lvPoints.EndUpdate();
            _suppressPointCheck = false;

            StretchPointsColumn();
            UpdatePointsSelHeader();
        }

        private void OnPointItemChecked(object sender, ItemCheckedEventArgs e)
        {
            if (_suppressPointCheck) return;
            var row = e.Item.Tag as PointRow;
            if (row != null) row.Checked = e.Item.Checked;
            UpdatePointsSelHeader();
            UpdatePointCount();
        }

        /// <summary>state==true 全选；false 全不选；null 反选（仅作用于当前显示的点）。</summary>
        private void SetAllPointsChecked(bool? state)
        {
            if (_lvPoints == null) return;
            _suppressPointCheck = true;
            _lvPoints.BeginUpdate();
            foreach (ListViewItem it in _lvPoints.Items)
            {
                bool v = state ?? !it.Checked;
                it.Checked = v;
                var row = it.Tag as PointRow;
                if (row != null) row.Checked = v;
            }
            _lvPoints.EndUpdate();
            _suppressPointCheck = false;
            UpdatePointsSelHeader();
            UpdatePointCount();
        }

        private void UpdatePointsSelHeader()
        {
            if (_lblPointSel == null) return;
            if (!_pointsLoaded)
            {
                _lblPointSel.Text = _ops.Count == 0 ? "已选 0 / 共 0 点" : "正在加载点…";
                _lblPointSel.ForeColor = SystemColors.GrayText;
                return;
            }
            int shown = 0, sel = 0;
            PointType filter = GetPtType();
            foreach (var row in _pointRows)
            {
                if (filter != PointType.All && row.Pt.Type != filter) continue;
                shown++;
                if (row.Checked) sel++;
            }
            _lblPointSel.Text = "已选 " + sel + " / 共 " + shown + " 点";
            _lblPointSel.ForeColor = sel > 0 ? Theme.StatusOkFg : Theme.BtnDanger;
        }

        private void StretchPointsColumn()
        {
            if (_lvPoints == null || _lvPoints.Columns.Count < 2) return;
            int w = _lvPoints.ClientSize.Width - _lvPoints.Columns[1].Width - 4;
            if (w > 80) _lvPoints.Columns[0].Width = w;
        }

        private static string PtTypeLabel(PointType t)
        {
            switch (t)
            {
                case PointType.WeldPoint: return "焊点";
                case PointType.PathPoint: return "路径点";
                case PointType.ContinuousPoint: return "连续点";
                default: return "其他";
            }
        }

        // ════════════════════════════════════════════════════════════
        //  TCP 选项加载与选择
        // ════════════════════════════════════════════════════════════
        /// <summary>根据当前操作枚举可用 TCP，填充下拉。后台线程 + InvokeOnPs。</summary>
        private void ReloadTcpOptions()
        {
            if (_cmbTcp == null) return;
            int token = ++_tcpLoadToken;

            // 取一个代表性操作（首个）枚举其工具的 TCP
            OperationInfo repOp = (_ops != null && _ops.Count > 0) ? _ops[0] : null;
            if (repOp == null)
            {
                ResetTcpCombo(null);
                return;
            }

            ThreadPool.QueueUserWorkItem(delegate (object s)
            {
                var names = new List<string>();
                var logBuf = new List<string>();
                Action<string> collect = delegate (string m) { lock (logBuf) logBuf.Add(m); };
                try
                {
                    _svc.InvokeOnPs(delegate ()
                    {
                        var opts = PsReader.EnumerateTcpOptions(repOp, collect);
                        foreach (var o in opts) if (!names.Contains(o.Name)) names.Add(o.Name);
                    });
                }
                catch (Exception ex) { collect("[TCP] 枚举异常：" + ex.Message); }

                UI(delegate ()
                {
                    if (token != _tcpLoadToken) return;
                    foreach (var line in logBuf) Log(line);
                    ResetTcpCombo(names);
                });
            });
        }

        /// <summary>用枚举到的名字重建下拉（保留当前选择，追加“自定义坐标…”）。</summary>
        private void ResetTcpCombo(List<string> names)
        {
            if (_cmbTcp == null) return;
            string prev = _cmbTcp.SelectedItem as string;

            _cmbTcp.BeginUpdate();
            _cmbTcp.Items.Clear();
            _cmbTcp.Items.Add(PsReader.DefaultTcpLabel);   // 默认项始终在首位
            if (names != null)
                foreach (var n in names)
                    if (n != PsReader.DefaultTcpLabel && !_cmbTcp.Items.Contains(n))
                        _cmbTcp.Items.Add(n);
            _cmbTcp.Items.Add(TcpCustomLabel);             // 自定义项始终在末尾
            _cmbTcp.EndUpdate();

            int idx = (prev != null) ? _cmbTcp.Items.IndexOf(prev) : -1;
            _cmbTcp.SelectedIndex = idx >= 0 ? idx : 0;
        }

        private void OnTcpChoiceChanged(object sender, EventArgs e)
        {
            if (_cmbTcp == null) return;
            string sel = _cmbTcp.SelectedItem as string;

            bool custom = (sel == TcpCustomLabel);
            if (custom) EnsureTcpFrameCombo();   // 懒加载：首次选自定义时才创建坐标控件
            if (_tcpCustomPanel != null) _tcpCustomPanel.Visible = custom;

            if (custom)
            {
                _tcpChoiceName = null;
                // _tcpCustomMatrix 由用户在坐标控件中拾取后回填（OnTcpFrameValidSet）
                if (_lblTcpInfo != null) _lblTcpInfo.Text = "TCP：自定义（请在下方拾取坐标系）";
            }
            else
            {
                _tcpCustomMatrix = null;
                if (_tcpFrameCombo != null) try { _tcpFrameCombo.Clear(); } catch { }
                if (sel == PsReader.DefaultTcpLabel || string.IsNullOrEmpty(sel))
                {
                    _tcpChoiceName = null;   // 默认：不覆盖，沿用 GunInfo 默认 TCP
                    if (_lblTcpInfo != null) _lblTcpInfo.Text = "TCP：默认（机器人当前TCP）";
                }
                else
                {
                    _tcpChoiceName = sel;
                    if (_lblTcpInfo != null) _lblTcpInfo.Text = "TCP：" + sel;
                }
            }
        }

        // ════════════════════════════════════════════════════════════
        //  Grid 事件
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
            if (!(_listAdapter is GridAdapter) || _objGrid == null) return;

            var gridObjs = new List<ITxObject>();
            int n = 0;
            try { n = _objGrid.Count; }
            catch (Exception ex) { Log("[警告] 读取 _objGrid.Count 失败：" + ex.Message, LogLevel.Warn); return; }

            for (int i = 0; i < n; i++)
            {
                ITxObject txo = null;
                try { txo = _objGrid.GetObject(i); }
                catch { /* 单行读取失败跳过 */ }
                if (txo != null) gridObjs.Add(txo);
            }

            ThreadPool.QueueUserWorkItem(delegate (object s)
            {
                var newOps = new List<OperationInfo>();
                var seenNames = new HashSet<string>();
                try
                {
                    _svc.InvokeOnPs(delegate ()
                    {
                        foreach (var obj in gridObjs)
                        {
                            List<OperationInfo> parsed = null;
                            try { parsed = PsReader.ParsePickedObjectToOperations(obj); }
                            catch { }
                            if (parsed == null) continue;
                            foreach (var op in parsed)
                                if (seenNames.Add(op.Name)) newOps.Add(op);
                        }
                    });
                }
                catch (Exception ex)
                { UI(delegate () { Log("[错误] 同步操作列表异常：" + ex.Message, LogLevel.Error); }); return; }

                UI(delegate ()
                {
                    bool firstFill = (_ops.Count == 0 && newOps.Count > 0);
                    _ops = newOps;
                    RefreshHintFromOps();
                    ReloadPointsList();
                    UpdatePointCount();
                    Log("[PS] 列表" + (firstFill ? "已同步" : "更新") + "：" + newOps.Count + " 个操作", LogLevel.Ps);
                    if (firstFill) OnOpsFirstFilled();
                    UpdateGunInfo();
                });
            });
        }

        /// <summary>首次填充时的副作用：解析工具/坐标信息。</summary>
        private void OnOpsFirstFilled()
        {
            LogOperationTools();
            if (_refFrameName == "世界坐标系" && _refFrameMatrix == null)
                AutoLoadRefFrameFromOps();
        }

        private void RefreshHintFromOps()
        {
            if (_lblListHint == null) return;
            _lblListHint.Text = _ops.Count > 0 ? "共 " + _ops.Count + " 个操作" : "列表为空";
            _lblListHint.ForeColor = _ops.Count > 0 ? Theme.StatusOkFg : SystemColors.GrayText;
        }

        // ════════════════════════════════════════════════════════════
        //  焊钳信息显示（修复：原代码错误地赋值给 Form.Name；
        //                 改为后台线程解析，UI 线程回填）
        // ════════════════════════════════════════════════════════════
        private void UpdateGunInfo()
        {
            if (_lblGunInfo == null) return;

            if (_ops == null || _ops.Count == 0)
            {
                _lblGunInfo.Text = "焊钳：（待选取）";
                _lblGunInfo.ForeColor = SystemColors.GrayText;
                return;
            }

            // 先显示加载中状态，避免阻塞 UI
            _lblGunInfo.Text = "焊钳：解析中...";
            _lblGunInfo.ForeColor = SystemColors.GrayText;

            var opsSnapshot = new List<OperationInfo>(_ops);

            ThreadPool.QueueUserWorkItem(delegate (object s)
            {
                var names = new List<string>();
                var seen = new HashSet<string>();
                try
                {
                    _svc.InvokeOnPs(delegate ()
                    {
                        foreach (var op in opsSnapshot)
                        {
                            string toolName = null;
                            try { toolName = PsReader.GetToolNameFromOperation(op); }
                            catch { }
                            if (!string.IsNullOrEmpty(toolName) && seen.Add(toolName))
                                names.Add(toolName);
                        }
                    });
                }
                catch (Exception ex)
                {
                    UI(delegate () { Log("[警告] 解析焊钳失败：" + ex.Message, LogLevel.Warn); });
                }

                UI(delegate ()
                {
                    if (_lblGunInfo == null) return;
                    if (names.Count == 0)
                    {
                        _lblGunInfo.Text = "焊钳：（无可用信息）";
                        _lblGunInfo.ForeColor = SystemColors.GrayText;
                    }
                    else
                    {
                        _lblGunInfo.Text = "焊钳：" + string.Join(", ", names);
                        _lblGunInfo.ForeColor = SystemColors.GrayText;
                    }
                });
            });
        }

        // ════════════════════════════════════════════════════════════
        //  控件工厂
        // ════════════════════════════════════════════════════════════
        private static Label MkHeaderLabel(string text, TxColor bg)
        {
            return new FlatColorLabel
            {
                Text = text,
                Dock = DockStyle.Fill,
                BgColor = bg.Color,
                ForeColor = TxColor.TxColorWhite.Color,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold),
                Margin = new Padding(1, 0, 1, 0)
            };
        }

        private static GroupBox MkCard(string title, TxColor titleColor)
        {
            return new ColoredGroupBox
            {
                Text = title,
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold),
                TitleColor = titleColor.Color,
                Margin = new Padding(2, 2, 2, 4),
                Padding = new Padding(8, 6, 8, 4)
            };
        }

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
                Font = SystemFonts.MessageBoxFont,
                Padding = new Padding(0, 2, 0, 0)
            };
        }

        private static Label MkLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Font = SystemFonts.MessageBoxFont,
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(0, 7, 4, 0)
            };
        }

        // FlowLayoutPanel(TopDown) 不会拉伸子控件宽度，也不会让 AutoSize 标签按容器宽换行。
        // 以下两个助手在容器尺寸变化时实时校正，兼容任意 DPI/缩放。

        /// <summary>让子控件宽度始终填满 flow 内容区。</summary>
        private void FillWidthInFlow(FlowLayoutPanel flow, Control child)
        {
            if (flow == null || child == null) return;
            EventHandler h = delegate (object s, EventArgs ev)
            {
                int w = flow.ClientSize.Width - flow.Padding.Horizontal - child.Margin.Horizontal;
                if (w > 0 && child.Width != w) child.Width = w;
            };
            flow.ClientSizeChanged += h;
            flow.SizeChanged += h;
            h(flow, EventArgs.Empty);
        }

        /// <summary>让 AutoSize 标签按 flow 内容区宽度换行（高度自适应，不被裁剪）。</summary>
        private void WrapLabelInFlow(FlowLayoutPanel flow, Label lbl)
        {
            if (flow == null || lbl == null) return;
            lbl.AutoSize = true;
            EventHandler h = delegate (object s, EventArgs ev)
            {
                int w = flow.ClientSize.Width - flow.Padding.Horizontal - lbl.Margin.Horizontal;
                if (w > 20) lbl.MaximumSize = new Size(w, 0);
            };
            flow.ClientSizeChanged += h;
            flow.SizeChanged += h;
            h(flow, EventArgs.Empty);
        }

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

        private static Button MkFuncButton(string text, Color bgColor)
        {
            return new FlatColorButton
            {
                Text = text,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Height = 26,
                Font = SystemFonts.MessageBoxFont,
                BgColor = bgColor,
                BorderColor = bgColor,
                Margin = new Padding(0, 2, 4, 2),
                Padding = new Padding(8, 2, 8, 2)
            };
        }

        /// <summary>用 TextRenderer 测量，避免 CreateGraphics 在窗体未显示时不准。</summary>
        private static void AutoFitComboBoxWidth(ComboBox cb)
        {
            if (cb == null || cb.Items.Count == 0) return;
            int maxW = 0;
            foreach (var item in cb.Items)
            {
                int w = TextRenderer.MeasureText(item.ToString(), cb.Font).Width;
                if (w > maxW) maxW = w;
            }
            cb.Width = maxW + 28;
        }

        private static void AutoFitNumericWidth(NumericUpDown nud)
        {
            if (nud == null) return;
            string maxText = nud.Maximum.ToString("F" + nud.DecimalPlaces);
            int w = TextRenderer.MeasureText(maxText, nud.Font).Width;
            nud.Width = w + 26;
        }

        // ════════════════════════════════════════════════════════════
        //  PS 数据加载 / 列表管理
        // ════════════════════════════════════════════════════════════
        private void OnPickSel(object sender, EventArgs e)
        {
            Log("[PS] 从PS当前选中拾取...", LogLevel.Ps);
            _btnPickFromSel.Enabled = false;
            ThreadPool.QueueUserWorkItem(delegate
            {
                List<OperationInfo> ops = null;
                try { ops = _svc.LoadFromSelection(new Action<string>(delegate (string m) { Log(m); })); }
                catch (Exception ex) { UI(delegate () { Log("[错误] " + ex.Message, LogLevel.Error); }); }
                var final = ops ?? new List<OperationInfo>();

                UI(delegate ()
                {
                    _listAdapter.AddOperations(final);
                    _btnPickFromSel.Enabled = true;
                });
            });
        }

        private void OnClearOps(object sender, EventArgs e)
        {
            _listAdapter.Clear();
            _ops.Clear();
            if (_objEditOp != null) _objEditOp.Object = null;
            if (_lblListHint != null)
            {
                _lblListHint.Text = "列表为空";
                _lblListHint.ForeColor = SystemColors.GrayText;
            }
            ReloadPointsList();
            UpdatePointCount();
            UpdateGunInfo();
        }

        /// <summary>合并新操作到 _ops（按名称去重，O(n) 而非 O(n²)）。</summary>
        private void MergeOpsByName(IEnumerable<OperationInfo> incoming)
        {
            if (incoming == null) return;
            var existing = new HashSet<string>();
            foreach (var op in _ops) existing.Add(op.Name);
            foreach (var op in incoming)
                if (existing.Add(op.Name)) _ops.Add(op);
        }

        /// <summary>仅 ListView 回落模式下生效（Grid 模式由 SyncOpsFromGrid 统一处理）。</summary>
        private void RefreshOpList()
        {
            if (!(_listAdapter is ListViewAdapter)) return;

            if (_lvOps != null)
            {
                _lvOps.Items.Clear();
                _lvOps.Columns[0].Width = -2;
                foreach (OperationInfo op in _ops)
                {
                    var item = new ListViewItem(op.Name) { Checked = true, Tag = op };
                    _lvOps.Items.Add(item);
                }
            }
            if (_lblListHint != null)
            {
                _lblListHint.Text = _ops.Count > 0 ? "共 " + _ops.Count + " 个操作" : "列表为空";
                _lblListHint.ForeColor = _ops.Count > 0 ? Theme.StatusOkFg : SystemColors.GrayText;
            }
            if (_ops.Count > 0)
            {
                Log("[PS] 加载 " + _ops.Count + " 个操作", LogLevel.Ps);
                OnOpsFirstFilled();
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
                    try { _svc.InvokeOnPs(delegate () { toolName = PsReader.GetToolNameFromOperation(op); }); }
                    catch { }
                    string on = op.Name;
                    string tn = toolName;
                    UI(delegate ()
                    { Log("[工具] " + on + " → " + (string.IsNullOrEmpty(tn) ? "未绑定" : tn)); });
                }
            });
        }

        private void AutoLoadRefFrameFromOps()
        {
            var ops = new List<OperationInfo>(_ops);
            var ptType = GetPtType();
            var useMfg = _chkUseMfgName.Checked;

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
                        Log("[坐标] 未找到焊点的零件/外观绑定，保持世界坐标系（可在 PS 中选择 Component/Frame 手动覆盖）", LogLevel.Coord);
                        return;
                    }

                    if (_refFrameName != "世界坐标系" || _refFrameMatrix != null) return;

                    if (rfLocal.NeedsUserChoice)
                    {
                        var picked = AskUserPickRefFrame(rfLocal);
                        if (picked == null)
                        {
                            Log("[坐标] 用户取消，保持世界坐标系", LogLevel.Coord);
                            return;
                        }
                        _refFrameName = picked.Item1;
                        _refFrameMatrix = picked.Item2;
                        UpdateRefCoordStatus(picked.Item1, false);
                        Log("[坐标] 用户选择参考坐标：" + picked.Item1, LogLevel.Coord);
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
                    Log("[坐标] 自动获取参考坐标：" + rfLocal.Name + "（来源=" + srcLabel + pointHint + "）", LogLevel.Coord);
                });
            });
        }

        /// <summary>
        /// 当零件与夹具坐标不一致时弹窗让用户选择。
        /// 返回 (显示名, 4x4矩阵)；取消返回 null。
        /// </summary>
        private Tuple<string, double[]> AskUserPickRefFrame(PsReader.RefFrameResult rf)
        {
            var options = new List<Tuple<string, double[], string>>();
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
                dlg.MinimizeBox = false;
                dlg.MaximizeBox = false;
                dlg.Width = 520;
                dlg.Height = 340;

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

                // 用 FlowLayoutPanel + RightToLeft 让按钮永远靠右
                var pnlBtn = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    Height = 44,
                    FlowDirection = FlowDirection.RightToLeft,
                    Padding = new Padding(10, 8, 10, 6)
                };
                var btnOk = MkFuncButton("确定", Theme.BtnPrimary);
                btnOk.DialogResult = DialogResult.OK;
                var btnCancel = MkFuncButton("取消", Theme.BtnDanger);
                btnCancel.DialogResult = DialogResult.Cancel;
                pnlBtn.Controls.Add(btnCancel);
                pnlBtn.Controls.Add(btnOk);
                dlg.AcceptButton = btnOk;
                dlg.CancelButton = btnCancel;

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
            return _listAdapter.GetCheckedOps();
        }

        /// <summary>导出守卫：列表为空时弹提示并返回 false。</summary>
        private bool EnsureOpsSelected(out List<OperationInfo> ops)
        {
            ops = GetCheckedOps();
            if (ops.Count == 0)
            {
                MessageBox.Show("请先添加操作到列表。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            return true;
        }

        /// <summary>
        /// 导出守卫（点级）：要求已加载点且至少勾选一个点。
        /// 校验并产出：真实 ops（含全部点，下游按白名单过滤）+ 选中点白名单。
        /// </summary>
        private bool EnsureExportSelection(out List<OperationInfo> ops, out HashSet<string> selKeys)
        {
            ops = null; selKeys = null;
            if (_ops == null || _ops.Count == 0)
            {
                MessageBox.Show("请先添加操作到列表。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            if (!_pointsLoaded)
            {
                MessageBox.Show("点列表正在加载，请稍候再导出。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            selKeys = BuildSelectedKeys();
            if (selKeys.Count == 0)
            {
                MessageBox.Show("请至少勾选一个点。", "提示",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            ops = GetSelectedRealOps(selKeys);
            return true;
        }

        /// <summary>从列表控件的实际勾选状态构建选中点白名单（地面真值）。</summary>
        private HashSet<string> BuildSelectedKeys()
        {
            var keys = new HashSet<string>();
            if (_lvPoints == null) return keys;
            foreach (ListViewItem it in _lvPoints.Items)
            {
                var row = it.Tag as PointRow;
                if (it.Checked && row != null)
                    keys.Add(CatiaBridge.PointKey(row.Op.Name, row.Pt.Name));
            }
            return keys;
        }

        /// <summary>返回 _ops 中至少含一个选中点的操作（保留其完整 Points，下游按白名单过滤）。</summary>
        private List<OperationInfo> GetSelectedRealOps(HashSet<string> keys)
        {
            var order = new List<OperationInfo>();
            foreach (var op in _ops)
            {
                bool any = false;
                if (op.Points != null)
                    foreach (var pt in op.Points)
                        if (keys.Contains(CatiaBridge.PointKey(op.Name, pt.Name))) { any = true; break; }
                if (any) order.Add(op);
            }
            return order;
        }

        /// <summary>
        /// 按操作分组的副本，Points 仅含勾选点。供没有白名单过滤支持的通道（Excel）使用。
        /// </summary>
        private List<OperationInfo> GetCheckedOpsClones()
        {
            var map = new Dictionary<OperationInfo, OperationInfo>();
            var order = new List<OperationInfo>();
            if (_lvPoints == null) return order;
            foreach (ListViewItem it in _lvPoints.Items)
            {
                var row = it.Tag as PointRow;
                if (!it.Checked || row == null) continue;
                OperationInfo clone;
                if (!map.TryGetValue(row.Op, out clone))
                {
                    clone = new OperationInfo
                    {
                        Name = row.Op.Name,
                        TypeLabel = row.Op.TypeLabel,
                        PsObject = row.Op.PsObject,
                        Gun = row.Op.Gun,
                        Points = new List<PointInfo>()
                    };
                    map[row.Op] = clone;
                    order.Add(clone);
                }
                clone.Points.Add(row.Pt);
            }
            return order;
        }

        private void UpdatePointCount()
        {
            try
            {
                if (_lblPointCount == null) return;

                if (_ops.Count == 0)
                {
                    _lblPointCount.Text = "将导出点数量：0（列表为空）";
                    _lblPointCount.ForeColor = Theme.BtnDanger;
                    UpdatePointsSelHeader();
                    return;
                }
                if (!_pointsLoaded)
                {
                    _lblPointCount.Text = "将导出点数量：统计中…";
                    _lblPointCount.ForeColor = SystemColors.GrayText;
                    return;
                }

                int n = 0;
                PointType filter = GetPtType();
                foreach (var row in _pointRows)
                {
                    if (!row.Checked) continue;
                    if (filter != PointType.All && row.Pt.Type != filter) continue;
                    n++;
                }
                _lblPointCount.Text = "将导出点数量：" + n;
                _lblPointCount.ForeColor = n > 0 ? Theme.StatusOkFg : Theme.BtnDanger;
                UpdatePointsSelHeader();
            }
            catch { }
        }

        private PointType GetPtType()
        {
            switch (_cmbPointType.SelectedIndex)
            {
                case 0: return PointType.WeldPoint;
                case 1: return PointType.PathPoint;
                case 2: return PointType.ContinuousPoint;
                default: return PointType.All;
            }
        }

        // ════════════════════════════════════════════════════════════
        //  事件处理
        // ════════════════════════════════════════════════════════════
        private void OnBrowseGun(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "选择焊枪数模";
                dlg.Filter = "CATIA文件|*.CATProduct;*.CATPart;*.cgr|所有文件|*.*";
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    _txtGunModel.Tag = dlg.FileName;
                    _txtGunModel.Text = Path.GetFileName(dlg.FileName);
                }
            }
        }

        private void OnExportExcel(object sender, EventArgs e)
        {
            HashSet<string> selKeys;
            if (!EnsureExportSelection(out _, out selKeys)) return;
            // Excel 通道无白名单过滤支持，改传仅含勾选点的副本
            var ops = GetCheckedOpsClones();

            SetBusy(true);
            Log("[Excel] 开始导出，参考坐标：" + _refFrameName, LogLevel.Excel);
            _svc.ExportExcelAsync(ops, GetPtType(), _chkUseMfgName.Checked,
                DefaultOut(), _refFrameMatrix, _refFrameName,
                new Action<string>(delegate (string m) { UI(delegate () { Log(m); }); }),
                new Action<bool, string>(delegate (bool ok, string result) {
                    UI(delegate ()
                    {
                        SetBusy(false);
                        if (ok)
                        {
                            Log("✓ Excel：" + result, LogLevel.Ok);
                            if (MessageBox.Show("成功：\n" + result + "\n\n打开？", "完成",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                                try { System.Diagnostics.Process.Start(result); } catch { }
                        }
                        else
                        {
                            Log("✗ 失败：" + result, LogLevel.Error);
                            MessageBox.Show("失败：\n" + result, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    });
                }));
        }

        // ════════════════════════════════════════════════════════════
        //  导出插枪 — 拆分为 解析CGR + 构造参数 + 启动异步
        // ════════════════════════════════════════════════════════════
        private void OnExportGun(object sender, EventArgs e)
        {
            HashSet<string> selKeys;
            if (!EnsureExportSelection(out _, out selKeys)) return;

            // 用「仅含勾选点」的副本导出：下游按 op.Points 计数即为勾选数；
            // 同时仍带 SelectedKeys 作为几何生成处的白名单兜底。
            var ops = GetCheckedOpsClones();
            _lastSelCount = selKeys.Count;

            // 第一步：为每个 op 解析 CGR 路径（用户可能取消）
            var perOpPaths = ResolveCgrPathsForOps(ops);
            if (perOpPaths == null) return;

            // 第二步：构造参数 + 启动异步导出
            var p = BuildGunExportParams(ops, perOpPaths);
            p.SelectedKeys = selKeys;
            StartGunExport(p);
        }

        /// <summary>
        /// 为每个 op 解析 CGR 文件路径。
        /// 返回 null 表示用户取消或失败，已弹窗/打日志通知。
        /// </summary>
        private Dictionary<string, string> ResolveCgrPathsForOps(List<OperationInfo> ops)
        {
            const double HIGH_CONFIDENCE = 0.75;

            string globalCustomPath = _chkCustomGun.Checked
                ? (_txtGunModel.Tag as string ?? _txtGunModel.Text)
                : null;
            bool hasGlobalCustom = !string.IsNullOrEmpty(globalCustomPath) && File.Exists(globalCustomPath);

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
                        lookup = PsReader.LookupCgrForOperation(op,
                            delegate (string m) { UI(delegate () { Log(m); }); });
                    });
                }
                catch (Exception ex) { Log("[错误] CGR 查找异常：" + ex.Message, LogLevel.Error); }

                if (lookup == null || lookup.Candidates == null || lookup.Candidates.Count == 0)
                {
                    string reason = lookup != null ? lookup.Reason : "未知原因";
                    string dir = lookup != null ? lookup.SameDir : null;
                    string toolName = (lookup != null && lookup.ToolName != null) ? lookup.ToolName : op.Name;

                    // 弹出可交互对话框：允许用户直接打开 JT 所在位置（同级目录），
                    // 或手动浏览选择一个 CGR/CATPart/CATProduct，或中止。
                    string manualPath = PromptCgrNotFound(op.Name, toolName, dir, reason);
                    if (string.IsNullOrEmpty(manualPath))
                    {
                        Log("[CGR] 已中止：" + reason, LogLevel.Warn);
                        return null;
                    }
                    perOpPaths[op.Name] = manualPath;
                    Log("[CGR] " + op.Name + " → " + Path.GetFileName(manualPath) + "（手动浏览）");
                    continue;
                }

                var best = lookup.Candidates[0];
                if (best.Similarity >= HIGH_CONFIDENCE)
                {
                    perOpPaths[op.Name] = best.Path;
                    bool exact = best.Similarity >= 0.999;
                    Log("[CGR] " + op.Name + " → " + Path.GetFileName(best.Path)
                        + (exact ? "（精确命中）"
                                 : ("（相似度 " + (best.Similarity * 100).ToString("F0") + "%）")));
                }
                else
                {
                    string chosen = PromptCgrSelection(op.Name, lookup.ToolName, lookup.SameDir, lookup.Candidates);
                    if (string.IsNullOrEmpty(chosen))
                    {
                        Log("[CGR] 用户取消，已中止导出", LogLevel.Warn);
                        return null;
                    }
                    perOpPaths[op.Name] = chosen;
                    Log("[CGR] " + op.Name + " → " + Path.GetFileName(chosen) + "（用户确认）");
                }
            }

            return perOpPaths;
        }

        private GunExportParams BuildGunExportParams(List<OperationInfo> ops, Dictionary<string, string> perOpPaths)
        {
            string globalCustomPath = _chkCustomGun.Checked
                ? (_txtGunModel.Tag as string ?? _txtGunModel.Text)
                : null;
            bool hasGlobalCustom = !string.IsNullOrEmpty(globalCustomPath) && File.Exists(globalCustomPath);

            return new GunExportParams
            {
                Operations = ops,
                ExportTCP = _chkExportTCP.Checked,
                GunOriginAtTCP = _chkGunOriginTCP.Checked,
                CustomModelPath = hasGlobalCustom ? globalCustomPath : null,
                PerOpModelPaths = perOpPaths,
                CustomProductName = string.IsNullOrWhiteSpace(_txtGunProductName.Text) ? null : _txtGunProductName.Text.Trim(),
                Format = ExportFormat.Xml3d,
                OutputPath = DefaultOut(),
                RefMatrix = _refFrameMatrix,
                RefName = _refFrameName,
                PointFilter = GetPtType(),
                UseMfgName = _chkUseMfgName.Checked,
                TcpName = _tcpCustomMatrix == null ? _tcpChoiceName : null,
                TcpCustomMatrix = _tcpCustomMatrix,
                ExportMode = (_cmbGunExportMode != null && _cmbGunExportMode.SelectedIndex == 1)
                    ? GunExportMode.IndependentNaming
                    : GunExportMode.SharedGeometry
            };
        }

        private void StartGunExport(GunExportParams p)
        {
            SetBusy(true);
            _svc.ExportGunsAsync(p,
                new Action<string>(delegate (string m) { UI(delegate () { Log(m); }); }),
                new Action<ExportProgress>(delegate (ExportProgress pg) { UI(delegate () { SetProgress(pg); }); }),
                new Action<bool, string>(delegate (bool ok, string msg) {
                    UI(delegate ()
                    {
                        SetBusy(false);
                        Log((ok ? "✓ " : "✗ ") + msg, ok ? LogLevel.Ok : LogLevel.Error);
                        ShowResult(ok, ok ? ("本次按勾选共导出 " + _lastSelCount + " 个点。\n\n" + msg) : msg);
                    });
                }));
        }

        // ════════════════════════════════════════════════════════════
        //  CGR 候选选择对话框
        //  - 双击列表项 = 使用选中
        //  - 鼠标悬停显示完整路径
        //  - 状态条显示当前选中的完整路径，避免文件名截断歧义
        //  - "打开 JT 所在位置" 按钮：用资源管理器打开同级目录
        // ════════════════════════════════════════════════════════════
        private string PromptCgrSelection(string opName, string toolName, string sameDir,
            List<PsReader.CgrMatch> candidates)
        {
            using (var dlg = new Form())
            {
                dlg.Text = "请确认 CGR 文件 — " + opName;
                dlg.Size = new Size(620, 440);
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.FormBorderStyle = FormBorderStyle.Sizable;
                dlg.MinimumSize = new Size(520, 360);
                dlg.MaximizeBox = false;
                dlg.MinimizeBox = false;
                dlg.BackColor = SystemColors.Control;

                var header = new Label
                {
                    Text = "工具名称：" + (toolName ?? "(未知)")
                        + "\n同级目录：" + (sameDir ?? "(未解析)")
                        + "\n未找到高相似度的 CGR 文件，请从以下候选中选择，或手动浏览：",
                    Dock = DockStyle.Top,
                    Height = 64,
                    Padding = new Padding(10, 8, 10, 0),
                    Font = SystemFonts.MessageBoxFont,
                    ForeColor = SystemColors.ControlText
                };

                // 状态条（底部上方）：显示当前选中候选的完整路径
                var lblPath = new Label
                {
                    Dock = DockStyle.Bottom,
                    Height = 22,
                    Padding = new Padding(10, 4, 10, 0),
                    Font = SystemFonts.MessageBoxFont,
                    ForeColor = SystemColors.GrayText,
                    AutoEllipsis = true,
                    Text = ""
                };

                var lb = new ListBox
                {
                    Dock = DockStyle.Fill,
                    IntegralHeight = false,
                    Font = SystemFonts.MessageBoxFont,
                    BorderStyle = BorderStyle.FixedSingle
                };
                foreach (var c in candidates)
                {
                    string sim = (c.Similarity * 100).ToString("F0");
                    lb.Items.Add(c.CandidateName + ".cgr   (相似度 " + sim + "%)");
                }
                if (lb.Items.Count > 0) lb.SelectedIndex = 0;

                // 候选项 tooltip：完整路径
                var tip = new ToolTip { AutoPopDelay = 8000, InitialDelay = 300, ReshowDelay = 200, ShowAlways = true };
                int lastHoverIdx = -1;
                lb.MouseMove += delegate (object s, MouseEventArgs e)
                {
                    int idx = lb.IndexFromPoint(e.Location);
                    if (idx == lastHoverIdx) return;
                    lastHoverIdx = idx;
                    if (idx >= 0 && idx < candidates.Count)
                        tip.SetToolTip(lb, candidates[idx].Path ?? "");
                    else
                        tip.SetToolTip(lb, "");
                };

                // 选中变化 → 同步状态条
                lb.SelectedIndexChanged += delegate
                {
                    int idx = lb.SelectedIndex;
                    if (idx >= 0 && idx < candidates.Count)
                        lblPath.Text = "路径：" + (candidates[idx].Path ?? "");
                    else
                        lblPath.Text = "";
                };
                if (lb.SelectedIndex >= 0 && lb.SelectedIndex < candidates.Count)
                    lblPath.Text = "路径：" + (candidates[lb.SelectedIndex].Path ?? "");

                var pnlBtns = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    Height = 44,
                    FlowDirection = FlowDirection.RightToLeft,
                    Padding = new Padding(10, 6, 10, 6),
                    BackColor = SystemColors.Control
                };

                bool browseFallback = false;
                var btnOk = MkFuncButton("使用选中", Theme.BtnPrimary);
                var btnBrowse = MkFuncButton("手动浏览...", Theme.BtnSecondary);
                var btnOpenDir = MkFuncButton("打开 JT 所在位置", Theme.BtnMuted);
                var btnCancel = MkFuncButton("取消", Theme.BtnDanger);

                btnOk.DialogResult = DialogResult.OK;
                btnCancel.DialogResult = DialogResult.Cancel;

                btnBrowse.Click += delegate
                {
                    browseFallback = true;
                    dlg.DialogResult = DialogResult.OK;
                    dlg.Close();
                };
                btnOpenDir.Enabled = !string.IsNullOrEmpty(sameDir) && Directory.Exists(sameDir);
                btnOpenDir.Click += delegate { OpenJtLocation(sameDir); };

                // 双击 = 使用选中
                lb.DoubleClick += delegate
                {
                    if (lb.SelectedIndex >= 0)
                    {
                        dlg.DialogResult = DialogResult.OK;
                        dlg.Close();
                    }
                };

                pnlBtns.Controls.Add(btnCancel);
                pnlBtns.Controls.Add(btnBrowse);
                pnlBtns.Controls.Add(btnOpenDir);
                pnlBtns.Controls.Add(btnOk);

                // Dock 顺序：先加 Fill，再底部，再顶部
                dlg.Controls.Add(lb);
                dlg.Controls.Add(pnlBtns);
                dlg.Controls.Add(lblPath);
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

                int sel = lb.SelectedIndex;
                if (sel < 0 || sel >= candidates.Count) return null;
                return candidates[sel].Path;
            }
        }

        // ════════════════════════════════════════════════════════════
        //  CGR 未找到对话框 — 提供"打开 JT 所在位置 / 手动浏览 / 取消"
        //  返回：用户手动选定的 CGR 路径；取消返回 null。
        // ════════════════════════════════════════════════════════════
        private string PromptCgrNotFound(string opName, string toolName, string sameDir, string reason)
        {
            using (var dlg = new Form())
            {
                dlg.Text = "未找到 CGR — " + opName;
                dlg.Size = new Size(560, 260);
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                dlg.MaximizeBox = false;
                dlg.MinimizeBox = false;
                dlg.BackColor = SystemColors.Control;

                var body = new Label
                {
                    Text = "操作：" + opName
                        + "\n工具名称：" + (toolName ?? "(未知)")
                        + "\n同级目录：" + (string.IsNullOrEmpty(sameDir) ? "(未解析)" : sameDir)
                        + "\n原因：" + (reason ?? "未知原因")
                        + "\n\n你可以：\n"
                        + "  · 打开 JT 所在位置，确认目录下是否存在 CGR 文件；\n"
                        + "  · 手动浏览选择一个 CGR / CATPart / CATProduct；\n"
                        + "  · 取消本次导出。",
                    Dock = DockStyle.Fill,
                    Padding = new Padding(14, 12, 14, 0),
                    Font = SystemFonts.MessageBoxFont,
                    ForeColor = SystemColors.ControlText
                };

                var pnlBtns = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    Height = 44,
                    FlowDirection = FlowDirection.RightToLeft,
                    Padding = new Padding(10, 6, 10, 6),
                    BackColor = SystemColors.Control
                };

                bool browse = false;
                var btnBrowse = MkFuncButton("手动浏览...", Theme.BtnPrimary);
                var btnOpenDir = MkFuncButton("打开 JT 所在位置", Theme.BtnSecondary);
                var btnCancel = MkFuncButton("取消", Theme.BtnDanger);
                btnCancel.DialogResult = DialogResult.Cancel;

                btnBrowse.Click += delegate
                {
                    browse = true;
                    dlg.DialogResult = DialogResult.OK;
                    dlg.Close();
                };
                btnOpenDir.Enabled = !string.IsNullOrEmpty(sameDir) && Directory.Exists(sameDir);
                btnOpenDir.Click += delegate { OpenJtLocation(sameDir); };

                pnlBtns.Controls.Add(btnCancel);
                pnlBtns.Controls.Add(btnOpenDir);
                pnlBtns.Controls.Add(btnBrowse);

                dlg.Controls.Add(body);
                dlg.Controls.Add(pnlBtns);
                dlg.AcceptButton = btnBrowse;
                dlg.CancelButton = btnCancel;

                if (dlg.ShowDialog(this) != DialogResult.OK) return null;
                if (!browse) return null;

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
        }

        // ════════════════════════════════════════════════════════════
        //  打开 JT 所在位置 — 优先用资源管理器选中目录内的 .jt/.cojt
        //  失败时回落到打开目录本身。
        // ════════════════════════════════════════════════════════════
        private void OpenJtLocation(string dir)
        {
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
            {
                MessageBox.Show("目录不可用或不存在：\n" + (dir ?? "(空)"),
                    "无法打开", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 在同级目录内寻找最有代表性的 JT 资源文件用于"选中"：
            //   1) *.jt   — 最常见的 PS 部件几何文件
            //   2) *.cojt — PS 复合 JT 资源（实际是目录或文件）
            string selectTarget = null;
            try
            {
                var jts = Directory.GetFiles(dir, "*.jt", SearchOption.TopDirectoryOnly);
                if (jts != null && jts.Length > 0) selectTarget = jts[0];
                if (selectTarget == null)
                {
                    var cojts = Directory.GetFileSystemEntries(dir, "*.cojt", SearchOption.TopDirectoryOnly);
                    if (cojts != null && cojts.Length > 0) selectTarget = cojts[0];
                }
            }
            catch { /* 忽略扫描异常，直接退化到打开目录 */ }

            try
            {
                if (selectTarget != null)
                {
                    // /select 让 explorer 高亮指定文件；路径必须用双引号包裹防止空格断裂
                    Process.Start("explorer.exe", "/select,\"" + selectTarget + "\"");
                    Log("[系统] 已打开 JT 位置（选中 " + Path.GetFileName(selectTarget) + "）", LogLevel.Info);
                }
                else
                {
                    Process.Start("explorer.exe", "\"" + dir + "\"");
                    Log("[系统] 已打开 JT 所在目录（未发现 .jt 文件）", LogLevel.Info);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("打开资源管理器失败：" + ex.Message,
                    "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log("[错误] 打开 JT 位置失败：" + ex.Message, LogLevel.Error);
            }
        }

        private void OnExportBall(object sender, EventArgs e)
        {
            HashSet<string> selKeys;
            if (!EnsureExportSelection(out _, out selKeys)) return;

            SetBusy(true);
            _lastSelCount = selKeys.Count;
            var ops = GetCheckedOpsClones();   // 仅含勾选点，使下游计数=勾选数
            var optMap = new[] { BallExportOption.TrajectoryAndBall, BallExportOption.TrajectoryOnly, BallExportOption.BallOnly };
            var p = new BallExportParams
            {
                Operations = ops,
                ExportToCurrentDoc = _cmbBallTarget.SelectedIndex == 0,
                Option = optMap[_cmbBallOption.SelectedIndex],
                BallDiameter = (double)_nudDiameter.Value,
                OutputPath = DefaultOut(),
                PointFilter = GetPtType(),
                UseMfgName = _chkUseMfgName.Checked,
                GeomSetName = string.IsNullOrWhiteSpace(_txtGeomSet.Text) ? "Geometry_Spheres" : _txtGeomSet.Text.Trim(),
                NamePrefix = string.IsNullOrWhiteSpace(_txtNamePrefix.Text) ? "SPHERE" : _txtNamePrefix.Text.Trim(),
                CustomPartName = string.IsNullOrWhiteSpace(_txtBallPartName.Text) ? null : _txtBallPartName.Text.Trim(),
                RefMatrix = _refFrameMatrix,
                RefName = _refFrameName,
                SelectedKeys = selKeys
            };
            _svc.ExportBallsAsync(p,
                new Action<string>(delegate (string m) { UI(delegate () { Log(m); }); }),
                new Action<ExportProgress>(delegate (ExportProgress pg) { UI(delegate () { SetProgress(pg); }); }),
                new Action<bool, string>(delegate (bool ok, string msg) {
                    UI(delegate ()
                    {
                        SetBusy(false);
                        Log((ok ? "✓ " : "✗ ") + msg, ok ? LogLevel.Ok : LogLevel.Error);
                        ShowResult(ok, ok ? ("本次按勾选共导出 " + _lastSelCount + " 个点。\n\n" + msg) : msg);
                    });
                }));
        }

        private void OnReset(object sender, EventArgs e)
        {
            _chkExportTCP.Checked = false;
            _chkGunOriginTCP.Checked = true;
            _chkCustomGun.Checked = false;
            _txtGunModel.Text = "";
            _txtGunModel.Tag = null;
            if (_txtGunProductName != null) _txtGunProductName.Text = "";
            if (_cmbGunExportMode != null) try { _cmbGunExportMode.SelectedIndex = 0; } catch { }
            // 复位 TCP 选择
            _tcpChoiceName = null;
            _tcpCustomMatrix = null;
            if (_tcpFrameCombo != null) try { _tcpFrameCombo.Clear(); } catch { }
            if (_cmbTcp != null && _cmbTcp.Items.Count > 0) _cmbTcp.SelectedIndex = 0;
            if (_tcpCustomPanel != null) _tcpCustomPanel.Visible = false;
            if (_lblTcpInfo != null) _lblTcpInfo.Text = "TCP：默认（机器人当前TCP）";
            _cmbPointType.SelectedIndex = 3;
            _cmbBallTarget.SelectedIndex = 0;
            _cmbBallOption.SelectedIndex = 0;
            _nudDiameter.Value = 10;
            _txtGeomSet.Text = "Geometry_Spheres";
            _txtNamePrefix.Text = "SPHERE";
            _chkUseMfgName.Checked = false;
            if (_txtBallPartName != null) _txtBallPartName.Text = "";
            _refFrameName = "世界坐标系";
            _refFrameMatrix = null;
            UpdateRefCoordStatus("世界坐标系", true);
            if (_frameCombo != null) try { _frameCombo.Clear(); } catch { }
            if (_objEditOp != null) _objEditOp.Object = null;
            _listAdapter.Clear();
            _ops.Clear();
            if (_lblListHint != null)
            {
                _lblListHint.Text = "列表为空";
                _lblListHint.ForeColor = SystemColors.GrayText;
            }
            ReloadPointsList();
            // 进度条用 SetBusy(false) 统一处理，文案与"导出结束"一致
            SetBusy(false);
            Log("[操作] 已复位所有设置");
        }

        // ════════════════════════════════════════════════════════════
        //  UI 辅助
        // ════════════════════════════════════════════════════════════
        private void SetBusy(bool busy)
        {
            _btnExportGun.Enabled = !busy;
            _btnExportBall.Enabled = !busy;
            _btnPickFromSel.Enabled = !busy;
            _btnExportExcel.Enabled = !busy;
            _progressBar.Visible = busy;
            if (busy)
            {
                _progressBar.Value = 0;
                _lblProgress.Text = "运行中...";
            }
            else
            {
                _progressBar.Value = 0;
                _lblProgress.Text = "就绪";
            }
        }

        private void SetProgress(ExportProgress p)
        {
            if (p.Total <= 0) return;
            _progressBar.Value = Math.Min((int)((double)p.Current / p.Total * 100), 100);
            _lblProgress.Text = p.Current + "/" + p.Total + "  " + p.CurrentItem;
        }

        // ── 日志 ─────────────────────────────────────────────────────
        // 旧 Log(string) 保留，用关键字猜测级别（向后兼容外部回调）；
        // 新代码尽量使用 Log(msg, level) 显式指定。
        private void Log(string msg)
        {
            Log(msg, GuessLevel(msg));
        }

        private void Log(string msg, LogLevel level)
        {
            if (_rtbLog == null || IsDisposed) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string, LogLevel>(Log), msg, level);
                return;
            }

            Color c;
            switch (level)
            {
                case LogLevel.Ok: c = Theme.LogOk; break;
                case LogLevel.Error: c = Theme.LogErr; break;
                case LogLevel.Warn: c = Theme.LogWarn; break;
                case LogLevel.Ps: c = Theme.LogPs; break;
                case LogLevel.Coord: c = Theme.LogCoord; break;
                case LogLevel.Excel: c = Theme.LogExcel; break;
                case LogLevel.Debug: c = SystemColors.GrayText; break;
                default: c = Theme.LogText; break;
            }

            _rtbLog.SelectionStart = _rtbLog.TextLength;
            _rtbLog.SelectionLength = 0;
            _rtbLog.SelectionColor = c;
            _rtbLog.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + msg + Environment.NewLine);

            // 折叠状态下出现错误时自动展开，避免用户漏看
            if (level == LogLevel.Error && !_logExpanded) ToggleLog();
            if (_rtbLog.Visible) try { _rtbLog.ScrollToCaret(); } catch { }
        }

        private static LogLevel GuessLevel(string msg)
        {
            if (string.IsNullOrEmpty(msg)) return LogLevel.Info;
            if (msg.StartsWith("✓") || msg.Contains("完成")) return LogLevel.Ok;
            if (msg.StartsWith("✗") || msg.Contains("失败") || msg.Contains("错误")) return LogLevel.Error;
            if (msg.Contains("⚠") || msg.Contains("[警告]")) return LogLevel.Warn;
            if (msg.StartsWith("[PS]")) return LogLevel.Ps;
            if (msg.StartsWith("[坐标]")) return LogLevel.Coord;
            if (msg.StartsWith("[Excel]")) return LogLevel.Excel;
            return LogLevel.Info;
        }

        private void UI(Action act)
        {
            if (IsDisposed) return;
            if (InvokeRequired) BeginInvoke(act); else act();
        }

        private void ShowResult(bool ok, string msg)
        {
            MessageBox.Show(msg, ok ? "完成" : "失败",
                MessageBoxButtons.OK,
                ok ? MessageBoxIcon.Information : MessageBoxIcon.Error);
        }

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
        //  窗体关闭 — 解绑事件 + 释放 PS 控件持有的对象引用
        // ════════════════════════════════════════════════════════════
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_frameCombo != null)
            {
                try { _frameCombo.ValidFrameSet -= new TxFrameComboBoxCtrl_ValidFrameSetEventHandler(OnFrameValidSet); } catch { }
                try { _frameCombo.InvalidFrameSet -= new TxFrameComboBoxCtrl_InvalidFrameSetEventHandler(OnFrameInvalidSet); } catch { }
                try { _frameCombo.Clear(); } catch { }   // 释放对 PS 对象的引用
            }
            if (_tcpFrameCombo != null)
            {
                try { _tcpFrameCombo.ValidFrameSet -= new TxFrameComboBoxCtrl_ValidFrameSetEventHandler(OnTcpFrameValidSet); } catch { }
                try { _tcpFrameCombo.InvalidFrameSet -= new TxFrameComboBoxCtrl_InvalidFrameSetEventHandler(OnTcpFrameInvalidSet); } catch { }
                try { _tcpFrameCombo.Clear(); } catch { }
            }
            if (_objEditOp != null)
            {
                try { _objEditOp.Picked -= new TxObjEditBoxCtrl_PickedEventHandler(OnObjEditPicked); } catch { }
                try { _objEditOp.Object = null; } catch { }
            }
            if (_objGrid != null)
            {
                try { _objGrid.ObjectInserted -= new TxObjGridCtrl_ObjectInsertedEventHandler(OnGridObjectInserted); } catch { }
                try { _objGrid.RowDeleted -= new TxObjGridCtrl_RowDeletedEventHandler(OnGridRowDeleted); } catch { }
                try { _objGrid.Objects = new TxObjectList(); } catch { }   // 释放对 PS 对象的引用
            }
            if (_lvPoints != null)
            {
                try { _lvPoints.ItemChecked -= new ItemCheckedEventHandler(OnPointItemChecked); } catch { }
            }
            if (_svc != null) _svc.Dispose();
            base.OnFormClosing(e);
        }

        private static string DefaultOut()
        { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "CatiaExport"); }

        private void InitializeComponent()
        {
            SuspendLayout();
            Name = "ExportGunForm";
            ResumeLayout(false);
        }

        // ════════════════════════════════════════════════════════════
        //  操作点列表行模型
        // ════════════════════════════════════════════════════════════
        private sealed class PointRow
        {
            public OperationInfo Op;       // 所属操作
            public PointInfo Pt;           // 点数据
            public bool Checked = true;    // 勾选状态（默认选中）
            public ListViewItem Item;      // 当前显示对应的列表项（被筛掉时为 null）
        }

        // ════════════════════════════════════════════════════════════
        //  IOpListAdapter — 抽象列表适配，消除 _useGrid 分支
        // ════════════════════════════════════════════════════════════
        private interface IOpListAdapter
        {
            /// <summary>添加一批 op（Grid: 推入控件触发同步；ListView: 直接刷新）。</summary>
            void AddOperations(List<OperationInfo> ops);
            /// <summary>清空控件内容。</summary>
            void Clear();
            /// <summary>获取被选中的 op 列表（Grid: 全部；ListView: 勾选项）。</summary>
            List<OperationInfo> GetCheckedOps();
        }

        /// <summary>Grid 模式：操作通过 _objGrid 推入，_ops 由 SyncOpsFromGrid 自动同步。</summary>
        private sealed class GridAdapter : IOpListAdapter
        {
            private readonly ExportGunForm _f;
            public GridAdapter(ExportGunForm f) { _f = f; }

            public void AddOperations(List<OperationInfo> ops)
            {
                if (_f._objGrid == null) return;
                int added = 0, skipped = 0;
                foreach (var op in ops)
                {
                    if (op.PsObject == null) { skipped++; continue; }
                    try { _f._objGrid.AppendObject(op.PsObject); added++; }
                    catch (Exception ex)
                    { _f.Log("[警告] AppendObject 失败：" + ex.Message, LogLevel.Warn); skipped++; }
                }
                _f.Log("[PS] 已推入 " + added + " 个对象到列表"
                    + (skipped > 0 ? "（跳过 " + skipped + "）" : ""), LogLevel.Ps);
                if (added > 0) _f.SyncOpsFromGrid();
            }

            public void Clear()
            {
                if (_f._objGrid == null) return;
                try { _f._objGrid.Objects = new TxObjectList(); }
                catch (Exception ex) { _f.Log("[警告] 清空 grid 失败：" + ex.Message, LogLevel.Warn); }
            }

            public List<OperationInfo> GetCheckedOps()
            { return new List<OperationInfo>(_f._ops); }
        }

        /// <summary>ListView 回落模式：勾选项即选中项。</summary>
        private sealed class ListViewAdapter : IOpListAdapter
        {
            private readonly ExportGunForm _f;
            public ListViewAdapter(ExportGunForm f) { _f = f; }

            public void AddOperations(List<OperationInfo> ops)
            {
                _f._ops = ops ?? new List<OperationInfo>();
                _f.RefreshOpList();
                _f.ReloadPointsList();
                _f.UpdatePointCount();
            }

            public void Clear()
            {
                if (_f._lvOps != null) _f._lvOps.Items.Clear();
            }

            public List<OperationInfo> GetCheckedOps()
            {
                var list = new List<OperationInfo>();
                if (_f._lvOps == null) return list;
                foreach (ListViewItem item in _f._lvOps.Items)
                    if (item.Checked && item.Tag is OperationInfo)
                        list.Add((OperationInfo)item.Tag);
                return list;
            }
        }

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
            /// <summary>边框色（本实现为无边框设计，保留属性以兼容统一 API）。</summary>
            public Color BorderColor { get; set; } = Color.Empty;
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
                        (BgColor.B + 255) / 2);
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

                using (var bgBrush = new SolidBrush(BackColor))
                    g.FillRectangle(bgBrush, ClientRectangle);

                var borderRect = new Rectangle(0, halfH, Width - 1, Height - halfH - 1);
                using (var borderPen = new Pen(BorderColor))
                    g.DrawRectangle(borderPen, borderRect);

                if (!string.IsNullOrEmpty(title))
                {
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