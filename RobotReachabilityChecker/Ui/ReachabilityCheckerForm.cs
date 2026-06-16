// ============================================================================
// ReachabilityCheckerForm.cs   —   分部类（主入口 + 状态 + 生命周期 + ILogger 实现）
//
// 这个表单原本一个文件 ~1900 行，按职责拆分为几个 partial 文件：
//   ReachabilityCheckerForm.cs        ← 本文件 — 字段、构造、Load/Closing、Log 实现
//   ReachabilityCheckerForm.Layout.cs ← UI 构建（BuildToolStrip / BuildCardsPanel / ...）
//   ReachabilityCheckerForm.Grid.cs   ← TxFlexGrid 构建、刷新、过滤
//   ReachabilityCheckerForm.Events.cs ← 事件处理（按钮、表格点击、双击、轮询）
//   ReachabilityCheckerForm.Helpers.cs ← UI 控件工厂（MkLabel / MkButton / AutoFit*）
// ============================================================================
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;
using Tecnomatix.Engineering.Ui.WPF;
using TxTools.Common;
using TxTools.RobotReachabilityChecker.Diagnostics;
using TxTools.RobotReachabilityChecker.Models;
using static TxTools.RobotReachabilityChecker.Ui.Theme;

namespace TxTools.RobotReachabilityChecker.Ui
{
    public partial class ReachabilityCheckerForm : TxForm, ILogger
    {
        // ====== 顶部工具栏（默认隐藏，仅检查时显示进度条）======
        private TxToolStrip _toolStrip;
        private ToolStripProgressBar _tsProgress;

        // ====== 页脚日志按钮（哑元，用于让 ToggleLogPanel 改文字）======
        private Button _btnFooterLog;

        // ====== 5 张并列卡片 ======
        private Panel _cardsPanel;

        // — Card1: 检查目标及筛选
        private TxObjEditBoxCtrl _txtOpNode;
        private CheckBox _chkHideNormal;
        private ComboBox _cbPointTypeFilter;
        private ComboBox _cbBrand;

        // — Card2: TCP余量
        private CheckBox _chkTcpXyz;
        private NumericUpDown _nudTcpMargin;

        // — Card3: 轴角度余量
        private CheckBox _chkJointMargin;
        private NumericUpDown _nudJointMarginDeg;

        // — Card4: 干涉检查（仅展示）
        private CheckBox _chkStaticInterference;
        private CheckBox _chkDynamicInterference;

        // ====== 结果表格（TxFlexGrid）======
        private TxFlexGrid _grid;

        // 列索引常量（在 Grid 文件中也使用）
        internal const int COL_IDX = 0, COL_BRAND = 1, COL_ROBOT = 2, COL_OP = 3,
                           COL_PT = 4, COL_TYPE = 5, COL_J1 = 6, COL_J2 = 7,
                           COL_J3 = 8, COL_J4 = 9, COL_J5 = 10, COL_J6 = 11,
                           COL_RESULT = 12, COL_NOTE = 13;

        // ====== 日志面板 ======
        private Panel _logPanel;
        private RichTextBox _logBox;
        private bool _logVisible = false;

        /// <summary>
        /// 详细日志开关：true 时打印每个点位的关节值提取细节、限位获取细节等
        /// 排查问题时设为 true，正常使用时保持 false 以减少日志噪音
        /// </summary>
        private bool _logVerbose = false;

        // ====== 底部 StatusStrip ======
        private StatusStrip _statusStrip;
        private ToolStripStatusLabel _lblStatus;

        // ====== 数据 ======
        private readonly List<RobotPathCheckTask> _tasks = new List<RobotPathCheckTask>();
        private RobotPathCheckTask _currentTask;

        // 用户通过 OP 节点拾取器选中的具体对象实例（而非按名字查找的副本）
        // 解决场景：场景中存在同名 Operation 时，按名字查找可能返回错的实例。
        private ITxObject _pickedOperation;

        // 当前操作关联的机器人实例（拾取/预览/检查时缓存，后续双击表格行驱动姿态时直接用）
        private TxRobot _currentRobot;

        // 缓存每行数据索引（TxFlexGrid 行 → PathPointResult），用于点击跳转
        private readonly Dictionary<int, PathPointResult> _rowToResult = new Dictionary<int, PathPointResult>();

        // 双击 → 三击 检测：记录上次双击的时间和行号
        // 在 _tripleClickWindowMs 内对同一行再次单击 → 触发三击（Robot Jog）
        private DateTime _lastDoubleClickTime = DateTime.MinValue;
        private int _lastDoubleClickRow = -1;
        private const int _tripleClickWindowMs = 500;
        private static readonly Size _designSize = new Size(1280, 780);
        private bool _dpiApplied;

        // =====================================================================
        // 构造与生命周期
        // =====================================================================
        public ReachabilityCheckerForm()
        {
            SemiModal = false;
            // 统一窗体规范 + DPI（套件唯一一处缩放设置，详见 FormUiKit）
            FormUiKit.InitStandardForm(this, "机器人路径点位检查",
                _designSize, new Size(960, 580));

            InitializeComponent();
            // 防御：InitStandardForm 已设唯一 Name 作为 TxForm 几何持久化键（消除串扰），
            // 这里在 InitializeComponent 之后再钉一次，确保 Layout 分部里没有把它改回空。
            this.Name = this.GetType().FullName;


            // 注意：InitializeComponent 所在的 Layout 分部文件里不要再设
            //       AutoScaleMode / AutoScaleDimensions / Font，统一交给 FormUiKit。
            // LoadRobotsAndOperations 移到 OnLoad，避免阻塞窗体首次显示
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            FormUiKit.ApplyDpiScaling(this, ref _dpiApplied, _designSize);

            // 全局护栏：UI 线程未处理异常 → 写日志而不是闪退
            // 注：仅对本插件 UI 线程生效；PS 主线程的 SEH 仍可能击穿
            Application.ThreadException += (s, ev) =>
            {
                try { Log($"[全局拦截] {ev.Exception?.GetType().Name}: {ev.Exception?.Message}", "ERR"); }
                catch { }
            };

            try { LoadRobotsAndOperations(); }
            catch (Exception ex) { Log($"加载异常: {ex.Message}", "ERR"); }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            try
            {
                if (_txtOpNode != null)
                {
                    _txtOpNode.Picked -= OnOpNodePicked;
                    _txtOpNode.ListenToPick = false;
                    _txtOpNode.UnInitialize();
                }
            }
            catch { }
            base.OnFormClosing(e);
        }

        // =====================================================================
        // ILogger 实现 — 服务层通过此接口写日志
        // =====================================================================
        public void Log(string message, string level = "INFO")
        {
            if (_logBox == null || _logBox.IsDisposed) return;

            // DEBUG 级别在非 verbose 模式下静默丢弃
            if (level == "DEBUG" && !_logVerbose) return;

            if (_logBox.InvokeRequired) { _logBox.BeginInvoke(new Action(() => Log(message, level))); return; }

            string line = $"[{DateTime.Now:HH:mm:ss.fff}] [{level}] {message}";
            System.Drawing.Color col = level == "ERR" ? TxClrLogErr.Color
                                    : level == "WARN" ? TxClrLogWarn.Color
                                    : level == "OK" ? TxClrLogOk.Color
                                                       : TxClrLogText.Color;

            _logBox.SelectionStart = _logBox.TextLength;
            _logBox.SelectionLength = 0;
            _logBox.SelectionColor = col;
            _logBox.AppendText(line + "\n");
            _logBox.SelectionColor = _logBox.ForeColor;
            _logBox.ScrollToCaret();

            if ((level == "ERR" || level == "WARN") && !_logVisible)
            {
                _logVisible = true; _logPanel.Visible = true;
                if (_btnFooterLog != null) _btnFooterLog.Text = "▼ 隐藏日志";
            }
        }

        private void SetStatus(string text)
        {
            if (_lblStatus != null) _lblStatus.Text = text;
        }

        private void SetStatus(string text, System.Drawing.Color _ignored) => SetStatus(text);
    }
}