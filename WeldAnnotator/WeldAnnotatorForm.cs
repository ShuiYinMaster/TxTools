// =============================================================================
// WeldAnnotatorForm.cs  v3  —  焊点标注截图导出插件
// C# 7.3 / Tecnomatix.Engineering SDK
//
// GUI 布局（参考 RobotReachabilityChecker 风格）：
//   TxToolStrip          ← 顶部操作按钮
//   TabControl（顶部）    ← Tab1:快照控制  Tab2:标注信息
//   ┌──────────┬─────────────────────────────┐
//   │左侧卡片   │ TxFlexGrid（焊点列表）       │
//   │GroupBox  │                             │
//   │TxObjGrid │                             │
//   └──────────┴─────────────────────────────┘
//   StatusStrip + 可折叠日志面板
//
// 所有 PS 操作通过 PsReader 完成，Form 只处理 UI 逻辑。
// Excel 写入：dynamic COM（无需 Microsoft.Office.Interop.Excel 引用）。
// =============================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;
using C1.Win.C1FlexGrid;
using MyPlugin.ExportGun;     // PsReader, WeldAnnotationPoint
using TxTools.Common;

using D = System.Drawing;
using DI = System.Drawing.Imaging;
using WF = System.Windows.Forms;

namespace WeldAnnotator
{
    // =========================================================================
    // 命令入口
    // =========================================================================
    public class WeldPointAnnotatorCmd : TxButtonCommand
    {
        public override string Name { get { return ".焊点标注截图"; } }
        public override string Category { get { return "My Plugins"; } }
        public override string Tooltip { get { return "焊点标注截图 → Excel"; } }
        public override string Description { get { return "将PS视口截图与焊点标注导出到活动Excel"; } }
        public override string Bitmap { get { return "image.WeldAnnotator.bmp"; } }
        public override string LargeBitmap { get { return "image.WeldAnnotator.png"; } }

        public override void Execute(object param)
        {
            try
            {
                WeldAnnotatorForm form = new WeldAnnotatorForm();
                form.Show();
                // 打开后立即把焦点还给 PS 主窗口：用户无需先点 PS 视图就可操作
                try
                {
                    // SetForegroundWindow 需要目标窗口句柄；拿 PS 主窗口 (TxApplication)
                    dynamic app = TxApplication.ActiveDocument;
                    // 若 SDK 暴露 TxApplication.MainWindow / MainForm 则用之；否则退而求其次用 Process
                    System.Diagnostics.Process ps = System.Diagnostics.Process.GetCurrentProcess();
                    if (ps != null && ps.MainWindowHandle != IntPtr.Zero)
                        SetForegroundWindow(ps.MainWindowHandle);
                }
                catch { }
            }
            catch (Exception ex)
            {
                TxMessageBox.ShowModal("插件启动失败：" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
    }

    // =========================================================================
    // 标注文字命名模式
    // =========================================================================
    public enum LabelNamingMode
    {
        Sequence = 0,   // 按顺序号（1, 2, 3, ...）
        PointName = 1,   // 按焊点名称
        Prefix = 2,   // 前缀 + 序号
        Suffix = 3,   // 序号 + 后缀
        SeqAndName = 4    // 序号 + 焊点名称（如 "1-via2098"）
    }

    // =========================================================================
    // 标注样式
    // =========================================================================
    public class AnnotationStyle
    {
        public D.Color DotColor = D.Color.Red;
        public int DotRadius = 8;
        public D.Color LineColor = D.Color.Red;
        public float LineWidth = 1.5f;
        public D.Color BoxBorderColor = D.Color.Black;
        public D.Color BoxFillColor = D.Color.White;
        public D.Color TextColor = D.Color.Black;
        public D.Font TextFont = new D.Font("Arial", 9f, D.FontStyle.Regular);
        public int BoxPadding = 4;
        public int OffsetX = 40;
        public int OffsetY = 40;

        // 分类 → Excel MsoAutoShapeType 形状 ID。默认按图例约定。
        // 可用形状号（msoAutoShapeType 常用值）：
        //   1  矩形, 4  菱形, 5  圆角矩形, 7  等腰三角形, 9  椭圆（圆）,
        //   10 正六边形, 12 五角星, 20 环（胶囊类），其他可按需扩展
        public System.Collections.Generic.Dictionary<string, int> CategoryShapes
            = new System.Collections.Generic.Dictionary<string, int>
            {
                { "二层板点焊",         9  },   // 实心圆
                { "二层板补焊",         9  },   // 空心圆（Filled=false）
                { "三层板及以上点焊",   7  },   // 实心三角
                { "三层板及以上补焊",   7  },   // 空心三角
                { "焊缝",              1  },   // 矩形（会拉成细长条）
                { "CO2焊点",           1  },   // 实心矩形
                { "螺母",              10 },   // 实心六边形
                { "螺钉螺栓",          1  },   // 空心矩形
                { "胶",                5  },   // 圆角矩形（胶囊）
                { "强度校验点",         12 },   // 实心五角星
                { "重要特性",          12 },   // 空心五角星
                { "关键特性",          12 }    // 实心五角星（加粗）
            };

        // 分类 → 是否实心（true=填充颜色，false=空心仅描边）
        public System.Collections.Generic.Dictionary<string, bool> CategoryFilled
            = new System.Collections.Generic.Dictionary<string, bool>
            {
                { "二层板点焊",         true  },
                { "二层板补焊",         false },
                { "三层板及以上点焊",   true  },
                { "三层板及以上补焊",   false },
                { "焊缝",              true  },
                { "CO2焊点",           true  },
                { "螺母",              true  },
                { "螺钉螺栓",          false },
                { "胶",                false },
                { "强度校验点",         true  },
                { "重要特性",          false },
                { "关键特性",          true  }
            };
    }

    // =========================================================================
    // 主窗体
    // =========================================================================
    public class WeldAnnotatorForm : TxForm
    {
        // ── 顶部卡片（三张独立 GroupBox） ─────────────────────────────────────
        // 卡片①：快照控制
        private WF.Button _btnSnap, _btnRestore, _btnShowOnly, _btnShowAll;
        private WF.Label _lblSnapStatus;
        // 卡片②：导出选项
        private WF.Label _lblCount;
        private WF.CheckBox _chkNewSheet;
        private WF.CheckBox _chkWriteList;
        private WF.CheckBox _chkAutoThick;
        // 卡片③：操作
        private WF.Button _btnExport, _btnClear, _btnStyle;

        // 进度条现居底部状态栏
        private ToolStripProgressBar _progress;

        // ── 左侧卡片：TxObjGridCtrl ────────────────────────────────────────────
        private WF.GroupBox _leftCard;
        private TxObjGridCtrl _opGrid;
        private WF.Label _lblOpHint;

        // ── TxFlexGrid 焊点列表 ───────────────────────────────────────────────
        private TxFlexGrid _grid;

        // ── 状态栏 + 日志 ─────────────────────────────────────────────────────
        private StatusStrip _status;
        private ToolStripStatusLabel _lblStatus;
        private WF.Panel _logPanel;
        private RichTextBox _logBox;
        private bool _logVisible;

        // ── 列索引 ────────────────────────────────────────────────────────────
        private const int C_IDX = 0;
        private const int C_NAME = 1;
        private const int C_OP = 2;
        private const int C_X = 3;
        private const int C_Y = 4;
        private const int C_Z = 5;
        private const int C_VIS = 6;
        private const int C_TYPE = 7;   // 焊点分类（下拉：焊点/补焊点/强度校验点）

        // ── 卡片④：标注内容 ──────────────────────────────────────────────────
        private WF.ComboBox _cmbLabelMode;   // 命名模式：序号/名称/前缀+序号/序号+后缀
        private WF.TextBox _txtPrefix;      // 前缀文本
        private WF.TextBox _txtSuffix;      // 后缀文本
        private WF.Label _lblPrefix, _lblSuffix;

        // ── 数据 ─────────────────────────────────────────────────────────────
        private List<WeldAnnotationPoint> _points = new List<WeldAnnotationPoint>();
        private string _selOpName = null;
        private ITxObject _selOpObj = null;
        private List<Tuple<ITxObject, bool>> _dispSnapshot = null;
        private AnnotationStyle _style = new AnnotationStyle();

        // 焊点分类（key = 焊点在 _points 中的 Index 字段；默认 "焊点"）
        private System.Collections.Generic.Dictionary<int, string> _categories
            = new System.Collections.Generic.Dictionary<int, string>();

        // 记录用户手动编辑过类型的焊点 Index，自动检测板厚时不覆盖这些
        private System.Collections.Generic.HashSet<int> _categoriesManuallyEdited
            = new System.Collections.Generic.HashSet<int>();

        // 请将此行添加到你的 Form 类成员变量中（如果在循环或多次调用外已有，可忽略）：
        private WF.ToolTip _sharedToolTip = new WF.ToolTip();

        private static readonly string[] CATEGORY_OPTIONS = {
            "二层板点焊", "二层板补焊",
            "三层板及以上点焊", "三层板及以上补焊",
            "焊缝", "CO2焊点",
            "螺母", "螺钉螺栓",
            "胶",
            "强度校验点", "重要特性", "关键特性"
        };
        private const string DEFAULT_CATEGORY = "二层板点焊";

        // =========================================================================
        // 私有：本插件独占的窗口尺寸/位置（不与其他 TxForm 共享缓存）
        private static readonly D.Size _myDefaultSize = new D.Size(1000, 620);
        private static readonly D.Size _myMinimumSize = new D.Size(820, 480);
        private bool _dpiApplied;

        public WeldAnnotatorForm()
        {
            SemiModal = false;
            // 统一窗体规范 + DPI + 唯一 Name（ExportGun 同款防串扰方案，详见 FormUiKit）。
            // InitStandardForm 内部按类型全名设 this.Name，使本窗体在 TxForm 的几何
            // 持久化里拥有独立存储键 —— 既消除跨插件尺寸串扰，又能各自记住上次大小。
            FormUiKit.InitStandardForm(this, "焊点标注截图导出",
                _myDefaultSize, _myMinimumSize);

            // —— 保留本插件特有设置 ——
            // 1) 中文字体：Microsoft YaHei UI，解决 PS 宿主下 Flat+Tahoma 的中文按钮乱码
            //    （这是之前专门定的修复，勿改回系统默认字体）
            Font = new D.Font("Microsoft YaHei UI", 9f);
            ShowInTaskbar = true;
            // 2) 默认停靠右上角，避免居中盖住 PS 主视图。仅作首次打开默认位置，
            //    之后用户的移动/缩放由 TxForm 按本窗体独立键各自记忆。
            PlaceTopRight();

            // Dock 顺序：Fill 必须最先 Add（占剩余空间），边缘 Dock 控件后 Add。
            BuildMainArea();    // Dock=Fill
            BuildLogPanel();    // Dock=Bottom
            BuildStatusBar();   // Dock=Bottom
            BuildCardPanel();   // Dock=Left — 左侧控制栏

            this.FormClosing += WeldAnnotatorForm_FormClosing;


        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            FormUiKit.ApplyDpiScaling(this, ref _dpiApplied, _myDefaultSize);
        }

        /// <summary>
        /// 把窗口放到主屏右上角（首次打开的默认位置，不盖住 PS 主视图）。
        /// 不再强制重设 Size —— 尺寸由 TxForm 按本窗体唯一 Name 独立持久化，
        /// 这才是消除跨插件串扰的正解（ExportGun 同款）。
        /// </summary>
        private void PlaceTopRight()
        {
            try
            {
                StartPosition = FormStartPosition.Manual;
                var scr = Screen.PrimaryScreen.WorkingArea;
                Location = new D.Point(Math.Max(scr.Left, scr.Right - Width - 20),
                                       Math.Max(scr.Top, scr.Top + 60));
            }
            catch { }
        }

        private void WeldAnnotatorForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                bool hasHide = PsReader.HasUnrestoredHide;
                bool hasSnap = _dispSnapshot != null && _dispSnapshot.Count > 0;
                if (!hasHide && !hasSnap) return;

                string msg =
                    "检测到场景中存在本插件产生的未恢复显示状态：\n\n" +
                    (hasHide ? "  • 有对象被'仅显示操作外观'隐藏，未恢复\n" : "") +
                    (hasSnap ? "  • 有已拍摄的快照，未恢复\n" : "") +
                    "\n关闭窗口前是否恢复？\n" +
                    "  [是] 恢复显示后关闭\n" +
                    "  [否] 直接关闭（场景保持当前状态）\n" +
                    "  [取消] 返回窗口";

                var r = MessageBox.Show(this, msg, "未恢复的显示状态",
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (r == DialogResult.Cancel) { e.Cancel = true; return; }
                if (r == DialogResult.Yes)
                {
                    // 优先用 _dispSnapshot（能完全还原初始场景），否则 RestoreFromLastHide
                    if (hasSnap)
                        PsReader.RestoreDisplayStates(_dispSnapshot, s => Log("INFO", s));
                    else
                        PsReader.RestoreFromLastHide(s => Log("INFO", s));
                }
            }
            catch (Exception ex)
            {
                Log("ERR", "关闭前恢复检查异常：" + ex.Message);
            }
        }

        // ── PS宿主回调：打开时不抢焦点，便于用户继续在 PS 视图中操作 ──────
        public override void OnInitTxForm()
        {
            base.OnInitTxForm();
        }

        // =========================================================================
        // 构建 UI
        // =========================================================================

        // ── 顶部卡片区（四个独立 GroupBox） ─────────────────────────────────
        //
        // 布局要点（避免 WinForms 常见坑）：
        //   1) cardHost 不用 AutoSize — Dock=Top 指定高度更稳定
        //   2) cardHost WrapContents=true 让窗口变窄时卡片自动换行
        //   3) GroupBox 用 AutoSize=GrowAndShrink 单独撑开宽度
        //   4) GroupBox 内放一个 Dock=Fill 的 Panel 容纳内部 FlowLayoutPanel，
        //      让内部控件能相对 GroupBox 的客户区垂直居中（Anchor 在
        //      FlowLayoutPanel 里无效，所以必须用 Panel + 手动定位）
        // ════════════════════════════════════════════════════════════════════
        //  左侧控制栏：单列堆叠，从上到下依次为
        //    ① 操作节点 (OP Selector) — TxObjGridCtrl 直接放这里
        //    ② 显示控制
        //    ③ 导出设置（合并旧"导出选项"+"标注内容"）
        //    ④ 主操作 + 状态信息
        //  好处：
        //    - 焊点表格获得全部右侧空间（原来被顶部 4 张卡片挤占）
        //    - 单列布局，工作流从上到下顺序清晰
        //    - 卡片宽度一致（固定 250px），视觉规整
        // ════════════════════════════════════════════════════════════════════
        private void BuildCardPanel()
        {

            // 左侧滚动容器（窗口高度不够时整列可滚）
            WF.Panel rail = new WF.Panel
            {
                Dock = DockStyle.Left,
                Width = 270,
                AutoScroll = true,
                BackColor = D.SystemColors.Control,
                Padding = new Padding(6, 4, 6, 4)
            };

            // FlowLayoutPanel 垂直堆叠
            WF.FlowLayoutPanel stack = new WF.FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                BackColor = D.Color.Transparent
            };

            // 卡片①：操作节点（OP）—— TxObjGridCtrl 移入此卡
            WF.GroupBox opCard = MakeRailCard("操作节点 (OP)");
            _opGrid = new TxObjGridCtrl
            {
                Dock = DockStyle.Top,
                Height = 110,
                ListenToPick = true,
                EnableMultipleSelection = false,
                EnableRecurringObjects = false
            };
            _opGrid.ObjectInserted += new TxObjGridCtrl_ObjectInsertedEventHandler(OnOpInserted);
            _opGrid.RowDeleted += new TxObjGridCtrl_RowDeletedEventHandler(OnOpDeleted);
            _lblOpHint = new WF.Label
            {
                Text = "在 PS 中选 OP 或视口选中",
                Dock = DockStyle.Bottom,
                Height = 18,
                ForeColor = D.Color.Gray,
                Font = new D.Font("Microsoft YaHei UI", 8f),
                TextAlign = D.ContentAlignment.MiddleCenter
            };
            opCard.Controls.Add(_lblOpHint);
            opCard.Controls.Add(_opGrid);

            // 卡片②：显示控制（2×2 按钮 + 状态行）
            WF.GroupBox snapCard = MakeRailCard("显示控制");
            _btnSnap = MkRailBtn("拍摄快照", BtnSnap_Click, D.Color.FromArgb(0, 112, 192));
            _btnRestore = MkRailBtn("恢复快照", BtnRestore_Click, D.Color.FromArgb(84, 130, 53));
            _btnShowOnly = MkRailBtn("仅显示外观", BtnShowOnly_Click, D.Color.FromArgb(197, 90, 17));
            _btnShowAll = MkRailBtn("显示全部", BtnShowAll_Click, D.Color.FromArgb(100, 100, 100));
            _btnRestore.Enabled = false;
            _lblSnapStatus = new WF.Label
            {
                Text = "",
                AutoSize = false,
                Height = 16,
                Width = 240,
                ForeColor = D.Color.Gray,
                Font = new D.Font("Microsoft YaHei UI", 8f),
                Margin = new Padding(0, 4, 0, 0),
                TextAlign = D.ContentAlignment.MiddleLeft
            };
            FillRailCardGrid(snapCard, 2,
                new WF.Control[] { _btnSnap, _btnRestore, _btnShowOnly, _btnShowAll },
                new WF.Control[] { _lblSnapStatus });

            // 卡片③：导出设置（合并旧"导出选项" + "标注内容"）
            WF.GroupBox cfgCard = MakeRailCard("导出设置");
            _chkNewSheet = new WF.CheckBox
            {
                Text = "写入新 Sheet",
                Checked = false,
                AutoSize = true,
                Margin = new Padding(0, 2, 0, 2)
            };
            _chkWriteList = new WF.CheckBox
            {
                Text = "附焊点数据表",
                Checked = false,
                AutoSize = true,
                Margin = new Padding(0, 2, 0, 2)
            };
            new WF.ToolTip().SetToolTip(_chkWriteList,
                "勾选后，在截图下方附加焊点数据表（序号/名称/操作/XYZ/视口状态/类型）。");
            _chkAutoThick = new WF.CheckBox
            {
                Text = "自动检测板厚",
                Checked = false,
                AutoSize = true,
                Margin = new Padding(0, 2, 0, 2)
            };
            new WF.ToolTip().SetToolTip(_chkAutoThick,
                "按焊点绑定的不重复零件数自动设置类型：\n" +
                "  2 个零件 → 二层板点焊\n" +
                "  3+ 个零件 → 三层板及以上点焊\n" +
                "已手动修改过类型的焊点不会被覆盖。");
            _chkAutoThick.CheckedChanged += (s, e) => ApplyAutoThicknessIfEnabled();

            // 命名模式区
            WF.Label lblMode = new WF.Label
            {
                Text = "标注命名：",
                AutoSize = true,
                Margin = new Padding(0, 8, 0, 2),
                Font = new D.Font("Microsoft YaHei UI", 8.5f, D.FontStyle.Bold)
            };
            _cmbLabelMode = new WF.ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 200,
                Margin = new Padding(0, 0, 0, 2)
            };
            _cmbLabelMode.Items.AddRange(new object[]
            { "按序号", "按焊点名", "前缀+序号", "序号+后缀", "序号+焊点名" });
            _cmbLabelMode.SelectedIndex = 0;
            _cmbLabelMode.SelectedIndexChanged += (s, e) => UpdateLabelCardEnabled();

            // 前后缀紧凑放一行
            WF.FlowLayoutPanel pfxRow = new WF.FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 2)
            };
            _lblPrefix = new WF.Label { Text = "前缀", AutoSize = true, Margin = new Padding(0, 7, 2, 0) };
            _txtPrefix = new WF.TextBox { Width = 60, Margin = new Padding(0, 4, 8, 0) };
            _lblSuffix = new WF.Label { Text = "后缀", AutoSize = true, Margin = new Padding(0, 7, 2, 0) };
            _txtSuffix = new WF.TextBox { Width = 60, Margin = new Padding(0, 4, 0, 0) };
            pfxRow.Controls.AddRange(new WF.Control[] { _lblPrefix, _txtPrefix, _lblSuffix, _txtSuffix });

            // 把所有控件按顺序放入卡片
            WF.FlowLayoutPanel cfgInner = new WF.FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Padding = new Padding(2, 2, 2, 2)
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
            WF.GroupBox actCard = MakeRailCard("操作");

            _btnExport = MkRailBtnWide("导出到 Excel", BtnExport_Click, D.Color.FromArgb(0, 120, 215));
            _btnExport.Font = new D.Font("Microsoft YaHei UI", 10f, D.FontStyle.Bold);
            _btnExport.Height = 32;   // 主操作加大

            _btnClear = MkRailBtn("清空列表", BtnClear_Click, D.Color.FromArgb(183, 28, 28));
            _btnStyle = MkRailBtn("标注样式", BtnStyle_Click, D.Color.FromArgb(66, 66, 66));
            WF.Button btnMinimize = MkRailBtn("最小化",
                (s, e) => WindowState = FormWindowState.Minimized,
                D.Color.FromArgb(120, 120, 120));

            // 修复：复用类级别的 ToolTip 实例，防止多次生成卡片导致句柄泄漏
            _sharedToolTip.SetToolTip(btnMinimize, "最小化本窗口，以便在 PS 中调整视角、选择/隐藏对象。");

            // 修复：增加 Dock = DockStyle.Top 和 AutoSize = true，防止与上面的网格重叠或高 DPI 截断
            _lblCount = new WF.Label
            {
                Text = "焊点：0  视口内：0",
                AutoSize = true,
                Dock = DockStyle.Top,
                Margin = new Padding(0, 4, 0, 4),
                Font = new D.Font("Microsoft YaHei UI", 9f, D.FontStyle.Bold),
                ForeColor = D.Color.FromArgb(0, 70, 127),
                TextAlign = D.ContentAlignment.MiddleLeft
            };

            FillRailCardGrid(actCard, 1,
                new WF.Control[] { _btnExport },
                new WF.Control[] { /* 2x2 子区 */ });

            // 第二行手动加 2x2 副按钮 + 状态行
            WF.TableLayoutPanel subGrid = new WF.TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                RowCount = 2, // 显式声明 2 行
                Margin = new Padding(0, 4, 0, 0)
            };

            subGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            subGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            // 修复：使用 RowStyle 替代插入空 Label 占位，规范布局并减少句柄开销
            subGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            subGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            subGrid.Controls.Add(_btnClear, 0, 0);
            subGrid.Controls.Add(_btnStyle, 1, 0);
            subGrid.Controls.Add(btnMinimize, 0, 1);

            // 修复：调整 Add 顺序。在 WinForms 中，先 Add 的 Dock.Top 控件会在最上面
            actCard.Controls.Add(subGrid);
            actCard.Controls.Add(_lblCount);

            // 把所有卡片放进 stack
            stack.Controls.Add(opCard);
            stack.Controls.Add(snapCard);
            stack.Controls.Add(cfgCard);
            stack.Controls.Add(actCard);

            rail.Controls.Add(stack);
            Controls.Add(rail);
        }

        /// <summary>
        /// 创建左侧卡片：Dock=Top, 宽度固定 254，AutoSize 高度，ColoredGroupBox 自绘标题。
        /// </summary>
        private WF.GroupBox MakeRailCard(string title)
        {
            return new FormUiKit.ColoredGroupBox
            {
                Text = title,
                Dock = DockStyle.Top,
                Width = 254,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Font = new D.Font("Microsoft YaHei UI", 9f, D.FontStyle.Bold),
                HeaderColor = D.Color.FromArgb(0, 70, 127),
                ForeColor = D.Color.FromArgb(0, 70, 127),
                Padding = new Padding(6, 12, 6, 4),
                Margin = new Padding(0, 0, 0, 4),
                MinimumSize = new D.Size(254, 0)
            };
        }

        /// <summary>
        /// 卡片内的紧凑按钮（约 112×24，2 个并排正好填充卡片内宽 240）。
        /// </summary>
        private WF.Button MkRailBtn(string text, EventHandler h, D.Color fg)
        {
            FormUiKit.FlatColorButton b = new FormUiKit.FlatColorButton
            {
                Text = text,
                Width = 116,
                Height = 26,
                BgColor = fg,
                ForeColor = D.Color.White,
                BorderColor = fg,
                Font = new D.Font("Microsoft YaHei UI", 9f, D.FontStyle.Regular),
                Margin = new Padding(0, 2, 4, 2),
                Padding = new Padding(4, 2, 4, 2),
                FlatStyle = FlatStyle.Flat
            };
            b.Click += h;
            return b;
        }

        /// <summary>整宽按钮（240px，用于主操作如"导出到 Excel"）。</summary>
        private WF.Button MkRailBtnWide(string text, EventHandler h, D.Color fg)
        {
            FormUiKit.FlatColorButton b = new FormUiKit.FlatColorButton
            {
                Text = text,
                Width = 240,
                Height = 28,
                Dock = DockStyle.Top,
                BgColor = fg,
                ForeColor = D.Color.White,
                BorderColor = fg,
                Font = new D.Font("Microsoft YaHei UI", 10f, D.FontStyle.Bold),
                Padding = new Padding(4),
                FlatStyle = FlatStyle.Flat
            };
            b.Click += h;
            return b;
        }

        /// <summary>
        /// 把按钮按 cols 列均匀排列在 GroupBox 里，下面再可选地附加几个额外控件。
        /// </summary>
        private void FillRailCardGrid(WF.GroupBox card, int cols,
                                       WF.Control[] buttons, WF.Control[] extras)
        {
            WF.TableLayoutPanel tlp = new WF.TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = cols,
                RowCount = (buttons.Length + cols - 1) / cols
            };
            for (int c = 0; c < cols; c++)
                tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / cols));
            for (int i = 0; i < buttons.Length; i++)
                tlp.Controls.Add(buttons[i], i % cols, i / cols);
            card.Controls.Add(tlp);
            if (extras != null)
                foreach (var c in extras) card.Controls.Add(c);
        }

        /// <summary>
        /// 根据命名模式启用/禁用前缀、后缀输入框。
        /// </summary>
        private void UpdateLabelCardEnabled()
        {
            if (_cmbLabelMode == null) return;
            LabelNamingMode m = (LabelNamingMode)_cmbLabelMode.SelectedIndex;
            bool pfx = m == LabelNamingMode.Prefix;
            bool sfx = m == LabelNamingMode.Suffix;
            _lblPrefix.Enabled = _txtPrefix.Enabled = pfx;
            _lblSuffix.Enabled = _txtSuffix.Enabled = sfx;
        }

        // ── 主区域：左侧卡片 + TxFlexGrid ────────────────────────────────────
        private void BuildMainArea()
        {
            // 主区域：右侧整片用于 TxFlexGrid 焊点表格
            // 左侧 OP 选择 + 显示控制 + 设置 + 操作 已统一移至 BuildCardPanel 构建的 rail。
            WF.Panel main = new WF.Panel { Dock = DockStyle.Fill };

            // ── TxFlexGrid 焊点列表 ───────────────────────────────────────
            _grid = new TxFlexGrid
            {
                Dock = DockStyle.Fill,
                AllowEditing = true,                        // 启用编辑（下方只对 C_TYPE 开放）
                SelectionMode = SelectionModeEnum.Row
            };
            _grid.Rows.Fixed = 1;
            _grid.Cols.Count = 8;                            // 新增 C_TYPE
            _grid.Rows.Count = 1;

            string[] hdrs = { "#", "焊点名称", "操作名", "X(mm)", "Y(mm)", "Z(mm)", "视口", "类型" };
            int[] widths = { 36, 140, 110, 78, 78, 78, 52, 92 };
            for (int c = 0; c < hdrs.Length; c++)
            {
                _grid[0, c] = hdrs[c];
                _grid.Cols[c].Width = widths[c];
                _grid.Cols[c].AllowSorting = false;
                _grid.Cols[c].AllowEditing = (c == C_TYPE);
            }
            _grid.Cols[C_TYPE].ComboList = string.Join("|", CATEGORY_OPTIONS);
            _grid.Cols[C_TYPE].TextAlign = TextAlignEnum.CenterCenter;

            _grid.Styles[CellStyleEnum.Fixed].BackColor = D.Color.FromArgb(68, 114, 196);
            _grid.Styles[CellStyleEnum.Fixed].ForeColor = D.Color.White;
            _grid.Styles[CellStyleEnum.Fixed].Font = new D.Font("Microsoft YaHei UI", 9f, D.FontStyle.Bold);
            _grid.Styles[CellStyleEnum.Alternate].BackColor = D.Color.FromArgb(242, 242, 242);

            _grid.AfterEdit += Grid_AfterEdit;

            main.Controls.Add(_grid);
            Controls.Add(main);
        }

        /// <summary>
        /// Grid 单元格编辑完成后：若编辑的是"类型"列，把新值写回 _categories。
        /// </summary>
        private void Grid_AfterEdit(object sender, RowColEventArgs e)
        {
            try
            {
                if (e.Col != C_TYPE) return;
                if (e.Row < 1 || e.Row - 1 >= _points.Count) return;
                string newVal = _grid[e.Row, C_TYPE] as string;
                if (string.IsNullOrEmpty(newVal)) newVal = DEFAULT_CATEGORY;
                int key = _points[e.Row - 1].Index;
                _categories[key] = newVal;
                _categoriesManuallyEdited.Add(key);   // 标记为用户手动编辑，自动检测不再覆盖
                Log("INFO", string.Format(
                    "[Annotator] 点 [{0}] 类型改为 [{1}]", _points[e.Row - 1].Name, newVal));
            }
            catch (Exception ex) { Log("WARN", "Grid_AfterEdit: " + ex.Message); }
        }

        private void BuildStatusBar()
        {
            _status = new StatusStrip();
            _lblStatus = new ToolStripStatusLabel("就绪")
            {
                Spring = true,
                TextAlign = D.ContentAlignment.MiddleLeft
            };
            _progress = new ToolStripProgressBar
            {
                Width = 150,
                Visible = false,
                Alignment = ToolStripItemAlignment.Right
            };
            ToolStripButton tsLog = new ToolStripButton("日志 ▲");
            tsLog.Click += (s, e) => ToggleLog(tsLog);
            _status.Items.AddRange(new ToolStripItem[] { _lblStatus, _progress, tsLog });
            Controls.Add(_status);
        }

        private void BuildLogPanel()
        {
            _logPanel = new WF.Panel
            {
                Dock = DockStyle.Bottom,
                Height = 110,
                Visible = false
            };
            _logBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                BackColor = D.Color.FromArgb(30, 30, 30),
                ForeColor = D.Color.LightGray,
                Font = new D.Font("Consolas", 8f),
                ReadOnly = true
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

        // =========================================================================
        // TxObjGridCtrl 事件（自动确认，无需按钮）
        // =========================================================================

        private void OnOpInserted(object sender, TxObjGridCtrl_ObjectInsertedEventArgs e)
        {
            // 每次插入后立即读取第0个对象作为当前选中 OP
            try
            {
                ITxObject txo = _opGrid.GetObject(0);
                if (txo == null) return;

                // 接受任意可作为焊点容器的 ITxObject（TxWeldOperation /
                // TxCompoundOperation / 甚至 TxWeldPoint 等），由 FillPoints 兜底
                _selOpName = SafeGetName(txo);
                _selOpObj = txo;
                _lblOpHint.Text = _selOpName;
                _lblOpHint.ForeColor = D.Color.DarkGreen;
                SetStatus("已选择操作：" + _selOpName + "  正在读取焊点...");

                // 立即读取焊点并刷新网格（供用户预览）
                bool ok = ReadWeldPointsFromSelectedOp();
                ApplyAutoThicknessIfEnabled();
                RefreshGrid();
                if (ok)
                    SetStatus(string.Format(
                        "已选择 [{0}]，读取到 {1} 个焊点", _selOpName, _points.Count));
                else
                    SetStatus(string.Format(
                        "已选择 [{0}]，但未读取到焊点（可能该 OP 下无焊点）", _selOpName));
            }
            catch (Exception ex) { Log("WARN", "OnOpInserted：" + ex.Message); }
        }

        private void OnOpDeleted(object sender, TxObjGridCtrl_RowDeletedEventArgs e)
        {
            // 重新读取（可能还有其他行）
            try
            {
                if (_opGrid.Count > 0)
                {
                    ITxObject txo = _opGrid.GetObject(0);
                    if (txo != null)
                    {
                        _selOpName = SafeGetName(txo);
                        _selOpObj = txo;
                        _lblOpHint.Text = _selOpName;
                        _lblOpHint.ForeColor = D.Color.DarkGreen;
                        ReadWeldPointsFromSelectedOp();
                        RefreshGrid();
                        return;
                    }
                }
            }
            catch { }
            _selOpName = null;
            _selOpObj = null;
            _lblOpHint.Text = "选OP或视口选中";
            _lblOpHint.ForeColor = D.Color.Gray;
            _points.Clear();
            RefreshGrid();
        }

        // =========================================================================
        // 快照控制
        // =========================================================================

        private void BtnSnap_Click(object sender, EventArgs e)
        {
            SetStatus("拍摄快照中...");
            try
            {
                _dispSnapshot = PsReader.SnapshotDisplayStates(s => Log("INFO", s));
                int cnt = _dispSnapshot != null ? _dispSnapshot.Count : 0;
                if (cnt == 0)
                {
                    _lblSnapStatus.Text = "快照为空";
                    _lblSnapStatus.ForeColor = D.Color.Gray;
                    _btnRestore.Enabled = false;
                    SetStatus("快照为空");
                    return;
                }
                _lblSnapStatus.Text = string.Format("已记录 {0} 对象", cnt);
                _lblSnapStatus.ForeColor = D.Color.DarkGreen;
                _btnRestore.Enabled = true;
                SetStatus(string.Format("快照完成，记录 {0} 个对象显示状态", cnt));
            }
            catch (Exception ex)
            {
                Log("ERR", "拍摄快照失败：" + ex.Message);
                SetStatus("快照失败：" + ex.Message);
            }
        }

        private void BtnRestore_Click(object sender, EventArgs e)
        {
            SetStatus("恢复快照中...");
            try
            {
                if (_dispSnapshot != null && _dispSnapshot.Count > 0)
                {
                    PsReader.RestoreDisplayStates(_dispSnapshot, s => Log("INFO", s));
                    _lblSnapStatus.Text = "已恢复快照";
                    _lblSnapStatus.ForeColor = D.Color.DarkGreen;
                    SetStatus("快照已恢复");
                }
                else
                {
                    // 无快照时回退到"恢复本次隐藏"
                    PsReader.RestoreFromLastHide(s => Log("INFO", s));
                    _lblSnapStatus.Text = "已恢复本次隐藏";
                    _lblSnapStatus.ForeColor = D.Color.Gray;
                    SetStatus("已恢复本次被隐藏的对象");
                }
            }
            catch (Exception ex)
            {
                Log("ERR", "恢复失败：" + ex.Message);
                SetStatus("恢复失败：" + ex.Message);
            }
        }

        private void BtnShowOnly_Click(object sender, EventArgs e)
        {
            if (_selOpObj == null)
            {
                SetStatus("请先在左侧选择操作节点");
                return;
            }

            SetStatus("准备白名单...");

            // 1. 直接对操作跑一遍 FillPoints，拿到完整 PointInfo（含 AllAppearances.RawObject）
            OperationInfo op = new OperationInfo
            {
                Name = _selOpName ?? SafeGetName(_selOpObj),
                TypeLabel = _selOpObj.GetType().Name,
                PsObject = _selOpObj
            };
            try
            {
                PsReader.FillPoints(op, PointType.WeldPoint, false, s => Log("INFO", s));
            }
            catch (Exception ex)
            {
                Log("WARN", "FillPoints 异常：" + ex.Message);
            }

            if (op.Points == null || op.Points.Count == 0)
            {
                TxMessageBox.ShowModal(
                    "操作 [" + op.Name + "] 下未找到任何焊点。",
                    "未找到焊点", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SetStatus("未找到焊点");
                return;
            }

            // 2. 构建白名单
            //    注意：不把 _selOpObj（操作本身）加入 —— 操作是逻辑节点，其 Blank/Display
            //    对应的是路径/via点等辅助显示，并非 3D 模型；加入白名单反而会让操作的 via 点
            //    一直显示在截图中。
            var whitelist = new List<ITxObject>();

            // 机器人 + 工具（复用 PsReader 的内部逻辑通过公开的 FindRobotForOperation）
            TxRobot robot = null;
            try { robot = PsReader.FindRobotForOperation(_selOpObj); } catch { }
            if (robot != null) whitelist.Add(robot);
            if (robot != null)
            {
                ITxObject tool = null;
                try
                {
                    dynamic dr = robot;
                    try { tool = dr.ActiveTool as ITxObject; } catch { }
                    if (tool == null) try { tool = dr.CurrentTool as ITxObject; } catch { }
                }
                catch { }
                if (tool != null) whitelist.Add(tool);
            }

            // 焊点绑定的外观对象
            int boundCount = 0;
            foreach (PointInfo pi in op.Points)
            {
                if (pi == null || pi.AllAppearances == null) continue;
                foreach (AppearanceRef ar in pi.AllAppearances)
                {
                    if (ar != null && ar.RawObject != null)
                    {
                        whitelist.Add(ar.RawObject);
                        boundCount++;
                    }
                }
            }

            Log("INFO", string.Format(
                "[仅显示] 白名单：操作+机器人/工具 + 焊点绑定外观 {0} 个（焊点总数 {1}）",
                boundCount, op.Points.Count));

            if (boundCount == 0)
            {
                TxMessageBox.ShowModal(
                    "已读取 " + op.Points.Count + " 个焊点，但未能从中获取任何绑定的外观对象。\n\n" +
                    "可能原因：\n" +
                    "• 焊点未绑定零件/外观（PS 中右键焊点 → 检查 Assigned Parts）\n" +
                    "• FillAppearancesForOperation 在此操作类型下未触发\n\n" +
                    "为避免把场景中所有对象都隐藏，已取消操作。",
                    "白名单无外观", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SetStatus("已取消：未能获取焊点绑定外观");
                return;
            }

            // 3. 执行隐藏
            SetStatus("设置仅显示操作外观...");
            try
            {
                PsReader.HideAllExcept(whitelist, s => Log("INFO", s));
                _btnRestore.Enabled = true;
                _lblSnapStatus.Text = "已隐藏其他对象";
                _lblSnapStatus.ForeColor = D.Color.FromArgb(197, 90, 17);
                SetStatus("已仅显示操作绑定外观");
            }
            catch (InvalidOperationException iex)
            {
                Log("WARN", iex.Message);
                TxMessageBox.ShowModal(iex.Message, "无法执行",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SetStatus("已取消");
            }
            catch (Exception ex)
            {
                Log("ERR", "操作失败：" + ex.Message);
                SetStatus("失败：" + ex.Message);
            }
        }

        private void BtnShowAll_Click(object sender, EventArgs e)
        {
            SetStatus("恢复全部显示...");
            try
            {
                PsReader.ShowAllDevices(s => Log("INFO", s));
                SetStatus("已恢复全部显示");
            }
            catch (Exception ex)
            {
                Log("ERR", "失败：" + ex.Message);
                SetStatus("失败：" + ex.Message);
            }
        }

        // =========================================================================
        // 读取焊点 / 清空 / 导出
        // =========================================================================

        /// <summary>
        /// 从左侧 TxObjGridCtrl 中已选中的 OP 读取焊点。
        /// 调用 PsReader 原有逻辑：FillPoints 四级降级
        /// （L1 WeldOpEnum → L2 Locations → L3 MfgFeatures → L4 DeepWalk）。
        /// 成功时将结果写入 _points；失败或未选 OP 时返回 false（调用方自行提示）。
        /// </summary>
        private bool ReadWeldPointsFromSelectedOp()
        {
            if (_selOpObj == null)
            {
                _points.Clear();
                return false;
            }

            OperationInfo op = new OperationInfo
            {
                Name = _selOpName ?? SafeGetName(_selOpObj),
                TypeLabel = _selOpObj.GetType().Name,
                PsObject = _selOpObj
            };
            try
            {
                PsReader.FillPoints(op, PointType.WeldPoint, false, s => Log("INFO", s));
            }
            catch (Exception ex)
            {
                Log("WARN", string.Format("FillPoints [{0}] 异常：{1}", op.Name, ex.Message));
            }

            List<WeldAnnotationPoint> result = new List<WeldAnnotationPoint>();
            int idx = 1;
            if (op.Points != null)
            {
                foreach (PointInfo pi in op.Points)
                {
                    if (pi == null) continue;
                    if (pi.Type != PointType.WeldPoint) continue;
                    if (pi.Position == null || pi.Position.Length < 3) continue;

                    // 统计绑定的不重复零件数：优先按 ParentPartName 去重，
                    // 没有 ParentPartName 的退而用 Appearance Name 当作 key
                    int boundParts = 0;
                    if (pi.AllAppearances != null && pi.AllAppearances.Count > 0)
                    {
                        var seen = new System.Collections.Generic.HashSet<string>(
                            StringComparer.OrdinalIgnoreCase);
                        foreach (var ar in pi.AllAppearances)
                        {
                            if (ar == null) continue;
                            string key = !string.IsNullOrEmpty(ar.ParentPartName)
                                ? ar.ParentPartName
                                : (ar.Name ?? "");
                            if (!string.IsNullOrEmpty(key)) seen.Add(key);
                        }
                        boundParts = seen.Count;
                    }

                    result.Add(new WeldAnnotationPoint
                    {
                        Index = idx++,
                        Name = string.IsNullOrEmpty(pi.Name) ? ("P" + idx) : pi.Name,
                        OpName = op.Name ?? "(未知)",
                        X = pi.Position[0],
                        Y = pi.Position[1],
                        Z = pi.Position[2],
                        WorldTx = null,
                        BoundPartsCount = boundParts
                    });
                }
            }

            _points = result;

            // 保留已有的分类选择（修复：导出时调用此方法会清空已编辑的类型）
            // 仅对字典中不存在的 Index 设默认值；已存在的保持用户选择。
            foreach (WeldAnnotationPoint wp in _points)
                if (!_categories.ContainsKey(wp.Index))
                    _categories[wp.Index] = DEFAULT_CATEGORY;

            Log("INFO", string.Format(
                "[Annotator] 读取 [{0}] 共 {1} 个焊点", op.Name, result.Count));
            return result.Count > 0;
        }

        private static string SafeGetName(ITxObject o)
        {
            try { dynamic d = o; return (d.Name as string) ?? o.GetType().Name; }
            catch { return o != null ? o.GetType().Name : "(null)"; }
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            _points.Clear();
            _categories.Clear();
            _categoriesManuallyEdited.Clear();
            RefreshGrid();
            SetStatus("已清空");
        }

        /// <summary>
        /// 若"自动检测板厚"已勾选，按每个焊点的 BoundPartsCount 自动设置分类。
        /// 已被用户手动编辑过的焊点（在 _categoriesManuallyEdited 里）保持不变。
        /// </summary>
        private void ApplyAutoThicknessIfEnabled()
        {
            if (_chkAutoThick == null || !_chkAutoThick.Checked) return;
            if (_points == null || _points.Count == 0) return;

            int changed = 0;
            foreach (WeldAnnotationPoint pt in _points)
            {
                if (_categoriesManuallyEdited.Contains(pt.Index)) continue;
                string cat = CategoryByThickness(pt.BoundPartsCount);
                if (cat == null) continue;
                string old;
                _categories.TryGetValue(pt.Index, out old);
                if (old == cat) continue;
                _categories[pt.Index] = cat;
                changed++;
            }
            Log("INFO", string.Format(
                "[自动板厚] 已自动设置 {0} 个焊点的分类（按绑定零件数）", changed));
            RefreshGrid();
        }

        /// <summary>
        /// 板厚数 → 分类名映射。null 表示不修改（如 0 或 1，保留默认/已有值）。
        /// </summary>
        private static string CategoryByThickness(int parts)
        {
            if (parts <= 1) return null;            // 0 或 1：保留现状
            if (parts == 2) return "二层板点焊";
            return "三层板及以上点焊";                 // 3+
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            // 先从左侧已选 OP 读取最新焊点（保证数据与当前选择一致）
            if (_selOpObj == null)
            {
                SetStatus("请先在左侧选择操作节点（在PS中点击一个OP即可）");
                return;
            }

            SetStatus("读取焊点...");
            bool hasPts = ReadWeldPointsFromSelectedOp();
            // 读取后若自动检测开关已开，立即应用一次
            ApplyAutoThicknessIfEnabled();
            RefreshGrid();
            if (!hasPts)
            {
                SetStatus(string.Format(
                    "OP [{0}] 下未读取到任何焊点，无法导出。请确认该 OP 包含 TxWeldPoint。",
                    _selOpName));
                return;
            }

            _btnExport.Enabled = false;
            _progress.Visible = true;
            _progress.Value = 0;
            try
            {
                // 截图（3-arg 实现：Size.Empty 让 PsReader 自动读取视口实际尺寸）
                SetStatus("截图中...");
                _progress.Value = 15;
                D.Bitmap bmp = PsReader.CaptureActiveViewer(
                    D.Size.Empty, false, s => Log("INFO", s));
                if (bmp == null) throw new Exception("截图失败（CaptureActiveViewer 返回 null）");

                // 投影（计算每个焊点在截图中的像素坐标）
                SetStatus("计算焊点位置...");
                _progress.Value = 35;
                PsReader.ProjectPointsToScreen(
                    _points, bmp.Width, bmp.Height, s => Log("INFO", s));
                RefreshGrid();

                // 写入 Excel（纯净截图 + 可编辑 Shape 标注）
                SetStatus("写入Excel（含可编辑标注）...");
                _progress.Value = 55;
                ExportToActiveExcel(bmp, _style);

                _progress.Value = 100;
                int inVp = _points.Count(p => p.InViewport);
                SetStatus(string.Format("导出完成：{0} 个焊点（视口内 {1} 个）",
                    _points.Count, inVp));
                Log("INFO", string.Format(
                    "已写入活动Excel：{0} 个焊点，{1} 个视口内点已生成可编辑 Shape 标注",
                    _points.Count, inVp));
            }
            catch (Exception ex)
            {
                Log("ERR", "导出失败：" + ex.Message);
                SetStatus("导出失败：" + ex.Message);
            }
            finally
            {
                _btnExport.Enabled = true;
                _progress.Visible = false;
            }
        }

        private void BtnStyle_Click(object sender, EventArgs e)
        {
            AnnotationStyleForm sf = new AnnotationStyleForm(_style);
            if (sf.ShowDialog(this) == DialogResult.OK)
            {
                _style = sf.Result;
                Log("INFO", "样式已更新");
            }
        }

        // =========================================================================
        // Grid 刷新
        // =========================================================================

        private void RefreshGrid()
        {
            _grid.Rows.Count = _points.Count + 1;
            for (int i = 0; i < _points.Count; i++)
            {
                WeldAnnotationPoint pt = _points[i];
                int row = i + 1;
                _grid[row, C_IDX] = pt.Index.ToString();
                _grid[row, C_NAME] = pt.Name;
                _grid[row, C_OP] = pt.OpName;
                _grid[row, C_X] = pt.X.ToString("F2");
                _grid[row, C_Y] = pt.Y.ToString("F2");
                _grid[row, C_Z] = pt.Z.ToString("F2");
                _grid[row, C_VIS] = pt.InViewport ? "是" : (pt.ScreenX == 0 && pt.ScreenY == 0 ? "-" : "否");

                // 类型：若字典中没有，初始化为默认分类
                string cat;
                if (!_categories.TryGetValue(pt.Index, out cat) || string.IsNullOrEmpty(cat))
                {
                    cat = DEFAULT_CATEGORY;
                    _categories[pt.Index] = cat;
                }
                _grid[row, C_TYPE] = cat;

                CellStyle cs = _grid.Rows[row].Style ?? _grid.Styles.Add("r" + row);
                if (pt.InViewport)
                    cs.BackColor = D.Color.FromArgb(198, 239, 206);
                else if (pt.ScreenX != 0 || pt.ScreenY != 0)
                    cs.BackColor = D.Color.FromArgb(255, 235, 156);
                else
                    cs.BackColor = D.Color.Empty;
                _grid.Rows[row].Style = cs;
            }

            int inVp = _points.Count(p => p.InViewport);
            _lblCount.Text = _points.Count > 0
                ? string.Format("焊点：{0} 个  视口内：{1} 个", _points.Count, inVp)
                : "焊点：0 个";

            // 列宽按内容自动扩展（保留最小宽度避免过窄）
            try
            {
                int[] minWidths = { 36, 100, 90, 70, 70, 70, 50, 90 };
                for (int c = 0; c < _grid.Cols.Count && c < minWidths.Length; c++)
                {
                    _grid.AutoSizeCol(c);
                    if (_grid.Cols[c].Width < minWidths[c])
                        _grid.Cols[c].Width = minWidths[c];
                    // 加 12px 留白避免文字贴边
                    _grid.Cols[c].Width += 12;
                }
            }
            catch { /* AutoSizeCol 在某些 C1 版本签名不同；失败不影响功能 */ }
        }

        // =========================================================================
        // 标注文字内容计算（按用户选择的命名模式生成）
        // =========================================================================
        private string GetAnnotationLabel(WeldAnnotationPoint pt, int ordinal)
        {
            // ordinal = 1-based 视口内序号（从1开始）
            LabelNamingMode mode = (LabelNamingMode)
                (_cmbLabelMode != null ? _cmbLabelMode.SelectedIndex : 0);
            string name = pt.Name ?? "";
            string pfx = _txtPrefix != null ? (_txtPrefix.Text ?? "") : "";
            string sfx = _txtSuffix != null ? (_txtSuffix.Text ?? "") : "";
            switch (mode)
            {
                case LabelNamingMode.PointName: return name;
                case LabelNamingMode.Prefix: return pfx + ordinal.ToString();
                case LabelNamingMode.Suffix: return ordinal.ToString() + sfx;
                case LabelNamingMode.SeqAndName: return ordinal.ToString() + "-" + name;
                default: return ordinal.ToString();
            }
        }

        // =========================================================================
        // 写入活动 Excel：纯净截图 + 可编辑 Shape 标注
        //   Shape 分三类（dot / line / textbox），每个焊点独立，用户可在 Excel
        //   中对位置、大小、颜色、文字进行任意编辑。
        //   dynamic COM 调用，无需 Microsoft.Office.Interop.Excel 引用。
        // =========================================================================

        private void ExportToActiveExcel(D.Bitmap bmp, AnnotationStyle style)
        {
            string tmp = Path.GetTempFileName() + ".png";
            bmp.Save(tmp, DI.ImageFormat.Png);

            dynamic xlApp = null;
            bool savedScreenUpdating = true;
            try
            {
                try { xlApp = Marshal.GetActiveObject("Excel.Application"); }
                catch { throw new Exception("未检测到已打开的 Excel，请先打开工作簿"); }

                dynamic wb = xlApp.ActiveWorkbook;
                if (wb == null) throw new Exception("Excel 中没有活动工作簿");

                // 关闭屏幕刷新以加速大量 Shape 创建（finally 恢复）
                try { savedScreenUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; }
                catch { }

                // 目标 Sheet
                dynamic ws;
                if (_chkNewSheet.Checked)
                {
                    dynamic sheets = wb.Sheets;
                    ws = sheets.Add(
                        System.Reflection.Missing.Value,
                        sheets[sheets.Count],
                        System.Reflection.Missing.Value,
                        System.Reflection.Missing.Value);
                    ws.Name = "焊点标注_" + DateTime.Now.ToString("MMddHHmm");
                }
                else
                {
                    ws = wb.ActiveSheet;
                }

                // ── 读取活动 Excel 当前选中范围（用于"按选中单元格大小缩放"） ──
                //   规则：
                //     ① 单个单元格（且未合并）→ 不缩放，按截图自然尺寸输出，位置=该单元格左上角
                //     ② 多个单元格 / 合并后的单元格 → 缩放图像 Group 至选中区，纵横比保持
                //   失败时回落到 A1 + maxWidthPt 800 的旧行为。
                double targetLeft = -1;
                double targetTop = -1;
                double targetWidth = -1;     // -1 表示"不约束"
                double targetHeight = -1;
                bool scaleToTarget = false;
                try
                {
                    dynamic sel = xlApp.Selection;
                    int selCellCount = 0;
                    bool selMerged = false;
                    try { selCellCount = (int)sel.Cells.Count; } catch { }
                    try { selMerged = (bool)sel.MergeCells; } catch { }

                    targetLeft = (double)sel.Left;
                    targetTop = (double)sel.Top;
                    if (selCellCount == 1 && !selMerged)
                    {
                        // 单格未合并：仅作为定位锚，不缩放
                        scaleToTarget = false;
                    }
                    else
                    {
                        targetWidth = (double)sel.Width;
                        targetHeight = (double)sel.Height;
                        scaleToTarget = (targetWidth > 4 && targetHeight > 4);
                    }
                }
                catch
                {
                    targetLeft = -1; targetTop = -1; scaleToTarget = false;
                }

                // ── 插入纯净截图 ──────────────────────────────────────────────
                dynamic pics = ws.Pictures(System.Reflection.Missing.Value);
                dynamic pic = pics.Insert(tmp);
                // 初始放在 A1（无视所选位置），稍后整体平移 / 缩放 Group
                dynamic a1 = ws.Cells[1, 1];
                pic.Left = (double)a1.Left;
                pic.Top = (double)a1.Top;

                const double maxWidthPt = 800.0;
                if ((double)pic.Width > maxWidthPt)
                {
                    double r = maxWidthPt / (double)pic.Width;
                    pic.Width = maxWidthPt;
                    pic.Height = (double)pic.Height * r;
                }

                // 给截图一个稳定可寻址的名字，以便后面 Shapes.Range 组合时找到它
                string picName = "WPPicture_" + DateTime.Now.ToString("HHmmss");
                try { pic.Name = picName; } catch { /* Pictures.Insert 返回的对象命名可能受限 */ }

                // 收集所有要参与最终 Group 的 Shape 名称
                var shapeNamesForGroup = new List<string>();
                shapeNamesForGroup.Add(picName);

                // 截图最终位置/尺寸（Excel 点），以及 像素→点 换算比
                double picLeft = (double)pic.Left;
                double picTop = (double)pic.Top;
                double picWidth = (double)pic.Width;
                double picHeight = (double)pic.Height;
                double sx = picWidth / bmp.Width;
                double sy = picHeight / bmp.Height;

                // Office 常量（手写，避免引用 PIA）
                const int msoShapeOval = 9;
                const int msoTextOrientationHoriz = 1;
                const int msoFalse = 0;
                const int xlHAlignCenter = -4108;
                const int xlVAlignCenter = -4108;
                const int msoConnectorStraight = 1;
                const int xlFreeFloating = 3;   // 图片/Shape 不随单元格尺寸变化移动

                // 截图 Placement 设为 FreeFloating：
                // 防止后续写入数据表时 Columns.AutoFit 改变列宽导致截图和标注移位。
                try { pic.Placement = xlFreeFloating; } catch { }

                int oleDot = D.ColorTranslator.ToOle(style.DotColor);
                int oleLine = D.ColorTranslator.ToOle(style.LineColor);
                int oleBoxFill = D.ColorTranslator.ToOle(style.BoxFillColor);
                int oleBoxBorder = D.ColorTranslator.ToOle(style.BoxBorderColor);
                int oleText = D.ColorTranslator.ToOle(style.TextColor);

                // 标注尺寸（Excel 点，独立于截图分辨率；用户可在 Excel 中自由修改）
                const double DOT_R_PT = 5.0;
                const double BOX_W_PT = 28.0;
                const double BOX_H_PT = 18.0;

                int shapeCount = 0;
                int inVp = _points.Count(p => p.InViewport);
                int processed = 0;

                // ── 预计算所有视口内点的标签位置（避免重叠） ─────────────────
                //   思路：把每个焊点按所属"靠近的边"分到四组（上/下/左/右），
                //   每组沿边按焊点坐标排序，再以固定间距 gap 均匀摆放。
                //   这样即使多个焊点挤在一起，标签也会沿边整齐分布，互不遮挡。
                var placedPts = _points.Where(p => p.InViewport).ToList();
                int ordCount = placedPts.Count;
                var labelPos = new System.Collections.Generic.Dictionary<int, Tuple<double, double>>();
                // Key = pt.Index（而非 Count 循环变量），Value = (bx, by)

                if (ordCount > 0)
                {
                    // 为每个点算到四条边的距离，选最近的
                    var pointEdges = placedPts.Select((p, idx) =>
                    {
                        double cx = picLeft + p.ScreenX * sx;
                        double cy = picTop + p.ScreenY * sy;
                        double dL = cx - picLeft;
                        double dR = (picLeft + picWidth) - cx;
                        double dT = cy - picTop;
                        double dB = (picTop + picHeight) - cy;
                        double minD = Math.Min(Math.Min(dL, dR), Math.Min(dT, dB));
                        int side = (minD == dT) ? 0 : (minD == dR) ? 1 : (minD == dB) ? 2 : 3;
                        // 0=Top, 1=Right, 2=Bottom, 3=Left
                        return new { Pt = p, Ordinal = idx + 1, Cx = cx, Cy = cy, Side = side };
                    }).ToList();

                    // 按 side 分组
                    foreach (int sideId in new[] { 0, 1, 2, 3 })
                    {
                        var grp = pointEdges.Where(e => e.Side == sideId).ToList();
                        if (grp.Count == 0) continue;

                        // 上/下：按 Cx 升序；左/右：按 Cy 升序
                        if (sideId == 0 || sideId == 2)
                            grp = grp.OrderBy(e => e.Cx).ToList();
                        else
                            grp = grp.OrderBy(e => e.Cy).ToList();

                        // 沿该边均匀分布标签
                        //   上/下：x 均匀分布在 [picLeft + BOX_W_PT/2 .. picLeft + picWidth - BOX_W_PT/2]
                        //   左/右：y 均匀分布在 [picTop  + BOX_H_PT/2 .. picTop  + picHeight - BOX_H_PT/2]
                        int n = grp.Count;
                        for (int i = 0; i < n; i++)
                        {
                            var e = grp[i];
                            double bx, by;
                            double t = (n == 1) ? 0.5 : (i + 0.5) / n;   // 0..1 之间
                            if (sideId == 0)        // Top
                            {
                                bx = picLeft + 2 + t * (picWidth - 2 * BOX_W_PT - 4) + BOX_W_PT * 0;
                                bx = picLeft + t * (picWidth - BOX_W_PT);
                                by = picTop - BOX_H_PT - 4;   // 图片上方外侧
                            }
                            else if (sideId == 2)   // Bottom
                            {
                                bx = picLeft + t * (picWidth - BOX_W_PT);
                                by = picTop + picHeight + 4;  // 图片下方外侧
                            }
                            else if (sideId == 3)   // Left
                            {
                                bx = picLeft - BOX_W_PT - 4;  // 图片左方外侧
                                by = picTop + t * (picHeight - BOX_H_PT);
                            }
                            else                    // Right
                            {
                                bx = picLeft + picWidth + 4;
                                by = picTop + t * (picHeight - BOX_H_PT);
                            }
                            labelPos[e.Pt.Index] = Tuple.Create(bx, by);
                        }
                    }
                }

                int ordinal = 0;
                foreach (WeldAnnotationPoint pt in _points)
                {
                    if (!pt.InViewport) continue;
                    ordinal++;

                    // 焊点中心（Excel 点）
                    double cx = picLeft + pt.ScreenX * sx;
                    double cy = picTop + pt.ScreenY * sy;

                    // 编号框位置：从预计算表取
                    double bx, by;
                    if (labelPos.TryGetValue(pt.Index, out var bp))
                    { bx = bp.Item1; by = bp.Item2; }
                    else
                    { bx = cx + 32; by = cy - 32; }

                    dynamic dotShape = null;
                    dynamic boxShape = null;

                    // 查该点分类对应的 Excel 形状号 + 是否填充
                    string cat;
                    if (!_categories.TryGetValue(pt.Index, out cat) || string.IsNullOrEmpty(cat))
                        cat = DEFAULT_CATEGORY;
                    int shapeId = msoShapeOval;
                    if (style.CategoryShapes != null)
                        style.CategoryShapes.TryGetValue(cat, out shapeId);
                    bool filled = true;
                    if (style.CategoryFilled != null)
                        style.CategoryFilled.TryGetValue(cat, out filled);

                    // ── ① 焊点标记（分类形状；先创建，作为连接器起点） ──────
                    try
                    {
                        dotShape = ws.Shapes.AddShape(
                            shapeId,
                            cx - DOT_R_PT, cy - DOT_R_PT,
                            DOT_R_PT * 2.0, DOT_R_PT * 2.0);
                        if (filled)
                        {
                            // 实心：填充颜色 + 同色细边框
                            try { dotShape.Fill.ForeColor.RGB = oleDot; dotShape.Fill.Solid(); } catch { }
                            try { dotShape.Line.ForeColor.RGB = oleDot; dotShape.Line.Weight = 1.0; } catch { }
                        }
                        else
                        {
                            // 空心：无填充 + 粗描边
                            try { dotShape.Fill.Visible = msoFalse; } catch { }
                            try { dotShape.Line.ForeColor.RGB = oleDot; dotShape.Line.Weight = 1.75; } catch { }
                        }
                        try { dotShape.Placement = xlFreeFloating; } catch { }
                        try { dotShape.Name = "WPDot_" + pt.Index; shapeNamesForGroup.Add("WPDot_" + pt.Index); } catch { }
                        shapeCount++;
                    }
                    catch (Exception ex) { Log("WARN", "AddShape 失败: " + ex.Message); }

                    // ── ② 编号文本框（连接器终点） ──────────────────────────
                    try
                    {
                        // 按 label 长度估算文本框宽度：每个字符 ≈ 6.5pt，多留 6pt 边距
                        string label = GetAnnotationLabel(pt, ordinal);
                        double estW = Math.Max(BOX_W_PT, label.Length * 6.5 + 6.0);
                        boxShape = ws.Shapes.AddTextbox(
                            msoTextOrientationHoriz, bx, by, estW, BOX_H_PT);
                        try { boxShape.Fill.ForeColor.RGB = oleBoxFill; boxShape.Fill.Solid(); } catch { }
                        try { boxShape.Line.ForeColor.RGB = oleBoxBorder; boxShape.Line.Weight = 1.0; } catch { }
                        bool setOk = false;
                        try { boxShape.TextFrame.Characters().Text = label; setOk = true; } catch { }
                        if (!setOk) try { boxShape.TextFrame2.TextRange.Text = label; } catch { }
                        try { boxShape.TextFrame.HorizontalAlignment = xlHAlignCenter; } catch { }
                        try { boxShape.TextFrame.VerticalAlignment = xlVAlignCenter; } catch { }
                        try
                        {
                            boxShape.TextFrame.MarginLeft = 1.0;
                            boxShape.TextFrame.MarginRight = 1.0;
                            boxShape.TextFrame.MarginTop = 1.0;
                            boxShape.TextFrame.MarginBottom = 1.0;
                        }
                        catch { }
                        try
                        {
                            dynamic chars = boxShape.TextFrame.Characters();
                            chars.Font.Bold = true;
                            chars.Font.Size = 10;
                            chars.Font.Color = oleText;
                        }
                        catch { }
                        try { boxShape.Placement = xlFreeFloating; } catch { }
                        try { boxShape.Name = "WPBox_" + pt.Index; shapeNamesForGroup.Add("WPBox_" + pt.Index); } catch { }
                        shapeCount++;
                    }
                    catch (Exception ex) { Log("WARN", "AddTextbox 失败: " + ex.Message); }

                    // ── ③ 连接器（AddConnector + BeginConnect/EndConnect） ─────
                    //   AddConnector 比 AddLine 强：BeginConnect / EndConnect 把
                    //   两端绑定到 dot / textbox 的连接点，Excel 会自动重新布线，
                    //   即使用户拖动 dot 或 textbox，线始终跟随。
                    //   ConnectionSiteIndex=1 表示形状的首个连接点（一般是顶部），
                    //   之后调用 RerouteConnections 让 Excel 选择最近的连接点。
                    if (dotShape != null && boxShape != null)
                    {
                        try
                        {
                            dynamic conn = ws.Shapes.AddConnector(
                                msoConnectorStraight, cx, cy, bx, by);
                            try { conn.ConnectorFormat.BeginConnect(dotShape, 1); } catch (Exception ex) { Log("WARN", "BeginConnect: " + ex.Message); }
                            try { conn.ConnectorFormat.EndConnect(boxShape, 1); } catch (Exception ex) { Log("WARN", "EndConnect: " + ex.Message); }
                            try { conn.RerouteConnections(); } catch { }
                            try { conn.Line.ForeColor.RGB = oleLine; } catch { }
                            try { conn.Line.Weight = (double)style.LineWidth; } catch { }
                            try { conn.Placement = xlFreeFloating; } catch { }
                            try { conn.Name = "WPConn_" + pt.Index; shapeNamesForGroup.Add("WPConn_" + pt.Index); } catch { }
                            shapeCount++;
                        }
                        catch (Exception ex) { Log("WARN", "AddConnector 失败: " + ex.Message); }
                    }

                    processed++;
                    if (inVp > 0 && _progress != null)
                    {
                        try { _progress.Value = 55 + (int)(processed * 40.0 / inVp); } catch { }
                    }
                }

                // ── 组合所有 Shape 为一个整体 + 按选中单元格缩放 ─────────────
                //   1) 用 Shapes.Range(Array(...)).Group() 把图片 + 所有标注绑成单个 Shape；
                //   2) 若用户选了多个/合并单元格，整体平移 + 等比缩放到该区域；
                //   3) 若用户只选了一个单格，则仅平移到该单元格（不缩放）。
                //
                //   缩放对 Group 的所有子 Shape 自动等比例生效（包括坐标和尺寸），
                //   省去了我们自己重算每个 dot/textbox/connector 位置的麻烦。
                double finalPicWidth = (double)pic.Width;
                double finalPicHeight = (double)pic.Height;
                dynamic group = null;
                try
                {
                    if (shapeNamesForGroup.Count >= 2)
                    {
                        // Shapes.Range 接受 string 数组（VBA-style 名字数组）
                        object namesArr = shapeNamesForGroup.ToArray();
                        dynamic range = ws.Shapes.Range(namesArr);
                        group = range.Group();
                        try { group.Name = "WPGroup_" + DateTime.Now.ToString("HHmmss"); } catch { }
                        try { group.Placement = xlFreeFloating; } catch { }

                        // 平移：把 Group 左上角对到锚点
                        if (targetLeft >= 0 && targetTop >= 0)
                        {
                            try { group.Left = targetLeft; } catch { }
                            try { group.Top = targetTop; } catch { }
                        }

                        // 缩放：仅当用户选中了多格/合并格
                        if (scaleToTarget && targetWidth > 0 && targetHeight > 0)
                        {
                            double gw = (double)group.Width;
                            double gh = (double)group.Height;
                            double sX = targetWidth / gw;
                            double sY = targetHeight / gh;
                            // 保持纵横比，取较小的缩放比保证图整体放进去
                            double scale = Math.Min(sX, sY);
                            try { group.LockAspectRatio = true; } catch { }
                            try { group.Width = gw * scale; } catch { }
                            // Width 已锁纵横比，Height 会自动跟随；如未生效再补一刀
                            try
                            {
                                double curH = (double)group.Height;
                                if (Math.Abs(curH - gh * scale) > 1)
                                    group.Height = gh * scale;
                            }
                            catch { }
                        }
                        finalPicWidth = (double)group.Width;
                        finalPicHeight = (double)group.Height;
                        Log("INFO", string.Format(
                            "[Annotator] 已组合 {0} 个 Shape 为单组" +
                            (scaleToTarget ? "，已缩放至选中单元格" : "") +
                            "，位置 ({1:F0},{2:F0})",
                            shapeNamesForGroup.Count,
                            targetLeft >= 0 ? targetLeft : 0,
                            targetTop >= 0 ? targetTop : 0));
                    }
                }
                catch (Exception ex)
                {
                    Log("WARN", "[Annotator] Group/缩放失败: " + ex.Message + "（保留为独立 Shape）");
                }

                // 数据表的相对定位需要"实际"图像高度（缩放后的）
                picHeight = finalPicHeight;
                picWidth = finalPicWidth;
                // picLeft/picTop 用于数据表起始行号估算；若已平移到选中区则用新位置
                if (targetLeft >= 0) picLeft = targetLeft;
                if (targetTop >= 0) picTop = targetTop;

                // ── 可选：焊点数据表（图片下方） ────────────────────────────
                if (_chkWriteList.Checked && _points.Count > 0)
                {
                    int ds = (int)Math.Ceiling(picHeight / 15.0) + 3;
                    string[] hdr = { "#", "焊点名称", "操作名", "X(mm)", "Y(mm)", "Z(mm)", "视口内", "类型" };
                    for (int c = 0; c < hdr.Length; c++)
                    {
                        dynamic cell = ws.Cells[ds, c + 1];
                        cell.Value2 = hdr[c];
                        cell.Font.Bold = true;
                        cell.Interior.Color = D.ColorTranslator.ToOle(D.Color.FromArgb(68, 114, 196));
                        cell.Font.Color = D.ColorTranslator.ToOle(D.Color.White);
                    }
                    for (int i = 0; i < _points.Count; i++)
                    {
                        WeldAnnotationPoint pt = _points[i];
                        int row = ds + 1 + i;
                        string catRow;
                        if (!_categories.TryGetValue(pt.Index, out catRow) || string.IsNullOrEmpty(catRow))
                            catRow = DEFAULT_CATEGORY;
                        ((dynamic)ws.Cells[row, 1]).Value2 = pt.Index;
                        ((dynamic)ws.Cells[row, 2]).Value2 = pt.Name;
                        ((dynamic)ws.Cells[row, 3]).Value2 = pt.OpName;
                        ((dynamic)ws.Cells[row, 4]).Value2 = pt.X;
                        ((dynamic)ws.Cells[row, 5]).Value2 = pt.Y;
                        ((dynamic)ws.Cells[row, 6]).Value2 = pt.Z;
                        ((dynamic)ws.Cells[row, 7]).Value2 = pt.InViewport ? "是" : "否";
                        ((dynamic)ws.Cells[row, 8]).Value2 = catRow;
                        if (i % 2 == 1)
                        {
                            dynamic rng = ws.Range[ws.Cells[row, 1], ws.Cells[row, 8]];
                            rng.Interior.Color =
                                D.ColorTranslator.ToOle(D.Color.FromArgb(242, 242, 242));
                        }
                    }
                    dynamic dr2 = ws.Range[ws.Cells[ds, 1], ws.Cells[ds + _points.Count, 8]];
                    dr2.Columns.AutoFit();
                }

                ws.Activate();
                xlApp.Visible = true;
                Log("INFO", string.Format(
                    "[Annotator] Excel 写入完成：{0} 个 Shape，{1} 个视口内点；{2}",
                    shapeCount, inVp,
                    _chkWriteList.Checked ? "附焊点数据表" : "未附数据表"));
            }
            finally
            {
                if (xlApp != null)
                {
                    try { xlApp.ScreenUpdating = savedScreenUpdating; } catch { }
                }
                try { File.Delete(tmp); } catch { }
            }
        }

        // =========================================================================
        // 辅助
        // =========================================================================

        private void SetStatus(string msg)
        {
            _lblStatus.Text = msg;
            _status.Refresh();
        }

        private void Log(string level, string msg)
        {
            if (_logBox == null) return;
            D.Color c = level == "ERR" ? D.Color.Tomato :
                        level == "WARN" ? D.Color.Yellow : D.Color.LightGray;
            _logBox.SelectionColor = c;
            _logBox.AppendText(string.Format("[{0}] {1}: {2}\r\n",
                DateTime.Now.ToString("HH:mm:ss"), level, msg));
            _logBox.ScrollToCaret();
            if (level != "INFO" && !_logVisible)
            {
                _logPanel.Visible = _logVisible = true;
            }
        }
    }

    // =========================================================================
    // 标注样式设置窗口
    // =========================================================================
    public class AnnotationStyleForm : WF.Form
    {
        public AnnotationStyle Result { get; private set; }

        // msoAutoShapeType 常用形状（显示名 → 形状ID）
        private static readonly Tuple<string, int>[] SHAPE_OPTIONS = new[]
        {
            Tuple.Create("圆形",       9),   // msoShapeOval
            Tuple.Create("矩形",       1),   // msoShapeRectangle
            Tuple.Create("三角形",     7),   // msoShapeIsoscelesTriangle
            Tuple.Create("菱形",       4),   // msoShapeDiamond
            Tuple.Create("圆角矩形",   5),   // msoShapeRoundedRectangle
            Tuple.Create("五角星",    12),   // msoShape5pointStar
        };

        public AnnotationStyleForm(AnnotationStyle src)
        {
            // 深拷贝 CategoryShapes，否则直接赋引用会在用户点"取消"后仍然修改外部
            var catCopy = new System.Collections.Generic.Dictionary<string, int>();
            if (src.CategoryShapes != null)
                foreach (var kv in src.CategoryShapes) catCopy[kv.Key] = kv.Value;

            Result = new AnnotationStyle
            {
                DotColor = src.DotColor,
                DotRadius = src.DotRadius,
                LineColor = src.LineColor,
                LineWidth = src.LineWidth,
                BoxBorderColor = src.BoxBorderColor,
                BoxFillColor = src.BoxFillColor,
                TextColor = src.TextColor,
                TextFont = src.TextFont,
                BoxPadding = src.BoxPadding,
                OffsetX = src.OffsetX,
                OffsetY = src.OffsetY,
                CategoryShapes = catCopy
            };
            Text = "标注样式设置";
            FormBorderStyle = WF.FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = WF.FormStartPosition.CenterParent;
            Font = new D.Font("Microsoft YaHei UI", 9f);
            // AutoSize：窗口根据控件内容自动撑开，不再被固定 Size 截断
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            MinimumSize = new D.Size(320, 0);
            Build();
        }

        private void Build()
        {
            TableLayoutPanel t = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(8, 6, 8, 6)
            };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            void Row(string lbl, WF.Control ctrl)
            {
                t.Controls.Add(new WF.Label
                {
                    Text = lbl,
                    AutoSize = true,
                    Anchor = AnchorStyles.Left,
                    Margin = new Padding(0, 6, 8, 6)
                });
                ctrl.Margin = new Padding(0, 3, 0, 3);
                t.Controls.Add(ctrl);
            }

            // 添加分组标题行（跨两列）
            void Section(string title)
            {
                WF.Label header = new WF.Label
                {
                    Text = title,
                    AutoSize = true,
                    Font = new D.Font("Microsoft YaHei UI", 9.5f, D.FontStyle.Bold),
                    ForeColor = D.Color.FromArgb(0, 70, 127),
                    Margin = new Padding(0, 10, 0, 4)
                };
                int rowIdx = t.RowCount;
                t.SetColumnSpan(header, 2);
                t.Controls.Add(header, 0, rowIdx);
                t.RowCount = rowIdx + 1;
            }

            WF.Button Clr(D.Color init, Action<D.Color> set)
            {
                WF.Button b = new WF.Button { Width = 80, Height = 22, BackColor = init };
                b.Click += (s, e) =>
                {
                    using (WF.ColorDialog cd = new WF.ColorDialog { Color = b.BackColor })
                        if (cd.ShowDialog() == DialogResult.OK)
                        { b.BackColor = cd.Color; set(cd.Color); }
                };
                return b;
            }

            NumericUpDown Nud(decimal v, decimal mn, decimal mx, decimal inc, int dp, Action<decimal> fn)
            {
                NumericUpDown n = new NumericUpDown
                {
                    Width = 75,
                    Minimum = mn,
                    Maximum = mx,
                    Value = v,
                    Increment = inc,
                    DecimalPlaces = dp
                };
                n.ValueChanged += (s, e) => fn(n.Value);
                return n;
            }

            // ── 基础样式 ──────────────────────────────────────────────────
            Section("基础样式");
            Row("圆点颜色：", Clr(Result.DotColor, c => Result.DotColor = c));
            Row("圆点半径(px)：", Nud(Result.DotRadius, 2, 30, 1, 0, v => Result.DotRadius = (int)v));
            Row("引线颜色：", Clr(Result.LineColor, c => Result.LineColor = c));
            Row("引线宽度(px)：", Nud((decimal)Result.LineWidth, 1, 10, 0.5m, 1, v => Result.LineWidth = (float)v));
            Row("框边框颜色：", Clr(Result.BoxBorderColor, c => Result.BoxBorderColor = c));
            Row("框填充颜色：", Clr(Result.BoxFillColor, c => Result.BoxFillColor = c));
            Row("编号颜色：", Clr(Result.TextColor, c => Result.TextColor = c));
            Row("框内边距(px)：", Nud(Result.BoxPadding, 1, 20, 1, 0, v => Result.BoxPadding = (int)v));
            Row("水平偏移(px)：", Nud(Result.OffsetX, 10, 300, 5, 0, v => Result.OffsetX = (int)v));
            Row("垂直偏移(px)：", Nud(Result.OffsetY, 10, 300, 5, 0, v => Result.OffsetY = (int)v));

            // ── 分类形状 ──────────────────────────────────────────────────
            Section("分类形状（Excel 标记形状）");
            foreach (string cat in new[] { "焊点", "补焊点", "强度校验点" })
            {
                string catLocal = cat;   // 闭包捕获
                WF.ComboBox cmb = new WF.ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Width = 110
                };
                foreach (var opt in SHAPE_OPTIONS) cmb.Items.Add(opt.Item1);

                int currentId;
                if (!Result.CategoryShapes.TryGetValue(catLocal, out currentId))
                    currentId = 9;   // 默认圆
                int selIdx = 0;
                for (int i = 0; i < SHAPE_OPTIONS.Length; i++)
                    if (SHAPE_OPTIONS[i].Item2 == currentId) { selIdx = i; break; }
                cmb.SelectedIndex = selIdx;

                cmb.SelectedIndexChanged += (s, e) =>
                {
                    int idx = cmb.SelectedIndex;
                    if (idx >= 0 && idx < SHAPE_OPTIONS.Length)
                        Result.CategoryShapes[catLocal] = SHAPE_OPTIONS[idx].Item2;
                };
                Row(catLocal + "：", cmb);
            }

            Controls.Add(t);

            // ── 按钮条 ────────────────────────────────────────────────────
            WF.Panel bp = new WF.Panel { Dock = DockStyle.Bottom, Height = 42 };
            FormUiKit.FlatColorButton ok = new FormUiKit.FlatColorButton
            {
                Text = "确定", Width = 80, Height = 26, DialogResult = DialogResult.OK, Anchor = AnchorStyles.Right,
                BgColor = D.Color.FromArgb(0, 100, 167), ForeColor = D.Color.White,
                BorderColor = D.Color.FromArgb(0, 100, 167), FlatStyle = FlatStyle.Flat
            };
            FormUiKit.FlatColorButton can = new FormUiKit.FlatColorButton
            {
                Text = "取消", Width = 80, Height = 26, DialogResult = DialogResult.Cancel, Anchor = AnchorStyles.Right,
                BgColor = D.Color.FromArgb(120, 124, 135), ForeColor = D.Color.White,
                BorderColor = D.Color.FromArgb(120, 124, 135), FlatStyle = FlatStyle.Flat
            };
            bp.Resize += (s, e) =>
            {
                can.Left = bp.ClientSize.Width - can.Width - 12;
                can.Top = 8;
                ok.Left = can.Left - ok.Width - 8;
                ok.Top = 8;
            };
            ok.Click += (s, e) => Close();
            can.Click += (s, e) => Close();
            bp.Controls.Add(ok);
            bp.Controls.Add(can);
            Controls.Add(bp);
            AcceptButton = ok;
            CancelButton = can;
        }
    }
}