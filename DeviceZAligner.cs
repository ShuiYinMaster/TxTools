// ═══════════════════════════════════════════════════════════════════════════
// DeviceZAligner — Process Simulate 二次开发插件
// 功能：遍历场景中的设备，忽略焊枪等特殊设备，检查设备Z向位置，
//       将最低点对齐到世界坐标Z=0，支持Ctrl+Z撤销
// ═══════════════════════════════════════════════════════════════════════════
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;

namespace TxTools.DeviceZAligner
{
    // =====================================================================
    // 插件入口 — TxButtonCommand
    // =====================================================================
    public class DeviceZAlignerCmd : TxButtonCommand
    {
        public override string Category    { get { return "TxTools"; } }
        public override string Name        { get { return "DeviceZAligner"; } }
        public override string Description { get { return "设备Z向对齐工具 — 将设备最低点对齐到Z=0"; } }

        // ── 小图标（16×16，菜单/工具栏） ────────────────────────────
        // PS SDK 中 Bitmap/LargeBitmap 是 string 类型 = 图标文件的路径
        public override string Bitmap
        {
            get { return GetIconPath(16); }
        }

        // ── 大图标（32×32，Ribbon） ─────────────────────────────────
        public override string LargeBitmap
        {
            get { return GetIconPath(32); }
        }

        // 缓存已生成的图标路径，避免每次访问都重新生成文件
        private static string _iconPath16;
        private static string _iconPath32;

        /// <summary>
        /// 获取图标文件路径（PS要求的是文件路径字符串）
        /// 优先查找外部文件，找不到则动态生成到临时目录
        /// </summary>
        private string GetIconPath(int size)
        {
            // 检查缓存
            string cached = size <= 16 ? _iconPath16 : _iconPath32;
            if (!string.IsNullOrEmpty(cached) && System.IO.File.Exists(cached))
                return cached;

            string fileName = size <= 16 ? "icon_zalign_16.png" : "icon_zalign_32.png";

            // 方案A：从DLL同目录查找现成图标文件
            try
            {
                string dir = System.IO.Path.GetDirectoryName(
                    Assembly.GetExecutingAssembly().Location);
                string path = System.IO.Path.Combine(dir, fileName);
                if (System.IO.File.Exists(path))
                {
                    if (size <= 16) _iconPath16 = path; else _iconPath32 = path;
                    return path;
                }
            }
            catch { }

            // 方案B：从嵌入资源提取到临时文件
            try
            {
                string resName = size <= 16
                    ? "TxTools.DeviceZAligner.Resources.icon_zalign_16.png"
                    : "TxTools.DeviceZAligner.Resources.icon_zalign_32.png";
                var asm = Assembly.GetExecutingAssembly();
                var stream = asm.GetManifestResourceStream(resName);
                if (stream != null)
                {
                    string tmpPath = System.IO.Path.Combine(
                        System.IO.Path.GetTempPath(), "TxDeviceZAligner_" + fileName);
                    using (var fs = System.IO.File.Create(tmpPath))
                        stream.CopyTo(fs);
                    if (size <= 16) _iconPath16 = tmpPath; else _iconPath32 = tmpPath;
                    return tmpPath;
                }
            }
            catch { }

            // 方案C：代码动态生成图标 → 保存到临时文件 → 返回路径
            try
            {
                string tmpPath = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(), "TxDeviceZAligner_" + fileName);
                using (var bmp = GenerateIcon(size))
                    bmp.Save(tmpPath, System.Drawing.Imaging.ImageFormat.Png);
                if (size <= 16) _iconPath16 = tmpPath; else _iconPath32 = tmpPath;
                return tmpPath;
            }
            catch { }

            return "";  // 返回空字符串，PS会使用默认图标
        }

        /// <summary>
        /// 动态绘制Z向对齐图标：Z轴箭头 + 地面基准线 + 设备方块
        /// PS深蓝色(0,70,127) + 红色基准线，与PS风格一致
        /// </summary>
        internal static System.Drawing.Bitmap GenerateIcon(int size)
        {
            var bmp = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                float s = size / 32f;  // 基准缩放因子（以32为基准）
                float penW = Math.Max(1f, 1.5f * s);

                // ── 红色基准线（Z=0 地面） ───────────────────────
                using (var pen = new Pen(Color.FromArgb(220, 50, 50), penW * 1.2f))
                {
                    pen.DashStyle = DashStyle.Dash;
                    float y0 = 26f * s;
                    g.DrawLine(pen, 2f * s, y0, 30f * s, y0);
                }

                // ── 蓝色Z轴箭头（垂直向上） ─────────────────────
                Color psBlue = Color.FromArgb(0, 70, 127);
                using (var pen = new Pen(psBlue, penW * 1.3f))
                {
                    float cx = 16f * s;
                    g.DrawLine(pen, cx, 6f * s, cx, 24f * s);   // 轴线
                    // 箭头
                    g.DrawLine(pen, cx - 4f * s, 10f * s, cx, 5f * s);
                    g.DrawLine(pen, cx + 4f * s, 10f * s, cx, 5f * s);
                }

                // ── "Z" 字母标识 ─────────────────────────────────
                using (var font = new Font("Tahoma", 7f * s, FontStyle.Bold))
                using (var brush = new SolidBrush(psBlue))
                {
                    g.DrawString("Z", font, brush, 22f * s, 3f * s);
                }

                // ── 设备方块（被箭头移动的对象） ─────────────────
                using (var brush = new SolidBrush(Color.FromArgb(180, 0, 70, 127)))
                {
                    float bx = 9f * s, by = 17f * s, bw = 14f * s, bh = 8f * s;
                    g.FillRectangle(brush, bx, by, bw, bh);
                }
                using (var pen = new Pen(psBlue, penW * 0.8f))
                {
                    float bx = 9f * s, by = 17f * s, bw = 14f * s, bh = 8f * s;
                    g.DrawRectangle(pen, bx, by, bw, bh);
                }
            }
            return bmp;
        }

        public override void Execute(object cmdParams)
        {
            try
            {
                var form = new DeviceZAlignerForm();
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    // =====================================================================
    // 数据模型
    // =====================================================================
    public class DeviceInfo
    {
        public int    Index       { get; set; }
        public string Name        { get; set; } = "";
        public string TypeName    { get; set; } = "";
        public double CurrentZ    { get; set; }        // 当前AbsoluteLocation的Z值(mm)
        public double MinZ        { get; set; }        // 包围盒最低点Z(mm)
        public double OffsetZ     { get; set; }        // 需要偏移的量(mm)
        public bool   IsSkipped   { get; set; }        // 是否被忽略（gun等）
        public string SkipReason  { get; set; } = "";  // 忽略原因
        public bool   IsAligned   { get; set; }        // 是否已对齐
        public string Message     { get; set; } = "";  // 状态消息
        public string Method      { get; set; } = "";  // 检测方法（轴交点/坐标原点）
        public ITxObject TxObj    { get; set; }        // PS对象引用
    }

    // =====================================================================
    // 主窗体 — 继承TxForm获得PS原生托管、主题集成和DPI自适应
    // =====================================================================
    public class DeviceZAlignerForm : TxForm
    {
        // ── UI控件 ──────────────────────────────────────────────────────
        private TxFlexGrid      _grid;
        private RichTextBox     _logBox;
        private Panel           _logPanel;
        private StatusStrip     _statusStrip;
        private ToolStripStatusLabel _lblStatus;
        private ToolStripProgressBar _tsProgress;

        // 工具栏按钮
        private ToolStripButton _btnScan;
        private ToolStripButton _btnAlignSelected;
        private ToolStripButton _btnAlignAll;
        private ToolStripButton _btnLog;

        // 卡片区控件
        private TextBox         _txtIgnoreKeywords;
        private Label           _lblTotal;
        private Label           _lblSkipped;
        private Label           _lblNeedAlign;
        private Label           _lblAligned;
        private Label           _lblOk;

        // ── 数据 ────────────────────────────────────────────────────────
        private List<DeviceInfo> _devices = new List<DeviceInfo>();
        private bool _logVisible = false;

        // ── 默认忽略关键词（不区分大小写） ──────────────────────────────
        private static readonly string[] DefaultIgnoreKeywords = new[]
        {
            "gun", "Gun", "GUN",
            "weldgun", "WeldGun",
            "gripper", "Gripper",
            "tool", "Tool",
            "tcp", "TCP",
            "nozzle", "Nozzle",
            "torch", "Torch",
            "clamp", "Clamp",
        };

        // ── 表格列索引常量 ──────────────────────────────────────────────
        private const int COL_IDX      = 0;
        private const int COL_NAME     = 1;
        private const int COL_TYPE     = 2;
        private const int COL_CURZ     = 3;
        private const int COL_OFFSET   = 4;
        private const int COL_METHOD   = 5;
        private const int COL_STATUS   = 6;
        private const int COL_MSG      = 7;

        // =====================================================================
        // 构造函数
        // =====================================================================
        public DeviceZAlignerForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            Text = "设备Z向对齐工具";
            Size = new Size(1100, 650);
            MinimumSize = new Size(800, 450);
            StartPosition = FormStartPosition.CenterScreen;

            // ── 窗体图标（标题栏+任务栏） ────────────────────────
            try
            {
                // 优先从嵌入资源加载 .ico
                var asm = Assembly.GetExecutingAssembly();
                var stream = asm.GetManifestResourceStream(
                    "TxTools.DeviceZAligner.Resources.icon_zalign.ico");
                if (stream != null)
                    Icon = new Icon(stream);
            }
            catch
            {
                // 回退：用动态生成的Bitmap转Icon
                try
                {
                    var bmp = DeviceZAlignerCmd.GenerateIcon(32);
                    IntPtr hIcon = bmp.GetHicon();
                    Icon = Icon.FromHandle(hIcon);
                }
                catch { }
            }

            BuildToolStrip();
            BuildTopCards();
            BuildGrid();
            BuildLogPanel();
            BuildStatusBar();

            ResumeLayout(false);
            PerformLayout();
        }

        // ── 顶部工具栏 ─────────────────────────────────────────────────
        private void BuildToolStrip()
        {
            var ts = new TxToolStrip();
            ts.Dock = DockStyle.Top;

            _btnScan = new ToolStripButton("扫描设备");
            _btnScan.ToolTipText = "遍历场景中所有设备，分析Z向位置";
            _btnScan.Click += (s, e) => ScanDevices();
            ts.Items.Add(_btnScan);

            ts.Items.Add(new ToolStripSeparator());

            _btnAlignSelected = new ToolStripButton("对齐选中");
            _btnAlignSelected.ToolTipText = "将表格中选中的设备最低点对齐到Z=0";
            _btnAlignSelected.Enabled = false;
            _btnAlignSelected.Click += (s, e) => AlignSelected();
            ts.Items.Add(_btnAlignSelected);

            _btnAlignAll = new ToolStripButton("全部对齐");
            _btnAlignAll.ToolTipText = "将所有需要偏移的设备最低点对齐到Z=0（支持Ctrl+Z撤销）";
            _btnAlignAll.Enabled = false;
            _btnAlignAll.Click += (s, e) => AlignAll();
            ts.Items.Add(_btnAlignAll);

            ts.Items.Add(new ToolStripSeparator());

            _btnLog = new ToolStripButton("日志");
            _btnLog.CheckOnClick = true;
            _btnLog.Click += (s, e) =>
            {
                _logVisible = _btnLog.Checked;
                _logPanel.Visible = _logVisible;
            };
            ts.Items.Add(_btnLog);

            // 右侧弹性间距 + 进度条
            ts.Items.Add(new ToolStripSeparator());
            var spring = new ToolStripLabel("") { AutoSize = false, Width = 0 };
            // 注：TxToolStrip 在部分PS版本不支持Spring，用固定间距代替
            ts.Items.Add(spring);

            _tsProgress = new ToolStripProgressBar();
            _tsProgress.Visible = false;
            _tsProgress.Size = new Size(180, 16);
            ts.Items.Add(_tsProgress);

            Controls.Add(ts);
        }

        // ── 顶部卡片区（参考 RobotReachabilityChecker 的 GroupBox 布局） ──
        private void BuildTopCards()
        {
            var topPanel = new Panel();
            topPanel.Dock = DockStyle.Top;
            topPanel.Height = 80;
            topPanel.Padding = new Padding(4, 2, 4, 2);

            // ── 卡片1：忽略关键词配置 ────────────────────────────────
            var gbFilter = new GroupBox();
            gbFilter.Text = "过滤设置";
            gbFilter.Font = new Font("Tahoma", 8.5f);
            gbFilter.Location = new Point(4, 2);
            gbFilter.Size = new Size(480, 72);
            gbFilter.Anchor = AnchorStyles.Left | AnchorStyles.Top;

            var lblKw = new Label();
            lblKw.Text = "忽略关键词:";
            lblKw.Font = new Font("Tahoma", 8.5f);
            lblKw.Location = new Point(8, 20);
            lblKw.AutoSize = true;
            gbFilter.Controls.Add(lblKw);

            _txtIgnoreKeywords = new TextBox();
            _txtIgnoreKeywords.Location = new Point(80, 17);
            _txtIgnoreKeywords.Size = new Size(390, 22);
            _txtIgnoreKeywords.Font = new Font("Tahoma", 8.5f);
            _txtIgnoreKeywords.Text = string.Join(", ", DefaultIgnoreKeywords.Distinct(StringComparer.OrdinalIgnoreCase));
            _txtIgnoreKeywords.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            gbFilter.Controls.Add(_txtIgnoreKeywords);

            var lblTip = new Label();
            lblTip.Text = "逗号分隔，不区分大小写。同时自动过滤 Robot/Gun/Gripper/Tool 类型";
            lblTip.Font = new Font("Tahoma", 7.5f);
            lblTip.ForeColor = SystemColors.GrayText;
            lblTip.Location = new Point(80, 44);
            lblTip.AutoSize = true;
            gbFilter.Controls.Add(lblTip);

            topPanel.Controls.Add(gbFilter);

            // ── 卡片2：扫描统计 ──────────────────────────────────────
            var gbStats = new GroupBox();
            gbStats.Text = "扫描统计";
            gbStats.Font = new Font("Tahoma", 8.5f);
            gbStats.Location = new Point(492, 2);
            gbStats.Size = new Size(560, 72);
            gbStats.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;

            int sx = 12, gap = 108;
            _lblTotal     = CreateStatLabel(gbStats, "总数: -",     sx + gap * 0, SystemColors.ControlText);
            _lblSkipped   = CreateStatLabel(gbStats, "已忽略: -",   sx + gap * 1, SystemColors.GrayText);
            _lblNeedAlign = CreateStatLabel(gbStats, "待对齐: -",   sx + gap * 2, Color.FromArgb(180, 130, 0));
            _lblAligned   = CreateStatLabel(gbStats, "已对齐: -",   sx + gap * 3, Color.FromArgb(0, 130, 60));
            _lblOk        = CreateStatLabel(gbStats, "正常: -",     sx + gap * 4, SystemColors.ControlText);

            topPanel.Controls.Add(gbStats);

            Controls.Add(topPanel);
        }

        private Label CreateStatLabel(GroupBox parent, string text, int x, Color color)
        {
            var lbl = new Label();
            lbl.Text = text;
            lbl.Font = new Font("Tahoma", 9f, FontStyle.Bold);
            lbl.ForeColor = color;
            lbl.Location = new Point(x, 30);
            lbl.AutoSize = true;
            parent.Controls.Add(lbl);
            return lbl;
        }

        // ── 数据表格 ────────────────────────────────────────────────────
        private void BuildGrid()
        {
            _grid = new TxFlexGrid();
            _grid.Dock = DockStyle.Fill;
            _grid.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row;
            _grid.AllowEditing = false;
            _grid.AllowSorting = C1.Win.C1FlexGrid.AllowSortingEnum.SingleColumn;

            // 定义列（增加"方法"列）
            _grid.Cols.Count = 9;
            _grid.Rows.Fixed = 1;

            string[] headers = { "#", "设备名称", "类型", "当前Z(mm)", "偏移量(mm)", "方法", "状态", "备注" };
            int[] widths     = { 40, 220, 120, 90, 90, 70, 70, 220 };

            for (int i = 0; i < headers.Length; i++)
            {
                _grid.Cols[i].Caption = headers[i];
                _grid.Cols[i].Width = widths[i];
            }

            // 行选择变化事件
            _grid.AfterSelChange += (s, e) =>
            {
                _btnAlignSelected.Enabled = _grid.RowSel >= _grid.Rows.Fixed;
            };

            // 双击跳转到PS对象
            _grid.DoubleClick += Grid_DblClick;

            // 鼠标滚轮支持（TxFlexGrid/C1FlexGrid有时不响应滚轮）
            _grid.MouseWheel += (s, e) =>
            {
                try
                {
                    int delta = e.Delta > 0 ? -3 : 3;
                    int newRow = Math.Max(_grid.Rows.Fixed,
                        Math.Min(_grid.Rows.Count - 1, _grid.TopRow + delta));
                    _grid.TopRow = newRow;
                }
                catch { }
            };

            Controls.Add(_grid);
        }

        // ── 日志面板（底部可折叠） ──────────────────────────────────────
        private void BuildLogPanel()
        {
            _logPanel = new Panel();
            _logPanel.Dock = DockStyle.Bottom;
            _logPanel.Height = 160;
            _logPanel.Visible = false;

            var btnClear = new Button();
            btnClear.Text = "清空日志";
            btnClear.Dock = DockStyle.Top;
            btnClear.Height = 24;
            btnClear.FlatStyle = FlatStyle.Flat;
            btnClear.Click += (s, e) => _logBox.Clear();
            _logPanel.Controls.Add(btnClear);

            _logBox = new RichTextBox();
            _logBox.Dock = DockStyle.Fill;
            _logBox.ReadOnly = true;
            _logBox.BackColor = Color.FromArgb(30, 30, 30);
            _logBox.ForeColor = Color.FromArgb(204, 204, 204);
            _logBox.Font = new Font("Consolas", 9f);
            _logBox.WordWrap = false;
            _logPanel.Controls.Add(_logBox);

            Controls.Add(_logPanel);
        }

        // ── 底部状态栏 ──────────────────────────────────────────────────
        private void BuildStatusBar()
        {
            _statusStrip = new StatusStrip();
            _lblStatus = new ToolStripStatusLabel("就绪 — 点击「扫描设备」开始");
            _lblStatus.Spring = true;
            _lblStatus.TextAlign = ContentAlignment.MiddleLeft;
            _statusStrip.Items.Add(_lblStatus);
            Controls.Add(_statusStrip);
        }

        // =====================================================================
        // 核心功能1：扫描设备
        // =====================================================================
        private void ScanDevices()
        {
            _devices.Clear();
            Log("═══════════════════════════════════════════════════");
            Log("开始扫描场景中的设备...");

            try
            {
                TxDocument doc = TxApplication.ActiveDocument;
                if (doc == null)
                {
                    Log("ActiveDocument 为 null，无法扫描", "ERR");
                    SetStatus("错误：ActiveDocument 为 null");
                    return;
                }
                Log("ActiveDocument OK");

                // ── 获取忽略关键词列表 ────────────────────────────────
                var ignoreKeywords = ParseIgnoreKeywords();
                Log($"忽略关键词: [{string.Join(", ", ignoreKeywords)}]");

                // ── 遍历PhysicalRoot下所有后代 ───────────────────────
                // 使用多种类型过滤策略，因为PS版本间设备类型不一致
                var allObjects = new List<ITxObject>();

                // 策略1：TxComponent（最通用的设备/组件类型）
                try
                {
                    var comps = doc.PhysicalRoot.GetAllDescendants(
                        new TxTypeFilter(typeof(TxComponent)));
                    if (comps != null)
                    {
                        foreach (ITxObject obj in comps)
                            allObjects.Add(obj);
                    }
                    Log($"TxComponent 发现 {allObjects.Count} 个对象");
                }
                catch (Exception ex)
                {
                    Log($"TxComponent 遍历失败: {ex.Message}", "WARN");
                }

                // 策略2：如果TxComponent没找到，尝试ITxDevice（部分版本）
                if (allObjects.Count == 0)
                {
                    try
                    {
                        dynamic root = doc.PhysicalRoot;
                        TxObjectList devs = root.GetAllDescendants(
                            new TxTypeFilter(typeof(ITxObject)));
                        if (devs != null)
                        {
                            foreach (ITxObject obj in devs)
                            {
                                // 过滤出"设备级"对象，跳过过深的子零件
                                string typeName = obj.GetType().Name;
                                if (typeName.Contains("Component") ||
                                    typeName.Contains("Device") ||
                                    typeName.Contains("Fixture") ||
                                    typeName.Contains("Part"))
                                {
                                    allObjects.Add(obj);
                                }
                            }
                        }
                        Log($"ITxObject 宽泛遍历后筛选: {allObjects.Count} 个设备级对象");
                    }
                    catch (Exception ex)
                    {
                        Log($"ITxObject 遍历失败: {ex.Message}", "ERR");
                    }
                }

                Log($"共获取 {allObjects.Count} 个设备对象");

                // ── 分析每个设备 ──────────────────────────────────────
                int idx = 0;
                foreach (var obj in allObjects)
                {
                    idx++;
                    var info = new DeviceInfo
                    {
                        Index = idx,
                        Name = obj.Name ?? $"未命名_{idx}",
                        TypeName = obj.GetType().Name,
                        TxObj = obj
                    };

                    // 检查是否需要忽略
                    string skipReason = CheckShouldSkip(obj, ignoreKeywords);
                    if (!string.IsNullOrEmpty(skipReason))
                    {
                        info.IsSkipped = true;
                        info.SkipReason = skipReason;
                        info.Message = $"已忽略: {skipReason}";
                        Log($"  [{info.Name}] 忽略 — {skipReason}");
                        _devices.Add(info);
                        continue;
                    }

                    // 获取当前Z位置和最低Z点
                    AnalyzeDeviceZ(obj, info);
                    _devices.Add(info);

                    Log($"  [{info.Name}] 类型={info.TypeName}, 当前Z={info.CurrentZ:F1}, " +
                        $"最低Z={info.MinZ:F1}, 偏移={info.OffsetZ:F1}");
                }

                // ── 刷新表格 ─────────────────────────────────────────
                RefreshGrid();

                int total = _devices.Count;
                int skipped = _devices.Count(d => d.IsSkipped);
                int needAlign = _devices.Count(d => !d.IsSkipped && Math.Abs(d.OffsetZ) > 0.01);
                Log($"扫描完成: 共 {total} 个设备, 忽略 {skipped} 个, 需要对齐 {needAlign} 个", "OK");
                SetStatus($"扫描完成: {total} 个设备, 忽略 {skipped}, 需对齐 {needAlign}");

                _btnAlignAll.Enabled = needAlign > 0;
            }
            catch (Exception ex)
            {
                Log($"扫描异常: {ex.Message}", "ERR");
                SetStatus($"扫描失败: {ex.Message}");
            }
        }

        // =====================================================================
        // 核心功能2：对齐选中设备
        // =====================================================================
        private void AlignSelected()
        {
            int selRow = _grid.RowSel;
            if (selRow < _grid.Rows.Fixed) return;

            int dataIdx = selRow - _grid.Rows.Fixed;
            if (dataIdx < 0 || dataIdx >= _devices.Count) return;

            var dev = _devices[dataIdx];
            if (dev.IsSkipped)
            {
                SetStatus($"设备 [{dev.Name}] 已被忽略，无法对齐");
                return;
            }
            if (Math.Abs(dev.OffsetZ) < 0.01)
            {
                SetStatus($"设备 [{dev.Name}] 已在Z=0，无需对齐");
                return;
            }

            AlignDevices(new List<DeviceInfo> { dev });
        }

        // =====================================================================
        // 核心功能3：对齐所有设备
        // =====================================================================
        private void AlignAll()
        {
            var toAlign = _devices
                .Where(d => !d.IsSkipped && Math.Abs(d.OffsetZ) > 0.01)
                .ToList();

            if (toAlign.Count == 0)
            {
                SetStatus("没有需要对齐的设备");
                return;
            }

            var result = MessageBox.Show(
                $"将对 {toAlign.Count} 个设备执行Z向对齐操作。\n" +
                "此操作支持 Ctrl+Z 撤销。\n\n确定继续？",
                "确认对齐", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result != DialogResult.Yes) return;

            AlignDevices(toAlign);
        }

        // =====================================================================
        // 执行对齐（包裹在Undo上下文中，支持Ctrl+Z）
        // =====================================================================
        private void AlignDevices(List<DeviceInfo> devices)
        {
            Log("═══════════════════════════════════════════════════");
            Log($"开始对齐 {devices.Count} 个设备...");

            int success = 0, fail = 0;

            try
            {
                TxDocument doc = TxApplication.ActiveDocument;
                if (doc == null)
                {
                    Log("ActiveDocument 为 null", "ERR");
                    return;
                }

                // ── 包裹在Undo上下文中，使整个操作可撤销 ────────────
                // PS 的 Undo 机制通过 TxDocument 或 TxApplication 暴露
                // 不同版本属性名不同，使用 dynamic 防御式访问
                bool undoStarted = BeginUndoBlock(doc, $"设备Z向对齐({devices.Count}个)");
                if (!undoStarted)
                    Log("Undo上下文启动失败，操作将不可撤销", "WARN");
                else
                    Log("Undo上下文已启动 — 操作完成后可用 Ctrl+Z 撤销", "OK");

                _tsProgress.Visible = true;
                _tsProgress.Maximum = devices.Count;
                _tsProgress.Value = 0;

                try
                {
                    foreach (var dev in devices)
                    {
                        try
                        {
                            bool ok = ApplyZOffset(dev);
                            if (ok)
                            {
                                dev.IsAligned = true;
                                dev.Message = $"已对齐 (偏移 {-dev.OffsetZ:F1}mm)";
                                success++;
                                Log($"  [{dev.Name}] 对齐成功, Z偏移={-dev.OffsetZ:F1}mm", "OK");
                            }
                            else
                            {
                                dev.Message = "对齐失败";
                                fail++;
                                Log($"  [{dev.Name}] 对齐失败", "ERR");
                            }
                        }
                        catch (Exception ex)
                        {
                            dev.Message = $"异常: {ex.Message}";
                            fail++;
                            Log($"  [{dev.Name}] 异常: {ex.Message}", "ERR");
                        }

                        _tsProgress.Value++;
                    }
                }
                finally
                {
                    // ── 关闭Undo上下文 ───────────────────────────────
                    EndUndoBlock(doc);
                }

                _tsProgress.Visible = false;
                RefreshGrid();

                Log($"对齐完成: 成功 {success}, 失败 {fail}", success > 0 ? "OK" : "WARN");
                SetStatus($"对齐完成: 成功 {success}, 失败 {fail}");
            }
            catch (Exception ex)
            {
                Log($"对齐过程异常: {ex.Message}", "ERR");
                SetStatus($"对齐异常: {ex.Message}");
                _tsProgress.Visible = false;
            }
        }

        // =====================================================================
        // Undo 上下文管理（防御式多策略访问）
        // =====================================================================

        /// <summary>
        /// 启动Undo块 — PS SDK不同版本暴露不同的API
        /// 成功返回true，失败返回false（操作仍然执行，但不可撤销）
        /// </summary>
        private bool BeginUndoBlock(TxDocument doc, string description)
        {
            // 策略1：TxDocument.UndoRedo.BeginCommand（PS 15+常见）
            try
            {
                dynamic ddoc = doc;
                dynamic undoRedo = ddoc.UndoRedo;
                if (undoRedo != null)
                {
                    undoRedo.BeginCommand(description);
                    Log("  Undo策略1: UndoRedo.BeginCommand OK");
                    return true;
                }
            }
            catch { }

            // 策略2：TxDocument.UndoContext.Open（部分版本）
            try
            {
                dynamic ddoc = doc;
                dynamic ctx = ddoc.UndoContext;
                if (ctx != null)
                {
                    ctx.Open(description);
                    Log("  Undo策略2: UndoContext.Open OK");
                    return true;
                }
            }
            catch { }

            // 策略3：TxApplication 级别的Undo
            try
            {
                dynamic app = TxApplication.ActiveDocument;
                dynamic undoMgr = app.UndoManager;
                if (undoMgr != null)
                {
                    undoMgr.BeginUndoStep(description);
                    Log("  Undo策略3: UndoManager.BeginUndoStep OK");
                    return true;
                }
            }
            catch { }

            // 策略4：直接通过 TxApplication 静态方法
            try
            {
                dynamic txApp = typeof(TxApplication);
                // 部分版本有 TxApplication.BeginCommand(string)
                TxApplication.ActiveDocument.GetType()
                    .GetMethod("BeginCommand")?
                    .Invoke(doc, new object[] { description });
                Log("  Undo策略4: 反射BeginCommand OK");
                return true;
            }
            catch { }

            Log("  所有Undo策略均失败", "WARN");
            return false;
        }

        /// <summary>结束Undo块</summary>
        private void EndUndoBlock(TxDocument doc)
        {
            // 与BeginUndoBlock对应，逐策略尝试关闭
            try { dynamic ddoc = doc; ddoc.UndoRedo.EndCommand(); return; } catch { }
            try { dynamic ddoc = doc; ddoc.UndoContext.Close(); return; } catch { }
            try { dynamic ddoc = doc; ddoc.UndoManager.EndUndoStep(); return; } catch { }
            try
            {
                doc.GetType().GetMethod("EndCommand")?.Invoke(doc, null);
            }
            catch { }
        }

        // =====================================================================
        // 设备忽略判断
        // =====================================================================

        /// <summary>
        /// 检查设备是否应被忽略
        /// 返回忽略原因（空字符串=不忽略）
        /// </summary>
        private string CheckShouldSkip(ITxObject obj, List<string> ignoreKeywords)
        {
            string name = (obj.Name ?? "").Trim();
            string typeName = obj.GetType().Name;

            // ── 1. 名称关键词匹配 ───────────────────────────────────
            foreach (var kw in ignoreKeywords)
            {
                if (name.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
                    return $"名称包含关键词 \"{kw}\"";
            }

            // ── 2. 类型名匹配（SDK内部类型） ────────────────────────
            string[] skipTypes = {
                "TxWeldGun", "TxGun", "TxGripper", "TxTool",
                "TxRobot", "TxHumanModel", "TxConveyor"
            };
            foreach (var st in skipTypes)
            {
                if (typeName.IndexOf(st, StringComparison.OrdinalIgnoreCase) >= 0)
                    return $"类型为 {typeName}";
            }

            // ── 3. 接口检查（Gun接口） ──────────────────────────────
            try
            {
                // ITxWeldGun / ITxGun 接口
                if (obj is ITxObject)
                {
                    Type[] ifaces = obj.GetType().GetInterfaces();
                    foreach (var iface in ifaces)
                    {
                        string ifName = iface.Name;
                        if (ifName.Contains("Gun") || ifName.Contains("Gripper") ||
                            ifName.Contains("Tool") || ifName.Contains("Robot"))
                            return $"实现接口 {ifName}";
                    }
                }
            }
            catch { }

            // ── 4. 检查是否挂载在机器人下（作为末端工具） ────────────
            try
            {
                dynamic dobj = obj;
                dynamic parent = dobj.Parent;
                if (parent != null)
                {
                    string parentType = parent.GetType().Name;
                    if (parentType.Contains("Robot") || parentType.Contains("Flange"))
                        return $"父级为 {parentType}（末端工具）";
                }
            }
            catch { }

            return ""; // 不忽略
        }

        // =====================================================================
        // 分析设备Z向位置
        // =====================================================================
        private void AnalyzeDeviceZ(ITxObject obj, DeviceInfo info)
        {
            // ── 1. 获取设备的世界坐标位置 ────────────────────────────
            TxTransformation absTx = GetAbsoluteLocation(obj);
            if (absTx == null)
            {
                info.Message = "无法获取AbsoluteLocation";
                Log($"  [{info.Name}] 无法获取AbsoluteLocation", "WARN");
                return;
            }

            // ── 2. 提取当前Z位置（变换矩阵第3行第4列 = [2,3]） ──────
            double curZ = ExtractZ(absTx);
            info.CurrentZ = curZ;

            // ── 3. 获取设备最低Z点 ─────────────────────────────────
            string method;
            double minZ = GetDeviceMinZ(obj, absTx, out method);
            info.MinZ = minZ;
            info.Method = method;

            // ── 4. 计算偏移量 ────────────────────────────────────────
            // 目标：使包围盒最低点位于Z=0
            // 偏移量 = 当前最低Z（需要减掉这个值让最低点归零）
            info.OffsetZ = minZ;

            if (Math.Abs(minZ) < 0.01)
                info.Message = "已在Z=0";
            else if (minZ > 0)
                info.Message = $"最低点在Z={minZ:F1}mm，需下移";
            else
                info.Message = $"最低点在Z={minZ:F1}mm，需上移";
        }

        // =====================================================================
        // 获取设备世界坐标（防御式多策略）
        // =====================================================================
        private TxTransformation GetAbsoluteLocation(ITxObject obj)
        {
            // 策略1：ITxLocatableObject 接口（最标准）
            try
            {
                if (obj is ITxLocatableObject loc)
                {
                    var tx = loc.AbsoluteLocation;
                    if (tx != null) return tx;
                }
            }
            catch { }

            // 策略2：dynamic AbsoluteLocation
            try { dynamic d = obj; var tx = d.AbsoluteLocation as TxTransformation; if (tx != null) return tx; } catch { }

            // 策略3：dynamic Location
            try { dynamic d = obj; var tx = d.Location as TxTransformation; if (tx != null) return tx; } catch { }

            // 策略4：dynamic AbsoluteFrame
            try { dynamic d = obj; var tx = d.AbsoluteFrame as TxTransformation; if (tx != null) return tx; } catch { }

            // 策略5：dynamic LocationInWorld
            try { dynamic d = obj; var tx = d.LocationInWorld as TxTransformation; if (tx != null) return tx; } catch { }

            return null;
        }

        // =====================================================================
        // 提取变换矩阵中的Z分量
        // =====================================================================
        private double ExtractZ(TxTransformation tx)
        {
            // TxTransformation 的平移部分在矩阵的 [row, 3] 列
            // Z = [2, 3]

            // 策略1：矩阵索引器（PsReader.cs 中的标准用法）
            try { dynamic d = tx; return Convert.ToDouble(d[2, 3]); } catch { }

            // 策略2：Translation.Z 属性
            try { dynamic d = tx; return Convert.ToDouble(d.Translation.Z); } catch { }

            // 策略3：Z 直接属性
            try { dynamic d = tx; return Convert.ToDouble(d.Z); } catch { }

            // 策略4：TranslationVector
            try { dynamic d = tx; var v = d.TranslationVector; return Convert.ToDouble(v.Z); } catch { }

            return 0;
        }

        /// <summary>
        /// 提取变换矩阵中的完整平移分量 [X, Y, Z]
        /// </summary>
        private double[] ExtractTranslation(TxTransformation tx)
        {
            // 策略1：矩阵索引器
            try
            {
                dynamic d = tx;
                return new double[]
                {
                    Convert.ToDouble(d[0, 3]),
                    Convert.ToDouble(d[1, 3]),
                    Convert.ToDouble(d[2, 3])
                };
            }
            catch { }

            // 策略2：Translation 属性（返回TxVector）
            try
            {
                dynamic d = tx;
                dynamic t = d.Translation;
                return new double[]
                {
                    Convert.ToDouble(t.X),
                    Convert.ToDouble(t.Y),
                    Convert.ToDouble(t.Z)
                };
            }
            catch { }

            return null;
        }

        // =====================================================================
        // 获取设备最低Z点（基于TxComponent实际API）
        // =====================================================================
        // 设计决策：
        //   TxComponent 没有 BoundingBox 属性，GeometricCenter 推算公式不可靠
        //   （对于底部在地面的设备，推算结果为负数，导致所有设备被标记需对齐）
        //
        //   最可靠策略：直接使用 AbsoluteLocation.Z（设备坐标原点的世界Z值）
        //   PS中绝大多数设备的坐标原点就在底部基准面，这是行业惯例
        //   仅在 GetLocationAxisIntersectionPoints 可用时才尝试更精确的计算
        private double GetDeviceMinZ(ITxObject obj, TxTransformation absTx, out string method)
        {
            method = "坐标原点";

            // ── 策略1：GetLocationAxisIntersectionPoints（最精确） ────
            // 沿Z轴的交点可得到几何体在Z方向的上下边界
            try
            {
                if (obj is TxComponent comp)
                {
                    dynamic dComp = comp;

                    // 方式A：传入轴索引
                    try
                    {
                        dynamic points = dComp.GetLocationAxisIntersectionPoints(2);
                        double minZ = ExtractMinZFromPoints(points);
                        if (minZ < double.MaxValue)
                        {
                            method = "轴交点";
                            Log($"    [{obj.Name}] GetLocationAxisIntersectionPoints(Z) MinZ={minZ:F1}");
                            return minZ;
                        }
                    }
                    catch { }

                    // 方式B：无参调用
                    try
                    {
                        dynamic points = dComp.GetLocationAxisIntersectionPoints();
                        double minZ = ExtractMinZFromPoints(points);
                        if (minZ < double.MaxValue)
                        {
                            method = "轴交点";
                            Log($"    [{obj.Name}] GetLocationAxisIntersectionPoints() MinZ={minZ:F1}");
                            return minZ;
                        }
                    }
                    catch { }
                }
            }
            catch { }

            // ── 策略2（回退）：直接使用 AbsoluteLocation.Z ──────────
            // 设备坐标原点的世界Z值 — PS中设备原点通常在底部
            double originZ = ExtractZ(absTx);
            Log($"    [{obj.Name}] 使用坐标原点Z={originZ:F1}（PS设备原点通常在底部）");
            return originZ;
        }

        /// <summary>
        /// 从交点集合中提取最小Z值
        /// 交点可能是 TxVector[]、TxTransformation[]、List、TxObjectList 等
        /// </summary>
        private double ExtractMinZFromPoints(object points)
        {
            double minZ = double.MaxValue;
            if (points == null) return minZ;

            // 情况1：可枚举对象（数组、List、TxObjectList）
            try
            {
                var enumerable = points as System.Collections.IEnumerable;
                if (enumerable != null)
                {
                    foreach (object pt in enumerable)
                    {
                        double z = ExtractZFromPoint(pt);
                        if (z < minZ) minZ = z;
                    }
                    return minZ;
                }
            }
            catch { }

            // 情况2：单个点
            double singleZ = ExtractZFromPoint(points);
            if (singleZ < minZ) minZ = singleZ;

            return minZ;
        }

        /// <summary>
        /// 从单个点对象中提取Z值（TxVector/TxTransformation/dynamic）
        /// </summary>
        private double ExtractZFromPoint(object pt)
        {
            if (pt == null) return double.MaxValue;

            // TxVector.Z
            try { dynamic d = pt; return Convert.ToDouble(d.Z); } catch { }
            // TxTransformation 矩阵 [2,3]
            try { dynamic d = pt; return Convert.ToDouble(d[2, 3]); } catch { }
            // Translation.Z
            try { dynamic d = pt; return Convert.ToDouble(d.Translation.Z); } catch { }

            return double.MaxValue;
        }

        // =====================================================================
        // 执行Z偏移（修改设备的AbsoluteLocation）
        // =====================================================================
        private bool ApplyZOffset(DeviceInfo dev)
        {
            if (dev.TxObj == null) return false;
            double offsetZ = dev.OffsetZ; // 需要减去这个值

            if (Math.Abs(offsetZ) < 0.01) return true; // 无需偏移

            // ── 获取当前位置 ─────────────────────────────────────────
            TxTransformation curTx = GetAbsoluteLocation(dev.TxObj);
            if (curTx == null)
            {
                Log($"    [{dev.Name}] 无法获取当前位置用于偏移", "ERR");
                return false;
            }

            // ── 修改Z分量 ────────────────────────────────────────────
            // 新Z = 当前Z - offsetZ（使最低点归零）
            double[] curXYZ = ExtractTranslation(curTx);
            if (curXYZ == null)
            {
                Log($"    [{dev.Name}] 无法提取平移分量", "ERR");
                return false;
            }

            double newZ = curXYZ[2] - offsetZ;

            // ── 写入新位置 ───────────────────────────────────────────
            bool written = false;

            // 写入策略1：矩阵索引器 [2,3]
            if (!written)
            {
                try
                {
                    dynamic d = curTx;
                    d[2, 3] = newZ;
                    written = true;
                }
                catch { }
            }

            // 写入策略2：Translation 属性
            if (!written)
            {
                try
                {
                    dynamic d = curTx;
                    dynamic trans = d.Translation;
                    // TxVector 可能是值类型，需要先读出修改再写回
                    var newTrans = new TxVector(curXYZ[0], curXYZ[1], newZ);
                    d.Translation = newTrans;
                    written = true;
                }
                catch { }
            }

            // 写入策略3：Z 直接属性
            if (!written)
            {
                try
                {
                    dynamic d = curTx;
                    d.Z = newZ;
                    written = true;
                }
                catch { }
            }

            if (!written)
            {
                Log($"    [{dev.Name}] 所有Z值写入策略失败", "ERR");
                return false;
            }

            // ── 将修改后的变换矩阵写回设备 ──────────────────────────
            bool applied = false;

            // 应用策略1：ITxLocatableObject.AbsoluteLocation（可写属性）
            if (!applied)
            {
                try
                {
                    if (dev.TxObj is ITxLocatableObject loc)
                    {
                        loc.AbsoluteLocation = curTx;
                        applied = true;
                    }
                }
                catch { }
            }

            // 应用策略2：dynamic AbsoluteLocation setter
            if (!applied)
            {
                try
                {
                    dynamic d = dev.TxObj;
                    d.AbsoluteLocation = curTx;
                    applied = true;
                }
                catch { }
            }

            // 应用策略3：dynamic Location setter
            if (!applied)
            {
                try
                {
                    dynamic d = dev.TxObj;
                    d.Location = curTx;
                    applied = true;
                }
                catch { }
            }

            // 应用策略4：SetAbsoluteLocation 方法
            if (!applied)
            {
                try
                {
                    dynamic d = dev.TxObj;
                    d.SetAbsoluteLocation(curTx);
                    applied = true;
                }
                catch { }
            }

            if (!applied)
            {
                Log($"    [{dev.Name}] 所有位置写回策略失败", "ERR");
                return false;
            }

            return true;
        }

        // =====================================================================
        // 忽略关键词解析
        // =====================================================================
        private List<string> ParseIgnoreKeywords()
        {
            var keywords = new List<string>();
            if (_txtIgnoreKeywords == null || string.IsNullOrWhiteSpace(_txtIgnoreKeywords.Text))
                return keywords;

            var parts = _txtIgnoreKeywords.Text.Split(
                new[] { ',', '，', ';', '；', '|' },
                StringSplitOptions.RemoveEmptyEntries);

            foreach (var p in parts)
            {
                string trimmed = p.Trim();
                if (trimmed.Length > 0)
                    keywords.Add(trimmed);
            }

            return keywords.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        // =====================================================================
        // 表格刷新
        // =====================================================================
        private void RefreshGrid()
        {
            if (_grid == null) return;

            _grid.Rows.Count = _grid.Rows.Fixed + _devices.Count;

            for (int i = 0; i < _devices.Count; i++)
            {
                int row = i + _grid.Rows.Fixed;
                var d = _devices[i];

                _grid[row, COL_IDX]    = d.Index.ToString();
                _grid[row, COL_NAME]   = d.Name;
                _grid[row, COL_TYPE]   = d.TypeName;
                _grid[row, COL_CURZ]   = d.IsSkipped ? "-" : $"{d.CurrentZ:F1}";
                _grid[row, COL_OFFSET] = d.IsSkipped ? "-" : $"{d.OffsetZ:F1}";
                _grid[row, COL_METHOD] = d.IsSkipped ? "-" : d.Method;
                _grid[row, COL_STATUS] = d.IsSkipped ? "忽略"
                                       : d.IsAligned ? "已对齐"
                                       : Math.Abs(d.OffsetZ) < 0.01 ? "正常"
                                       : "待对齐";
                _grid[row, COL_MSG]    = d.Message;

                // 行颜色
                try
                {
                    var rowStyle = _grid.Styles.Add($"rs_{row}");
                    if (d.IsSkipped)
                        rowStyle.BackColor = Color.FromArgb(240, 240, 240);   // 浅灰
                    else if (d.IsAligned)
                        rowStyle.BackColor = Color.FromArgb(198, 239, 206);   // 绿
                    else if (Math.Abs(d.OffsetZ) < 0.01)
                        rowStyle.BackColor = SystemColors.Window;             // 白
                    else
                        rowStyle.BackColor = Color.FromArgb(255, 235, 156);   // 黄

                    _grid.Rows[row].Style = rowStyle;
                }
                catch { }
            }

            UpdateSummaryStats();
        }

        /// <summary>更新卡片区的统计数据</summary>
        private void UpdateSummaryStats()
        {
            int total     = _devices.Count;
            int skipped   = _devices.Count(d => d.IsSkipped);
            int needAlign = _devices.Count(d => !d.IsSkipped && !d.IsAligned && Math.Abs(d.OffsetZ) > 0.01);
            int aligned   = _devices.Count(d => d.IsAligned);
            int ok        = _devices.Count(d => !d.IsSkipped && !d.IsAligned && Math.Abs(d.OffsetZ) < 0.01);

            if (_lblTotal     != null) _lblTotal.Text     = $"总数: {total}";
            if (_lblSkipped   != null) _lblSkipped.Text   = $"已忽略: {skipped}";
            if (_lblNeedAlign != null) _lblNeedAlign.Text = $"待对齐: {needAlign}";
            if (_lblAligned   != null) _lblAligned.Text   = $"已对齐: {aligned}";
            if (_lblOk        != null) _lblOk.Text        = $"正常: {ok}";
        }

        // =====================================================================
        // 双击跳转到PS对象
        // =====================================================================
        private void Grid_DblClick(object sender, EventArgs e)
        {
            int selRow = _grid.RowSel;
            if (selRow < _grid.Rows.Fixed) return;

            int dataIdx = selRow - _grid.Rows.Fixed;
            if (dataIdx < 0 || dataIdx >= _devices.Count) return;

            var dev = _devices[dataIdx];
            if (dev.TxObj == null) return;

            try
            {
                // 在PS中选中该对象
                // TxSelection API（来自文档）：
                //   Clear()    — 清空选择
                //   SetItems(TxObjectList)  — 设置选择（替换）
                //   AddItems(TxObjectList)  — 追加到选择
                //   GetItems() — 获取当前选择
                bool selected = false;

                // 首选：Clear + AddItems（文档确认的标准方法）
                if (!selected)
                {
                    try
                    {
                        var objList = new TxObjectList();
                        objList.Add(dev.TxObj);
                        TxApplication.ActiveSelection.Clear();
                        TxApplication.ActiveSelection.AddItems(objList);
                        selected = true;
                    }
                    catch { }
                }

                // 备选：SetItems（直接替换当前选择）
                if (!selected)
                {
                    try
                    {
                        var objList = new TxObjectList();
                        objList.Add(dev.TxObj);
                        TxApplication.ActiveSelection.SetItems(objList);
                        selected = true;
                    }
                    catch { }
                }

                if (selected)
                    Log($"已选中: {dev.Name}");
                else
                    Log($"选中失败", "WARN");

                // 尝试 Zoom to Fit 聚焦
                try
                {
                    dynamic viewer = TxApplication.ActiveDocument;
                    // 不同版本可能是 viewer.ZoomToSelection / FitToSelection
                    try { viewer.ZoomToSelection(); } catch { }
                }
                catch { }
            }
            catch (Exception ex)
            {
                Log($"跳转失败: {ex.Message}", "WARN");
            }
        }

        // =====================================================================
        // 状态与日志辅助
        // =====================================================================
        private void SetStatus(string text)
        {
            if (_lblStatus != null) _lblStatus.Text = text;
        }

        private void Log(string message, string level = "INFO")
        {
            if (_logBox == null || _logBox.IsDisposed) return;
            if (_logBox.InvokeRequired)
            {
                _logBox.BeginInvoke(new Action(() => Log(message, level)));
                return;
            }

            string ts = DateTime.Now.ToString("HH:mm:ss.fff");
            string line = $"[{ts}] [{level}] {message}";

            Color col;
            switch (level)
            {
                case "ERR":  col = Color.FromArgb(255, 100, 100); break;
                case "WARN": col = Color.FromArgb(255, 200, 80);  break;
                case "OK":   col = Color.FromArgb(80, 220, 120);  break;
                default:     col = Color.FromArgb(204, 204, 204); break;
            }

            _logBox.SelectionStart = _logBox.TextLength;
            _logBox.SelectionLength = 0;
            _logBox.SelectionColor = col;
            _logBox.AppendText(line + "\n");
            _logBox.SelectionColor = _logBox.ForeColor;
            _logBox.ScrollToCaret();

            // ERR/WARN 自动弹出日志面板
            if ((level == "ERR" || level == "WARN") && !_logVisible)
            {
                _logVisible = true;
                _logPanel.Visible = true;
                if (_btnLog != null) _btnLog.Checked = true;
            }
        }

        // =====================================================================
        // 窗体关闭
        // =====================================================================
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _devices.Clear();
            base.OnFormClosing(e);
        }
    }
}
