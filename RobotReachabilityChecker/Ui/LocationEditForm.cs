using System;
using System.Drawing;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui.WPF;
using static TxTools.RobotReachabilityChecker.Ui.Theme;

namespace TxTools.RobotReachabilityChecker.Ui
{
    // =========================================================================
    // 点位编辑对话框 — 基于 TxPlacementCollisionControl API
    //
    // 核心属性（来自 PS API 文档）：
    //   LinearValue           — 当前平移值（mm）
    //   AngularValue          — 当前旋转值（deg）
    //   LinearStepSize        — 每次移动步长（mm）
    //   AngularStepSize       — 每次旋转步长（deg）
    //   SelectedAxis          — 当前激活轴（X/Y/Z/RX/RY/RZ）
    //   ShowCollisionButtons  — 显示/隐藏碰撞检测按钮
    //
    // 核心方法：
    //   Manipulate(delta)     — 执行一次位移/旋转，delta 为增量值
    //   MoveOneStepLinear()   — 按 LinearStepSize 移动一步
    //   MoveOneStepAngular()  — 按 AngularStepSize 旋转一步
    //   Reset()               — 重置控件显示值
    //   UnInitialize()        — 释放内部资源（关闭前调用）
    //
    // 核心事件：
    //   DeltaChanged          — 值发生变化时触发
    //   MovementTypeChanged   — 激活轴切换时触发
    //   RunToCollision        — 碰撞按钮按下时触发
    //
    // 注意：依赖 WindowsFormsIntegration.dll（提供 ElementHost）。
    // =========================================================================
    internal class LocationEditForm : Form
    {
        private readonly ITxRoboticLocationOperation _locOp;
        private readonly string _pointName;

        private System.Windows.Forms.Integration.ElementHost _host;
        private TxPlacementCollisionControl _ctrl;

        private Label _lblInfo;
        private Label _lblStatus;
        private Button _btnReset;
        private Button _btnClose;

        public LocationEditForm(ITxRoboticLocationOperation locOp, string pointName)
        {
            _locOp = locOp;
            _pointName = pointName;
            Text = "编辑点位 — " + pointName;
            StartPosition = FormStartPosition.CenterParent;
            Size = new System.Drawing.Size(480, 400);
            MinimumSize = new System.Drawing.Size(420, 340);
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            BackColor = SystemColors.Control;
            Build();
        }

        private void Build()
        {
            // 顶部标题栏
            _lblInfo = new Label
            {
                Text = "点位: " + _pointName,
                Dock = DockStyle.Top,
                Height = 26,
                Padding = new Padding(8, 5, 0, 0),
                Font = new System.Drawing.Font(SystemFonts.DefaultFont, FontStyle.Bold),
                ForeColor = TxClrAccent.Color,
                BackColor = TxClrEditHeader.Color
            };

            // 底部状态 + 按钮栏
            var btmPanel = new Panel { Dock = DockStyle.Bottom, Height = 36 };

            _lblStatus = new Label
            {
                Text = "使用控件箭头调整点位位置和姿态",
                AutoSize = false,
                Width = 260,
                Height = 22,
                Location = new System.Drawing.Point(6, 7),
                ForeColor = SystemColors.GrayText,
                Font = SystemFonts.DefaultFont
            };

            _btnReset = new Button
            {
                Text = "还原",
                Width = 72,
                Height = 26,
                FlatStyle = FlatStyle.System,
                Font = SystemFonts.DefaultFont,
                Anchor = AnchorStyles.Right | AnchorStyles.Top
            };
            _btnReset.Click += BtnReset_Click;

            _btnClose = new Button
            {
                Text = "关闭",
                Width = 72,
                Height = 26,
                FlatStyle = FlatStyle.System,
                Font = SystemFonts.DefaultFont,
                Anchor = AnchorStyles.Right | AnchorStyles.Top
            };
            _btnClose.Click += (s, e) => Close();

            btmPanel.Controls.AddRange(new Control[] { _lblStatus, _btnReset, _btnClose });
            btmPanel.Resize += (s, e) =>
            {
                _btnClose.Location = new System.Drawing.Point(btmPanel.Width - 78, 5);
                _btnReset.Location = new System.Drawing.Point(btmPanel.Width - 156, 5);
            };
            _btnClose.Location = new System.Drawing.Point(392, 5);
            _btnReset.Location = new System.Drawing.Point(314, 5);

            // ElementHost: 承载 TxPlacementCollisionControl（WPF 控件）
            _host = new System.Windows.Forms.Integration.ElementHost
            {
                Dock = DockStyle.Fill,
                BackColor = SystemColors.Control,
                BackColorTransparent = false
            };

            _ctrl = new TxPlacementCollisionControl();
            _ctrl.InitializeComponent();

            _ctrl.ShowCollisionButtons = true;
            _ctrl.LinearStepSize = 1.0;   // mm
            _ctrl.AngularStepSize = 1.0;  // deg

            _ctrl.DeltaChanged += Ctrl_DeltaChanged;
            _ctrl.MovementTypeChanged += Ctrl_MovementTypeChanged;
            _ctrl.RunToCollision += Ctrl_RunToCollision;

            _host.Child = _ctrl;

            Controls.Add(_host);
            Controls.Add(btmPanel);
            Controls.Add(_lblInfo);
        }

        // DeltaChanged: sender 是控件本身
        // LinearValue / AngularValue 是 double（当前激活轴的标量值）
        private void Ctrl_DeltaChanged(object sender, EventArgs e)
        {
            if (_ctrl == null || _lblStatus == null) return;
            try
            {
                double linVal = _ctrl.LinearValue;
                double angVal = _ctrl.AngularValue;

                string axis = "";
                try { axis = _ctrl.SelectedAxis.ToString(); } catch { }

                _lblStatus.Text = string.IsNullOrEmpty(axis)
                    ? string.Format("平移={0:F2}mm  旋转={1:F2}°", linVal, angVal)
                    : string.Format("[{0}] 平移={1:F2}mm  旋转={2:F2}°", axis, linVal, angVal);

                // 控件内部会自动调用 Manipulate 驱动 PS 场景中的点位，此处无需重复调用
            }
            catch { }
        }

        private void Ctrl_MovementTypeChanged(object sender, EventArgs e)
        {
            if (_ctrl == null || _lblInfo == null) return;
            try
            {
                string axis = _ctrl.SelectedAxis.ToString();
                _lblInfo.Text = "点位: " + _pointName
                    + (string.IsNullOrEmpty(axis) ? "" : "  [" + axis + "]");
            }
            catch { }
        }

        private void Ctrl_RunToCollision(object sender, EventArgs e)
        {
            if (_lblStatus != null) _lblStatus.Text = "⚠ 碰撞检测触发";
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            if (_ctrl == null) return;
            try
            {
                _ctrl.Reset();
                if (_lblStatus != null) _lblStatus.Text = "已重置";
            }
            catch { }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            try
            {
                if (_ctrl != null)
                {
                    _ctrl.UnInitialize();
                    _ctrl.DeltaChanged -= Ctrl_DeltaChanged;
                    _ctrl.MovementTypeChanged -= Ctrl_MovementTypeChanged;
                    _ctrl.RunToCollision -= Ctrl_RunToCollision;
                }
                if (_host != null) { _host.Child = null; _host.Dispose(); }
            }
            catch { }
            base.OnFormClosing(e);
        }
    }
}
