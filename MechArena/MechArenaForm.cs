using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;
using TxTools.Common;

using Timer = System.Windows.Forms.Timer;

namespace TxTools.MechArena
{
    /// <summary>
    /// MechArena 主窗体 —— 采用 TxTools 统一 GUI 规范（FormUiKit）。
    /// 布局：Header(28) + Body(左右两栏) + Bottom(46 按钮栏)。
    /// 右栏血量列表为手写 Dock=Fill 面板（不依赖 MkCard）。
    /// 视角：右键/中键按住拖动 = 绕玩家旋转（全局 Cursor.Position 采样，
    /// 不要求鼠标在本窗体上）；Esc 解除焦点锁定，F 重新锁定。
    /// </summary>
    public class MechArenaForm : TxForm
    {
        // ---------- 单例 ----------
        private static MechArenaForm _instance;
        public static MechArenaForm Instance
        {
            get { return _instance ?? (_instance = new MechArenaForm()); }
        }

        // ---------- 设计尺寸（96-DPI 裸像素）----------
        private static readonly Size DesignSize = new Size(600, 520);
        private static readonly Size MinSize = new Size(500, 420);
        private bool _dpiApplied;

        // ---------- 游戏状态 ----------
        private MechArenaEngine _engine;
        private readonly Timer _tick;
        private readonly HashSet<Keys> _keysDown = new HashSet<Keys>();
        private bool _fireRequested;

        // 鼠标视角采集（全局，无需窗体焦点）
        private Point _lastMousePos;
        private bool _mouseWasDragging;

        // 自动重启倒计时（GameOver 后启用）
        private const double RESTART_DELAY = 3.0;   // 秒
        private double _restartCountdown = -1;

        // ---------- 控件 ----------
        private FormUiKit.FlatColorLabel _lblStatus;
        private FormUiKit.FlatColorLabel _lblRobotList;
        private FormUiKit.FlatColorLabel _lblBanner;   // GAME OVER / VICTORY / 倒计时
        private FormUiKit.FlatColorButton _btnStart;
        private FormUiKit.FlatColorButton _btnStop;

        // =====================================================================
        public MechArenaForm()
        {
            FormUiKit.InitStandardForm(this,
                "MechArena  |  机器人竞技场",
                DesignSize, MinSize, sizable: true);

            try { this.SemiModal = false; } catch { }

            BuildUi();

            this.KeyPreview = true;
            this.KeyDown += OnKeyDown;
            this.KeyUp += OnKeyUp;

            _tick = new Timer { Interval = 50 };
            _tick.Tick += OnTick;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            FormUiKit.ApplyDpiScaling(this, ref _dpiApplied, DesignSize);
        }

        // =====================================================================
        //  UI 构建
        // =====================================================================
        private void BuildUi()
        {
            var header = BuildHeader();
            var body = BuildBody();
            var bottom = BuildBottom();
            var root = FormUiKit.BuildRoot(header, body, bottom);
            Controls.Add(root);
        }

        private Control BuildHeader()
        {
            return new FormUiKit.FlatColorLabel
            {
                Text = "  MechArena  —  机器人竞技场",
                Dock = DockStyle.Fill,
                Font = FormUiKit.TitleFont,
                BackColor = FormUiKit.TitleBack,
                ForeColor = FormUiKit.TitleFore,
                TextAlign = ContentAlignment.MiddleLeft,
            };
        }

        private Control BuildBody()
        {
            var body = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = Padding.Empty,
                BackColor = SystemColors.Control
            };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 280));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            body.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // === 左侧：帮助 + 状态 ===
            var leftCol = FormUiKit.BuildCardColumn(272);

            FlowLayoutPanel helpContent;
            var helpCard = FormUiKit.MkCard("操作说明", 260, 140, out helpContent);
            var helpLabel = new FormUiKit.FlatColorLabel
            {
                Text =
                    "W A S D      平面移动\r\n" +
                    "Space        锁定最近敌人开火\r\n" +
                    "右键/中键拖动 旋转视角\r\n" +
                    "PgUp/PgDn    视角远近\r\n" +
                    "Esc          解除视角锁定\r\n" +
                    "F            重新锁定玩家",
                AutoSize = true,
                Font = MonoFont(),
                BackColor = Color.Transparent,
                MaximumSize = new Size(240, 0)
            };
            helpContent.Controls.Add(helpLabel);
            leftCol.Controls.Add(helpCard);

            FlowLayoutPanel statusContent;
            var statusCard = FormUiKit.MkCard("游戏状态", 260, 180, out statusContent);
            _lblStatus = new FormUiKit.FlatColorLabel
            {
                Text = "未启动 —— 点击【开始】",
                AutoSize = true,
                Font = MonoFont(),
                BackColor = Color.Transparent,
                MaximumSize = new Size(240, 0)
            };
            statusContent.Controls.Add(_lblStatus);

            _lblBanner = new FormUiKit.FlatColorLabel
            {
                Text = "",
                AutoSize = true,
                Font = new Font(MonoFont().FontFamily, 10f, FontStyle.Bold),
                BackColor = Color.Transparent,
                ForeColor = Color.FromArgb(180, 40, 40),
                MaximumSize = new Size(240, 0),
                Margin = new Padding(3, 6, 3, 3)
            };
            statusContent.Controls.Add(_lblBanner);

            leftCol.Controls.Add(statusCard);

            // === 右侧：敌人血量列表（手写面板，Dock=Fill 占满右栏）===
            var rightCol = BuildRobotListPanel();

            body.Controls.Add(leftCol, 0, 0);
            body.Controls.Add(rightCol, 1, 0);
            return body;
        }

        /// <summary>
        /// 右栏敌人血量面板：标题条 + AutoScroll 内容区 + AutoSize 标签。
        /// </summary>
        private Control BuildRobotListPanel()
        {
            var outer = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(4, 6, 8, 6),
                BackColor = SystemColors.Control
            };

            var card = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            var title = new FormUiKit.FlatColorLabel
            {
                Text = "  敌人血量",
                Dock = DockStyle.Top,
                Height = 24,
                Font = new Font(FormUiKit.BaseFont, FontStyle.Bold),
                BackColor = Color.FromArgb(240, 242, 245),
                ForeColor = Color.FromArgb(60, 60, 60),
                TextAlign = ContentAlignment.MiddleLeft
            };

            var scroll = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.White
            };

            _lblRobotList = new FormUiKit.FlatColorLabel
            {
                Text = "(未启动)",
                AutoSize = true,
                Font = MonoFont(),
                BackColor = Color.White,
                Location = new Point(6, 4)
            };
            scroll.Controls.Add(_lblRobotList);

            // 注意添加顺序：先 Fill 后 Top，否则 Top 会盖住 Fill 的起始区域
            card.Controls.Add(scroll);
            card.Controls.Add(title);
            outer.Controls.Add(card);
            return outer;
        }

        private Control BuildBottom()
        {
            var bottom = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(4, 6, 4, 6),
                BackColor = SystemColors.Control
            };

            _btnStop = (FormUiKit.FlatColorButton)FormUiKit.MkButton("停 止", primary: false, width: 100, height: 30);
            _btnStop.Enabled = false;
            _btnStop.Click += (s, e) => StopGame();

            _btnStart = (FormUiKit.FlatColorButton)FormUiKit.MkButton("▶  开 始", primary: true, width: 120, height: 30);
            _btnStart.Click += (s, e) => StartGame();

            bottom.Controls.Add(_btnStop);
            bottom.Controls.Add(_btnStart);
            return bottom;
        }

        private static Font MonoFont()
        {
            try { return new Font("Consolas", FormUiKit.BaseFont.Size); }
            catch { return new Font(FontFamily.GenericMonospace, 9f); }
        }

        // =====================================================================
        //  Start / Stop / Restart
        // =====================================================================
        private void StartGame()
        {
            try
            {
                _engine = new MechArenaEngine();
                _engine.Initialize();
                _tick.Start();

                _btnStart.Enabled = false;
                _btnStop.Enabled = true;
                _keysDown.Clear();
                _fireRequested = false;
                _restartCountdown = -1;
                _mouseWasDragging = false;
                _lblBanner.Text = "";
                this.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show("初始化失败:\r\n" + ex.Message,
                    "MechArena", MessageBoxButtons.OK, MessageBoxIcon.Error);
                try { _engine?.Dispose(); } catch { }
                _engine = null;
            }
        }

        private void StopGame()
        {
            _tick.Stop();
            try { _engine?.Dispose(); } catch { }
            _engine = null;
            _keysDown.Clear();
            _fireRequested = false;
            _restartCountdown = -1;
            _mouseWasDragging = false;

            _btnStart.Enabled = true;
            _btnStop.Enabled = false;
            _lblStatus.Text = "已停止";
            _lblRobotList.Text = "(未启动)";
            _lblBanner.Text = "";
        }

        /// <summary>
        /// GameOver 倒计时结束后调用 —— 完整清理旧引擎并重新初始化。
        /// </summary>
        private void RestartGame()
        {
            _restartCountdown = -1;

            try { _engine?.Dispose(); } catch { }
            _engine = null;
            _keysDown.Clear();
            _fireRequested = false;

            try
            {
                _engine = new MechArenaEngine();
                _engine.Initialize();
                _lblBanner.Text = "";
                this.Focus();
            }
            catch (Exception ex)
            {
                _tick.Stop();
                _lblStatus.Text = "重启失败: " + ex.Message;
                _btnStart.Enabled = true;
                _btnStop.Enabled = false;
            }
        }

        // =====================================================================
        //  鼠标视角采集：右键或中键按住时取全局光标增量。
        //  Control.MouseButtons / Cursor.Position 均为进程级全局，
        //  不要求鼠标位于本窗体之上（在 3D 视口里拖动也生效）。
        //  按下的第一帧只记录位置不产生增量，避免视角跳变。
        // =====================================================================
        private void SampleMouseDelta(out int dx, out int dy)
        {
            dx = 0; dy = 0;
            bool dragging =
                (Control.MouseButtons & MouseButtons.Right) == MouseButtons.Right ||
                (Control.MouseButtons & MouseButtons.Middle) == MouseButtons.Middle;

            if (dragging)
            {
                var p = Cursor.Position;
                if (_mouseWasDragging)
                {
                    dx = p.X - _lastMousePos.X;
                    dy = p.Y - _lastMousePos.Y;
                }
                _lastMousePos = p;
            }
            _mouseWasDragging = dragging;
        }

        // =====================================================================
        //  Tick
        // =====================================================================
        private void OnTick(object sender, EventArgs e)
        {
            if (_engine == null) return;

            try
            {
                if (_restartCountdown < 0)
                {
                    int mdx, mdy;
                    SampleMouseDelta(out mdx, out mdy);
                    _engine.Update(_keysDown, _fireRequested, mdx, mdy);
                    _fireRequested = false;
                }

                UpdateStatusLabel();
                UpdateRobotList();

                // --- GameOver 自动重启 ---
                if (_engine.GameOver)
                {
                    if (_restartCountdown < 0) _restartCountdown = RESTART_DELAY;
                    _restartCountdown -= 0.05;
                    int sec = (int)Math.Ceiling(_restartCountdown);
                    if (sec < 0) sec = 0;
                    _lblBanner.ForeColor = Color.FromArgb(200, 40, 40);
                    _lblBanner.Text = string.Format(
                        "★ GAME OVER ★\r\n最终分数: {0}\r\n{1} 秒后自动重开",
                        _engine.Score, sec);

                    if (_restartCountdown <= 0)
                        RestartGame();
                }
                // --- 胜利：停下等玩家 ---
                else if (_engine.Victory)
                {
                    _tick.Stop();
                    _lblBanner.ForeColor = Color.FromArgb(40, 140, 60);
                    _lblBanner.Text = "☆ VICTORY ☆\r\n最终分数: " + _engine.Score;
                    _btnStart.Enabled = true;
                    _btnStop.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                _lblStatus.Text = "游戏循环异常:\r\n" + ex.Message;
            }
        }

        // =====================================================================
        //  状态显示（文本变化才写入，减少 AutoSize 标签闪烁）
        // =====================================================================
        private void UpdateStatusLabel()
        {
            if (_engine == null) return;

            int alive = 0;
            foreach (var e in _engine.Enemies) if (e.Health > 0) alive++;

            var sb = new StringBuilder();
            sb.AppendLine("HP    : " + Bar(_engine.Player.Health, 100, 14));
            sb.AppendLine("时间  : " + _engine.Time.ToString("0.0") + " s");
            sb.AppendLine("分数  : " + _engine.Score);
            sb.AppendLine("敌人  : " + alive + " / " + _engine.Enemies.Count +
                          "   子弹: " + _engine.Projectiles.Count);
            sb.AppendLine("Boss  : " + (_engine.Boss.Active ? "★ 合体 ★" : "未激活"));
            sb.AppendLine("视角  : " + (_engine.CameraFocusLocked
                          ? "锁定玩家 (Esc解除)" : "自由 (F锁定)"));
            sb.AppendLine("干涉集: " + _engine.CollisionStatus);
            string s = sb.ToString();
            if (_lblStatus.Text != s) _lblStatus.Text = s;
        }

        private void UpdateRobotList()
        {
            if (_engine == null) { _lblRobotList.Text = "(未启动)"; return; }
            if (_engine.Enemies.Count == 0) { _lblRobotList.Text = "(场景内无机器人)"; return; }

            var sb = new StringBuilder();
            for (int i = 0; i < _engine.Enemies.Count; i++)
            {
                var e = _engine.Enemies[i];
                bool bossMember = _engine.Boss.Active && _engine.Boss.Members.Contains(e);
                int maxHp = bossMember ? MechArenaEngine.BOSS_MEMBER_HP : MechArenaEngine.ENEMY_MAX_HP;

                string tag = e.Health <= 0 ? "☠"
                           : bossMember ? "★"
                                           : " ";
                string name = ShortName(e.RobotName, 12);
                sb.AppendLine(string.Format("{0} #{1:D2} {2,-12} {3} {4,3}/{5}  [{6}]",
                    tag, i + 1, name,
                    Bar(Math.Max(0, e.Health), maxHp, 10),
                    Math.Max(0, e.Health), maxHp,
                    e.MoveTargetInfo));   // 调试：当前移动目标（本体/容器/全部失效）
            }
            string s = sb.ToString();
            if (_lblRobotList.Text != s) _lblRobotList.Text = s;
        }

        private static string ShortName(string s, int maxLen)
        {
            if (string.IsNullOrEmpty(s)) return "?";
            if (s.Length <= maxLen) return s;
            return s.Substring(0, maxLen - 1) + "…";
        }

        private static string Bar(int cur, int max, int width)
        {
            if (max <= 0) max = 1;
            if (cur < 0) cur = 0;
            int seg = Math.Max(0, Math.Min(width, cur * width / max));
            return "[" + new string('█', seg) + new string('·', width - seg) + "]";
        }

        // =====================================================================
        //  键盘输入
        //  Esc：解除视角焦点锁定（不再停止游戏 —— 停止用【停 止】按钮）
        //  F  ：重新锁定玩家
        // =====================================================================
        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                _engine?.SetCameraFocus(false);
                return;
            }
            if (e.KeyCode == Keys.F)
            {
                _engine?.SetCameraFocus(true);
                return;
            }
            if (e.KeyCode == Keys.Space) _fireRequested = true;
            _keysDown.Add(e.KeyCode);
        }

        private void OnKeyUp(object sender, KeyEventArgs e)
        {
            _keysDown.Remove(e.KeyCode);
        }

        // =====================================================================
        //  关闭
        // =====================================================================
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            try { StopGame(); } catch { }
            base.OnFormClosing(e);
            _instance = null;
        }
    }
}