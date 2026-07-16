using System;
using System.Drawing;
using System.Windows.Forms;
using Tecnomatix.Engineering;

namespace TxTools.SnakeGame
{
    /// <summary>
    /// 贪吃蛇 UI 窗体：Timer 驱动游戏循环，方向键控制移动方向，空格暂停。
    /// 游戏几何全部渲染在 PS 3D 视图中：蛇头/蛇身/食物均是 CreateSolidBox 生成的实体。
    /// </summary>
    public class SnakeGameForm : Form
    {
        // ===== 游戏配置（可按需要调整）=====
        private const int GridHalfExtent = 10;      // 21×21 网格：[-10, 10]
        private const double CellSize = 60.0;       // 每格 60 mm，全场约 1.26 m × 1.26 m
        private const int InitialTickMs = 200;
        private const int MinTickMs = 60;
        private const int SpeedupEveryScore = 5;    // 每吃 5 个食物加一次速
        private const int SpeedupStepMs = 20;

        private readonly SnakeGameEngine _engine;
        private readonly SnakeWorld _world;
        private readonly Timer _timer;

        // UI
        private Label _lblScore;
        private Label _lblLength;
        private Label _lblState;
        private Label _lblSpeed;
        private Label _lblHelp;
        private Button _btnStart;
        private Button _btnPause;
        private Button _btnClear;
        private TextBox _txtLog;

        public SnakeGameForm()
        {
            _engine = new SnakeGameEngine(GridHalfExtent);
            _world = new SnakeWorld(CellSize, AppendLog);

            _timer = new Timer();
            _timer.Interval = InitialTickMs;
            _timer.Tick += TimerOnTick;

            InitUi();

            KeyPreview = true;
            KeyDown += Form_KeyDown;
            FormClosing += (s, e) => Cleanup();
            // 点击窗体任意空白区抢回焦点，确保方向键持续响应
            MouseClick += (s, e) => { if (!Focused) Focus(); };
        }

        private void InitUi()
        {
            Text = "TxTools · 贪吃蛇";
            Name = GetType().FullName;    // 记忆：Name 唯一化，避免 TxForm 持久化 key 串扰
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(460, 380);
            Size = new Size(480, 420);
            BackColor = SystemColors.Control;

            // 顶部信息 + 按钮区
            var top = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 100,
                ColumnCount = 4,
                RowCount = 3,
                Padding = new Padding(10, 8, 10, 8),
            };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
            top.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            top.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
            top.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));

            _lblScore = MakeLabel("得分: 0");
            _lblLength = MakeLabel("长度: 1");
            _lblState = MakeLabel("状态: 空闲");
            _lblSpeed = MakeLabel("速度: " + InitialTickMs + " ms");
            _lblHelp = MakeLabel("方向键 / WASD 控制，空格暂停");

            top.Controls.Add(_lblScore, 0, 0);
            top.Controls.Add(_lblLength, 1, 0);
            top.Controls.Add(_lblState, 2, 0);
            top.Controls.Add(_lblSpeed, 3, 0);
            top.Controls.Add(_lblHelp, 0, 1);
            top.SetColumnSpan(_lblHelp, 4);

            _btnStart = new Button { Text = "开始 / 重开", Dock = DockStyle.Fill };
            _btnStart.Click += (s, e) => StartNewGame();
            _btnPause = new Button { Text = "暂停 / 继续", Dock = DockStyle.Fill };
            _btnPause.Click += (s, e) => TogglePause();
            _btnClear = new Button { Text = "清除几何", Dock = DockStyle.Fill };
            _btnClear.Click += (s, e) => ClearAll();

            top.Controls.Add(_btnStart, 0, 2);
            top.Controls.Add(_btnPause, 1, 2);
            top.Controls.Add(_btnClear, 2, 2);

            // 日志区
            _txtLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.FromArgb(32, 32, 32),
                ForeColor = Color.Gainsboro,
                Font = new Font("Consolas", 9f),
                WordWrap = false,
            };
            var container = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10, 6, 10, 10) };
            container.Controls.Add(_txtLog);

            Controls.Add(container);
            Controls.Add(top);
        }

        private static Label MakeLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 9f, FontStyle.Regular),
            };
        }

        // ============ 游戏控制 ============

        private void StartNewGame()
        {
            try
            {
                _timer.Stop();
                _world.ClearAll();
                _engine.Reset();

                AppendLog("=== 新游戏 ===");
                AppendLog(string.Format(
                    "网格 [{0},{1}]×[{0},{1}], 格宽 {2:F0} mm, 食物起始 {3}",
                    -GridHalfExtent, GridHalfExtent, CellSize, _engine.FoodCell));

                // 场景原点生成蛇头
                _world.CreateSnakeBox(_engine.SnakeCells[0], "SnakeHead");
                // 生成食物
                _world.CreateFood(_engine.FoodCell);
                // 建立干涉集：head vs (food + body)
                _world.SetupCollisionPair();

                _engine.Start();
                _timer.Interval = InitialTickMs;
                _timer.Start();
                RefreshLabels();
                Focus();          // 确保方向键立即响应
            }
            catch (Exception ex)
            {
                AppendLog("StartNewGame 异常：" + ex.Message);
            }
        }

        private void TogglePause()
        {
            _engine.TogglePause();
            if (_engine.State == GameState.Running) _timer.Start();
            else _timer.Stop();
            RefreshLabels();
            Focus();          // 按钮点击后焦点回到窗体，方向键继续生效
        }

        private void ClearAll()
        {
            _timer.Stop();
            _world.ClearAll();
            _engine.Reset();
            RefreshLabels();
            AppendLog("已清除所有游戏几何。");
            Focus();          // 按钮点击后焦点回到窗体
        }

        private void Cleanup()
        {
            try
            {
                _timer.Stop();
                _world.ClearAll();
            }
            catch { }
        }

        // ============ 游戏循环 ============

        private void TimerOnTick(object sender, EventArgs e)
        {
            // 每帧确保焦点在窗体上，避免 PS 3D 视图抢走焦点后方向键失效
            if (!Focused) { try { Focus(); } catch { } }

            try
            {
                int result = _engine.Tick();

                if (result == -1)
                {
                    _timer.Stop();
                    RefreshLabels();
                    AppendLog("游戏结束！最终得分：" + _engine.Score);
                    MessageBox.Show(this,
                        "游戏结束！\r\n\r\n最终得分：" + _engine.Score + "\r\n蛇长：" + _engine.SnakeCells.Count,
                        "贪吃蛇", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (result == 1)
                {
                    // 吃到食物：
                    //   1) 引擎已把新 head 插入 [0]，长度 = 旧长度 + 1
                    //   2) 现有 SnakeBoxes 比 SnakeCells 少 1，我们在末尾 append 一个新 box
                    //      初始放在当前 tail cell（这个 cell 就是"未被删掉的旧尾"），随后被统一移动
                    //   3) 食物 box 平移到新食物位置（复用同一个 box，避免频繁 create/delete）
                    //   4) 新 box 加入干涉集 SecondList
                    int tailIdx = _engine.SnakeCells.Count - 1;
                    var tailCell = _engine.SnakeCells[tailIdx];
                    var newBox = _world.CreateSnakeBox(tailCell, "SnakeBody_" + tailIdx);
                    _world.AddToCollisionSecond(newBox);
                    _world.MoveFoodTo(_engine.FoodCell);

                    AppendLog(string.Format("吃到食物！长度 = {0}, 新食物 = {1}",
                        _engine.SnakeCells.Count, _engine.FoodCell));

                    // 逐段提速
                    if (_engine.Score > 0 && _engine.Score % SpeedupEveryScore == 0)
                    {
                        int nextInterval = Math.Max(MinTickMs, _timer.Interval - SpeedupStepMs);
                        if (nextInterval != _timer.Interval)
                        {
                            _timer.Interval = nextInterval;
                            AppendLog("速度提升 → " + nextInterval + " ms");
                        }
                    }
                }

                // 同步所有蛇 box 到当前网格坐标
                int n = Math.Min(_engine.SnakeCells.Count, _world.SnakeBoxes.Count);
                for (int i = 0; i < n; i++)
                {
                    _world.MoveSnakeBoxTo(i, _engine.SnakeCells[i]);
                }

                // 强制刷新 PS 3D 视图，使画面与移动同步
                try { TxApplication.RefreshDisplay(); } catch { }

                RefreshLabels();
            }
            catch (Exception ex)
            {
                _timer.Stop();
                AppendLog("Tick 异常：" + ex.Message);
            }
        }

        // ============ 输入 ============

        private void Form_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Up:
                case Keys.W:
                    _engine.PendingDir = Direction.Up; e.Handled = true; break;
                case Keys.Down:
                case Keys.S:
                    _engine.PendingDir = Direction.Down; e.Handled = true; break;
                case Keys.Left:
                case Keys.A:
                    _engine.PendingDir = Direction.Left; e.Handled = true; break;
                case Keys.Right:
                case Keys.D:
                    _engine.PendingDir = Direction.Right; e.Handled = true; break;
                case Keys.Space:
                    TogglePause();
                    e.Handled = true;
                    break;
            }
        }

        // ============ UI 辅助 ============

        private void RefreshLabels()
        {
            _lblScore.Text = "得分: " + _engine.Score;
            _lblLength.Text = "长度: " + _engine.SnakeCells.Count;
            _lblSpeed.Text = "速度: " + _timer.Interval + " ms";
            _lblState.Text = "状态: " + StateText(_engine.State);
        }

        private static string StateText(GameState s)
        {
            switch (s)
            {
                case GameState.Idle: return "空闲";
                case GameState.Running: return "运行中";
                case GameState.Paused: return "已暂停";
                case GameState.Over: return "游戏结束";
                default: return s.ToString();
            }
        }

        private void AppendLog(string msg)
        {
            if (_txtLog == null) return;
            try
            {
                if (_txtLog.InvokeRequired)
                {
                    _txtLog.BeginInvoke(new Action(() => AppendLog(msg)));
                    return;
                }
                var line = DateTime.Now.ToString("HH:mm:ss.fff") + "  " + msg + Environment.NewLine;
                _txtLog.AppendText(line);
            }
            catch { }
        }
    }
}
