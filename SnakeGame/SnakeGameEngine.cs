using System;
using System.Collections.Generic;

namespace TxTools.SnakeGame
{
    public enum Direction { Up, Down, Left, Right }
    public enum GameState { Idle, Running, Paused, Over }

    /// <summary>网格格子坐标（不含浮点，避免累积误差）。</summary>
    public struct Cell : IEquatable<Cell>
    {
        public int X;
        public int Y;
        public Cell(int x, int y) { X = x; Y = y; }
        public bool Equals(Cell other) { return X == other.X && Y == other.Y; }
        public override bool Equals(object obj) { return obj is Cell && Equals((Cell)obj); }
        public override int GetHashCode() { return (X * 397) ^ Y; }
        public static bool operator ==(Cell a, Cell b) { return a.Equals(b); }
        public static bool operator !=(Cell a, Cell b) { return !a.Equals(b); }
        public override string ToString() { return "(" + X + "," + Y + ")"; }
    }

    /// <summary>
    /// 贪吃蛇纯逻辑引擎：只处理网格坐标、方向、食物、越界与自碰撞，
    /// 不依赖 PS SDK，方便单独调试。
    /// </summary>
    public class SnakeGameEngine
    {
        private readonly int _halfExtent;     // 网格半边长；例如 10 → 21×21 网格 [-10,10]
        private readonly Random _rng = new Random();

        public List<Cell> SnakeCells { get; private set; }
        public Cell FoodCell { get; private set; }
        public Direction CurrentDir { get; private set; }
        public Direction PendingDir { get; set; }
        public GameState State { get; private set; }
        public int HalfExtent { get { return _halfExtent; } }
        public int Score { get { return Math.Max(0, SnakeCells.Count - 1); } }

        public SnakeGameEngine(int halfExtent)
        {
            _halfExtent = halfExtent;
            SnakeCells = new List<Cell>();
            State = GameState.Idle;
        }

        public void Reset()
        {
            SnakeCells.Clear();
            SnakeCells.Add(new Cell(0, 0));       // 起始位置：场景原点
            CurrentDir = Direction.Right;
            PendingDir = Direction.Right;
            SpawnFood();
            State = GameState.Idle;
        }

        public void Start()
        {
            if (SnakeCells.Count == 0) Reset();
            State = GameState.Running;
        }

        public void TogglePause()
        {
            if (State == GameState.Running) State = GameState.Paused;
            else if (State == GameState.Paused) State = GameState.Running;
        }

        /// <summary>
        /// 推进一格。返回值：
        ///   0 = 正常移动；
        ///   1 = 吃到食物（长度已 +1，新食物已刷新）；
        ///  -1 = 游戏结束（越界或撞自己）。
        /// </summary>
        public int Tick()
        {
            if (State != GameState.Running) return 0;

            // 应用待生效方向，禁止反向自杀
            if (!IsOpposite(PendingDir, CurrentDir))
                CurrentDir = PendingDir;

            var head = SnakeCells[0];
            var newHead = Next(head, CurrentDir);

            // 判断吃食物
            bool willEat = (newHead == FoodCell);

            // 自碰撞：如果吃食物则尾巴不动，检查全部；否则尾巴会腾出来，跳过最后一节
            int checkCount = willEat ? SnakeCells.Count : SnakeCells.Count - 1;
            for (int i = 0; i < checkCount; i++)
            {
                if (SnakeCells[i] == newHead)
                {
                    State = GameState.Over;
                    return -1;
                }
            }

            // 推进
            SnakeCells.Insert(0, newHead);
            if (willEat)
            {
                SpawnFood();
                return 1;
            }
            SnakeCells.RemoveAt(SnakeCells.Count - 1);
            return 0;
        }

        // ===== 内部工具 =====

        private void SpawnFood()
        {
            // 随机尝试
            for (int i = 0; i < 500; i++)
            {
                var c = new Cell(
                    _rng.Next(-_halfExtent, _halfExtent + 1),
                    _rng.Next(-_halfExtent, _halfExtent + 1));
                if (!SnakeCells.Contains(c)) { FoodCell = c; return; }
            }
            // 兜底：线性扫描一个空位
            for (int x = -_halfExtent; x <= _halfExtent; x++)
            {
                for (int y = -_halfExtent; y <= _halfExtent; y++)
                {
                    var c = new Cell(x, y);
                    if (!SnakeCells.Contains(c)) { FoodCell = c; return; }
                }
            }
            // 极端情况：蛇占满整场，视为胜利/游戏结束
            State = GameState.Over;
        }

        private static Cell Next(Cell c, Direction d)
        {
            switch (d)
            {
                case Direction.Up: return new Cell(c.X, c.Y + 1);
                case Direction.Down: return new Cell(c.X, c.Y - 1);
                case Direction.Left: return new Cell(c.X - 1, c.Y);
                case Direction.Right: return new Cell(c.X + 1, c.Y);
                default: return c;
            }
        }

        private static bool IsOpposite(Direction a, Direction b)
        {
            return (a == Direction.Up && b == Direction.Down)
                || (a == Direction.Down && b == Direction.Up)
                || (a == Direction.Left && b == Direction.Right)
                || (a == Direction.Right && b == Direction.Left);
        }
    }
}
