using System;
using System.Collections.Generic;

namespace TxTools.AutoPathPlanner
{
    /// <summary>
    /// 焊点顺序优化 (v6.0) —— 节拍收益最大的一项。
    ///
    /// 现状: 焊接顺序 = 焊点在操作里的原始顺序 (往往是画图/导入顺序, 极其绕)。
    /// 优化: 以**关节空间节拍**为边权跑 TSP 启发式, 重排焊接顺序。
    ///
    /// 算法: 最近邻构造初解 → 2-opt (段反转) → Or-opt (段搬移) 交替迭代至收敛。
    ///
    /// ═══════════════════════════════════════════════════════════════
    /// 为什么边权必须用关节代价而不是笛卡尔距离:
    ///   两个焊点 TCP 只差 100mm, 但需要翻腕 90° → 实际很慢
    ///   另两个焊点 TCP 差 600mm, 但只是 J1 转一点 → 实际很快
    /// 用欧氏距离排序会排出一个"看起来短、跑起来慢"的顺序。
    /// ═══════════════════════════════════════════════════════════════
    ///
    /// 工艺约束 (调用方可配):
    ///   LockFirst / LockLast — 锁定首/尾若干个焊点不参与重排
    ///                          (定位焊必须先焊的场景)
    /// </summary>
    public sealed class WeldOrderOptimizer
    {
        private readonly Action<string> _log;

        /// <summary>锁定开头 N 个焊点 (不参与重排) — 定位焊</summary>
        public int LockFirst = 0;

        /// <summary>锁定末尾 N 个焊点</summary>
        public int LockLast = 0;

        /// <summary>最大迭代轮次</summary>
        public int MaxRounds = 60;

        /// <summary>相对改进小于此值视为收敛</summary>
        public double ConvergeEps = 1e-4;

        public WeldOrderOptimizer(Action<string> log)
        {
            _log = log ?? delegate { };
        }

        /// <summary>
        /// v6.3 并行求解: 先把代价矩阵拿到手 (调用方在主线程算好, SDK 不可并行),
        /// 之后的 2-opt/Or-opt 是**纯数学**, 可以吃满多核。
        ///
        /// matrix[i,j] = 焊点 i → j 的节拍代价 (秒)。
        /// 多起点并行: 从不同起点跑最近邻+局部搜索, 取最优 —— 既提质量又用满 CPU。
        /// </summary>
        public int[] SolveParallel(double[,] matrix, int n, Func<int, double> homeCost)
        {
            if (n <= 3 || matrix == null) return Identity(n);

            int lo = Math.Max(0, LockFirst);
            int hi = n - Math.Max(0, LockLast);
            if (hi - lo <= 3)
            {
                _log("  [焊序] 可重排区间过小, 跳过");
                return Identity(n);
            }

            Func<int, int, double> cost = delegate (int a, int b) { return matrix[a, b]; };

            int[] orig = Identity(n);
            double costOrig = TourCost(orig, cost, homeCost);

            // ---- 多起点并行 ----
            // 每个起点独立跑 [最近邻构造 → 2-opt → Or-opt], 互不干扰 (纯数学, 线程安全)
            int starts = Math.Min(hi - lo, Environment.ProcessorCount * 2);
            var results = new int[starts][];
            var scores = new double[starts];

            System.Threading.Tasks.Parallel.For(0, starts, delegate (int s)
            {
                int seed = lo + s;
                int[] t = NearestNeighborFrom(n, lo, hi, seed, cost);
                double best = TourCost(t, cost, homeCost);

                for (int round = 0; round < MaxRounds; round++)
                {
                    double before = best;
                    best = TwoOpt(t, lo, hi, cost, homeCost, best);
                    best = OrOpt(t, lo, hi, cost, homeCost, best);
                    if (before - best < before * ConvergeEps) break;
                }

                results[s] = t;
                scores[s] = best;
            });

            // ---- 取最优 ----
            int bi = 0;
            for (int i = 1; i < starts; i++)
                if (scores[i] < scores[bi]) bi = i;

            int[] tour = results[bi];
            double bestScore = scores[bi];

            if (bestScore >= costOrig - 1e-9)
            {
                _log(string.Format(
                    "  [焊序] 原顺序已足够好 ({0:F1}s), {1} 个起点最优 {2:F1}s — 保持原序",
                    costOrig, starts, bestScore));
                return orig;
            }

            double save = costOrig - bestScore;
            _log(string.Format(
                "  [焊序] 重排完成: {0:F1}s → {1:F1}s (省 {2:F1}s, {3:P0}) [{4} 起点并行, {5} 核]",
                costOrig, bestScore, save, save / Math.Max(1e-9, costOrig),
                starts, Environment.ProcessorCount));

            if (LockFirst > 0 || LockLast > 0)
                _log(string.Format("  [焊序] 锁定 首{0} 尾{1} 个焊点不参与重排",
                    LockFirst, LockLast));

            return tour;
        }

        /// <summary>从指定起点做最近邻构造</summary>
        private static int[] NearestNeighborFrom(int n, int lo, int hi, int seed,
            Func<int, int, double> cost)
        {
            var tour = Identity(n);
            var pool = new List<int>();
            for (int i = lo; i < hi; i++)
                if (i != seed) pool.Add(i);

            int cur = seed;
            int w = lo;
            tour[w++] = cur;

            while (pool.Count > 0)
            {
                int bi = 0;
                double bc = double.MaxValue;
                for (int k = 0; k < pool.Count; k++)
                {
                    double c = cost(cur, pool[k]);
                    if (c < bc) { bc = c; bi = k; }
                }
                cur = pool[bi];
                pool.RemoveAt(bi);
                tour[w++] = cur;
            }
            return tour;
        }

        /// <summary>
        /// 求解新顺序。
        ///
        /// cost(i,j): 焊点 i → j 的节拍代价 (秒)。调用方用 CycleCost.PtpTime 提供,
        ///            拿不到关节值时可回退笛卡尔距离。
        /// homeCost(i): home → 焊点 i 的代价 (用于确定起点); 传 null 则以原首点为起点。
        ///
        /// 返回: 新的索引顺序 (长度 = n)。失败或无改进返回原顺序。
        /// </summary>
        public int[] Solve(int n, Func<int, int, double> cost, Func<int, double> homeCost)
        {
            if (n <= 3 || cost == null) return Identity(n);

            int lo = Math.Max(0, LockFirst);
            int hi = n - Math.Max(0, LockLast);   // [lo, hi) 为可重排区间
            if (hi - lo <= 3)
            {
                _log("  [焊序] 可重排区间过小, 跳过");
                return Identity(n);
            }

            // ---- 原顺序代价 (基线) ----
            int[] orig = Identity(n);
            double costOrig = TourCost(orig, cost, homeCost);

            // ---- 最近邻构造初解 (仅在可重排区间内) ----
            int[] tour = NearestNeighbor(n, lo, hi, cost, homeCost);

            // ---- 2-opt + Or-opt 交替 ----
            double best = TourCost(tour, cost, homeCost);
            int round = 0;
            for (; round < MaxRounds; round++)
            {
                double before = best;

                best = TwoOpt(tour, lo, hi, cost, homeCost, best);
                best = OrOpt(tour, lo, hi, cost, homeCost, best);

                if (before - best < before * ConvergeEps) break;
            }

            // ---- 只有确实更优才采纳 ----
            if (best >= costOrig - 1e-9)
            {
                _log(string.Format(
                    "  [焊序] 原顺序已足够好 ({0:F1}s), 优化结果 {1:F1}s — 保持原序",
                    costOrig, best));
                return orig;
            }

            double save = costOrig - best;
            _log(string.Format(
                "  [焊序] 重排完成: {0:F1}s → {1:F1}s (省 {2:F1}s, {3:P0}), {4} 轮收敛",
                costOrig, best, save, save / Math.Max(1e-9, costOrig), round + 1));

            if (LockFirst > 0 || LockLast > 0)
                _log(string.Format("  [焊序] 锁定 首{0} 尾{1} 个焊点不参与重排",
                    LockFirst, LockLast));

            return tour;
        }

        // ════════════════════════════════════════════════════════════

        private static int[] Identity(int n)
        {
            var a = new int[n];
            for (int i = 0; i < n; i++) a[i] = i;
            return a;
        }

        /// <summary>路径总代价 (含 home → 首点)</summary>
        private static double TourCost(int[] tour, Func<int, int, double> cost,
            Func<int, double> homeCost)
        {
            double t = 0;
            if (homeCost != null && tour.Length > 0) t += homeCost(tour[0]);
            for (int i = 0; i < tour.Length - 1; i++)
                t += cost(tour[i], tour[i + 1]);
            return t;
        }

        /// <summary>最近邻构造 (在 [lo,hi) 区间内, 从 lo 位置的原始点出发)</summary>
        private static int[] NearestNeighbor(int n, int lo, int hi,
            Func<int, int, double> cost, Func<int, double> homeCost)
        {
            var tour = Identity(n);
            var pool = new List<int>();
            for (int i = lo; i < hi; i++) pool.Add(tour[i]);

            // 起点: 若有 home 代价, 选离 home 最近的; 否则用原首点
            int startIdx = 0;
            if (homeCost != null && lo == 0)
            {
                double bestH = double.MaxValue;
                for (int k = 0; k < pool.Count; k++)
                {
                    double h = homeCost(pool[k]);
                    if (h < bestH) { bestH = h; startIdx = k; }
                }
            }

            int cur = pool[startIdx];
            pool.RemoveAt(startIdx);

            int w = lo;
            tour[w++] = cur;

            while (pool.Count > 0)
            {
                int bi = 0;
                double bc = double.MaxValue;
                for (int k = 0; k < pool.Count; k++)
                {
                    double c = cost(cur, pool[k]);
                    if (c < bc) { bc = c; bi = k; }
                }
                cur = pool[bi];
                pool.RemoveAt(bi);
                tour[w++] = cur;
            }

            return tour;
        }

        /// <summary>2-opt: 反转 [i+1, j] 段, 消除交叉</summary>
        private static double TwoOpt(int[] tour, int lo, int hi,
            Func<int, int, double> cost, Func<int, double> homeCost, double best)
        {
            bool improved = true;
            while (improved)
            {
                improved = false;
                for (int i = lo; i < hi - 2; i++)
                {
                    for (int j = i + 2; j < hi; j++)
                    {
                        Reverse(tour, i + 1, j);
                        double c = TourCost(tour, cost, homeCost);
                        if (c < best - 1e-9)
                        {
                            best = c;
                            improved = true;
                        }
                        else
                        {
                            Reverse(tour, i + 1, j);  // 回滚
                        }
                    }
                }
            }
            return best;
        }

        /// <summary>Or-opt: 把长度 1~3 的连续段搬到别处 (2-opt 覆盖不到的改进)</summary>
        private static double OrOpt(int[] tour, int lo, int hi,
            Func<int, int, double> cost, Func<int, double> homeCost, double best)
        {
            for (int segLen = 1; segLen <= 3; segLen++)
            {
                bool improved = true;
                while (improved)
                {
                    improved = false;
                    for (int i = lo; i + segLen <= hi; i++)
                    {
                        var seg = new int[segLen];
                        Array.Copy(tour, i, seg, 0, segLen);

                        for (int j = lo; j <= hi - segLen; j++)
                        {
                            if (j >= i - segLen && j <= i + segLen) continue;

                            int[] cand = MoveSegment(tour, lo, hi, i, segLen, j);
                            double c = TourCost(cand, cost, homeCost);
                            if (c < best - 1e-9)
                            {
                                Array.Copy(cand, tour, tour.Length);
                                best = c;
                                improved = true;
                                break;
                            }
                        }
                        if (improved) break;
                    }
                }
            }
            return best;
        }

        /// <summary>把 tour[at..at+len) 段搬到位置 to (返回新数组)</summary>
        private static int[] MoveSegment(int[] tour, int lo, int hi, int at, int len, int to)
        {
            var result = (int[])tour.Clone();

            var seg = new int[len];
            Array.Copy(tour, at, seg, 0, len);

            // 收集剩余 (仅可重排区间)
            var rest = new List<int>();
            for (int i = lo; i < hi; i++)
            {
                if (i >= at && i < at + len) continue;
                rest.Add(tour[i]);
            }

            // 插入位置换算到 rest 的下标
            int insertAt = to - lo;
            if (to > at) insertAt -= len;
            insertAt = Math.Max(0, Math.Min(rest.Count, insertAt));
            rest.InsertRange(insertAt, seg);

            for (int i = 0; i < rest.Count && lo + i < hi; i++)
                result[lo + i] = rest[i];

            return result;
        }

        private static void Reverse(int[] a, int i, int j)
        {
            while (i < j)
            {
                int t = a[i]; a[i] = a[j]; a[j] = t;
                i++; j--;
            }
        }
    }
}
