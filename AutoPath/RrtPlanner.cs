using System;
using System.Collections.Generic;

namespace TxTools.AutoPathPlanner
{
    /// <summary>
    /// RRT-Connect 双向树规划器 (纯算法层)。
    ///
    /// 相比单树RRT: 从起点和终点同时生长两棵树, 交替 [延伸一步 → 贪心互连],
    /// 相遇即成功。目标口袋由目标树自己探索, 两树在中间任意位置会师即可 ——
    /// 专治"树在长但连不到目标邻域"的窄通道失败模式。
    ///
    /// 接口与单树版兼容: Plan(start, goal, bounds)。
    /// </summary>
    public sealed class RrtPlanner
    {
        public double StepSize = 80.0;
        public double GoalTolerance = 80.0;   // 兼容保留 (Connect版不依赖)
        public double GoalBias = 0.15;        // 兼容保留 (Connect版无需目标偏置)
        public int MaxIterations = 3000;
        public double EdgeCheckResolution = 25.0;
        public int SmoothingAttempts = 60;

        /// <summary>停滞检测: 该迭代数后两树仍未生长则提前放弃</summary>
        public int StagnationCheckIters = 500;

        /// <summary>桥接停滞: 连续 N 次迭代桥接距离无改进(改进量 < 1mm) 则提前放弃</summary>
        public int StagnationBridgeIters = 800;

        /// <summary>最近一次 Plan 失败的原因 (成功时为 null)</summary>
        public string LastFailReason;

        /// <summary>
        /// 本次 Plan 探索到的自由状态 (两树全部节点, 上限截断)。
        /// 失败时同样有效 — 供经验中继池复用, 让探索成本跨段摊销。
        /// </summary>
        public List<Vec3> LastExploredFreeStates = new List<Vec3>();

        /// <summary>
        /// 种子状态 (经验中继池注入): Plan 开始时尝试挂接到两棵树,
        /// 把跨段探索成果直接变成树的初始枝干。
        /// </summary>
        public List<Vec3> SeedStates;

        /// <summary>状态有效性判定: true = 无碰撞可用</summary>
        public Func<Vec3, bool> IsStateFree;

        /// <summary>中止钩子: 返回true时 Plan 立即返回 null</summary>
        public Func<bool> ShouldAbort;

        /// <summary>进度回调 (迭代数, 两树节点总数)</summary>
        public Action<int, int> OnProgress;

        private readonly Random _rng;

        public RrtPlanner(int seed = 20260702) { _rng = new Random(seed); }

        private sealed class Node
        {
            public Vec3 P;
            public int Parent; // -1 = 根
        }

        public List<Vec3> Plan(Vec3 start, Vec3 goal, Aabb bounds)
        {
            if (IsStateFree == null)
                throw new InvalidOperationException("IsStateFree 未设置");
            LastFailReason = null;
            LastExploredFreeStates = new List<Vec3>();

            var treeStart = new List<Node> { new Node { P = start, Parent = -1 } }; // 根=起点
            var treeGoal = new List<Node> { new Node { P = goal, Parent = -1 } };   // 根=终点
            bool extendStartTree = true;

            _corrStart = start; _corrGoal = goal;
            _bridgeMid = Vec3.Lerp(start, goal, 0.5);
            _bridgeDist = Vec3.Distance(start, goal);
            _failStreak = 0;
            _stagnationCounter = 0;
            _lastBridgeDist = _bridgeDist;

            // ---- 种子注入: 中继池点挂接到最近的树节点 ----
            if (SeedStates != null)
            {
                int attached = 0;
                foreach (var seed in SeedStates)
                {
                    if (attached >= 12) break;
                    if (TryAttachSeed(treeStart, seed) || TryAttachSeed(treeGoal, seed))
                        attached++;
                }
            }

            try
            {
                for (int iter = 0; iter < MaxIterations; iter++)
                {
                    if (ShouldAbort != null && ShouldAbort())
                    {
                        LastFailReason = "用户中止";
                        return null;
                    }
                    if (iter == StagnationCheckIters
                        && treeStart.Count + treeGoal.Count <= 2)
                    {
                        LastFailReason = string.Format(
                            "停滞: {0}次迭代两树零生长 (端点邻域被包围或姿态无法逆解)", iter);
                        return null;
                    }
                    // 桥接停滞: 连续 N 次迭代桥接距离无改进 → 提前放弃
                    if (iter > StagnationBridgeIters && _stagnationCounter > StagnationBridgeIters / 2)
                    {
                        LastFailReason = string.Format(
                            "桥接停滞: 连续{0}次迭代桥接距离无改进 (双树无法靠近)", _stagnationCounter);
                        return null;
                    }
                    if (OnProgress != null && iter % 200 == 0)
                        OnProgress(iter, treeStart.Count + treeGoal.Count);

                    var ext = extendStartTree ? treeStart : treeGoal;
                    var oth = extendStartTree ? treeGoal : treeStart;

                    // ---- 1. 延伸树 ext 一步 (混合采样) ----
                    Vec3 sample = MixedSample(bounds);
                    int nExt = Extend(ext, sample);

                    if (nExt >= 0)
                    {
                        UpdateBridge(ext[nExt].P, oth);

                        // ---- 2. 贪心互连: 树 oth 反复向新节点推进 ----
                        bool reached;
                        int nOth = Connect(oth, ext[nExt].P, out reached);

                        if (reached)
                        {
                            // 会师: 组装 start树(根→会师点) + goal树(会师点→根)
                            int meetStart = extendStartTree ? nExt : nOth;
                            int meetGoal = extendStartTree ? nOth : nExt;

                            var pathS = Backtrack(treeStart, meetStart); // start → 会师
                            var pathG = Backtrack(treeGoal, meetGoal);   // goal → 会师
                            pathG.Reverse();                             // 会师 → goal

                            var full = new List<Vec3>(pathS);
                            for (int i = 1; i < pathG.Count; i++)        // 跳过重复会师点
                                full.Add(pathG[i]);

                            return SmoothThenSubdivide(full);
                        }
                    }

                    extendStartTree = !extendStartTree; // 交替生长
                }

                LastFailReason = "迭代耗尽未会师";
                return null;
            }
            finally
            {
                // 无论成败, 收割两树自由节点供中继池复用 (均匀截断至200)
                HarvestFreeStates(treeStart);
                HarvestFreeStates(treeGoal);
            }
        }

        /// <summary>向目标延伸一步; 成功返回新节点索引, 失败-1</summary>
        private int Extend(List<Node> tree, Vec3 target)
        {
            int nearest = 0;
            double bestDist = double.MaxValue;
            for (int i = 0; i < tree.Count; i++)
            {
                double d = Vec3.Distance(tree[i].P, target);
                if (d < bestDist) { bestDist = d; nearest = i; }
            }

            Vec3 from = tree[nearest].P;
            Vec3 dir = target - from;
            double dist = dir.Length();
            if (dist < 1e-6) return -1;

            // 自适应步长: 连续受挫 → 缩步以钻窄通道 (下限20mm)
            double step = _failStreak >= 50 ? Math.Max(20.0, StepSize * 0.5) : StepSize;
            Vec3 newP = dist <= step ? target : from + dir.Normalized() * step;

            if (!IsStateFree(newP)) { _failStreak++; return -1; }
            if (!IsEdgeFree(from, newP)) { _failStreak++; return -1; }

            _failStreak = 0;
            tree.Add(new Node { P = newP, Parent = nearest });
            return tree.Count - 1;
        }

        // ---- 混合采样基础设施 ----
        private Vec3 _corrStart, _corrGoal, _bridgeMid;
        private double _bridgeDist;
        private int _failStreak;
        private int _stagnationCounter;
        private double _lastBridgeDist;

        /// <summary>
        /// 混合采样: 40% 均匀盒内 (全局探索) / 30% 走廊偏置 (start→goal 圆柱内,
        /// 高效覆盖最可能的通道区域) / 30% 桥接偏置 (双树最近对中点邻域,
        /// 直接攻击两树间的缺口)。
        /// </summary>
        private Vec3 MixedSample(Aabb bounds)
        {
            double u = _rng.NextDouble();
            if (u < 0.4)
                return bounds.Sample(_rng);

            if (u < 0.7)
            {
                // 走廊: 轴向均匀 + 径向高斯 (σ = 2×步长)
                Vec3 axial = Vec3.Lerp(_corrStart, _corrGoal, _rng.NextDouble());
                double sigma = StepSize * 2.0;
                return new Vec3(
                    axial.X + Gaussian() * sigma,
                    axial.Y + Gaussian() * sigma,
                    axial.Z + Gaussian() * sigma);
            }

            // 桥接: 最近对中点邻域 (σ = 步长, 但不小于缺口的一半)
            double s = Math.Max(StepSize, _bridgeDist * 0.5);
            return new Vec3(
                _bridgeMid.X + Gaussian() * s,
                _bridgeMid.Y + Gaussian() * s,
                _bridgeMid.Z + Gaussian() * s);
        }

        /// <summary>新节点入树后更新双树最近对 (桥接采样锚点) + 停滞跟踪</summary>
        private void UpdateBridge(Vec3 newNode, List<Node> otherTree)
        {
            for (int i = 0; i < otherTree.Count; i++)
            {
                double d = Vec3.Distance(newNode, otherTree[i].P);
                if (d < _bridgeDist)
                {
                    _bridgeDist = d;
                    _bridgeMid = Vec3.Lerp(newNode, otherTree[i].P, 0.5);
                }
            }
            // 停滞跟踪: 桥接距离改进 < 1mm → 计数累加, 否则清零
            if (_bridgeDist < _lastBridgeDist - 1.0)
                _stagnationCounter = 0;
            else
                _stagnationCounter++;
            _lastBridgeDist = _bridgeDist;
        }

        /// <summary>标准正态随机数 (Box-Muller)</summary>
        private double Gaussian()
        {
            double u1 = 1.0 - _rng.NextDouble();
            double u2 = _rng.NextDouble();
            return Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Sin(2.0 * Math.PI * u2);
        }

        /// <summary>种子挂接: 与树上最近节点连边可行则入树</summary>
        private bool TryAttachSeed(List<Node> tree, Vec3 seed)
        {
            int nearest = 0;
            double best = double.MaxValue;
            for (int i = 0; i < tree.Count; i++)
            {
                double d = Vec3.Distance(tree[i].P, seed);
                if (d < best) { best = d; nearest = i; }
            }
            if (best > StepSize * 6) return false; // 太远, 挂接边验证不划算
            if (!IsStateFree(seed)) return false;
            if (!IsEdgeFree(tree[nearest].P, seed)) return false;
            tree.Add(new Node { P = seed, Parent = nearest });
            return true;
        }

        /// <summary>贪心互连: 反复向 target 推进直至到达或受阻</summary>
        private int Connect(List<Node> tree, Vec3 target, out bool reached)
        {
            reached = false;
            int last = -1;
            for (int guard = 0; guard < 256; guard++) // 防死循环
            {
                int n = Extend(tree, target);
                if (n < 0) return last; // 受阻
                last = n;
                if (Vec3.Distance(tree[n].P, target) < 1e-6)
                {
                    reached = true;
                    return n;
                }
            }
            return last;
        }

        /// <summary>连边离散碰撞检测</summary>
        public bool IsEdgeFree(Vec3 a, Vec3 b)
        {
            double len = Vec3.Distance(a, b);
            int steps = Math.Max(1, (int)Math.Ceiling(len / EdgeCheckResolution));
            for (int s = 1; s <= steps; s++)
            {
                if (!IsStateFree(Vec3.Lerp(a, b, (double)s / steps)))
                    return false;
            }
            return true;
        }

        private static List<Vec3> Backtrack(List<Node> tree, int leaf)
        {
            var path = new List<Vec3>();
            int cur = leaf;
            while (cur >= 0)
            {
                path.Add(tree[cur].P);
                cur = tree[cur].Parent;
            }
            path.Reverse(); // 根在前
            return path;
        }

        /// <summary>平滑后长边细分 (流水线: Smooth → Subdivide)</summary>
        private List<Vec3> SmoothThenSubdivide(List<Vec3> path)
        {
            var smoothed = Smooth(path);
            return SubdivideLongEdges(smoothed);
        }

        /// <summary>shortcut 平滑: 可直连的两点间剔除中间节点</summary>
        private List<Vec3> Smooth(List<Vec3> path)
        {
            if (path.Count <= 2) return path;
            var pts = new List<Vec3>(path);

            for (int attempt = 0; attempt < SmoothingAttempts && pts.Count > 2; attempt++)
            {
                int i = _rng.Next(0, pts.Count - 2);
                int j = _rng.Next(i + 2, pts.Count);
                if (IsEdgeFree(pts[i], pts[j]))
                    pts.RemoveRange(i + 1, j - i - 1);
            }
            return pts;
        }

        /// <summary>长边细分: 对平滑后仍超过 maxEdgeLen 的边段自动插入中间点,
        /// 确保机器人运动不会跳过大距离 (每段 ≤ maxEdgeLen)。
        /// 插入点仅需位置有效性检查 (姿态在后续 BuildTransitPoints 中择优锁定)。
        /// </summary>
        public double MaxEdgeLength = 200.0;

        private List<Vec3> SubdivideLongEdges(List<Vec3> path)
        {
            if (path.Count <= 2 || MaxEdgeLength <= 0) return path;
            var result = new List<Vec3>();
            result.Add(path[0]);
            for (int i = 1; i < path.Count; i++)
            {
                double dist = Vec3.Distance(path[i - 1], path[i]);
                if (dist > MaxEdgeLength)
                {
                    int nSteps = (int)Math.Ceiling(dist / MaxEdgeLength);
                    for (int s = 1; s < nSteps; s++)
                    {
                        Vec3 mid = Vec3.Lerp(path[i - 1], path[i], (double)s / nSteps);
                        // 只做位置有效性检查 (姿态由 WeldPathPlanner.BuildTransitPoints 择优)
                        if (IsStateFree(mid))
                            result.Add(mid);
                    }
                }
                result.Add(path[i]);
            }
            return result;
        }

        private void HarvestFreeStates(List<Node> tree)
        {
            int cap = 100; // 每棵树上限
            int stride = Math.Max(1, tree.Count / cap);
            for (int i = 1; i < tree.Count; i += stride) // 跳过根(即端点本身)
                LastExploredFreeStates.Add(tree[i].P);
        }
    }
}
