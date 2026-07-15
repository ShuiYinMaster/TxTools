using System;
using System.Collections.Generic;

namespace TxTools.AutoPathPlanner
{
    /// <summary>
    /// 枪坐标系定向过渡搜索 (领域知识版树生长)。
    /// 坐标约定 (WeldSpotAllocator 确认): X=前后, Y=左右, Z=上下。
    ///
    /// 语义澄清 (v5.4 — 之前混淆过, 这里写死):
    ///   进/出枪点 = 沿枪 **-Z**, 把枪嘴从板件上提起 (20mm 量级即可)
    ///   后撤点   = 沿枪 **-X**, 把整个枪体从夹具里抽出来 (可能需要 300~800mm!)
    ///   两者完全不同。**抬升/横移之前必须先后撤到枪体脱离夹具**。
    ///
    /// 生长顺序 (模拟人工示教挪枪):
    ///   1. 后撤: 沿 -X 大步进逐步检查, 撤不动即为最大深度
    ///   2. 门形直连: start → A后撤 → B后撤 → goal (纯 -X 平移, 枪体沿轴抽出, 最安全)
    ///   3. 不通则后撤 + 竖直检索: +Z(上)优先, -Z(下)其次
    ///   4. 再后撤 + 横向: ±Y
    ///
    /// ═══════════════════════════════════════════════════════════════
    /// v5.4 关键修复 (症状: "后撤40mm就抬升 → 抬升时枪体扫过夹具")
    ///
    ///   ① 侧向机动最小后撤门槛 (MinBackoutForSide)
    ///      原来 dA 从 BackoutStep(40) 由浅到深, 40mm 就开始试抬升 —— 枪体还整个
    ///      埋在夹具里。IsEdgeFree 只采样 TCP 点位, 看不见枪体, 于是判"通过";
    ///      到阶段2关节扫掠才发现撞, 此时修复很被动。
    ///      现在: 侧向机动 (抬升/横移/对角) 必须 dA >= MinBackoutForSide。
    ///      门形直连 (纯 -X) 不受此限 —— 沿枪轴抽出不会扫掠。
    ///
    ///   ② 动态扫掠验证前移 (IsEdgeSafeDynamic)
    ///      候选路径组装完成后, 用**关节扫掠**再复验一遍 —— 捕捉 IsEdgeFree
    ///      看不见的枪体扫掠。不通过就继续加深后撤, 而不是等阶段2去修。
    ///      点查询仍用便宜的静态检查, 只有最终候选才做昂贵动态验证, 性能可控。
    ///
    ///   ③ 深度上限放宽 (BackoutMax 320 → 800), 步进 40 → 60
    ///
    ///   ④ 侧向机动优先深后撤 (PreferDeepBackoutForSide)
    ///      侧向检索的深度由**深到浅**遍历, 优先返回后撤充分的解。
    /// ═══════════════════════════════════════════════════════════════
    /// </summary>
    public sealed class GunFrameSearch
    {
        public double BackoutStep = 60.0;    // 后撤步进 (v5.4: 40 → 60)
        public double BackoutMax = 800.0;    // 最大后撤深度 (v5.4: 320 → 800)
        public double SideStep = 60.0;       // 竖直/横向检索步进
        public double SideMax = 360.0;       // 竖直/横向最大偏移

        /// <summary>
        /// v5.4 侧向机动 (抬升/横移/对角) 的最小后撤门槛 (mm)。
        /// 后撤不足此值时不允许任何侧向动作 —— 枪体尚未脱离夹具, 侧移必然扫掠。
        /// 设 0 关闭该约束 (退回旧行为)。
        /// </summary>
        public double MinBackoutForSide = 200.0;

        /// <summary>v5.4 侧向检索深度由深到浅遍历 (优先返回后撤充分的解)</summary>
        public bool PreferDeepBackoutForSide = true;

        public Func<Vec3, bool> IsFree;
        public Func<Vec3, Vec3, bool> IsEdgeFree;

        /// <summary>
        /// v5.4 可选: 动态关节扫掠边验证 (a→b 关节插值全程无干涉)。
        /// 候选路径组装完成后逐边复验, 捕捉 IsEdgeFree 看不见的**枪体扫掠**。
        /// null = 不做动态验证 (退回旧行为)。
        /// </summary>
        public Func<Vec3, Vec3, bool> IsEdgeSafeDynamic;

        public Func<bool> ShouldAbort;
        public Action<string> Log = delegate { };

        /// <summary>可选: 受阻点的状态描述 (无逆解/有干涉), 用于诊断日志</summary>
        public Func<Vec3, string> DescribeBlock;

        /// <summary>动态验证预算 (昂贵, 限次)</summary>
        public int DynamicVerifyBudget = 24;

        private int _dynUsed;
        private int _dynRejected;

        /// <summary>
        /// 定向搜索: 成功返回中间过渡点序列, 失败返回 null。
        ///
        /// 三阶段:
        ///   阶段1 - 深度剖面: A/B 两侧各自独立扫描最大无阻后撤深度 (纯点查询)
        ///   阶段2a - 门形直连 (纯 -X, 由浅到深, 短路径优先)
        ///   阶段2b - 后撤 + 侧向机动 (受 MinBackoutForSide 门槛, 由深到浅)
        ///   候选路径组装后做 [静态连边 → 动态关节扫掠] 双重验证
        /// </summary>
        public List<Vec3> Search(Vec3 start, Vec3 goal,
            Vec3 backA, Vec3 backB, Vec3 zAxis, Vec3 yAxis)
        {
            _dynUsed = 0;
            _dynRejected = 0;

            // ---- 阶段1: 双侧深度剖面 ----
            double maxA = ScanBackout(start, backA, "A");
            double maxB = ScanBackout(goal, backB, "B");
            if (maxA < BackoutStep && maxB < BackoutStep)
            {
                Log("      定向: 双侧均无后撤空间");
                return null;
            }

            // 检索方向: 上 → 下 → 左 → 右 → 上左对角 → 上右对角
            Vec3 diag1 = (zAxis + yAxis).Normalized();
            Vec3 diag2 = (zAxis + yAxis * -1.0).Normalized();
            Vec3[] sideDirs = { zAxis, zAxis * -1.0, yAxis, yAxis * -1.0, diag1, diag2 };
            string[] sideNames = { "上", "下", "左", "右", "对角上左", "对角上右" };

            int budget = 60; // 候选路径静态连边验证预算

            // ---- 阶段2a: 门形直连 (纯 -X 后撤平移, 不受侧向门槛限制) ----
            // 沿枪轴抽出 → 平移 → 沿枪轴插入。枪体沿自身轴线运动, 不扫掠周边。
            // 由浅到深 (短路径优先, 节拍友好)。
            for (double dA = BackoutStep; dA <= maxA + 1e-6; dA += BackoutStep)
            {
                if (Abort()) return null;
                Vec3 pa = start + backA * dA;

                foreach (double dB in BuildDepthCandidates(dA, maxB))
                {
                    if (Abort()) return null;
                    Vec3 pb = goal + backB * dB;

                    List<Vec3> path;
                    if (budget-- > 0 && TryAssemble(start, goal, out path, pa, pb))
                    {
                        Log(string.Format("      定向: 门形直连@后撤A{0:F0}/B{1:F0}mm 通过", dA, dB));
                        LogDynSummary();
                        return path;
                    }
                }
            }

            // ---- 阶段2b: 后撤 + 侧向机动 ----
            // v5.4: 侧向机动必须先后撤到 MinBackoutForSide 以上, 且优先深后撤。
            var depths = BuildSideDepths(maxA);
            if (depths.Count == 0)
            {
                Log(string.Format(
                    "      定向: A侧最大后撤 {0:F0}mm < 侧向门槛 {1:F0}mm — 枪体无法脱困, 侧向机动跳过",
                    maxA, MinBackoutForSide));
                LogDynSummary();
                return null;
            }

            foreach (double dA in depths)
            {
                if (Abort()) return null;
                Vec3 pa = start + backA * dA;

                foreach (double dB in BuildDepthCandidatesForSide(dA, maxB))
                {
                    if (Abort()) return null;
                    Vec3 pb = goal + backB * dB;

                    for (int s = 0; s < sideDirs.Length; s++)
                    {
                        bool dirBlocked = false;
                        for (double h = SideStep; h <= SideMax && !dirBlocked; h += SideStep)
                        {
                            if (Abort()) return null;

                            Vec3 qa = pa + sideDirs[s] * h;
                            if (!IsFree(qa)) { dirBlocked = true; continue; }
                            Vec3 qb = pb + sideDirs[s] * h;
                            if (!IsFree(qb)) continue;

                            List<Vec3> path;
                            if (budget-- > 0 && TryAssemble(start, goal, out path, pa, qa, qb, pb))
                            {
                                Log(string.Format("      定向: 后撤A{0:F0}/B{1:F0} + {2}{3:F0}mm 通过",
                                    dA, dB, sideNames[s], h));
                                LogDynSummary();
                                return path;
                            }
                            if (budget <= 0)
                            {
                                Log("      定向: 预算耗尽");
                                LogDynSummary();
                                return null;
                            }
                        }
                    }
                }
            }

            Log(string.Format("      定向搜索未通 (剖面 A={0:F0}mm B={1:F0}mm)", maxA, maxB));
            LogDynSummary();
            return null;
        }

        /// <summary>
        /// v5.4 侧向机动的深度候选: 必须 >= MinBackoutForSide。
        /// PreferDeepBackoutForSide=true 时由深到浅 (优先充分后撤的解)。
        /// </summary>
        private List<double> BuildSideDepths(double maxA)
        {
            var list = new List<double>();
            double floor = Math.Max(BackoutStep, MinBackoutForSide);
            for (double d = floor; d <= maxA + 1e-6; d += BackoutStep)
                list.Add(d);

            // A侧最大深度本身也是候选 (可能不落在步进网格上)
            if (maxA >= floor
                && (list.Count == 0 || Math.Abs(list[list.Count - 1] - maxA) > 1e-6))
                list.Add(maxA);

            if (PreferDeepBackoutForSide) list.Reverse();
            return list;
        }

        /// <summary>阶段1: 沿后撤方向逐步扫描, 返回最大无阻深度</summary>
        private double ScanBackout(Vec3 origin, Vec3 back, string side)
        {
            double reached = 0;
            for (double d = BackoutStep; d <= BackoutMax; d += BackoutStep)
            {
                if (Abort()) return reached;
                Vec3 p = origin + back * d;
                if (!IsFree(p))
                {
                    string why = DescribeBlock != null ? " [" + DescribeBlock(p) + "]" : "";
                    Log(string.Format("      定向: {0}侧后撤剖面 {1:F0}mm 受阻{2}", side, d, why));
                    break;
                }
                reached = d;
            }
            if (reached > 0)
                Log(string.Format("      定向: {0}侧后撤可达 {1:F0}mm", side, reached));
            return reached;
        }

        /// <summary>门形直连的B侧深度候选: 与A对称 → B侧最深 → 最浅 (去重)</summary>
        private double[] BuildDepthCandidates(double dA, double maxB)
        {
            var list = new List<double>();
            double sym = Math.Min(dA, maxB);
            if (sym >= BackoutStep) list.Add(sym);
            if (maxB >= BackoutStep && !list.Contains(maxB)) list.Add(maxB);
            if (BackoutStep <= maxB && !list.Contains(BackoutStep)) list.Add(BackoutStep);
            if (list.Count == 0) list.Add(0); // B侧无空间: pb=goal 本身
            return list.ToArray();
        }

        /// <summary>
        /// v5.4 侧向机动的B侧深度候选: 同样受 MinBackoutForSide 约束。
        /// B侧退不够门槛时退到能退的最深 (总比不退好), 最终由动态验证把关。
        /// </summary>
        private double[] BuildDepthCandidatesForSide(double dA, double maxB)
        {
            var list = new List<double>();
            double floor = Math.Max(BackoutStep, MinBackoutForSide);

            double sym = Math.Min(dA, maxB);
            if (sym >= floor) list.Add(sym);
            if (maxB >= floor && !list.Contains(maxB)) list.Add(maxB);
            if (list.Count == 0 && maxB >= BackoutStep) list.Add(maxB);
            if (list.Count == 0) list.Add(0);
            return list.ToArray();
        }

        /// <summary>
        /// 组装候选路径 → 静态连边验证 → 动态关节扫掠复验。
        /// v5.4: 静态通过后还要过动态这一关, 捕捉枪体扫掠。
        /// </summary>
        private bool TryAssemble(Vec3 start, Vec3 goal, out List<Vec3> mids, params Vec3[] waypoints)
        {
            mids = null;

            // ---- 静态连边 (便宜, TCP 点位采样) ----
            Vec3 prev = start;
            foreach (var w in waypoints)
            {
                if (!IsEdgeFree(prev, w)) return false;
                prev = w;
            }
            if (!IsEdgeFree(prev, goal)) return false;

            // ---- 动态关节扫掠复验 (昂贵, 限预算) ----
            // v5.4 核心: IsEdgeFree 只看 TCP, 看不见枪体扫过夹具。
            // 关节扫掠把整个枪体的运动包络算进去。
            if (IsEdgeSafeDynamic != null && _dynUsed < DynamicVerifyBudget)
            {
                _dynUsed++;
                prev = start;
                foreach (var w in waypoints)
                {
                    if (!IsEdgeSafeDynamic(prev, w))
                    {
                        _dynRejected++;
                        return false;   // 静态通过但动态撞 → 拒绝, 继续加深后撤
                    }
                    prev = w;
                }
                if (!IsEdgeSafeDynamic(prev, goal))
                {
                    _dynRejected++;
                    return false;
                }
            }

            mids = new List<Vec3>(waypoints);
            return true;
        }

        private bool Abort()
        {
            return ShouldAbort != null && ShouldAbort();
        }

        private void LogDynSummary()
        {
            if (IsEdgeSafeDynamic == null || _dynUsed == 0) return;
            if (_dynRejected > 0)
                Log(string.Format("      定向: 动态复验 {0} 次, 拒绝 {1} 条静态可行但枪体扫掠的候选",
                    _dynUsed, _dynRejected));
        }
    }
}
