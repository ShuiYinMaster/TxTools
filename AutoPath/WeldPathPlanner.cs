using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Olp;

namespace TxTools.AutoPathPlanner
{
    public enum PlanPointKind { Approach, Weld, Retract, Transit } // FallbackGuess 已移除: 失败段不再插点

    public sealed class PlanPoint
    {
        public Vec3 Position;
        public TxVector RpyZyx;               // 弧度 (兜底)
        public TxTransformation AbsTransform; // 完整位姿 (整矩阵姿态传递优先)
        public PlanPointKind Kind;
        public string Source = "";
        public string CustomName; // 非空则Via直接用此名 (如 "home")
        public bool Verified;
        public string Note = "";
    }

    public sealed class PlanningReport
    {
        public int OperationCount;
        public int WeldCount;
        public int InsertedViaCount;
        public int DirectHits;
        public int QuickPlanHits;
        public int RelayHits;
        public int RrtInvocations;
        public int RrtSuccesses;
        public int FailedSegments;
        public int ClearanceSegments; // 净空绕行段 (端点干涉但已生成高位中继)
        public int DynamicViolations;
        public int DynamicRepairs;
        public int RefineRounds;
        public int PrunedVias;        // v6.0: 捷径化删除的过渡点数
        public int ReorderedOps;      // v6.0: 焊序被重排的操作数
        public int GunOpeningsWritten;// v6.0: 写入开口的 Via 数
        public int CollisionQueries;
        public TimeSpan Elapsed;
        public List<string> Warnings = new List<string>();
    }

    /// <summary>
    /// 编排层 v2: 在【所选焊接操作内部】进行过渡点规划与原位插入。
    ///
    /// 与 v1 的区别:
    ///   - 不再全场景扫描焊点、不再新建大操作
    ///   - 对每个选中操作: 按树序取其焊点子项，只在相邻焊点间规划
    ///   - 过渡 Via 直接创建在该操作内，并重排到正确的序列位置
    ///
    /// 过渡段规划级联不变: L0直连 → L1中点抬升 → L2 RRT → L3兜底
    /// </summary>
    public sealed class WeldPathPlanner
    {
        // ---- 可配置参数 ----
        /// <summary>进/出枪最大距离 (沿枪-Z, 默认20mm)</summary>
        public double ApproachRetractDistance = 20.0;
        /// <summary>进/出枪最小距离 (区间下限, 默认10mm)</summary>
        public double ApproachRetractMin = 10.0;
        public bool UseWorldZForApproach = false;
        public bool GenerateApproachRetract = true;
        public double RrtStepSize = 50.0;      // 精度提升 (80→50)
        public int RrtMaxIterations = 5000;    // 迭代上限提升 (3000→5000)
        public double RrtGoalBias = 0.15;
        public double EdgeCheckResolution = 12.0; // 连边检测加密 (25→12), 防薄板隧穿
        // 枪坐标系定向搜索 (L1)
        public double GunBackoutStep = 60.0;   // v5.4: 40 → 60
        public double GunBackoutMax = 800.0;   // L1 后撤搜索面 (320→800)
        public double GunSideStep = 60.0;
        public double GunSideMax = 800.0;      // L1 竖直/横向搜索面 (360→800)

        /// <summary>
        /// v5.4 侧向机动 (抬升/横移) 的最小后撤门槛 (mm)。
        /// 后撤不足此值不允许抬升 —— 枪体还埋在夹具里, 一抬就扫。
        /// 症状: 日志出现 "后撤A40/B40 + 上60mm 通过", 静态过了但动态撞。
        /// </summary>
        public double GunMinBackoutForSide = 200.0;
        /// <summary>动态干涉检查 (成功段最终链的关节扫掠终验)</summary>
        public bool DynamicCheckEnabled = true;

        // ---- v5.5: 以下参数透传给 CollisionWorld (原本只能改代码) ----
        /// <summary>关节步进量子(°): 扫掠步数 = 关节最大跨度/此值</summary>
        public double DynamicJointQuantum = 4.0;
        /// <summary>笛卡尔量子(mm): 扫掠步数 = TCP位移/此值, 与关节判据取大者</summary>
        public double DynamicCartesianQuantum = 15.0;
        /// <summary>单边扫掠采样上限</summary>
        public int MaxSweepSteps = 64;
        /// <summary>构型突变阈值(°)</summary>
        public double ConfigJumpThreshold = 120.0;
        /// <summary>姿态变体搜索开关</summary>
        public bool OrientationVariantsEnabled = true;
        /// <summary>每点位姿态变体尝试上限</summary>
        public int MaxVariantTries = 13;
        /// <summary>共线剪枝开关 (v6.0 起为"路径捷径化"开关)</summary>
        public bool PruneEnabled = true;

        // ---- v6.0 节拍优化 ----
        /// <summary>焊点顺序优化 (2-opt/Or-opt, 关节代价) — 会改变焊接顺序!</summary>
        public bool WeldOrderOptEnabled = false;
        /// <summary>焊序优化: 锁定开头 N 个焊点 (定位焊)</summary>
        public int WeldOrderLockFirst = 0;
        /// <summary>焊序优化: 锁定末尾 N 个焊点</summary>
        public int WeldOrderLockLast = 0;

        // ---- v6.0 焊钳外部轴 ----
        /// <summary>把焊钳开口写入 Via 点的外部轴 (需焊钳注册为外部轴)</summary>
        public bool GunAxisWriteEnabled = true;
        /// <summary>进/出枪点的目标开口 (mm) — 够越过翻边即可</summary>
        public double TargetGunOpening = 30.0;
        /// <summary>过渡点按需最小开口搜索 (关闭则统一用 TransitGunOpening)</summary>
        public bool AdaptiveGunOpening = true;
        /// <summary>过渡点开口 (AdaptiveGunOpening=false 时使用)</summary>
        public double TransitGunOpening = 60.0;

        /// <summary>
        /// v6.4 焊钳开口方向手动覆盖 (自动探测失败时的兜底):
        ///   0 = 自动探测 / +1 = 强制正向 / -1 = 强制负向
        /// </summary>
        public int GunOpenDirectionOverride = 0;

        /// <summary>v6.4 手动指定最大开口幅值 (mm); &lt;=0 = 用探测/默认值</summary>
        public double GunMaxOpeningOverride = 0;

        // ---- v6.0 运动参数 ----
        /// <summary>过渡点运动类型设为 Joint (PTP, 各关节最高效) — 节拍关键</summary>
        public bool SetTransitMotionJoint = true;

        // ---- v6.3 性能 ----
        /// <summary>位姿查询缓存 (碰撞查询是唯一瓶颈, 且 SDK 无法并行 → 只能靠"少查")</summary>
        public bool QueryCacheEnabled = true;
        /// <summary>缓存位置量化精度 (mm)</summary>
        public double CacheQuantum = 0.5;

        public double SampleBoundsInflateXy = 300.0;
        public double SampleBoundsInflateZUp = 400.0;
        public double SampleBoundsInflateZDown = 100.0;

        private readonly Action<string> _log;
        private CollisionWorld _world;
        private RrtPlanner _rrt;
        private CycleCost _cost;        // v6.0 关节节拍代价
        private int _viaOpeningCount;   // v6.0 已写开口的 Via 计数

        /// <summary>v6.3 进度上报 (可选)</summary>
        public PlanningProgress Progress;
        private PlanningProgress _progress { get { return Progress; } }
        private GunAxisService _gunAxis; // v6.0 焊钳外部轴

        /// <summary>取消钩子: 由界面停止按钮驱动</summary>
        public Func<bool> IsCancelled;

        /// <summary>
        /// 经验中继池: 本操作内所有规划(含失败RRT的树节点)验证过的自由状态。
        /// 同一操作的过渡段共享一个工作口袋, 一次探索的成本跨段摊销。
        /// </summary>
        private readonly List<Vec3> _relayPool = new List<Vec3>();
        private bool _homePoseMissing; // 本操作内 home POSE 缺失只警告一次

        private void ThrowIfCancelled()
        {
            if (IsCancelled != null && IsCancelled())
                throw new OperationCanceledException();
        }

        public WeldPathPlanner(Action<string> log) { _log = log ?? delegate { }; }

        // ================================================================
        //  主入口: 对所选操作逐个规划
        //  机器人解析: 以操作的 .Robot 属性为准 (点位检查插件的既定做法 —
        //  场景中可能有多台同名机器人，按名字/顺序抓取会拿错工位)
        // ================================================================
        public PlanningReport ExecuteForOperations(
            TxRobot fallbackRobot,
            List<ITxObject> selectedOps)
        {
            var report = new PlanningReport { OperationCount = selectedOps.Count };
            var sw = System.Diagnostics.Stopwatch.StartNew();

            _log("========== RRT 自动路径规划 (操作内模式) ==========");
            _log(string.Format("选中操作数: {0}", selectedOps.Count));

            // 每台实际用到的机器人一套 (干涉集 + 碰撞世界)，按引用缓存
            var setCache = new Dictionary<TxRobot, CollisionSetService>();
            var worldCache = new Dictionary<TxRobot, CollisionWorld>();
            var costCache = new Dictionary<TxRobot, CycleCost>();   // v6.0

            _rrt = new RrtPlanner
            {
                StepSize = RrtStepSize,
                GoalTolerance = RrtStepSize,
                GoalBias = RrtGoalBias,
                MaxIterations = RrtMaxIterations,
                EdgeCheckResolution = EdgeCheckResolution,
                ShouldAbort = delegate { return IsCancelled != null && IsCancelled(); },
                OnProgress = delegate (int iter, int nodeCount)
                {
                    if (iter > 0)
                        _log(string.Format("    RRT: 迭代 {0}, 树节点 {1}", iter, nodeCount));
                    Application.DoEvents(); // SDK 单线程，保持UI响应
                }
            };

            try
            {
                int opIdx = -1;
                foreach (var op in selectedOps)
                {
                    opIdx++;
                    ThrowIfCancelled();

                    // ---- 机器人解析: op.Robot 优先 ----
                    TxRobot robot = ResolveRobotOf(op);
                    if (robot == null)
                    {
                        robot = fallbackRobot;
                        _log(string.Format("  [警告] 操作 {0} 未关联机器人, 回退使用界面选择: {1}",
                            GetNameSafe(op), robot != null ? robot.Name : "无"));
                    }
                    if (robot == null)
                    {
                        report.Warnings.Add("操作 " + GetNameSafe(op) + " 无法解析机器人，跳过");
                        continue;
                    }
                    LogRobotIdentity(robot);

                    // ---- 该机器人的干涉集/碰撞世界 (缓存) ----
                    if (!worldCache.ContainsKey(robot))
                    {
                        // v4.10: 纯附着模式 —— 只复用用户已建好的干涉集, 绝不新建。
                        // 找不到就跳过本操作 (不偷偷新建, 避免堆积)。
                        var cs = CollisionSetService.CreateRobotVsWorld(
                            robot, null, null, _log, forceNew: false, attachOnly: true);

                        if (!cs.IsReady)
                        {
                            string warn = string.Format(
                                "机器人 '{0}' 未找到干涉集 — 请先在主界面点击\"自动创建干涉集\"按钮建立, 再开始规划",
                                robot.Name);
                            _log("  [错误] " + warn + " → 跳过本操作");
                            report.Warnings.Add(warn);
                            cs.Dispose();
                            continue;
                        }

                        setCache[robot] = cs;
                        var w = new CollisionWorld(robot, cs, _log);

                        // v5.5: GUI 参数全量透传
                        w.DynamicCheckEnabled = DynamicCheckEnabled;
                        w.DynamicJointQuantum = DynamicJointQuantum;
                        w.DynamicCartesianQuantum = DynamicCartesianQuantum;
                        w.MaxSweepSteps = MaxSweepSteps;
                        w.ConfigJumpThreshold = ConfigJumpThreshold;
                        w.OrientationVariantsEnabled = OrientationVariantsEnabled;
                        w.MaxVariantTries = MaxVariantTries;
                        w.CacheEnabled = QueryCacheEnabled;      // v6.3
                        w.CacheQuantum = CacheQuantum;

                        w.ApplyGunOpenPoses(); // 焊枪张开后再进行任何碰撞查询
                        cs.RecaptureBaseline(); // 张开后的形态重拍常驻接触基线
                        worldCache[robot] = w;

                        // v6.0: 关节节拍代价模型 (每台机器人一份)
                        if (!costCache.ContainsKey(robot))
                            costCache[robot] = new CycleCost(robot, _log);
                    }
                    _world = worldCache[robot];
                    _rrt.IsStateFree = _world.IsPositionFree;
                    _cost = costCache.ContainsKey(robot) ? costCache[robot] : null;

                    // v6.0: 焊钳外部轴探测 —— 每个操作重探 (活动枪 op.Gun 随操作变)
                    _gunAxis = null;
                    if (GunAxisWriteEnabled)
                    {
                        var ga = new GunAxisService(_log);
                        // v6.4: 覆盖必须在 Probe 之前设 —— ProbeSignConvention 会读它
                        ga.OpenDirectionOverride = GunOpenDirectionOverride;
                        ga.MaxOpeningOverride = GunMaxOpeningOverride;
                        ga.Probe(robot, op);
                        // v6.3: 开口变了 = 枪包络变了 → 失效位姿缓存
                        var wRef = _world;
                        ga.OnGeometryChanged = delegate { wRef.InvalidateCache(); };
                        _gunAxis = ga;
                    }

                    if (_progress != null)
                    {
                        _progress.SetOperationScope(opIdx, selectedOps.Count);
                        _progress.Enter(PlanStage.Init);
                    }

                    PlanSingleOperation(op, report);
                }

                _log(string.Format(
                    "\n规划完成: 插入 {0} 个Via (直连 {1} / 定向 {2} / 中继 {3} / RRT {4}成功·{5}次 / 失败段 {6} / 净空绕行 {7} / 动态违例 {8}·修复 {9}·精修{10}轮 / 共线剪枝 {11})",
                    report.InsertedViaCount, report.DirectHits, report.QuickPlanHits,
                    report.RelayHits, report.RrtSuccesses, report.RrtInvocations,
                    report.FailedSegments, report.ClearanceSegments,
                    report.DynamicViolations, report.DynamicRepairs,
                    report.RefineRounds, report.PrunedVias));
                report.GunOpeningsWritten = _viaOpeningCount;
                if (_progress != null) _progress.Done();

                if (report.ReorderedOps > 0)
                    _log(string.Format("焊序优化: {0} 个操作已重排", report.ReorderedOps));
                if (report.GunOpeningsWritten > 0)
                    _log(string.Format("焊钳开口: {0} 个 Via 已写入外部轴", report.GunOpeningsWritten));

                if (report.FailedSegments > 0)
                    report.Warnings.Add(string.Format(
                        "共 {0} 个过渡段规划失败 (未插过渡点), 明细见上方警告", report.FailedSegments));
            }
            catch (OperationCanceledException)
            {
                _log("\n[中止] 用户停止了规划 — 已插入的Via保留，机器人姿态将恢复");
                report.Warnings.Add("规划被用户中止");
            }
            finally
            {
                foreach (var w in worldCache.Values)
                {
                    report.CollisionQueries += w.QueryCount;
                    if (QueryCacheEnabled) _log("  [性能] " + w.CacheStats);
                    w.Dispose();  // 恢复机器人姿态 + 删除探针 (必做, 与干涉集去留无关)
                }
                foreach (var cs in setCache.Values)
                {
                    // v4.10: 干涉集持久化 — 规划完不删除, 便于用户复用与在
                    // Collision Viewer 中检视。规划采用附着模式, 只会复用现有干涉集,
                    // 不会新建 (找不到则跳过该操作)。这里的 cs 都是复用来的。
                    cs.KeepPairOnDispose = true;
                    cs.Dispose(); // 只释放引用, 不 Delete 碰撞对
                }
                sw.Stop();
                report.Elapsed = sw.Elapsed;
                _log(string.Format("耗时 {0:F1}s, 碰撞查询 {1} 次",
                    sw.Elapsed.TotalSeconds, report.CollisionQueries));
                try { TxApplication.RefreshDisplay(); } catch { }
            }

            return report;
        }

        /// <summary>
        /// 从操作解析执行机器人: op.Robot → 首个焊点的 .Robot。
        /// </summary>
        private TxRobot ResolveRobotOf(ITxObject op)
        {
            try
            {
                var r = ((dynamic)op).Robot as TxRobot;
                if (r != null) return r;
            }
            catch { }

            // 部分结构里 Robot 挂在焊点位置上
            try
            {
                var locs = CollectWeldLocationsOf(op);
                if (locs.Count > 0)
                {
                    var r = ((dynamic)locs[0]).Robot as TxRobot;
                    if (r != null) return r;
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 机器人身份日志: 名称 + HashCode + 基座位置 + 同名计数告警
        /// (场景中曾出现 15 台同名 KR210_R2700-2)
        /// </summary>
        private void LogRobotIdentity(TxRobot robot)
        {
            string basePos = "?";
            try
            {
                var t = ((ITxLocatableObject)robot).AbsoluteLocation.Translation;
                basePos = string.Format("({0:F0},{1:F0},{2:F0})", t.X, t.Y, t.Z);
            }
            catch { }
            _log(string.Format("  机器人(.Robot解析): '{0}' (HashCode={1}) 基座 {2}",
                robot.Name, robot.GetHashCode(), basePos));

            try
            {
                var filter = new TxTypeFilter(typeof(TxRobot));
                TxObjectList all = TxApplication.ActiveDocument
                    .PhysicalRoot.GetAllDescendants(filter);
                int sameName = 0;
                foreach (ITxObject r in all)
                    if (r is TxRobot && r.Name == robot.Name) sameName++;
                if (sameName > 1)
                    _log(string.Format("  [提示] 场景中存在 {0} 台同名机器人 '{1}' — 已按操作.Robot精确解析",
                        sameName, robot.Name));
            }
            catch { }
        }

        // ================================================================
        //  单个操作内规划
        // ================================================================
        private void PlanSingleOperation(ITxObject op, PlanningReport report)
        {
            string opName = GetNameSafe(op);
            _log(string.Format("\n===== 操作: {0} =====", opName));
            _relayPool.Clear(); // 中继池按操作隔离
            _homePoseMissing = false;

            // ---- 按树序收集该操作内的焊点 ----
            var weldLocs = CollectWeldLocationsOf(op);
            if (weldLocs.Count == 0)
            {
                _log("  无焊点子项，跳过");
                return;
            }
            report.WeldCount += weldLocs.Count;
            _log(string.Format("  焊点数: {0}", weldLocs.Count));

            // ---- 提取各焊点位姿 ----
            var weldPoints = new List<PlanPoint>();
            var weldObjs = new List<TxWeldLocationOperation>();
            foreach (var loc in weldLocs)
            {
                PlanPoint p = ExtractWeldPoint(loc);
                if (p == null)
                {
                    _log(string.Format("  [跳过] {0}: 位姿获取失败", loc.Name));
                    continue;
                }
                weldPoints.Add(p);
                weldObjs.Add(loc);
            }

            // ---- 自检门控: 在第一个焊点的进枪位姿处测试完整摆位链路 ----
            // 逆解摆位失败说明链路故障(而非真实障碍)，继续规划只会产出垃圾
            if (weldPoints.Count > 0)
            {
                PlanPoint first = weldPoints[0];
                Vec3 firstOff = first.Position + GetOffsetDir(first) * ApproachRetractDistance;
                _world.SetOrientation(first.AbsTransform, first.RpyZyx);
                ITxObject refLoc = weldObjs.Count > 0 ? (ITxObject)weldObjs[0] : null;
                string verdict = _world.SelfTest(firstOff, first.RpyZyx, refLoc);
                _log(string.Format("  自检({0} 进枪点): {1}", first.Source, verdict));
                if (verdict == "逆解摆位失败")
                {
                    string warn = string.Format(
                        "操作 {0} 自检失败: 焊点进枪位姿无法完成逆解摆位，已跳过 — 见上方[诊断]日志",
                        opName);
                    _log("  [中止] " + warn);
                    report.Warnings.Add(warn);
                    return;
                }
            }

            // ---- v6.1 焊点顺序优化 ----
            // 必须在 IK 后端标定 (上面的自检) **之后** 跑 ——
            // 标定决定 Destination 用局部坐标还是世界坐标; 标定前摆位会全部失败,
            // 关节姿态读不出来, 焊序只能退回笛卡尔距离, 关节代价模型形同虚设。
            // (v6.0 就栽在这: 日志报 "0/25 个焊点关节姿态可读")
            if (WeldOrderOptEnabled && weldPoints.Count > 4)
                OptimizeWeldOrder(weldPoints, weldObjs, report, op);

            // ---- 自适应进出枪偏移解析 ----
            // 焊点坐标系Z轴方向在焊点间常不一致 (投影焊点一半朝内一半朝外),
            // 盲用 -toolZ 会让部分偏移点扎进钣金 → 端点无效 → 级联全灭。
            // 按候选序列逐个验证，取第一个有效方向; 全无效退回焊点原位。
            // ================================================================
            //  阶段一: 进出枪点全量生成 (工艺约束, 排除在路径规划之外)
            //  严格沿焊点 -Z 生成, 距离在 [Min,Max] 区间优先取无干涉;
            //  全区间干涉也照样生成 (取最大距离), 仅标记+警告 — 不挪点、不换向、不换姿态。
            // ================================================================
            var offsets = new Vec3[weldPoints.Count];
            var offsetFree = new bool[weldPoints.Count];
            int freeCount = 0;
            for (int i = 0; i < weldPoints.Count; i++)
            {
                ThrowIfCancelled();
                offsetFree[i] = GenerateOffsetPoint(weldPoints[i], out offsets[i]);
                if (offsetFree[i]) freeCount++;
                else
                    report.Warnings.Add(string.Format(
                        "焊点 {0} 的进/出枪点(严格-Z)存在干涉或不可达, 已按工艺位置强制生成, 需人工确认",
                        weldPoints[i].Source));
            }
            _log(string.Format("  进出枪点生成: {0}/{1} 无干涉 ({2} 个按工艺-Z强制生成)",
                freeCount, weldPoints.Count, weldPoints.Count - freeCount));

            // ================================================================
            //  结构骨架优先 (用户既定流程):
            //   ①创建两个 home Via → 移到 POSE home 位姿
            //   ②点序重排: 一个 home 到操作最前, 另一个到最末
            //   ③创建进出枪点 (含首焊点进枪), 锚定到位
            //   ④规划中间过渡点并插入
            //  骨架先固定, 重排只跟结构点打交道, 与规划结果解耦。
            // ================================================================
            _log("\n  == 结构骨架构建 ==");

            // ---- home 位姿解析 (POSE home 或既有 home 复用) ----
            HomeInfo home = ResolveHomeInfo(op, report);

            // ① 创建两个 home Via (都命名 home), ② 移到首末
            ITxObject homeFirst = null, homeLast = null;
            if (home != null)
            {
                var hp = new PlanPoint
                {
                    Position = home.Pos, RpyZyx = home.Rpy,
                    Kind = PlanPointKind.Transit, Source = "home", CustomName = "home",
                    Verified = true
                };
                if (!home.ReuseFirst) homeFirst = CreateViaAppended(op, hp);
                else homeFirst = home.ExistingFirst;
                if (!home.ReuseLast) homeLast = CreateViaAppended(op, hp);
                else homeLast = home.ExistingLast;

                // ② 点序: homeFirst → 操作最前, homeLast → 操作最末
                if (homeFirst != null && !home.ReuseFirst)
                    MoveChildToFront(op, homeFirst);
                if (homeLast != null && !home.ReuseLast)
                    MoveChildToEnd(op, homeLast);
                _log(string.Format("  home: 首={0} 末={1}",
                    homeFirst != null ? "OK" : "无", homeLast != null ? "OK" : "无"));
            }

            // ③ 创建进出枪点, 锚定: 进枪在焊点前, 出枪在焊点后
            //    首焊点进枪点: 锚到 homeFirst 之后 (即首焊点之前)
            var approachVia = new ITxObject[weldPoints.Count];
            var retractVia = new ITxObject[weldPoints.Count];
            if (GenerateApproachRetract)
            {
                for (int i = 0; i < weldPoints.Count; i++)
                {
                    ThrowIfCancelled();
                    var ap = new PlanPoint
                    {
                        Position = offsets[i], RpyZyx = weldPoints[i].RpyZyx,
                        Kind = PlanPointKind.Approach, Source = weldPoints[i].Source,
                        Verified = offsetFree[i]
                    };
                    var rp = new PlanPoint
                    {
                        Position = offsets[i], RpyZyx = weldPoints[i].RpyZyx,
                        Kind = PlanPointKind.Retract, Source = weldPoints[i].Source,
                        Verified = offsetFree[i]
                    };
                    // 进枪: 移到焊点之前
                    approachVia[i] = CreateViaAppended(op, ap);
                    if (approachVia[i] != null)
                    {
                        MoveChildBefore(op, approachVia[i], weldObjs[i]);
                        report.InsertedViaCount++;
                    }
                    // 出枪: 插到焊点之后
                    retractVia[i] = InsertViaAfter(op, rp, weldObjs[i]);
                    if (retractVia[i] != null) report.InsertedViaCount++;
                }
            }

            // ================================================================
            //  ④ 阶段1: 规划中间过渡点 + 构建精修链 (骨架已固定)
            //  链节点携带其已插入的锚对象 (home/进枪/焊点/出枪), 过渡点待插
            // ================================================================
            _log("\n  == 阶段1: 静态快速规划 ==");
            if (_progress != null) _progress.Enter(PlanStage.StaticPlan);
            var chain = new List<ChainNode>();

            // home(首) 锚
            if (homeFirst != null)
            {
                _world.SetOrientation(null, home.Rpy);
                bool hf = _world.ClassifyStrict(home.Pos) == CollisionWorld.PositionState.Free;
                chain.Add(new ChainNode { WeldObj = homeFirst,
                    Point = new PlanPoint { Position = home.Pos, RpyZyx = home.Rpy,
                        Kind = PlanPointKind.Transit, Source = "home", Verified = hf } });

                // home → 首进枪 过渡
                Vec3 firstGoal = GenerateApproachRetract ? offsets[0] : weldPoints[0].Position;
                var lead = PlanTransit(
                    chain[0].Point, weldPoints[0], home.Pos, firstGoal, hf, offsetFree[0], report);
                if (lead != null) foreach (var t in lead) chain.Add(new ChainNode { Point = t });
            }

            for (int i = 0; i < weldPoints.Count; i++)
            {
                ThrowIfCancelled();
                if (GenerateApproachRetract && approachVia[i] != null)
                    chain.Add(new ChainNode { WeldObj = approachVia[i],
                        Point = new PlanPoint { Position = offsets[i], RpyZyx = weldPoints[i].RpyZyx,
                            Kind = PlanPointKind.Approach, Source = weldPoints[i].Source,
                            Verified = offsetFree[i] } });

                chain.Add(new ChainNode { WeldObj = weldObjs[i], Point = weldPoints[i] });

                if (GenerateApproachRetract && retractVia[i] != null)
                    chain.Add(new ChainNode { WeldObj = retractVia[i],
                        Point = new PlanPoint { Position = offsets[i], RpyZyx = weldPoints[i].RpyZyx,
                            Kind = PlanPointKind.Retract, Source = weldPoints[i].Source,
                            Verified = offsetFree[i] } });

                if (i >= weldPoints.Count - 1) break;

                PlanPoint a = weldPoints[i], b = weldPoints[i + 1];
                var transitPoints = PlanTransit(a, b, offsets[i], offsets[i + 1],
                    offsetFree[i], offsetFree[i + 1], report);
                if (transitPoints != null)
                    foreach (var t in transitPoints) chain.Add(new ChainNode { Point = t });
            }

            // 末焊点 → home(末) 过渡
            if (homeLast != null && weldPoints.Count > 0)
            {
                int last = weldPoints.Count - 1;
                _world.SetOrientation(null, home.Rpy);
                bool hl = _world.ClassifyStrict(home.Pos) == CollisionWorld.PositionState.Free;
                Vec3 tailStart = GenerateApproachRetract ? offsets[last] : weldPoints[last].Position;
                var tail = PlanTransit(weldPoints[last],
                    new PlanPoint { Position = home.Pos, RpyZyx = home.Rpy, Source = "home" },
                    tailStart, home.Pos, offsetFree[last], hl, report);
                if (tail != null) foreach (var t in tail) chain.Add(new ChainNode { Point = t });

                chain.Add(new ChainNode { WeldObj = homeLast,
                    Point = new PlanPoint { Position = home.Pos, RpyZyx = home.Rpy,
                        Kind = PlanPointKind.Transit, Source = "home", Verified = hl } });
            }

            // ---- 阶段2: 逐帧动态精修 (完整链) ----
            if (DynamicCheckEnabled) DynamicRefine(chain, report);
            else _log("\n  (阶段2 动态精修已关闭)");

            // ---- 阶段2.5: 路径捷径化 (v6.0, 取代共线剪枝) ----
            if (PruneEnabled) ShortcutPath(chain, report);

            // ---- 过渡点插入: 每个过渡点锚到链上前一个已存在对象之后 ----
            // 骨架点(home/进枪/焊点/出枪)已就位, 只需把过渡点(WeldObj==null)插进去
            _log("\n  == 过渡点插入 ==");
            if (_progress != null) _progress.Enter(PlanStage.ViaInsert);
            ITxObject anchor = null;
            foreach (var n in chain)
            {
                if (n.WeldObj != null) { anchor = n.WeldObj; continue; }
                if (anchor == null) continue; // 链首过渡点无锚(理论上前面有home): 跳过
                ThrowIfCancelled();
                ITxObject via = InsertViaAfter(op, n.Point, anchor);
                if (via != null) { anchor = via; report.InsertedViaCount++; }
            }
        }

        /// <summary>
        /// v6.0 路径捷径化 (Shortcutting) —— 取代原来的"共线剪枝"。
        ///
        /// ═══════════════════════════════════════════════════════════
        /// 共线剪枝只能删"恰好落在直线上"的点, 太弱。
        /// L1 定向搜索产出的是**直角门形** (退出→平移→插入), 拐点根本不共线,
        /// 一个都删不掉 —— 但那条直角路径完全可以削成一条斜线。
        ///
        /// 真 shortcutting: 对每个起点 i, 从**最远的** j 往回试直连;
        /// 只要 i→j 这条边安全 (静态采样 + 动态关节扫掠), 就把 (i,j) 之间的
        /// 过渡点全部删掉。这是 RRT 后处理的标准操作, 收益远大于共线剪枝。
        ///
        /// 三重保障:
        ///   1. 只删过渡点 (WeldObj == null) —— 焊点/进出枪/home 骨架点绝不动,
        ///      跨越骨架点的捷径直接禁止 (否则会跳过焊点!)
        ///   2. 每条捷径都做 [静态边 + 动态关节扫掠] 双重验证
        ///   3. 用**关节节拍**判断是否真的更优 —— 捷径在笛卡尔上更短, 但如果
        ///      需要翻腕, 节拍反而更长, 这种捷径要拒绝
        /// ═══════════════════════════════════════════════════════════
        /// </summary>
        private void ShortcutPath(List<ChainNode> chain, PlanningReport report)
        {
            if (chain == null || chain.Count < 3) return;

            _log("\n  == 路径捷径化 ==");
            if (_progress != null) _progress.Enter(PlanStage.Shortcut);

            int removedTotal = 0;
            int rejectedSlower = 0;
            int before = chain.Count;

            // ---- 贪心捷径: 每个 i 从最远的 j 往回试 ----
            int i = 0;
            while (i < chain.Count - 2)
            {
                ThrowIfCancelled();
                if (_progress != null && (i & 3) == 0)
                    _progress.Update((double)i / Math.Max(1, chain.Count - 2),
                        string.Format("捷径化: {0}/{1}", i, chain.Count - 2));

                // j 的上界: 不能跨越骨架点 (焊点/进枪/出枪/home)
                int maxJ = chain.Count - 1;
                for (int k = i + 1; k <= maxJ; k++)
                {
                    if (chain[k].WeldObj != null) { maxJ = k; break; }
                }

                bool cut = false;
                // 从最远往回试 —— 一次削掉尽可能多的点
                for (int j = maxJ; j >= i + 2; j--)
                {
                    ThrowIfCancelled();

                    // (i, j) 之间必须全是过渡点
                    bool allTransit = true;
                    for (int k = i + 1; k < j; k++)
                    {
                        if (chain[k].WeldObj != null) { allTransit = false; break; }
                    }
                    if (!allTransit) continue;

                    if (!SegmentStillSafe(chain[i].Point, chain[j].Point)) continue;

                    // 节拍判据: 捷径必须真的更快
                    if (!ShortcutIsFaster(chain, i, j)) { rejectedSlower++; continue; }

                    int cnt = j - i - 1;
                    chain.RemoveRange(i + 1, cnt);
                    removedTotal += cnt;
                    cut = true;
                    break;
                }

                if (!cut) i++;
                // cut 成功时留在原地, 继续从 i 尝试更远的捷径
            }

            report.PrunedVias += removedTotal;

            if (removedTotal > 0)
            {
                _log(string.Format("    捷径化: 删除 {0} 个过渡点 ({1} → {2} 节点)",
                    removedTotal, before, chain.Count));
                if (rejectedSlower > 0)
                    _log(string.Format("    拒绝 {0} 条几何更短但节拍更慢的捷径 (翻腕代价)",
                        rejectedSlower));
            }
            else
            {
                _log("    无可削减 (路径已紧凑)");
            }
        }

        /// <summary>
        /// 捷径节拍判据: i→j 直连 是否比 i→...→j 原路径更快。
        ///
        /// 用关节 PTP 时间, 不用笛卡尔距离 —— 一条 TCP 更短但需要翻腕的捷径,
        /// 实际节拍可能更长。关节值读不到时回退笛卡尔距离比较。
        /// </summary>
        private bool ShortcutIsFaster(List<ChainNode> chain, int i, int j)
        {
            if (_cost == null) return true;   // 无代价模型 → 只按几何变短就接受

            try
            {
                // 原路径节拍
                double orig = 0;
                for (int k = i; k < j; k++)
                {
                    double t = SegmentTime(chain[k].Point, chain[k + 1].Point);
                    if (t < 0) return true;    // 拿不到 → 不阻塞捷径
                    orig += t;
                }

                // 捷径节拍
                double sc = SegmentTime(chain[i].Point, chain[j].Point);
                if (sc < 0) return true;

                return sc < orig - 1e-6;
            }
            catch { return true; }
        }

        /// <summary>两点间 PTP 节拍 (秒); 关节不可读返回 -1</summary>
        private double SegmentTime(PlanPoint a, PlanPoint b)
        {
            if (_cost == null) return -1;
            try
            {
                TxPoseData ja, jb;
                if (!_world.TryCaptureJointPose(a.Position, a.RpyZyx, out ja)) return -1;
                if (!_world.TryCaptureJointPose(b.Position, b.RpyZyx, out jb)) return -1;
                return _cost.PtpTime(ja, jb);
            }
            catch { return -1; }
        }

        /// <summary>共线容差 (mm) — 捷径化后的收尾清理仍会用到</summary>
        public double CollinearTolerance = 3.0;

        /// <summary>
        /// 捷径复验: a→b 直连是否安全 (静态边采样 + 动态关节扫掠)。
        /// 动态检查不可用时只做静态。
        /// </summary>
        private bool SegmentStillSafe(PlanPoint a, PlanPoint b)
        {
            try
            {
                _world.SetOrientationBlend(a.RpyZyx, b.RpyZyx, a.Position, b.Position);

                // 静态: 边采样
                if (!_rrt.IsEdgeFree(a.Position, b.Position)) return false;

                // 动态: 关节扫掠 (捕捉枪体扫掠)
                if (_world.DynamicCheckEnabled)
                {
                    TxPoseData ja, jb;
                    if (!_world.TryCaptureJointPose(a.Position, a.RpyZyx, out ja)) return false;
                    if (!_world.TryCaptureJointPose(b.Position, b.RpyZyx, out jb)) return false;
                    if (_world.CheckJointMotion(ja, jb) != null) return false;
                }
                return true;
            }
            catch { return false; }
        }
        // ════════════════════════════════════════════════════════════
        //  v6.0 焊点顺序优化
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 以关节节拍为边权重排焊点顺序 (2-opt + Or-opt)。
        /// 原地修改 weldPoints / weldObjs (两者必须同步重排!)。
        ///
        /// 边权用 PtpTime (关节空间), 不用笛卡尔距离 —— 见 CycleCost 的说明。
        /// 关节值拿不到时回退笛卡尔距离 (仍比不优化好)。
        /// </summary>
        private void OptimizeWeldOrder(
            List<PlanPoint> weldPoints, List<TxWeldLocationOperation> weldObjs,
            PlanningReport report, ITxObject op)
        {
            int n = weldPoints.Count;
            _log(string.Format("\n  == 焊点顺序优化 ({0} 点) ==", n));
            if (_progress != null) _progress.Enter(PlanStage.WeldOrder);

            // ---- 预计算各焊点的关节姿态 (代价矩阵的输入) ----
            var poses = new TxPoseData[n];
            int okPose = 0;
            for (int i = 0; i < n; i++)
            {
                ThrowIfCancelled();
                TxPoseData jp;
                if (_world.TryCaptureJointPose(weldPoints[i].Position, weldPoints[i].RpyZyx, out jp))
                {
                    poses[i] = jp;
                    okPose++;
                }
                if (_progress != null)
                    _progress.Update(0.5 * (i + 1) / n, "焊序: 捕获关节姿态");
            }

            bool useJoint = _cost != null && okPose == n;
            if (!useJoint)
            {
                _log(string.Format(
                    "  [焊序] {0}/{1} 个焊点关节姿态可读 — 退回笛卡尔距离度量 (精度下降)",
                    okPose, n));
            }

            // ---- 预计算代价矩阵 (SDK 部分, 必须主线程) ----
            // 之后的 2-opt/Or-opt 是纯数学 → 可以并行吃满多核
            var matrix = new double[n, n];
            int cells = n * n;
            int doneCells = 0;

            for (int i = 0; i < n; i++)
            {
                ThrowIfCancelled();
                for (int j = 0; j < n; j++)
                {
                    if (i == j) { matrix[i, j] = 0; doneCells++; continue; }

                    double t = -1;
                    if (useJoint) t = _cost.PtpTime(poses[i], poses[j]);
                    if (t < 0)
                        t = Vec3.Distance(weldPoints[i].Position, weldPoints[j].Position) / 1000.0;

                    matrix[i, j] = t;
                    doneCells++;
                }
                if (_progress != null)
                    _progress.Update(0.5 + 0.4 * doneCells / cells, "焊序: 代价矩阵");
            }

            var opt = new WeldOrderOptimizer(_log)
            {
                LockFirst = Math.Max(0, Math.Min(WeldOrderLockFirst, n - 2)),
                LockLast = Math.Max(0, Math.Min(WeldOrderLockLast, n - 2))
            };

            if (_progress != null) _progress.Update(0.9, "焊序: 并行 2-opt");
            int[] order = opt.SolveParallel(matrix, n, null);

            // ---- 应用新顺序 (两个列表必须同步!) ----
            bool changed = false;
            for (int i = 0; i < n; i++)
                if (order[i] != i) { changed = true; break; }

            if (!changed) return;
            report.ReorderedOps++;

            var newPts = new List<PlanPoint>(n);
            var newObjs = new List<TxWeldLocationOperation>(n);
            foreach (int idx in order)
            {
                newPts.Add(weldPoints[idx]);
                newObjs.Add(weldObjs[idx]);
            }
            weldPoints.Clear(); weldPoints.AddRange(newPts);
            weldObjs.Clear(); weldObjs.AddRange(newObjs);

            _log("  [焊序] ⚠ 焊接顺序已改变 — 请确认工艺允许 (定位焊/防变形约束可用锁定选项)");

            // ---- 按新顺序重排 PS 树里的焊点 ----
            ReorderWeldLocationsInPs(op, newObjs);
        }

        /// <summary>
        /// 把 PS 操作树里的焊点按新顺序物理重排。
        ///
        /// v6.1 修复: 之前用 weldObj.Parent 取有序容器 —— 失败 ("父容器非
        /// ITxOrderedObjectCollection")。正确做法是直接用**操作本身** (op),
        /// 它就是有序容器 (CreateViaAfter / MoveChildToFront 一直在这么用)。
        /// </summary>
        private void ReorderWeldLocationsInPs(ITxObject op, List<TxWeldLocationOperation> ordered)
        {
            if (ordered == null || ordered.Count == 0 || op == null) return;

            var container = op as ITxOrderedObjectCollection;
            if (container == null)
            {
                _log("  [焊序] 操作非 ITxOrderedObjectCollection — PS 树顺序未变更");
                _log("         (规划路径已按新顺序生成, 但操作树里焊点仍是原序)");
                return;
            }

            int moved = 0, failed = 0;
            ITxObject prev = null;
            foreach (var w in ordered)
            {
                try
                {
                    // prev=null → 移到首位; 否则移到 prev 之后
                    container.AddObjectAfter(w, (ITxOperation)prev);
                    prev = w;
                    moved++;
                }
                catch (Exception ex)
                {
                    failed++;
                    if (failed == 1)
                        _log("  [焊序] PS 树重排出错: " + ex.Message);
                }
            }

            if (failed == 0)
                _log(string.Format("  [焊序] PS 树已重排 ({0} 个焊点)", moved));
            else
                _log(string.Format("  [焊序] PS 树重排: 成功 {0} / 失败 {1} — 顺序可能不完整",
                    moved, failed));
        }

        // ════════════════════════════════════════════════════════════
        //  v6.0 焊钳外部轴 + 运动参数写入
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 给一个 Via 点写入 [焊钳开口 + 运动参数]。
        ///
        /// 焊钳开口策略:
        ///   进/出枪点 → TargetGunOpening (够越过翻边即可, 默认 30mm)
        ///   过渡点   → 按需最小开口 (从小到大试, 第一个不干涉的就用)
        ///
        /// 为什么不无脑全开:
        ///   开口大 = 伺服行程长 = 慢 (吃节拍); 且枪臂张开 = 包络更大 = 更容易撞。
        ///
        /// 运动参数:
        ///   MotionType = Joint (PTP, 各关节最高效) — 过渡点用 Linear 是节拍杀手
        ///
        /// inheritFrom: 相邻焊点, 用于继承导轨/变位机等非枪轴的工艺值。
        /// </summary>
        private void ApplyViaMotionAndGun(ITxObject via, PlanPoint p, ITxObject inheritFrom)
        {
            if (via == null) return;

            // ---- ① 运动类型: Joint (PTP) ----
            if (SetTransitMotionJoint)
            {
                try
                {
                    dynamic v = via;
                    var mp = new TxRoboticLocationOperationMotionParameters();
                    mp.MotionType = TxMotionType.Joint;
                    v.MotionParameters = mp;
                }
                catch
                {
                    // 备用: 直接属性
                    try { ((dynamic)via).MotionType = TxMotionType.Joint; }
                    catch { }
                }
            }

            // ---- ② 焊钳开口 (外部轴) ----
            if (_gunAxis == null || !_gunAxis.HasGunAxis) return;

            double opening;
            if (p.Kind == PlanPointKind.Approach || p.Kind == PlanPointKind.Retract)
            {
                opening = TargetGunOpening;
            }
            else if (AdaptiveGunOpening)
            {
                opening = FindMinOpeningAt(p);
                if (double.IsNaN(opening)) opening = TransitGunOpening;  // 全试不通 → 用默认
            }
            else
            {
                opening = TransitGunOpening;
            }

            if (_gunAxis.WriteOpening(via, opening, inheritFrom))
                _viaOpeningCount++;
        }

        /// <summary>
        /// 按需最小开口: 在点 p 处从小到大试开口, 第一个不干涉的就用。
        /// 找不到返回 NaN。
        /// </summary>
        private double FindMinOpeningAt(PlanPoint p)
        {
            if (_gunAxis == null || !_gunAxis.HasGunAxis) return double.NaN;

            // v6.2: 阶梯是**开口幅值** (恒非负), 不是关节值 ——
            // 负向开启的枪由 GunAxisService.OpeningToJointValue 换算成负数
            double gmax = _gunAxis.MaxOpeningMagnitude;
            if (gmax <= 1e-6) gmax = 200.0;

            var ladder = new List<double>();
            foreach (double o in new[] { TargetGunOpening, 60.0, 100.0, 150.0, gmax })
            {
                double m = Math.Min(Math.Abs(o), gmax);
                if (!ladder.Any(x => Math.Abs(x - m) < 1e-6)) ladder.Add(m);
            }
            ladder.Sort();   // 小 → 大, 第一个安全的即返回

            _world.SetOrientation(null, p.RpyZyx);

            return _gunAxis.FindMinSafeOpening(
                delegate (double o) { _gunAxis.ApplyOpeningToDevice(o); },
                delegate { return _world.IsPositionFree(p.Position); },
                ladder.ToArray());
        }

        private sealed class HomeInfo
        {
            public Vec3 Pos;
            public TxVector Rpy;
            public bool ReuseFirst, ReuseLast;
            public ITxObject ExistingFirst, ExistingLast;
        }

        /// <summary>
        /// 把 child 移到操作最前: 基于 ITxOrderedObjectCollection.AddObjectAfter(child, null)。
        /// prevChild=null 的语义 = "插入为第一个子元素" (API 文档确认)。
        /// AddObjectAfter 不可用时回退 MoveChildAfter(child, 首元素前驱=null)。
        /// </summary>
        private void MoveChildToFront(ITxObject op, ITxObject child)
        {
            // 首选: ITxOrderedObjectCollection.AddObjectAfter(child, null) = 插入为第一个
            var ordered = op as ITxOrderedObjectCollection;
            if (ordered != null)
            {
                try { ordered.AddObjectAfter(child, null); return; }
                catch { }
            }
            // dynamic 回退 (部分运行时类型可能只暴露 MoveChildAfter 而非 AddObjectAfter)
            try { ((dynamic)op).AddObjectAfter((dynamic)child, null); return; } catch { }
            // 最终回退: 移到当前首元素之后 (语义: child 成为第一个)
            try { ((dynamic)op).MoveChildAfter(child, null); return; } catch { }
            try
            {
                ITxObject firstChild = GetBoundaryChild(op, true);
                if (firstChild != null && !ReferenceEquals(firstChild, child))
                    ((ITxOrderedObjectCollection)op).MoveChildAfter(child, firstChild);
            }
            catch { _log("  [提示] home(首) 置顶失败, 需人工调整"); }
        }

        /// <summary>
        /// 把 child 移到操作最末: MoveChildAfter(child, lastChild)。
        /// lastChild=null 时 MoveChildAfter 语义=插入为第一个(不是最后一个!),
        /// 所以必须显式找到最末子元素作为前驱。
        /// </summary>
        private void MoveChildToEnd(ITxObject op, ITxObject child)
        {
            var ordered = op as ITxOrderedObjectCollection;
            if (ordered != null)
            {
                ITxObject lastChild = GetBoundaryChild(op, false);
                if (lastChild != null && !ReferenceEquals(lastChild, child))
                {
                    try { ordered.MoveChildAfter(child, lastChild); return; }
                    catch { }
                }
            }
            // dynamic 回退
            try
            {
                ITxObject lastChild = GetBoundaryChild(op, false);
                if (lastChild != null && !ReferenceEquals(lastChild, child))
                    ((dynamic)op).MoveChildAfter(child, lastChild);
            }
            catch { _log("  [提示] home(末) 置底失败, 需人工调整"); }
        }

        /// <summary>
        /// home 位姿解析: 检测操作首/尾既有 home 可复用; 否则应用机器人 POSE "home" 读TCP。
        /// </summary>
        private HomeInfo ResolveHomeInfo(ITxObject op, PlanningReport report)
        {
            var info = new HomeInfo();

            // 既有首/尾 home 检测
            try
            {
                ITxObject bf = GetBoundaryChild(op, true);
                if (bf is TxRoboticViaLocationOperation
                    && string.Equals(bf.Name, "home", StringComparison.Ordinal))
                { info.ReuseFirst = true; info.ExistingFirst = bf; }
                ITxObject bl = GetBoundaryChild(op, false);
                if (bl is TxRoboticViaLocationOperation
                    && string.Equals(bl.Name, "home", StringComparison.Ordinal))
                { info.ReuseLast = true; info.ExistingLast = bl; }
            }
            catch { }

            // 位姿来源: 既有 home 或 POSE home
            if (info.ReuseFirst && info.ExistingFirst is ITxLocatableObject)
            {
                var abs = ((ITxLocatableObject)info.ExistingFirst).AbsoluteLocation;
                info.Pos = new Vec3(abs.Translation.X, abs.Translation.Y, abs.Translation.Z);
                try { info.Rpy = abs.RotationRPY_ZYX; } catch { }
                _log("  home: 复用既有 home 位姿");
                return info;
            }

            if (!_world.TryApplyNamedPose("home"))
            {
                report.Warnings.Add("未找到名为 home 的机器人POSE(区分大小写), 首尾home未生成");
                return null;
            }
            Vec3 pos; TxVector rpy;
            if (!_world.TryReadTcpPose(out pos, out rpy))
            {
                report.Warnings.Add("home 姿态下 TCP 读取失败, 首尾home未生成");
                return null;
            }
            info.Pos = pos; info.Rpy = rpy;
            _log(string.Format("  home: 由机器人POSE生成 @({0:F0},{1:F0},{2:F0})", pos.X, pos.Y, pos.Z));
            return info;
        }

        /// <summary>取操作的首/尾直接子元素 (顺序即执行序)</summary>
        private ITxObject GetBoundaryChild(ITxObject op, bool first)
        {
            try
            {
                var container = op as ITxObjectCollection;
                if (container == null) return null;
                var filter = new TxTypeFilter(typeof(ITxObject));
                TxObjectList kids = null;
                try { kids = ((dynamic)container).GetDirectDescendants(filter); }
                catch { }
                if (kids == null)
                    kids = container.GetAllDescendants(filter);
                if (kids == null) return null;

                ITxObject firstObj = null, lastObj = null;
                foreach (ITxObject k in kids)
                {
                    if (firstObj == null) firstObj = k;
                    lastObj = k;
                }
                return first ? firstObj : lastObj;
            }
            catch { return null; }
        }

        /// <summary>操作级点链节点: WeldObj 非null = 既有焊点锚, 否则为待创建Via</summary>
        private sealed class ChainNode
        {
            public PlanPoint Point;
            public ITxObject WeldObj;
        }

        // ================================================================
        //  过渡段规划级联 (L0-L3)
        // ================================================================
        private List<PlanPoint> PlanTransit(
            PlanPoint a, PlanPoint b, Vec3 start, Vec3 goal,
            bool startFree, bool goalFree, PlanningReport report)
        {
            _log(string.Format("\n  过渡段: {0} → {1}  (距离 {2:F0}mm)",
                a.Source, b.Source, Vec3.Distance(start, goal)));

            // 姿态混合: 起点=A姿态(出枪Via), 终点=B姿态(进枪Via), 中途按行进进度插值
            _world.SetOrientationBlend(a.RpyZyx, b.RpyZyx, start, goal);
            ThrowIfCancelled();

            // ---- 工艺端点门控 ----
            // 进出枪点是固定工艺位置(不挪), 但端点干涉≠放弃整段: 机器人仍要去该点,
            // 直连会穿越干涉区(观测到的路径穿模)。改为"净空绕行": 经上方安全中继,
            // 让路径大部分在净空, 只保留端点附近不可避免的最后一段。
            if (!startFree || !goalFree)
            {
                _log(string.Format("    [工艺端点干涉] 出枪({0}):{1}  进枪({2}):{3} → 净空绕行",
                    a.Source, startFree ? "OK" : "干涉", b.Source, goalFree ? "OK" : "干涉"));
                var clearance = PlanClearancePath(a, b, start, goal, report);
                if (clearance != null) return clearance;
                MarkSegmentFailed(a, b, "进/出枪工艺位置干涉且上方无净空", report);
                return null;
            }

            // ---- L0: 直连 ----
            if (_rrt.IsEdgeFree(start, goal))
            {
                _log("    L0 直连通过");
                report.DirectHits++;
                return new List<PlanPoint>();
            }

            // ---- L1: 枪坐标系定向搜索 (后撤-X → 竖直±Z → 横向±Y, 大步进) ----
            // 领域知识版树生长: 模拟人工示教挪枪的动作序列, 替代原中点抬升
            Vec3 mid = Vec3.Lerp(start, goal, 0.5);
            {
                var gfs = new GunFrameSearch
                {
                    BackoutStep = GunBackoutStep,
                    BackoutMax = GunBackoutMax,
                    SideStep = GunSideStep,
                    SideMax = GunSideMax,
                    MinBackoutForSide = GunMinBackoutForSide,   // v5.4
                    PreferDeepBackoutForSide = true,            // v5.4
                    IsFree = _world.IsPositionFree,
                    IsEdgeFree = _rrt.IsEdgeFree,

                    // v5.4 核心: 动态扫掠验证前移。
                    // IsEdgeFree 只采样 TCP 点位, 看不见枪体扫过夹具 ——
                    // "后撤40mm就抬升"这种解静态能过, 动态必撞。
                    // 在 L1 生成阶段就拒掉, 逼它继续加深后撤, 而不是留给阶段2去修。
                    IsEdgeSafeDynamic = DynamicCheckEnabled && _world.DynamicCheckEnabled
                        ? new Func<Vec3, Vec3, bool>(_world.IsEdgeSafeDynamic)
                        : null,

                    DescribeBlock = delegate (Vec3 p)
                    {
                        var st = _world.Classify(p);
                        return st == CollisionWorld.PositionState.NoIk ? "无逆解" : "有干涉";
                    },
                    ShouldAbort = delegate { return IsCancelled != null && IsCancelled(); },
                    Log = _log
                };

                // 枪轴: 后撤 = -X (前后轴); 竖直/横向取 A/B 两侧枪轴平均
                Vec3 backA = GetAxisOf(a.AbsTransform, 0) * -1.0;
                Vec3 backB = GetAxisOf(b.AbsTransform, 0) * -1.0;
                Vec3 zAxis = (GetAxisOf(a.AbsTransform, 2) + GetAxisOf(b.AbsTransform, 2)).Normalized();
                Vec3 yAxis = (GetAxisOf(a.AbsTransform, 1) + GetAxisOf(b.AbsTransform, 1)).Normalized();

                List<Vec3> directed = gfs.Search(start, goal, backA, backB, zAxis, yAxis);
                ThrowIfCancelled();
                if (directed != null)
                {
                    _log(string.Format("    L1 定向搜索通过: {0} 个过渡点", directed.Count));
                    report.QuickPlanHits++;
                    _relayPool.AddRange(directed); // 成功点入池
                    return BuildTransitPoints(directed, a, b, "定向搜索");
                }
            }

            // ---- L1.5: 经验中继 (复用本操作已验证的自由点) ----
            if (_relayPool.Count > 0)
            {
                var ranked = _relayPool
                    .OrderBy(p => Vec3.Distance(start, p) + Vec3.Distance(p, goal))
                    .Take(8);
                foreach (var relay in ranked)
                {
                    ThrowIfCancelled();
                    if (_rrt.IsEdgeFree(start, relay) && _rrt.IsEdgeFree(relay, goal))
                    {
                        _log(string.Format("    L1.5 经验中继通过 (池 {0} 点)", _relayPool.Count));
                        report.RelayHits++;
                        return BuildTransitPoints(new List<Vec3> { relay }, a, b, "经验中继");
                    }
                }
            }

            // ---- L2: RRT (端点已验证有效) ----
            _log("    L1 未通过，启动 RRT...");
            report.RrtInvocations++;

            Aabb bounds = Aabb.FromPoints(start, goal);
            // 将中继池已有的自由点纳入边界 (缩小采样范围、提高命中)
            if (_relayPool.Count > 0)
            {
                foreach (var rp in _relayPool)
                {
                    bounds = new Aabb
                    {
                        Min = new Vec3(Math.Min(bounds.Min.X, rp.X),
                            Math.Min(bounds.Min.Y, rp.Y), Math.Min(bounds.Min.Z, rp.Z)),
                        Max = new Vec3(Math.Max(bounds.Max.X, rp.X),
                            Math.Max(bounds.Max.Y, rp.Y), Math.Max(bounds.Max.Z, rp.Z))
                    };
                }
            }
            bounds = bounds.Inflate(SampleBoundsInflateXy, SampleBoundsInflateZUp, SampleBoundsInflateZDown);

            // 种子注入: 中继池按走廊距离排序取前12, 直接挂接为树的初始枝干
            _rrt.SeedStates = _relayPool.Count == 0 ? null
                : _relayPool.OrderBy(p => Vec3.Distance(start, p) + Vec3.Distance(p, goal))
                            .Take(12).ToList();

            List<Vec3> rrtPath = _rrt.Plan(start, goal, bounds);
            ThrowIfCancelled(); // RRT 因中止返回 null 时不落兜底，直接退出

            // 收割探索成果: 无论成败, 两树自由节点入池供后续段复用
            if (_rrt.LastExploredFreeStates.Count > 0)
                _relayPool.AddRange(_rrt.LastExploredFreeStates);

            if (rrtPath != null)
            {
                report.RrtSuccesses++;
                _log(string.Format("    L2 RRT 成功: 平滑后 {0} 点", rrtPath.Count));

                var mids = new List<Vec3>();
                for (int i = 1; i < rrtPath.Count - 1; i++) // 剔除首尾
                    mids.Add(rrtPath[i]);
                _relayPool.AddRange(mids);
                return BuildTransitPoints(mids, a, b, "RRT");
            }

            // ---- 全级失败: 不插入任何未验证点, 明确标记留待人工 ----
            _log("    L2 RRT 失败" +
                (_rrt.LastFailReason != null ? " (" + _rrt.LastFailReason + ")" : "") +
                " → 本段不插过渡点");
            MarkSegmentFailed(a, b,
                _rrt.LastFailReason != null ? _rrt.LastFailReason : "级联全部未通", report);
            return null; // null=段失败
        }

        /// <summary>
        /// 过渡点定稿: 对每个几何点在混合模式下执行姿态择优锁定,
        /// Via 携带实际采用的欧拉角 (变体命中时 ≠ 目标焊点姿态)。
        /// </summary>
        private List<PlanPoint> BuildTransitPoints(
            List<Vec3> mids, PlanPoint a, PlanPoint b, string note)
        {
            var pts = new List<PlanPoint>();
            foreach (var v in mids)
            {
                TxVector rpy;
                if (!_world.TryFindFreePose(v, out rpy) || rpy == null)
                    rpy = b.RpyZyx; // 理论上已验证过, 兜底取B姿态
                pts.Add(new PlanPoint
                {
                    Position = v, RpyZyx = rpy,
                    Kind = PlanPointKind.Transit, Verified = true,
                    Source = a.Source + "→" + b.Source,
                    Note = note
                });
            }
            return pts;
        }

        /// <summary>
        /// 阶段2 逐帧动态精修: 对完整点链多轮执行 [捕获→关节扫掠→修复→复检]。
        /// 每个节点按其最终 Via 姿态精确摆位捕获; 违例边走三策略修复阶梯;
        /// 无违例或一轮内无新修复即收敛 (最多3轮)。
        /// </summary>
        private void DynamicRefine(List<ChainNode> chain, PlanningReport report)
        {
            _log("\n  == 阶段2: 逐帧动态精修 ==");
            if (_progress != null) _progress.Enter(PlanStage.DynamicRefine);
            if (chain.Count < 2) return;

            // v5.2: 先自检关节读写能力 —— 不可用就明说, 别假装在跑
            if (!_world.SelfTestJointIO())
            {
                report.Warnings.Add("动态干涉检查不可用 (关节读写失败), 路径仅通过静态点位验证, 需人工复核运动过程");
                return;
            }

            const int MaxRounds = 3;
            // v5.3: 修复预算从固定 8 改为按链长动态分配 (原来 15 个违例只修了 5 个,
            // 剩下全部"预算耗尽"直接放弃)。每 6 个节点给 1 次, 下限 12 上限 40。
            int repairBudget = Math.Max(12, Math.Min(40, chain.Count / 6 + 12));
            var warnedEdges = new HashSet<PlanPoint>(); // 已警告边(按左节点), 跨轮不重复计数

            for (int round = 1; round <= MaxRounds; round++)
            {
                ThrowIfCancelled();
                report.RefineRounds = round;

                // ---- 捕获: 每节点按自身最终姿态摆位 ----
                // v5.2: 分段容错 —— 摆位失败的节点标记为"洞"(joints[i]=null),
                // 只跳过与洞相邻的边, 其余边照常做扫掠检测。
                // 旧行为是一个节点失败就 return, 导致整条链一次都没检 ——
                // 而失败节点往往正是"进/出枪点强制生成"的那几个无逆解点。
                var joints = new List<TxPoseData>();
                int holes = 0;
                var holeNames = new List<string>();
                foreach (var n in chain)
                {
                    ThrowIfCancelled();
                    TxPoseData jp;
                    if (!_world.TryCaptureJointPose(n.Point.Position, n.Point.RpyZyx, out jp))
                    {
                        joints.Add(null);   // 洞
                        holes++;
                        if (holeNames.Count < 6)
                            holeNames.Add(n.Point.Source + "(" + n.Point.Kind + ")");
                        continue;
                    }
                    joints.Add(jp);
                }

                if (holes > 0)
                {
                    _log(string.Format("    [精修R{0}] {1}/{2} 个节点摆位失败(无逆解), 跳过相邻边; 其余边正常检测",
                        round, holes, chain.Count));
                    _log("      失败节点: " + string.Join(", ", holeNames.ToArray())
                        + (holes > holeNames.Count ? " ..." : ""));
                }
                if (holes == chain.Count)
                {
                    _log("    [精修] 全部节点摆位失败 — 精修中止 (检查机器人可达性/干涉集)");
                    return;
                }

                // ---- 逐帧扫掠 + 修复 ----
                int violations = 0, repairs = 0, skippedEdges = 0, unfixableEdges = 0;
                var inserts = new List<KeyValuePair<int, List<PlanPoint>>>();

                for (int i = 0; i < chain.Count - 1; i++)
                {
                    ThrowIfCancelled();
                    if (_progress != null && (i & 7) == 0)
                        _progress.Update(
                            ((round - 1) + (double)i / Math.Max(1, chain.Count - 1)) / MaxRounds,
                            string.Format("精修 R{0}: 边 {1}/{2}", round, i, chain.Count - 1));
                    if (warnedEdges.Contains(chain[i].Point)) continue; // 已放弃的边

                    // v5.2: 跳过与"洞"相邻的边 (端点无逆解, 无法做关节扫掠)
                    if (joints[i] == null || joints[i + 1] == null)
                    {
                        skippedEdges++;
                        continue;
                    }

                    double tv;
                    string issue = _world.CheckJointMotion(joints[i], joints[i + 1], out tv);
                    if (issue == null) continue;

                    violations++;
                    report.DynamicViolations++;

                    // v6.1: 端点本身就干涉的边 —— 修不好, 别浪费预算。
                    // 这些是"进/出枪点严格-Z 存在干涉, 已按工艺位置强制生成"的点
                    // (Verified=false)。端点是烂的, 插再多 Via 也没用 ——
                    // v6.0 在这类边上反复砸预算, "三策略均失败" 刷了 14 次。
                    if (!chain[i].Point.Verified || !chain[i + 1].Point.Verified)
                    {
                        warnedEdges.Add(chain[i].Point);
                        unfixableEdges++;
                        report.Warnings.Add(string.Format(
                            "动态干涉(端点不可达): {0}→{1} ({2}) — 进/出枪点工艺位置本身干涉, "
                            + "非路径问题, 需调整焊点姿态或夹具",
                            chain[i].Point.Source, chain[i + 1].Point.Source, issue));
                        continue;
                    }

                    if (repairBudget-- <= 0)
                    {
                        warnedEdges.Add(chain[i].Point);
                        report.Warnings.Add(string.Format(
                            "动态干涉: {0}→{1} ({2}), 修复预算耗尽, 需人工复核",
                            chain[i].Point.Source, chain[i + 1].Point.Source, issue));
                        continue;
                    }

                    // 修复上下文: 最近前后焊点提供轴系, 边两端姿态做混合
                    PlanPoint aCtx = NearestWeld(chain, i, -1);
                    PlanPoint bCtx = NearestWeld(chain, i + 1, +1);
                    Vec3 pa = chain[i].Point.Position;
                    Vec3 pb = chain[i + 1].Point.Position;
                    _world.SetOrientationBlend(
                        chain[i].Point.RpyZyx, chain[i + 1].Point.RpyZyx, pa, pb);

                    Vec3 backDir = GetAxisOf(aCtx.AbsTransform, 0) * -1.0;
                    Vec3 zAxis = (GetAxisOf(aCtx.AbsTransform, 2)
                                + GetAxisOf(bCtx.AbsTransform, 2)).Normalized();
                    Vec3 yAxis = (GetAxisOf(aCtx.AbsTransform, 1)
                                + GetAxisOf(bCtx.AbsTransform, 1)).Normalized();

                    var repair = TryRepairJointMotion(pa, pb, joints[i], joints[i + 1],
                        tv, backDir, zAxis, yAxis, aCtx, bCtx);

                    if (repair != null)
                    {
                        repairs++;
                        report.DynamicRepairs++;
                        _log(string.Format("    [精修R{0}] {1}→{2} {3} → {4}",
                            round, chain[i].Point.Source, chain[i + 1].Point.Source,
                            issue, repair[0].Note));
                        inserts.Add(new KeyValuePair<int, List<PlanPoint>>(i + 1, repair));
                    }
                    else
                    {
                        warnedEdges.Add(chain[i].Point);
                        _log(string.Format("    [精修R{0}] {1}→{2} {3} → 三策略均失败",
                            round, chain[i].Point.Source, chain[i + 1].Point.Source, issue));
                        report.Warnings.Add(string.Format(
                            "动态干涉: {0}→{1} ({2}), 修复失败, 需人工复核",
                            chain[i].Point.Source, chain[i + 1].Point.Source, issue));
                    }
                }

                // ---- 应用修复插入 (倒序保索引) ----
                for (int k = inserts.Count - 1; k >= 0; k--)
                {
                    int at = Math.Min(inserts[k].Key, chain.Count);
                    var nodes = new List<ChainNode>();
                    foreach (var p in inserts[k].Value)
                        nodes.Add(new ChainNode { Point = p });
                    chain.InsertRange(at, nodes);
                }

                int checkedEdges = (chain.Count - 1) - skippedEdges - warnedEdges.Count;
                _log(string.Format(
                    "    第{0}轮: 检测边 {1} / 跳过 {2} / 违例 {3} / 修复 {4} / 端点不可达 {5}",
                    round, Math.Max(0, checkedEdges), skippedEdges, violations, repairs,
                    unfixableEdges));

                if (unfixableEdges > 0)
                    _log(string.Format(
                        "    [提示] {0} 条边因端点(进/出枪点)工艺位置本身干涉而无法修复 — "
                        + "这是焊点姿态/夹具问题, 不是路径问题", unfixableEdges));

                if (skippedEdges > 0)
                    report.Warnings.Add(string.Format(
                        "动态检查: {0} 条边因端点无逆解被跳过, 未做关节扫掠验证, 需人工确认",
                        skippedEdges));

                if (violations == 0)
                {
                    if (checkedEdges > 0) _log("    已检边动态干净 ✓");
                    else _log("    [警告] 无任何边可检 (全部端点无逆解) — 动态验证未生效");
                    return;
                }
                if (repairs == 0) { _log("    无新修复, 精修收敛 (残余违例见警告)"); return; }
            }
        }

        /// <summary>沿链向 dir 方向找最近的焊点节点 (修复轴系上下文)</summary>
        private PlanPoint NearestWeld(List<ChainNode> chain, int idx, int dir)
        {
            for (int i = idx; i >= 0 && i < chain.Count; i += dir)
                if (chain[i].WeldObj != null) return chain[i].Point;
            return chain[Math.Max(0, Math.Min(chain.Count - 1, idx))].Point;
        }

        /// <summary>
        /// 净空绕行: 端点在干涉区时, 经两端点上方的安全中继连接, 使路径大部分
        /// 走净空, 只在端点附近保留不可避免的进入段。上方也无净空则返回 null 交人工。
        /// </summary>
        private List<PlanPoint> PlanClearancePath(
            PlanPoint a, PlanPoint b, Vec3 start, Vec3 goal, PlanningReport report)
        {
            _world.SetOrientationBlend(a.RpyZyx, b.RpyZyx, start, goal);

            Vec3 zUp = (GetAxisOf(a.AbsTransform, 2) + GetAxisOf(b.AbsTransform, 2)).Normalized();
            Vec3 mid = Vec3.Lerp(start, goal, 0.5);

            // 上方净空中继搜索 (沿枪+Z 与 世界Z, 高度递增)
            Vec3 safe = default(Vec3);
            bool found = false;
            foreach (double h in new[] { 150.0, 250, 400, 600, 800 })
            {
                foreach (Vec3 dir in new[] { zUp, new Vec3(0, 0, 1) })
                {
                    ThrowIfCancelled();
                    Vec3 cand = mid + dir * h;
                    if (_world.IsPositionFree(cand)) { safe = cand; found = true; break; }
                }
                if (found) break;
            }
            if (!found) return null;

            bool startEdge = _rrt.IsEdgeFree(start, safe);
            bool goalEdge = _rrt.IsEdgeFree(safe, goal);

            report.ClearanceSegments++; // 净空绕行已生成中继, 不计为失败段
            report.Warnings.Add(string.Format(
                "净空绕行: {0}→{1} 已生成高位中继避免路径穿模({2}), 端点工艺位置干涉仍需人工消除",
                a.Source, b.Source,
                (startEdge && goalEdge) ? "两侧净空" : "端点侧仍有不可避免进入段"));

            return new List<PlanPoint>
            {
                new PlanPoint
                {
                    Position = safe, RpyZyx = b.RpyZyx,
                    Kind = PlanPointKind.Transit, Verified = startEdge && goalEdge,
                    Source = a.Source + "→" + b.Source,
                    Note = "净空绕行中继"
                }
            };
        }

        /// <summary>三策略修复阶梯 (返回 null = 全部失败)</summary>
        private List<PlanPoint> TryRepairJointMotion(
            Vec3 pa, Vec3 pb, TxPoseData ja, TxPoseData jb, double violationT,
            Vec3 backDir, Vec3 zAxis, Vec3 yAxis, PlanPoint a, PlanPoint b)
        {
            // ═══ 策略⓪ (v5.4, 最优先): 双侧沿枪 -X 深退门形 ═══
            // 违例的根因通常是"枪体还埋在夹具里就开始侧移/抬升"。
            // 最直接的修复不是在浅位置周围找避让点, 而是**先把枪沿自身轴线抽出来**。
            // 沿 -X 抽出时枪体沿自身轴线运动, 不扫掠周边 —— 这是人工示教的标准动作。
            //
            // 深度由浅到深试 (短路径优先), 但起点就是 GunMinBackoutForSide (200mm),
            // 不再从 40mm 这种"抽了等于没抽"的深度开始。
            {
                foreach (double d in new[] { 200.0, 300, 450, 600, 800 })
                {
                    ThrowIfCancelled();
                    Vec3 qa = pa + backDir * d;
                    Vec3 qb = pb + backDir * d;
                    TxPoseData jqa, jqb;
                    if (!_world.TryGetJointPoseAt(qa, out jqa)) continue;
                    if (!_world.TryGetJointPoseAt(qb, out jqb)) continue;

                    // 门形: a → a后撤 → b后撤 → b
                    if (_world.CheckJointMotion(ja, jqa) == null
                        && _world.CheckJointMotion(jqa, jqb) == null
                        && _world.CheckJointMotion(jqb, jb) == null)
                    {
                        return new List<PlanPoint>
                        {
                            MakeRepairPoint(qa, a, b, string.Format("动态修复(深退{0:F0}门形)", d)),
                            MakeRepairPoint(qb, a, b, string.Format("动态修复(深退{0:F0}门形)", d))
                        };
                    }

                    // 门形 + 抬升: 退出来之后再抬起平移 (进一步拉开净空)
                    foreach (double h in new[] { 150.0, 300, 500 })
                    {
                        Vec3 ra = qa + zAxis * h;
                        Vec3 rb = qb + zAxis * h;
                        TxPoseData jra, jrb;
                        if (!_world.TryGetJointPoseAt(ra, out jra)) continue;
                        if (!_world.TryGetJointPoseAt(rb, out jrb)) continue;

                        if (_world.CheckJointMotion(ja, jqa) == null
                            && _world.CheckJointMotion(jqa, jra) == null
                            && _world.CheckJointMotion(jra, jrb) == null
                            && _world.CheckJointMotion(jrb, jqb) == null
                            && _world.CheckJointMotion(jqb, jb) == null)
                        {
                            return new List<PlanPoint>
                            {
                                MakeRepairPoint(qa, a, b, "动态修复(深退+抬升门形)"),
                                MakeRepairPoint(ra, a, b, "动态修复(深退+抬升门形)"),
                                MakeRepairPoint(rb, a, b, "动态修复(深退+抬升门形)"),
                                MakeRepairPoint(qb, a, b, "动态修复(深退+抬升门形)")
                            };
                        }
                    }
                }
            }

            // ---- 策略①: 1/3+2/3 两点细分 ----
            {
                Vec3 p1 = Vec3.Lerp(pa, pb, 1.0 / 3.0);
                Vec3 p2 = Vec3.Lerp(pa, pb, 2.0 / 3.0);
                TxPoseData j1, j2;
                if (_world.TryGetJointPoseAt(p1, out j1)
                    && _world.TryGetJointPoseAt(p2, out j2)
                    && _world.CheckJointMotion(ja, j1) == null
                    && _world.CheckJointMotion(j1, j2) == null
                    && _world.CheckJointMotion(j2, jb) == null)
                {
                    return new List<PlanPoint>
                    {
                        MakeRepairPoint(p1, a, b, "动态修复(细分)"),
                        MakeRepairPoint(p2, a, b, "动态修复(细分)")
                    };
                }
            }

            // ---- 策略②: 违例点定位避让 ----
            // v5.3: 退让阶梯大幅加深。原来最大只退 80mm — 枪深埋在夹具/工件里时
            // 根本脱不了困 (图1场景: 出枪只退 20mm, 抬升时枪身仍在夹具内扫过)。
            // 现在按"人工示教挪枪"的直觉排序:
            //   先沿枪 -X 深退 (真正脱困方向, 40→600mm)
            //   再"深退 + 抬升"组合 (退出来再抬, 而不是原地抬)
            //   最后才是纯抬升/横移
            if (violationT >= 0)
            {
                double tv = Math.Max(0.15, Math.Min(0.85, violationT));
                Vec3 v = Vec3.Lerp(pa, pb, tv);

                var escapes = new List<Vec3>();

                // ② -a 纯深退 (沿枪-X 逐级加深)
                foreach (double d in new[] { 40.0, 80, 150, 250, 400, 600 })
                    escapes.Add(backDir * d);

                // ② -b 深退 + 抬升 组合 (脱困后再抬 — 这才是人工的做法)
                foreach (double d in new[] { 150.0, 300, 500 })
                    foreach (double h in new[] { 100.0, 200, 350 })
                        escapes.Add(backDir * d + zAxis * h);

                // ② -c 深退 + 横移
                foreach (double d in new[] { 200.0, 400 })
                    foreach (double s in new[] { 120.0, -120, 240, -240 })
                        escapes.Add(backDir * d + yAxis * s);

                // ② -d 纯抬升 / 横移 (原地脱困, 成功率低但便宜)
                foreach (double h in new[] { 60.0, 120, 240, 400 })
                    escapes.Add(zAxis * h);
                escapes.Add(zAxis * -60);
                foreach (double s in new[] { 60.0, -60, 150, -150 })
                    escapes.Add(yAxis * s);

                foreach (var e in escapes)
                {
                    ThrowIfCancelled();
                    Vec3 pc = v + e;
                    TxPoseData jc;
                    if (_world.TryGetJointPoseAt(pc, out jc)
                        && _world.CheckJointMotion(ja, jc) == null
                        && _world.CheckJointMotion(jc, jb) == null)
                    {
                        return new List<PlanPoint>
                        {
                            MakeRepairPoint(pc, a, b, "动态修复(避让)")
                        };
                    }
                }
            }

            // ---- 策略③: 笛卡尔中点兜底 ----
            {
                Vec3 mid = Vec3.Lerp(pa, pb, 0.5);
                TxPoseData jm;
                if (_world.TryGetJointPoseAt(mid, out jm)
                    && _world.CheckJointMotion(ja, jm) == null
                    && _world.CheckJointMotion(jm, jb) == null)
                {
                    return new List<PlanPoint> { MakeRepairPoint(mid, a, b, "动态修复(中点)") };
                }
            }

            return null;
        }

        private PlanPoint MakeRepairPoint(Vec3 pos, PlanPoint a, PlanPoint b, string note)
        {
            TxVector rpy;
            _world.TryFindFreePose(pos, out rpy);

            // v6.1: Source 取根名, 防止对修复点再修复时名字套娃
            // (v6.0 日志出现 "A→B→A→B→A→B..." 这种)
            string sa = RootSource(a.Source);
            string sb = RootSource(b.Source);

            return new PlanPoint
            {
                Position = pos,
                RpyZyx = rpy != null ? rpy : b.RpyZyx,
                Kind = PlanPointKind.Transit, Verified = true,
                Source = sa + "→" + sb,
                Note = note
            };
        }

        /// <summary>取 Source 的根名 (第一段), 防止 "A→B→A→B" 无限拼接</summary>
        private static string RootSource(string s)
        {
            if (string.IsNullOrEmpty(s)) return "?";
            int i = s.IndexOf('→');
            return i > 0 ? s.Substring(0, i) : s;
        }

        /// <summary>段失败登记: 计数 + 逐段警告 (取代旧兜底抬升点)</summary>
        private void MarkSegmentFailed(PlanPoint a, PlanPoint b, string reason, PlanningReport report)
        {
            report.FailedSegments++;
            report.Warnings.Add(string.Format(
                "过渡段 {0}→{1} 规划失败 ({2}), 未插入过渡点, 需人工处理", a.Source, b.Source, reason));
        }

        // ================================================================
        //  Via 插入 (原生 After API: 直接创建在 predecessor 之后)
        //  predecessor==null → 创建为操作第一个元素 (等效插到最前)
        // ================================================================
        private ITxObject InsertViaAfter(ITxObject op, PlanPoint p, ITxObject predecessor)
        {
            string viaName = !string.IsNullOrEmpty(p.CustomName)
                ? p.CustomName
                : string.Format("Via_{0}_{1}", p.Kind, p.Source)
                    .Replace(" ", "_").Replace("→", "_to_");
            try
            {
                var viaData = new TxRoboticViaLocationOperationCreationData(viaName);
                TxRoboticViaLocationOperation via = CreateViaAfter(op, viaData, predecessor);
                if (via == null) return null;

                // 写入位姿
                var t = new TxTransformation();
                t.Translation = new TxVector(p.Position.X, p.Position.Y, p.Position.Z);
                if (p.RpyZyx != null) t.RotationRPY_ZYX = p.RpyZyx;
                ((ITxLocatableObject)via).AbsoluteLocation = t;

                // v6.0: 运动参数 (Joint/PTP) + 焊钳开口 (外部轴)
                // predecessor 通常是相邻焊点或前一个 Via —— 非枪轴 (导轨/变位机)
                // 的工艺值从它继承, 避免漏填导致 PS 把导轨归零
                ApplyViaMotionAndGun(via, p, predecessor);

                return via;
            }
            catch (Exception ex)
            {
                _log(string.Format("  [警告] Via {0} 插入失败: {1}", viaName, ex.Message));
                return null;
            }
        }

        /// <summary>
        /// 原生前插: CreateRoboticViaLocationOperationAfter(data, predecessor)。
        /// predecessor=null → 插为操作第一个。API 不可用时回退 [创建+重排]。
        /// </summary>
        private TxRoboticViaLocationOperation CreateViaAfter(
            ITxObject op, TxRoboticViaLocationOperationCreationData viaData, ITxObject predecessor)
        {
            // 首选: 原生 After (强类型接收者)
            var contOp = op as TxContinuousRoboticOperation;
            if (contOp != null)
            {
                try { return contOp.CreateRoboticViaLocationOperationAfter(viaData, (ITxOperation)predecessor); }
                catch { }
            }
            // 运行时类型 (TxBaseWeldOperation 等): dynamic 接收者 + 强类型实参
            try
            {
                return (TxRoboticViaLocationOperation)((dynamic)op)
                    .CreateRoboticViaLocationOperationAfter(viaData, predecessor);
            }
            catch { }

            // 回退: 无 After 版本 → 追加创建后重排
            try
            {
                TxRoboticViaLocationOperation via = contOp != null
                    ? contOp.CreateRoboticViaLocationOperation(viaData)
                    : (TxRoboticViaLocationOperation)((dynamic)op)
                        .CreateRoboticViaLocationOperation(viaData);
                if (via != null && predecessor != null)
                    TryReorderAfter(op, via, predecessor);
                return via;
            }
            catch (Exception ex)
            {
                _log("  [警告] Via 创建失败: " + ex.Message);
                return null;
            }
        }

        /// <summary>创建 Via 并写入位姿, 追加到操作末尾 (不排序)</summary>
        private ITxObject CreateViaAppended(ITxObject op, PlanPoint p)
        {
            string viaName = !string.IsNullOrEmpty(p.CustomName)
                ? p.CustomName
                : string.Format("Via_{0}_{1}", p.Kind, p.Source)
                    .Replace(" ", "_").Replace("→", "_to_");
            try
            {
                var viaData = new TxRoboticViaLocationOperationCreationData(viaName);
                var contOp = op as TxContinuousRoboticOperation;
                TxRoboticViaLocationOperation via = contOp != null
                    ? contOp.CreateRoboticViaLocationOperation(viaData)
                    : (TxRoboticViaLocationOperation)((dynamic)op)
                        .CreateRoboticViaLocationOperation(viaData);
                if (via == null) return null;
                var t = new TxTransformation();
                t.Translation = new TxVector(p.Position.X, p.Position.Y, p.Position.Z);
                if (p.RpyZyx != null) t.RotationRPY_ZYX = p.RpyZyx;
                ((ITxLocatableObject)via).AbsoluteLocation = t;
                return via;
            }
            catch (Exception ex)
            {
                _log(string.Format("  [警告] Via {0} 创建失败: {1}", viaName, ex.Message));
                return null;
            }
        }

        /// <summary>
        /// 把 child 移到 target 之前: 先找 target 的前驱 prev, 然后 MoveChildAfter(child, prev)。
        /// prev=null 时 MoveChildAfter(child, null) = 插入为第一个 = 在 target 之前。
        /// 这消除了对不存在 API (MoveChildBefore) 的依赖。
        /// </summary>
        private void MoveChildBefore(ITxObject op, ITxObject child, ITxObject target)
        {
            if (ReferenceEquals(child, target)) return;

            // 找 target 的前驱 (prev=null = target 是第一个, child 插为新的第一个即在其前)
            ITxObject prev = FindPredecessor(op, target);

            var ordered = op as ITxOrderedObjectCollection;
            if (ordered != null)
            {
                try { ordered.MoveChildAfter(child, prev); return; }
                catch { }
            }
            try { ((dynamic)op).MoveChildAfter(child, prev); return; } catch { }
            // 兜底: AddObjectAfter (从操作树取出再插入)
            try { ((dynamic)op).AddObjectAfter((dynamic)child, prev); return; } catch { }

            _log(string.Format("  [提示] 引导点 '{0}' 前置排序失败, 需人工调整", GetNameSafe(child)));
        }

        /// <summary>找 target 在操作子序列中的前驱; null = target 是第一个子元素</summary>
        private ITxObject FindPredecessor(ITxObject op, ITxObject target)
        {
            try
            {
                var container = op as ITxObjectCollection;
                if (container == null) return null;
                var filter = new TxTypeFilter(typeof(ITxObject));
                TxObjectList kids = null;
                try { kids = ((dynamic)container).GetDirectDescendants(filter); }
                catch { }
                if (kids == null) kids = container.GetAllDescendants(filter);
                if (kids == null) return null;

                ITxObject prev = null;
                foreach (ITxObject k in kids)
                {
                    if (ReferenceEquals(k, target)) return prev;
                    prev = k;
                }
            }
            catch { }
            return null;
        }

        /// <summary>回退重排 (仅原生 After 不可用时用)</summary>
        private void TryReorderAfter(ITxObject op, ITxObject via, ITxObject predecessor)
        {
            try { ((dynamic)op).MoveChildAfter(via, predecessor); return; } catch { }
            try { ((dynamic)via).MoveLocationInto(op, predecessor); return; } catch { }
            _log(string.Format("  [提示] Via '{0}' 已创建但重排失败(位于末尾)", GetNameSafe(via)));
        }

        // ================================================================
        //  辅助
        // ================================================================
        private List<TxWeldLocationOperation> CollectWeldLocationsOf(ITxObject op)
        {
            var result = new List<TxWeldLocationOperation>();

            var self = op as TxWeldLocationOperation;
            if (self != null) { result.Add(self); return result; }

            try
            {
                var container = op as ITxObjectCollection;
                if (container != null)
                {
                    var filter = new TxTypeFilter(typeof(TxWeldLocationOperation));
                    TxObjectList descendants = container.GetAllDescendants(filter);
                    foreach (ITxObject d in descendants)
                    {
                        var w = d as TxWeldLocationOperation;
                        if (w != null) result.Add(w);
                    }
                }
            }
            catch { }
            return result;
        }

        private PlanPoint ExtractWeldPoint(TxWeldLocationOperation loc)
        {
            try
            {
                var locatable = loc as ITxLocatableObject;
                if (locatable == null) return null;
                TxTransformation abs = locatable.AbsoluteLocation;

                TxVector rpy = null;
                try { rpy = abs.RotationRPY_ZYX; } catch { }

                var p = new PlanPoint
                {
                    Position = new Vec3(abs.Translation.X, abs.Translation.Y, abs.Translation.Z),
                    RpyZyx = rpy,
                    AbsTransform = abs,
                    Kind = PlanPointKind.Weld,
                    Source = loc.Name
                };
                // 缓存工具Z轴到 Note 之外 — 通过变换即时计算，见 GetOffsetDir
                p.Note = SerializeZAxis(abs);
                return p;
            }
            catch { return null; }
        }

        /// <summary>进出枪偏移方向: 工具Z反向 或 世界Z正向</summary>
        private Vec3 GetOffsetDir(PlanPoint p)
        {
            if (UseWorldZForApproach) return new Vec3(0, 0, 1);
            Vec3 z = DeserializeZAxis(p.Note);
            return z * -1.0;
        }

        /// <summary>
        /// 自适应偏移点解析: 自动在多方向阶梯中搜索首个无干涉点。
        /// 方向优先级 (自动搜索, 无需用户勾选):
        ///   1. 焊点 -Z (枪后退, 主方向)
        ///   2. 焊点 +Z (枪前进)
        ///   3. 焊点 ±X (前后) — 按焊点坐标系
        ///   4. 焊点 ±Y (左右)
        ///   5. 世界 +Z (向上抬枪)
        /// 每个方向在 [Min, Max] 区间递减取首个无干涉。
        /// 全方向全距离干涉则取 -Z 最大距离强制生成, 返回 false。
        /// </summary>
        private bool GenerateOffsetPoint(PlanPoint weld, out Vec3 pos)
        {
            _world.SetOrientation(weld.AbsTransform, weld.RpyZyx);

            // 方向阶梯: 焊点局部系 → 世界系
            Vec3 negZ = DeserializeZAxis(weld.Note) * -1.0;  // 焊点 -Z (枪后退)
            Vec3 posZDir = DeserializeZAxis(weld.Note);       // 焊点 +Z
            Vec3 xAxis = GetAxisOf(weld.AbsTransform, 0);     // 焊点 X (前后)
            Vec3 yAxis = GetAxisOf(weld.AbsTransform, 1);     // 焊点 Y (左右)
            Vec3 worldZ = new Vec3(0, 0, 1);                  // 世界 Z

            Vec3[] directions = UseWorldZForApproach
                ? new[] { worldZ } // 用户手动选择世界Z时仅此方向
                : new[] { negZ, posZDir, xAxis * -1.0, xAxis, yAxis * -1.0, yAxis, worldZ };

            double dMax = ApproachRetractDistance;
            double dMin = Math.Min(ApproachRetractMin, dMax);
            double dMid = (dMax + dMin) / 2.0;
            double[] dists = dMax - dMin < 1e-6
                ? new[] { dMax }
                : new[] { dMax, dMid, dMin };

            foreach (Vec3 dir in directions)
            {
                foreach (double d in dists)
                {
                    Vec3 p = weld.Position + dir * d;
                    if (_world.ClassifyStrict(p) == CollisionWorld.PositionState.Free)
                    {
                        pos = p;
                        return true;
                    }
                }
            }

            // 全方向全距离失败: 取 -Z 最大距离强制生成
            pos = weld.Position + negZ * dMax;
            return false;
        }

        /// <summary>从变换矩阵提取指定列的轴向 (0=X,1=Y,2=Z); 失败退化为世界轴</summary>
        private static Vec3 GetAxisOf(TxTransformation t, int col)
        {
            try
            {
                if (t != null)
                    return new Vec3(t[0, col], t[1, col], t[2, col]).Normalized();
            }
            catch { }
            return col == 0 ? new Vec3(1, 0, 0)
                 : col == 1 ? new Vec3(0, 1, 0)
                 : new Vec3(0, 0, 1);
        }

        private static string SerializeZAxis(TxTransformation t)
        {
            try
            {
                return string.Format("{0:R};{1:R};{2:R}", t[0, 2], t[1, 2], t[2, 2]);
            }
            catch { return "0;0;1"; }
        }

        private static Vec3 DeserializeZAxis(string s)
        {
            try
            {
                var parts = s.Split(';');
                return new Vec3(
                    double.Parse(parts[0]),
                    double.Parse(parts[1]),
                    double.Parse(parts[2])).Normalized();
            }
            catch { return new Vec3(0, 0, 1); }
        }

        private static string GetNameSafe(ITxObject obj)
        {
            try { return obj.Name; }
            catch { return obj != null ? obj.GetType().Name : "?"; }
        }
    }
}
