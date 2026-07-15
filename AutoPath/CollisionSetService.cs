using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Tecnomatix.Engineering;

namespace TxTools.AutoPathPlanner
{
    /// <summary>
    /// 干涉集独立服务 — 可复用于任何 TxTools 插件。
    ///
    /// 职责: 创建碰撞对 → 查询当前是否碰撞 → Dispose 时删除碰撞对。
    /// API 基于 RobotReachabilityChecker 验证过的真实路径:
    ///   TxApplication.ActiveDocument.CollisionRoot
    ///   → CreateCollisionPair → HasCollidingObjects
    ///
    /// 典型用法:
    ///   using (var cs = CollisionSetService.CreateRobotVsWorld(robot, null, null, log))
    ///   {
    ///       // ...把机器人摆到某个姿态...
    ///       bool hit = cs.QueryColliding();
    ///   }
    ///
    /// ══════════════════════════════════════════════════════════════
    /// v5.1 绑定资源完整收集 (TxAgent 实测 API):
    ///
    ///   【问题】FirstList 只有 5 项, fupa 底座落在 SecondList → 机器人与
    ///   自己的底座永久报干涉。原因: CollectBoundResources 漏收了
    ///   PrimaryLocator (底座) / Toolbox / AttachmentParent / Links;
    ///   且靠 IsDescendantOf 父子链排除障碍 —— 但 fupa 不在机器人物理子树内。
    ///
    ///   【TxRobot 实测 API】(inspect_type + run_csharp 双重确认):
    ///     MountedTools    : TxObjectList        ← 焊钳 TxServoGun ×2
    ///     PrimaryLocator  : ITxObjectCollection ← "KR210_R2700-2-fupa" [Count=8] ★
    ///     Toolbox         : TxRobotToolbox      ← 常为 null, 必须判空
    ///     AttachmentParent: ITxLocatableObject  ← 滑轨/外部轴
    ///     Links           : TxObjectList        ← TxKinematicLink ×9
    ///     Controller / SimulatingOperations / RoboticPrograms / Signals /
    ///     PoseList / DrivingJoints / TCPF等Frame → 无几何, 不收
    ///     **不存在** DressPacks / ExternalAxes / MountedDressPacks (之前一直在瞎猜)
    ///
    ///   【双向硬保证】
    ///     FirstList  = CollectBoundResources() 全部绑定资源
    ///     SecondList = CollectWorldObstacles() 用 HashSet 引用级排除绑定资源集
    ///                  (不再依赖 IsDescendantOf 拓扑判断, fupa 这类靠这条兜住)
    ///     + SanitizeObstacles() 最终安全网, 保证 First ∩ Second = ∅
    /// ══════════════════════════════════════════════════════════════
    ///
    /// ══════════════════════════════════════════════════════════════
    /// v5.0 致命 bug 修复 (强类型 API 重写):
    ///
    ///   【根因】v4.2~v4.10 一直用 root.CollisionPairs / root.GetAllCollisionPairs(),
    ///   但 TxCollisionRoot 上**根本没有这两个成员** (inspect_type 实测确认)。
    ///   dynamic 访问抛异常 → 被 catch 吞掉 → pairs 恒为 null → 查找函数永远
    ///   return null → "永远未找到干涉集"。之前所有 First 侧反射、指纹匹配、
    ///   宽松兜底代码一行都没被执行过 —— 全在那个 return 之后。
    ///
    ///   【正确 API】(TxAgent inspect_type 实测):
    ///     TxCollisionRoot.Pairs           : ArrayList        ← 碰撞对列表
    ///     TxCollisionRoot.PairList        : TxObjectList     ← 同上, 另一形态
    ///     TxCollisionRoot.CheckCollisions : Boolean get set  ← 全局开关!
    ///     TxCollisionPair.FirstList       : TxObjectList get set
    ///     TxCollisionPair.SecondList      : TxObjectList get set
    ///     TxCollisionPair.Active          : Boolean get set
    ///
    ///   【改动】
    ///     - GetAllPairs(): 强类型读 root.Pairs (兜底 PairList)
    ///     - PairFirstListHasRobot(): 强类型读 pair.FirstList, 只看 First 忽略 Second,
    ///       按 名字 + 位置(1mm) 匹配 TxRobot
    ///     - 删除全部反射代码 (PairFirstSideMatches / CollectFirstSide... /
    ///       MakeFingerprint / Fp / LooksLikeSecondSide / FindAnyPairLoose 等)
    ///     - _pair 字段改强类型 TxCollisionPair; Active / Delete 直接调用
    ///     - 新建 pair 时若 root.CheckCollisions == false 则自动开启
    ///       (实测场景中该全局开关为 false, 会导致所有 pair 不参与计算)
    /// ══════════════════════════════════════════════════════════════
    ///
    /// v4.10 附着模式:
    ///   - CreateRobotVsWorld 新增 attachOnly 参数; BuildPair 新增同名参数
    ///   - 规划器改用 attachOnly=true: 只复用现有干涉集, 找不到跳过操作, 绝不新建
    ///     (修复"注释检测后规划反而每次新建堆积"的问题)
    ///   - 附着失败时先严格匹配 (FindExistingPairForRobot), 再宽松匹配
    ///     (FindAnyPairLoose: 任意侧同名机器人即命中, 不比位置)
    ///   - UI"自动创建"按钮不受影响 (仍可新建)
    ///
    /// v4.9 鲁棒化:
    ///   - MakeFingerprint 从 ITxLocatableObject cast 改为 dynamic 拿位置 (兼容更多 SDK 版本)
    ///   - CollectFirstSideRobotAndGunFingerprints 从"仅 TxObjectList 属性"改为
    ///     属性 + 无参 Get* 方法 + IEnumerable 兜底 (兼容 pair First 侧的各种承载类型)
    ///   - FindExistingPairForRobot / TryFindExistingPairName 新增可选 log 参数,
    ///     找不到时 dump 期望指纹 + pair 数量 + 前几个 pair 的 First 侧候选,
    ///     方便定位 (a) 期望没生成 (b) 反射提取失败 (c) 位置不匹配
    ///
    /// v4.8 修复 (严格 First 侧对比):
    ///   - FindExistingPairForRobot 从"仅比机器人身份"改为"机器人+焊钳指纹集合匹配"
    ///     · 只看 pair.First 侧, 忽略 Second
    ///     · 提取 First 侧的 TxRobot / 焊钳, 以 (名字, 位置) 指纹存储
    ///     · 期望集 = 目标 robot + 挂载焊钳, 每项必须在 First 找到 名字+位置(1mm) 匹配
    ///   - 用位置匹配代替 ReferenceEquals, 天然免疫"PS 每次访问返回不同代理"问题
    ///   - 抽出 CollectMountedTools 供 FindExistingPair 和 CollectBoundResources 复用
    ///
    /// v4.5 修复 (API 优化):
    ///   - CollectWorldObstacles / customObstacles 首层剔除 ITxComponent.IsEquipment
    ///     容器 (官方 API 语义: "可包含子组件的容器"), 避免父级 Cell/Group
    ///     以整个容器边界参与碰撞检测导致大量误报
    ///
    /// v4.4 简化:
    ///   - CollectWorldObstacles 一律剔除场景所有 TxRobot / 焊枪 子树,
    ///     无论检测方是否精确收集绑定资源, 都不会误报机器人/焊枪永久干涉
    ///   - customObstacles 追加同步剔除
    ///
    /// v4 新增:
    ///   - TryFindExistingPairName: UI 层入口, 检测已有干涉集并取名
    ///   - CreateRobotVsWorld(..., forceNew): 跳过复用检测, 强制新建
    ///   - KeepPairOnDispose: 保留本次创建的碰撞对 (持久干涉集)
    /// </summary>
    public sealed class CollisionSetService : IDisposable
    {
        private TxCollisionPair _pair; // v5.0: 强类型 (API 已由 inspect_type 确认)
        private readonly Action<string> _log;

        public bool IsReady { get { return _pair != null; } }
        public int CheckObjectCount { get; private set; }
        public int ObstacleObjectCount { get; private set; }

        private CollisionSetService(Action<string> log)
        {
            _log = log ?? delegate { };
        }

        // ================================================================
        //  工厂方法
        // ================================================================

        /// <summary>
        /// 通用入口: 任意检测方 vs 任意障碍方。
        /// 注意: 此入口不传 robot，不会检测已有干涉集。
        /// </summary>
        public static CollisionSetService Create(
            IEnumerable<ITxObject> checkObjects,
            IEnumerable<ITxObject> obstacles,
            Action<string> log)
        {
            var svc = new CollisionSetService(log);
            svc.BuildPair(checkObjects, obstacles); // robot=null → 不检测已有干涉集
            return svc;
        }

        /// <summary>
        /// 焊接场景常用入口: (机器人 + 全部绑定资源) vs 场景障碍。
        ///
        /// FirstList (检测方) = CollectBoundResources(robot) — v5.1 强类型收集:
        ///   robot 本体 / MountedTools(焊钳) / PrimaryLocator(**底座 fupa**) /
        ///   Toolbox 工具 / AttachmentParent(滑轨/外轴) / Links / 以上各项的 ITxComponent 后代
        ///
        /// SecondList (障碍方) = CollectWorldObstacles(check) — 硬保证不含任何绑定资源:
        ///   PhysicalRoot 全部 ITxComponent, 剔除 [绑定资源集(HashSet引用级) +
        ///   Equipment 容器 + 所有 TxRobot/焊枪子树]
        ///
        /// 双向保证: 绑定资源只在 First, 绝不在 Second。构建 pair 前还有一道
        /// 最终安全网 (SanitizeObstacles) 做交集校验并剔除。
        ///
        /// 已有干涉集检测: CollisionRoot 中已有以本机器人为 FirstList 的碰撞对则复用,
        /// Dispose 时不删除。
        ///
        /// forceNew=true 跳过复用检测, 强制新建 (与已有并存)。
        /// attachOnly=true 纯附着模式: 只复用现有干涉集, 找不到也不新建 (IsReady=false)。
        /// 二者互斥, 同时为 true 时 forceNew 优先。
        /// </summary>
        public static CollisionSetService CreateRobotVsWorld(
            TxRobot robot,
            IEnumerable<ITxObject> extraCheckObjects,
            IEnumerable<ITxObject> customObstacles,
            Action<string> log,
            bool forceNew = false,
            bool attachOnly = false)
        {
            var lg = log ?? delegate { };

            // ---- FirstList: 机器人全部绑定资源 ----
            var check = CollectBoundResources(robot);
            if (extraCheckObjects != null)
                check.AddRange(extraCheckObjects.Where(o => o != null && !check.Any(c => ReferenceEquals(c, o))));

            LogBoundResourceSummary(robot, check, lg);

            // ---- SecondList: 场景障碍 (硬排除全部绑定资源) ----
            List<ITxObject> obstacles = CollectWorldObstacles(check);

            // 追加用户自定义障碍 (与 CollectWorldObstacles 同一套剔除规则)
            if (customObstacles != null)
            {
                var checkSet0 = new HashSet<ITxObject>(ReferenceComparer.Instance);
                foreach (var c in check) if (c != null) checkSet0.Add(c);

                int added = 0, skippedInternal = 0, skippedRobotOrGun = 0, skippedEquipment = 0;
                foreach (var o in customObstacles.Where(o => o != null))
                {
                    // ① 绑定资源 (引用级硬排除)
                    if (checkSet0.Contains(o)) { skippedInternal++; continue; }

                    // ② Equipment 容器
                    if (IsEquipmentContainer(o)) { skippedEquipment++; continue; }

                    // ③④ TxRobot / 焊枪 子树
                    if (IsRobotOrGunSubtree(o)) { skippedRobotOrGun++; continue; }

                    // ① 补充: 绑定资源的父容器 / 子树
                    bool isInternal = false;
                    foreach (var c in check)
                    {
                        if (c == null) continue;
                        if (IsDescendantOf(o, c) || IsDescendantOf(c, o))
                        { isInternal = true; break; }
                    }
                    if (isInternal) { skippedInternal++; continue; }

                    if (!obstacles.Any(ob => ReferenceEquals(ob, o)))
                    { obstacles.Add(o); added++; }
                }
                if (added > 0)
                    lg(string.Format("  自定义障碍追加: {0} 项 (剔除 绑定资源{1}/机器人焊枪{2}/Equipment容器{3})",
                        added, skippedInternal, skippedRobotOrGun, skippedEquipment));
                else if (skippedInternal + skippedRobotOrGun + skippedEquipment > 0)
                    lg(string.Format("  [提示] 自定义障碍全部被剔除 (绑定资源{0}/机器人焊枪{1}/Equipment容器{2})",
                        skippedInternal, skippedRobotOrGun, skippedEquipment));
            }

            // ---- 最终安全网: 校验 obstacles 与 check 无交集 ----
            // 双保险 —— 即使上游有疏漏, 这里也保证绑定资源绝不进 SecondList。
            int removed = SanitizeObstacles(check, obstacles);
            if (removed > 0)
                lg(string.Format("  [安全网] 从障碍方剔除 {0} 项绑定资源 (上游漏网)", removed));

            lg(string.Format("  → FirstList {0} 项 / SecondList {1} 项",
                check.Count, obstacles.Count));

            var svc = new CollisionSetService(log);
            // forceNew=true → robot 传 null 屏蔽复用检测, 强制新建
            // attachOnly=true → 纯附着, 找不到不新建
            svc.BuildPair(check, obstacles, forceNew ? null : robot, attachOnly && !forceNew);
            return svc;
        }

        /// <summary>
        /// 最终安全网: 从 obstacles 中移除任何出现在 check 中的对象 (引用级)。
        /// 返回移除数量。保证 FirstList ∩ SecondList = ∅。
        /// </summary>
        private static int SanitizeObstacles(List<ITxObject> check, List<ITxObject> obstacles)
        {
            if (check == null || obstacles == null) return 0;

            var checkSet = new HashSet<ITxObject>(ReferenceComparer.Instance);
            foreach (var c in check)
                if (c != null) checkSet.Add(c);

            int before = obstacles.Count;
            obstacles.RemoveAll(o => o == null || checkSet.Contains(o));
            return before - obstacles.Count;
        }

        /// <summary>打印绑定资源收集摘要 (按类型分组, 便于核对 fupa/焊钳是否收全)</summary>
        private static void LogBoundResourceSummary(TxRobot robot, List<ITxObject> check, Action<string> log)
        {
            if (log == null || check == null) return;
            try
            {
                log(string.Format("  绑定资源 (FirstList): {0} 项", check.Count));

                // 按类型名分组统计
                var byType = new Dictionary<string, int>();
                foreach (var o in check)
                {
                    if (o == null) continue;
                    string tn = o.GetType().Name;
                    if (byType.ContainsKey(tn)) byType[tn]++;
                    else byType[tn] = 1;
                }
                var parts = new List<string>();
                foreach (var kv in byType.OrderByDescending(k => k.Value))
                    parts.Add(string.Format("{0}×{1}", kv.Key, kv.Value));
                log("    类型: " + string.Join(", ", parts.ToArray()));

                // 列出关键项 (机器人 / 焊钳 / 底座) 便于人工核对
                foreach (var o in check)
                {
                    if (o == null) continue;
                    bool key = (o is TxRobot) || IsGunType(o);
                    string nm = null;
                    try { nm = o.Name; } catch { }
                    // fupa / 底座 通常名字含 fupa / base
                    if (!key && !string.IsNullOrEmpty(nm))
                    {
                        string low = nm.ToLowerInvariant();
                        if (low.Contains("fupa") || low.Contains("base") || low.Contains("sockel"))
                            key = true;
                    }
                    if (key)
                        log(string.Format("    ★ {0} [{1}]", nm ?? "?", o.GetType().Name));
                }
            }
            catch { }
        }

        // ================================================================
        //  对象收集
        // ================================================================

        /// <summary>
        /// 收集机器人的全部**物理绑定资源** —— 随机器人一起运动的东西。
        /// 这些必须全部进 FirstList (检测方), 且绝不能出现在 SecondList (障碍方)。
        ///
        /// v5.1 强类型重写 (API 由 TxAgent inspect_type + run_csharp 实测确认):
        ///   ① robot 本体
        ///   ② robot.MountedTools    : TxObjectList  ← 焊钳 (TxServoGun 等)
        ///   ③ robot.PrimaryLocator  : ITxObjectCollection ← **底座/fupa**!
        ///      实测 KR210_R2700-2.PrimaryLocator = "KR210_R2700-2-fupa" [Count=8]
        ///      之前一直漏收 → fupa 落入障碍方 → 机器人与自己的底座报永久干涉
        ///   ④ robot.Toolbox.GetAllTools() : TxRobotWorkTool[]  (Toolbox 可能为 null!)
        ///   ⑤ robot.AttachmentParent : ITxLocatableObject  ← 机器人挂载的父体 (滑轨/外部轴)
        ///   ⑥ robot.Links           : TxObjectList  ← 运动学连杆 (可能持有几何)
        ///   ⑦ 以上各项的 ITxComponent 后代 (递归)
        ///
        /// 不收集 (无几何, 不参与碰撞):
        ///   Controller / SimulatingOperations / RoboticPrograms / Signals /
        ///   PoseList / DrivingJoints / TCPF-Toolframe-Baseframe-Referenceframe
        ///
        /// 已废弃: DressPacks / ExternalAxes / MountedDressPacks —— TxRobot 上
        /// **不存在这些属性** (inspect_type 确认), 之前的反射查找永远返回 null。
        /// 线缆包与外轴实际通过 PrimaryLocator / AttachmentParent / 物理子树体现。
        /// </summary>
        public static List<ITxObject> CollectBoundResources(TxRobot robot)
        {
            var result = new List<ITxObject>();
            if (robot == null) return result;

            // ① 机器人本体 + 其 ITxComponent 后代 (连杆几何等)
            result.Add(robot);
            AddDescendantComponents(robot, result);

            // ② MountedTools — 焊钳 (TxServoGun)
            try
            {
                TxObjectList tools = robot.MountedTools;
                if (tools != null)
                {
                    foreach (ITxObject tool in tools)
                    {
                        if (!AddUnique(result, tool)) continue;
                        AddDescendantComponents(tool, result);
                        // 工具上的二级挂载 (传感器等)
                        AddNestedMountedTools(tool, result, 0);
                    }
                }
            }
            catch { }

            // ③ PrimaryLocator — 机器人底座 / fupa (关键!之前一直漏收)
            try
            {
                ITxObjectCollection ploc = robot.PrimaryLocator;
                if (ploc != null)
                {
                    var plocObj = ploc as ITxObject;
                    if (plocObj != null)
                    {
                        AddUnique(result, plocObj);
                        AddDescendantComponents(plocObj, result);
                    }
                }
            }
            catch { }

            // ④ Toolbox 里注册的工具 (Toolbox 常为 null, 必须保护)
            try
            {
                TxRobotToolbox tb = robot.Toolbox;
                if (tb != null)
                {
                    TxRobotWorkTool[] allTools = tb.GetAllTools();
                    if (allTools != null)
                    {
                        foreach (var wt in allTools)
                        {
                            var o = wt as ITxObject;
                            if (o != null && AddUnique(result, o))
                                AddDescendantComponents(o, result);
                        }
                    }
                }
            }
            catch { }

            // ⑤ AttachmentParent — 机器人挂在什么上 (滑轨/外部轴/变位机)
            try
            {
                ITxLocatableObject ap = robot.AttachmentParent;
                if (ap != null)
                {
                    var apObj = ap as ITxObject;
                    if (apObj != null && AddUnique(result, apObj))
                        AddDescendantComponents(apObj, result);
                }
            }
            catch { }

            // ⑥ Links — 运动学连杆 (部分连杆直接持有几何)
            try
            {
                TxObjectList links = robot.Links;
                if (links != null)
                {
                    foreach (ITxObject lk in links)
                    {
                        if (AddUnique(result, lk))
                            AddDescendantComponents(lk, result);
                    }
                }
            }
            catch { }

            return result;
        }

        /// <summary>递归收集工具上的二级/三级挂载 (深度上限防环)</summary>
        private static void AddNestedMountedTools(ITxObject tool, List<ITxObject> result, int depth)
        {
            if (tool == null || depth > 4) return;
            try
            {
                dynamic dynTool = tool;
                var subMounted = dynTool.MountedTools as IEnumerable;
                if (subMounted == null) return;
                foreach (var sub in subMounted)
                {
                    var s = sub as ITxObject;
                    if (s == null || !AddUnique(result, s)) continue;
                    AddDescendantComponents(s, result);
                    AddNestedMountedTools(s, result, depth + 1);
                }
            }
            catch { }
        }

        /// <summary>将 ITxComponent 后代添加到列表 (去重)</summary>
        private static void AddDescendantComponents(ITxObject obj, List<ITxObject> result)
        {
            try
            {
                var container = obj as ITxObjectCollection;
                if (container == null) return;
                var filter = new TxTypeFilter(typeof(ITxComponent));
                TxObjectList descendants = container.GetAllDescendants(filter);
                if (descendants == null) return;
                foreach (ITxObject d in descendants)
                    AddUnique(result, d);
            }
            catch { }
        }

        /// <summary>去重添加: 仅引用不重复时才加入</summary>
        private static bool AddUnique(List<ITxObject> list, ITxObject item)
        {
            if (item == null) return false;
            for (int i = 0; i < list.Count; i++)
                if (ReferenceEquals(list[i], item)) return false;
            list.Add(item);
            return true;
        }

        /// <summary>仅收集直接挂载工具 (供外部复用, 如 CollisionWorld.ApplyGunOpenPoses)</summary>
        public static List<ITxObject> CollectMountedTools(TxRobot robot)
        {
            var result = new List<ITxObject>();
            if (robot == null) return result;
            try
            {
                TxObjectList mounted = robot.MountedTools;   // v5.1: 强类型
                if (mounted != null)
                {
                    foreach (ITxObject t in mounted)
                    {
                        if (t != null && !result.Any(r => ReferenceEquals(r, t)))
                            result.Add(t);
                    }
                }
            }
            catch { }
            return result;
        }


        /// <summary>
        /// 收集障碍方 (SecondList): PhysicalRoot 下所有 ITxComponent, 剔除:
        ///
        ///   ① **检测方绑定资源** (boundResources 及其子树)          ← v5.1 硬保证
        ///   ② Equipment 容器 (ITxComponent.IsEquipment == true)     (v4.5, 官方 API)
        ///   ③ 场景所有 TxRobot 及其子树                              (v4.4)
        ///   ④ 场景所有焊枪类 (类型名含 "Gun") 及其子树                (v4.4)
        ///
        /// v5.1 关键修复: 原先只靠 IsDescendantOf 双向父子链判断来排除检测方,
        /// 但机器人的底座 (PrimaryLocator / fupa) 与挂载父体 (AttachmentParent)
        /// **不在机器人的物理父子链上** —— IsDescendantOf 判不出来 →
        /// fupa 落入障碍方 → 机器人与自己的底座永久报干涉。
        ///
        /// 现在改为: 用 CollectBoundResources 收集到的完整集合做**引用级 HashSet 排除**,
        /// 只要是绑定资源 (或其子树), 一律不进障碍方。这是硬保证, 与拓扑关系无关。
        /// 同时保留 IsDescendantOf 双向检查, 覆盖"绑定资源的父容器"这类间接情况。
        /// </summary>
        public static List<ITxObject> CollectWorldObstacles(IList<ITxObject> boundResources)
        {
            var result = new List<ITxObject>();

            // 绑定资源引用集 — O(1) 排除, 硬保证不入障碍方
            var boundSet = new HashSet<ITxObject>(ReferenceComparer.Instance);
            if (boundResources != null)
            {
                foreach (var b in boundResources)
                    if (b != null) boundSet.Add(b);
            }

            try
            {
                var filter = new TxTypeFilter(typeof(ITxComponent)); // null filter 会 NRE
                TxObjectList allComps = TxApplication.ActiveDocument
                    .PhysicalRoot.GetAllDescendants(filter);

                foreach (ITxObject comp in allComps)
                {
                    // ① 绑定资源 — 引用级硬排除 (fupa / AttachmentParent 等靠这条)
                    if (boundSet.Contains(comp)) continue;

                    // ② Equipment 容器 (无独立几何的逻辑分组)
                    if (IsEquipmentContainer(comp)) continue;

                    // ③④ 场景所有 TxRobot / 焊枪 子树
                    if (IsRobotOrGunSubtree(comp)) continue;

                    // ① 补充: 绑定资源的父容器 / 子树 (拓扑关联)
                    bool excluded = false;
                    if (boundResources != null)
                    {
                        foreach (var ex in boundResources)
                        {
                            if (ex == null) continue;
                            if (IsDescendantOf(comp, ex)     // comp 在绑定资源子树内
                                || IsDescendantOf(ex, comp)) // comp 是绑定资源的祖先容器
                            {
                                excluded = true;
                                break;
                            }
                        }
                    }
                    if (!excluded) result.Add(comp);
                }
            }
            catch { }
            return result;
        }

        /// <summary>按引用比较的 ITxObject 相等器 — PS 对象不可靠 Equals/GetHashCode</summary>
        private sealed class ReferenceComparer : IEqualityComparer<ITxObject>
        {
            public static readonly ReferenceComparer Instance = new ReferenceComparer();
            public bool Equals(ITxObject a, ITxObject b) { return ReferenceEquals(a, b); }
            public int GetHashCode(ITxObject o)
            {
                return o == null ? 0 : System.Runtime.CompilerServices
                    .RuntimeHelpers.GetHashCode(o);
            }
        }

        /// <summary>
        /// 判定是否 Equipment 容器 —— 官方 API 语义:
        /// ITxComponent.IsEquipment = "组件可包含其他子组件"。
        /// 这类容器无独立几何, 只是逻辑分组, 参与碰撞检测会产生大量误报,
        /// 且其内部真实设备已被 GetAllDescendants 单独遍历到, 不会漏检。
        /// </summary>
        private static bool IsEquipmentContainer(ITxObject obj)
        {
            if (obj == null) return false;
            try
            {
                var comp = obj as ITxComponent;
                if (comp == null) return false;
                return comp.IsEquipment;
            }
            catch { return false; }
        }

        /// <summary>
        /// 判定 obj 本身或其祖先链是否包含 TxRobot / 焊枪 (类型名含 "Gun")。
        /// 命中 = obj 位于机器人或焊枪的子树 → 障碍方一律剔除。
        /// 焊枪匹配用类型名而非精确类型, 覆盖 TxWeldGun / TxServoGun /
        /// TxPneumaticGun / TxSpotGun 等各种子类, 不需知道具体 SDK 类名。
        /// </summary>
        private static bool IsRobotOrGunSubtree(ITxObject obj)
        {
            if (obj == null) return false;
            try
            {
                object cur = obj;
                for (int depth = 0; depth < 64; depth++)
                {
                    if (cur == null) return false;

                    // TxRobot: 强类型判断, 最可靠
                    if (cur is TxRobot) return true;

                    // 焊枪: 类型名字符串匹配 (含继承链)
                    var t = cur.GetType();
                    while (t != null && t != typeof(object))
                    {
                        if (t.Name.IndexOf("Gun", StringComparison.OrdinalIgnoreCase) >= 0)
                            return true;
                        t = t.BaseType;
                    }

                    // 走祖先 (v6.5: ITxObject 没有 Parent 属性! 用 Collection)
                    ITxObject parent = GetParent(cur as ITxObject);
                    if (parent == null) return false;
                    cur = parent;
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 祖先链遍历。
        ///
        /// v6.5 重要修复: **ITxObject 没有 Parent 属性** (文档确认) ——
        /// 之前 dynamic cur.Parent 每次都抛异常被 catch 吞掉, 这个函数
        /// **一直返回 false**, 祖先链排除从未真正生效过 (幸亏 v5.1 加的
        /// HashSet 引用级硬排除兜住了)。
        ///
        /// 正确路径: ITxObject.Collection (返回 ITxObjectCollection, 即父容器)。
        /// </summary>
        public static bool IsDescendantOf(ITxObject obj, ITxObject ancestor)
        {
            if (obj == null || ancestor == null) return false;

            ITxObject cur = obj;
            for (int depth = 0; depth < 64; depth++)
            {
                ITxObject parent = GetParent(cur);
                if (parent == null) return false;
                if (ReferenceEquals(parent, ancestor)) return true;
                cur = parent;
            }
            return false;
        }

        /// <summary>取父容器 (ITxObject.Collection, 强类型)</summary>
        private static ITxObject GetParent(ITxObject o)
        {
            if (o == null) return null;
            try { return o.Collection as ITxObject; }
            catch { return null; }
        }

        /// <summary>
        /// 检测 CollisionRoot 中已有碰撞对是否**以指定机器人为检测方**。
        ///
        /// v5.0 重写 (强类型 API, 由 TxAgent inspect_type 实测确认):
        ///   TxCollisionRoot.Pairs      : ArrayList        ← 碰撞对列表 (正确属性名!)
        ///   TxCollisionRoot.PairList   : TxObjectList     ← 同上, 另一种形态
        ///   TxCollisionPair.FirstList  : TxObjectList get set
        ///   TxCollisionPair.SecondList : TxObjectList get set
        ///   TxCollisionPair.Active     : Boolean get set
        ///
        /// 致命 bug 修复: v4.2~v4.10 一直用 root.CollisionPairs / root.GetAllCollisionPairs(),
        /// 这两个成员**根本不存在** → dynamic 抛异常被 catch 吞掉 → pairs 恒为 null
        /// → 函数永远 return null → "永远未找到干涉集"。之前所有 First 侧反射/指纹匹配
        /// 代码一行都没被执行过。改用正确的 Pairs 属性后, 强类型直读即可, 反射全部删除。
        ///
        /// 匹配语义: 只看 FirstList (忽略 SecondList), 命中条件 = FirstList 中存在
        /// 与目标 robot 名字相同、位置相同 (1mm 容差) 的 TxRobot。
        /// 用 名字+位置 而非 ReferenceEquals, 兼容 PS 可能返回不同 CLR 代理的情况;
        /// 同名多机器人 (8~18 台) 靠基座位置天然区分。
        /// </summary>
        private static TxCollisionPair FindExistingPairForRobot(TxRobot robot, Action<string> log = null)
        {
            if (robot == null) return null;

            var pairs = GetAllPairs();
            if (pairs.Count == 0)
            {
                if (log != null) log("  [干涉集] CollisionRoot.Pairs 为空 — 场景无任何碰撞对");
                return null;
            }

            foreach (var pair in pairs)
            {
                try
                {
                    if (PairFirstListHasRobot(pair, robot))
                        return pair;
                }
                catch { continue; }
            }

            if (log != null)
                log(string.Format("  [干涉集] {0} 个碰撞对均无匹配 (FirstList 中未找到 '{1}' @ 当前位置)",
                    pairs.Count, robot.Name));
            return null;
        }

        /// <summary>
        /// 读取 CollisionRoot 的全部碰撞对 (强类型)。
        /// 主路径 Pairs (ArrayList), 兜底 PairList (TxObjectList)。
        /// </summary>
        private static List<TxCollisionPair> GetAllPairs()
        {
            var result = new List<TxCollisionPair>();
            try
            {
                TxCollisionRoot root = TxApplication.ActiveDocument.CollisionRoot;
                if (root == null) return result;

                // 主路径: Pairs (ArrayList)
                try
                {
                    ArrayList arr = root.Pairs;
                    if (arr != null)
                    {
                        foreach (var o in arr)
                        {
                            var p = o as TxCollisionPair;
                            if (p != null) result.Add(p);
                        }
                    }
                }
                catch { }

                // 兜底: PairList (TxObjectList)
                if (result.Count == 0)
                {
                    try
                    {
                        TxObjectList lst = root.PairList;
                        if (lst != null)
                        {
                            foreach (ITxObject o in lst)
                            {
                                var p = o as TxCollisionPair;
                                if (p != null) result.Add(p);
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return result;
        }

        /// <summary>
        /// pair.FirstList 中是否存在与 robot 同名且同位置 (1mm 容差) 的 TxRobot。
        /// 只看 FirstList, 完全忽略 SecondList。
        /// </summary>
        private static bool PairFirstListHasRobot(TxCollisionPair pair, TxRobot robot)
        {
            if (pair == null || robot == null) return false;

            TxObjectList first;
            try { first = pair.FirstList; }
            catch { return false; }
            if (first == null || first.Count == 0) return false;

            string targetName;
            try { targetName = robot.Name; }
            catch { return false; }

            double rx, ry, rz;
            bool gotRobotPos = TryGetPos(robot, out rx, out ry, out rz);

            foreach (ITxObject o in first)
            {
                var r = o as TxRobot;
                if (r == null) continue;

                // 快路径: 引用相同
                if (ReferenceEquals(r, robot)) return true;

                // 名字必须一致
                string n;
                try { n = r.Name; }
                catch { continue; }
                if (!string.Equals(n, targetName, StringComparison.Ordinal)) continue;

                // 位置一致 (同名多机器人靠基座区分)
                if (!gotRobotPos) return true;  // 拿不到目标位置 → 退化为仅名字匹配
                double x, y, z;
                if (!TryGetPos(r, out x, out y, out z)) continue;
                if (Math.Abs(x - rx) < 1.0 && Math.Abs(y - ry) < 1.0 && Math.Abs(z - rz) < 1.0)
                    return true;
            }
            return false;
        }

        /// <summary>读取对象的世界坐标 (AbsoluteLocation.Translation)</summary>
        private static bool TryGetPos(ITxObject obj, out double x, out double y, out double z)
        {
            x = y = z = 0;
            if (obj == null) return false;
            try
            {
                var loc = obj as ITxLocatableObject;
                if (loc == null) return false;
                var t = loc.AbsoluteLocation.Translation;
                x = t.X; y = t.Y; z = t.Z;
                return true;
            }
            catch { return false; }
        }

        /// <summary>类型名字符串是否含 "Gun" (含继承链) — 覆盖 TxWeldGun/TxServoGun/TxPneumaticGun 等</summary>
        private static bool IsGunType(ITxObject obj)
        {
            if (obj == null) return false;
            var t = obj.GetType();
            while (t != null && t != typeof(object))
            {
                if (t.Name.IndexOf("Gun", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
                t = t.BaseType;
            }
            return false;
        }

        private static string TryGetName(object obj)
        {
            try { return ((dynamic)obj).Name as string ?? "?"; }
            catch { return "?"; }
        }

        /// <summary>
        /// UI 层入口: 检测 CollisionRoot 中是否已有以指定机器人为检测方 (FirstList) 的碰撞对。
        /// 无匹配返回 false; 有则返回 true 并输出碰撞对名 (供弹窗提示)。
        ///
        /// v5.0: 底层改用强类型 root.Pairs → pair.FirstList, 修复"永远找不到"的致命 bug。
        /// </summary>
        public static bool TryFindExistingPairName(TxRobot robot, out string pairName,
            Action<string> log = null)
        {
            pairName = null;
            TxCollisionPair pair = FindExistingPairForRobot(robot, log);
            if (pair == null) return false;
            pairName = TryGetName(pair);
            if (string.IsNullOrEmpty(pairName) || pairName == "?")
                pairName = "(未命名碰撞对)";
            return true;
        }

        // ================================================================
        //  碰撞对创建
        // ================================================================

        private readonly List<ITxObject> _checkObjects = new List<ITxObject>();
        private TxObjectList _checkList, _obstacleList;
        private int _fromLists; // 0=未探测 1=可用 -1=不可用
        private bool _ownedPair; // true = 本插件创建, Dispose 时删除; false = 复用已有, 不删除

        /// <summary>
        /// 设为 true 时, Dispose 保留本插件新建的碰撞对 (不 Delete)。
        /// 用于"持久干涉集"场景 —— UI"自动创建"按钮 / 规划完保留供复用。
        /// 对复用得来的碰撞对无影响 (复用的本来就不会被删)。
        /// </summary>
        public bool KeepPairOnDispose { get; set; }

        /// <summary>
        /// 构建/复用碰撞对。
        ///
        /// robot 参数:
        ///   非 null → 尝试 FindExistingPairForRobot 复用; attachOnly=false 时未命中则新建
        ///   null    → 不检测已有, 直接新建 (forceNew 场景)
        ///
        /// attachOnly 参数 (v4.10):
        ///   true  → 纯附着模式: 严格匹配命中则复用; 未命中则尝试宽松查找 (任意含该机器人
        ///           名字的 pair); 再不行才放弃 (不新建, _pair=null)。用于规划器 ——
        ///           规划不应偷偷新建干涉集, 只用用户已建好的。
        ///   false → 原行为: 未命中则新建。用于 UI"自动创建"按钮。
        /// </summary>
        private void BuildPair(IEnumerable<ITxObject> checkObjects, IEnumerable<ITxObject> obstacles,
            TxRobot robot = null, bool attachOnly = false)
        {
            try
            {
                var checkList = new TxObjectList();
                foreach (var o in checkObjects) { checkList.Add(o); _checkObjects.Add(o); }
                var obstacleList = new TxObjectList();
                foreach (var o in obstacles) obstacleList.Add(o);

                CheckObjectCount = checkList.Count;
                ObstacleObjectCount = obstacleList.Count;
                _checkList = checkList;       // FromLists 精确查询用
                _obstacleList = obstacleList;

                _log(string.Format("  干涉集: 检测方 {0} 项, 障碍方 {1} 项",
                    checkList.Count, obstacleList.Count));

                if (checkList.Count == 0 || obstacleList.Count == 0)
                {
                    _log("  [警告] 干涉集为空，跳过创建");
                    return;
                }

                // ---- 已有干涉集检测 (v5.0 强类型: root.Pairs → pair.FirstList) ----
                TxCollisionPair existingPair = FindExistingPairForRobot(robot, _log);

                if (existingPair != null)
                {
                    _pair = existingPair;
                    _ownedPair = false;
                    _log(string.Format("  干涉集复用: 已有碰撞对 '{0}', 不再新建",
                        TryGetName(existingPair)));
                }
                else if (attachOnly)
                {
                    // 纯附着模式: 找不到就放弃, 绝不新建 (避免规划偷偷堆积干涉集)
                    _log("  [附着] 未找到现有干涉集 — 规划不新建。请先点\"自动创建干涉集\"");
                    _pair = null;
                    _ownedPair = false;
                    return;
                }
                else
                {
                    // 已固化API (inspect_type 确认): 无参构造 + FirstList/SecondList 属性。
                    var data = new TxCollisionPairCreationData();
                    data.FirstList = checkList;
                    data.SecondList = obstacleList;

                    TxCollisionRoot root = TxApplication.ActiveDocument.CollisionRoot;
                    _pair = root.CreateCollisionPair(data);
                    _ownedPair = true;
                    _log("  干涉集创建成功");

                    // 激活 (v5.0: TxCollisionPair.Active 强类型属性, 已确认存在)
                    try
                    {
                        _pair.Active = true;
                        _log("  干涉集激活: pair.Active = true");
                    }
                    catch (Exception ex) { _log("  [警告] 激活失败: " + ex.Message); }

                    // 全局碰撞检查开关 —— 若关闭, 所有 pair 都不参与计算
                    try
                    {
                        if (!root.CheckCollisions)
                        {
                            root.CheckCollisions = true;
                            _log("  [提示] CollisionRoot.CheckCollisions 原为 false, 已开启");
                        }
                    }
                    catch { }
                }

                // ---- 基线快照 (仅诊断参考): 记录规划前姿态的接触状态 ----
                // 基线签名不含位置信息, 不能用于碰撞豁免 (会导致机器人移动后
                // 同对象对的碰撞被错误豁免)。仅用于日志对照, 帮助理解碰撞来源。
                DetermineQueryMode();
                if (_queryMode == 1)
                {
                    var baseSigs = QueryOurCollisionSignatures();
                    if (baseSigs != null)
                    {
                        _baseline = baseSigs;
                        _log(string.Format("  基线快照: {0} 项接触 (仅诊断参考, 不豁免)", _baseline.Count));
                        int shown = 0;
                        foreach (var s in _baseline)
                        {
                            if (shown++ >= 5) { _log("    ... 等"); break; }
                            _log("    常驻: " + s);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _log("  [警告] 干涉集创建/复用失败: " + ex.Message);
                _pair = null;
                _ownedPair = false;
            }
        }

        // ================================================================
        //  查询 / 清理
        // ================================================================

        /// <summary>
        /// 重拍基线快照 — 在改变了设备形态(如焊枪张开)后调用,
        /// 更新诊断参考基线 (仅用于日志对照, 不再用于碰撞豁免)。
        /// </summary>
        public void RecaptureBaseline()
        {
            if (_queryMode != 1) return;
            var sigs = QueryOurCollisionSignatures();
            if (sigs != null)
            {
                _baseline = sigs;
                _log(string.Format("  基线重拍: {0} 项接触 (仅诊断参考, 不豁免)", _baseline.Count));
            }
        }

        /// <summary>
        /// 当前场景状态下，本干涉集是否存在与检测方相关的碰撞。
        ///
        /// 查询模式 (定型于首次调用):
        ///   1 = 结果集直接判定 (主): GetCollidingObjects → 只统计涉及机器人/焊钳子树的
        ///       碰撞状态。不再做基线扣除 —— 基线签名不含位置信息, 机器人移动后
        ///       同对象对的碰撞产生相同签名被错误豁免 (根因: 全段恒报"直连通过")。
        ///       永久接触 (7轴/线缆包) 通过排除机器人子树对象于障碍方来处理,
        ///       而非基线扣除。
        ///   2 = 布尔 HasCollidingObjects(params)  — 无法过滤, 可能被场景既有碰撞污染
        ///   3 = 布尔 属性/无参                    — 同上
        ///  -1 = 全部失败 (已dump成员清单)
        /// </summary>
        public bool QueryColliding()
        {
            if (_pair == null && _fromLists == -1)
                return false; // 碰撞对与 FromLists 均不可用 → 退化为纯可达性
            try
            {
                if (_queryMode == 0) DetermineQueryMode();

                switch (_queryMode)
                {
                    case 1:
                    {
                        var sigs = QueryOurCollisionSignatures();
                        if (sigs == null) return false;
                        bool colliding = sigs.Count > 0;
                        if (_sampleLogBudget > 0)
                        {
                            _sampleLogBudget--;
                            int baselineMatch = 0;
                            foreach (var s in sigs) if (_baseline.Contains(s)) baselineMatch++;
                            _log(string.Format("    [干涉样本] 相关碰撞 {0} 项 (基线内 {1} / 新增 {2}) → {3}",
                                sigs.Count, baselineMatch, sigs.Count - baselineMatch,
                                colliding ? "碰撞" : "安全"));
                        }
                        return colliding;
                    }
                    case 2:
                    {
                        dynamic root2 = TxApplication.ActiveDocument.CollisionRoot;
                        return (bool)root2.HasCollidingObjects((dynamic)_queryParams);
                    }
                    case 3:
                    {
                        dynamic root3 = TxApplication.ActiveDocument.CollisionRoot;
                        try { return (bool)root3.HasCollidingObjects; }
                        catch { return (bool)root3.HasCollidingObjects(); }
                    }
                    default:
                        return false;
                }
            }
            catch (Exception ex)
            {
                if (_sampleLogBudget > 0)
                {
                    _sampleLogBudget--;
                    _log("    [干涉查询异常] " + ex.GetType().Name + ": " + ex.Message);
                }
                return false;
            }
        }

        // ---- 查询定型状态 ----
        private int _queryMode;
        private object _queryParams;
        private int _sampleLogBudget = 6;
        private bool _stateDumped;
        private HashSet<string> _baseline = new HashSet<string>();

        private void DetermineQueryMode()
        {
            if (_queryMode != 0) return;
            _queryParams = TryBuildQueryParams();

            // ---- 模式1 (主): 结果集过滤 ----
            var probe = QueryOurCollisionSignatures();
            if (probe != null)
            {
                _queryMode = 1;
                _log("  干涉查询定型: 模式1 结果集过滤 (仅统计机器人/焊钳相关, 基线扣除)");
                return;
            }

            // ---- 模式2/3: 布尔 (无过滤, 有污染风险) ----
            try
            {
                dynamic root = TxApplication.ActiveDocument.CollisionRoot;
                if (_queryParams != null)
                {
                    try
                    {
                        bool b = (bool)root.HasCollidingObjects((dynamic)_queryParams);
                        _queryMode = 2;
                        _log("  干涉查询定型: 模式2 布尔(params) [警告:无法过滤场景既有碰撞] (首值=" + b + ")");
                        return;
                    }
                    catch { }
                }
                try
                {
                    bool b2 = (bool)root.HasCollidingObjects;
                    _queryMode = 3;
                    _log("  干涉查询定型: 模式3 布尔属性 [警告:无法过滤] (首值=" + b2 + ")");
                    return;
                }
                catch { }
                try
                {
                    bool b3 = (bool)root.HasCollidingObjects();
                    _queryMode = 3;
                    _log("  干涉查询定型: 模式3 布尔无参 [警告:无法过滤] (首值=" + b3 + ")");
                    return;
                }
                catch { }
            }
            catch { }

            _queryMode = -1;
            _log("  [需确认] 干涉查询全部方式失败 — 成员清单:");
            DumpCollisionMembers();
        }

        /// <summary>
        /// 结果集查询: 返回"涉及检测方"的碰撞签名集合; API不可用返回null。
        /// 签名 = 参与对象名排序拼接, 用于与基线比对。
        /// </summary>
        private HashSet<string> QueryOurCollisionSignatures()
        {
            try
            {
                dynamic root = TxApplication.ActiveDocument.CollisionRoot;
                if (_queryParams == null) _queryParams = TryBuildQueryParams();

                object results = null;

                // 首选 (API参考表固化): GetCollidingObjectsFromLists — 精确双列表查询,
                // 免疫场景 Collision Viewer 既有碰撞对的污染
                if (_fromLists >= 0 && _checkList != null && _obstacleList != null)
                {
                    try
                    {
                        results = (object)root.GetCollidingObjectsFromLists(
                            (dynamic)_checkList, (dynamic)_obstacleList, (dynamic)_queryParams);
                        if (_fromLists == 0)
                        {
                            _fromLists = 1;
                            _log("  干涉查询升级: GetCollidingObjectsFromLists (精确双列表)");
                        }
                    }
                    catch
                    {
                        if (_fromLists == 0)
                        {
                            _fromLists = -1;
                            _log("  [信息] FromLists 不可用, 回退碰撞对查询");
                        }
                        results = null;
                    }
                }

                if (results == null)
                {
                    results = _queryParams != null
                        ? (object)root.GetCollidingObjects((dynamic)_queryParams)
                        : (object)root.GetCollidingObjects();
                }
                var sigs = new HashSet<string>();
                if (results == null) return sigs;

                IEnumerable states = null;
                try { states = ((dynamic)results).States as IEnumerable; } catch { }
                if (states == null) states = results as IEnumerable;
                if (states == null)
                {
                    if (!_stateDumped)
                    {
                        _stateDumped = true;
                        _log("    [诊断] 结果对象不可枚举: " + results.GetType().FullName);
                        DumpMembersOf(results.GetType(), "results");
                    }
                    return null;
                }

                foreach (object state in states)
                {
                    // Type 过滤 (API参考表固化: Unchecked=0/Separate=1/NearMiss=2/Contact=3/Collision=5)
                    // 结果集可能含 Separate/NearMiss 项, 只有 Type>=3 才算碰撞
                    int typeVal = ReadStateType(state);
                    if (typeVal >= 0 && typeVal < 3) continue;

                    var objs = ExtractStateObjects(state);
                    if (objs.Count == 0) continue;
                    if (!InvolvesCheckSet(objs)) continue;
                    sigs.Add(Signature(objs));
                }
                return sigs;
            }
            catch (Exception ex)
            {
                if (!_stateDumped)
                {
                    _stateDumped = true;
                    _log("    [诊断] 结果集查询不可用: " + ex.GetType().Name + ": " + ex.Message);
                }
                return null;
            }
        }

        /// <summary>读取碰撞状态 Type 枚举值; 不可读返回 -1 (按碰撞处理, 保守)</summary>
        private static int ReadStateType(object state)
        {
            try { return Convert.ToInt32(((dynamic)state).Type); }
            catch { return -1; }
        }

        /// <summary>
        /// 描述当前姿态下与检测方相关的碰撞 (用于动态违例消息, 指明涉事对象)。
        /// 不再扣除基线 —— 与 QueryColliding() 保持一致。
        /// </summary>
        public string DescribeFreshCollisions(int maxItems = 2)
        {
            try
            {
                var sigs = QueryOurCollisionSignatures();
                if (sigs == null) return "";
                var items = new List<string>();
                foreach (var s in sigs)
                {
                    items.Add(s);
                    if (items.Count >= maxItems) break;
                }
                return items.Count > 0 ? string.Join("; ", items) : "";
            }
            catch { return ""; }
        }

        /// <summary>提取碰撞状态涉及的对象: CollidingObjects 集合优先, 反射兜底</summary>
        private List<ITxObject> ExtractStateObjects(object state)
        {
            var result = new List<ITxObject>();
            try
            {
                IEnumerable objs = null;
                try { objs = ((dynamic)state).CollidingObjects as IEnumerable; } catch { }
                if (objs == null) { try { objs = ((dynamic)state).Objects as IEnumerable; } catch { } }
                if (objs != null)
                {
                    foreach (var o in objs)
                    {
                        var t = o as ITxObject;
                        if (t != null) result.Add(t);
                    }
                    if (result.Count > 0) return result;
                }

                // 反射兜底: 收集状态对象上所有 ITxObject 型属性值
                foreach (var prop in state.GetType().GetProperties())
                {
                    if (!typeof(ITxObject).IsAssignableFrom(prop.PropertyType)) continue;
                    try
                    {
                        var v = prop.GetValue(state, null) as ITxObject;
                        if (v != null) result.Add(v);
                    }
                    catch { }
                }

                if (result.Count == 0 && !_stateDumped)
                {
                    _stateDumped = true;
                    _log("    [诊断] 无法从碰撞状态提取对象 — 状态成员清单:");
                    DumpMembersOf(state.GetType(), "state");
                }
            }
            catch { }
            return result;
        }

        /// <summary>碰撞是否涉及检测方 (机器人/焊钳本体或其子树)</summary>
        private bool InvolvesCheckSet(List<ITxObject> objs)
        {
            foreach (var o in objs)
            {
                foreach (var c in _checkObjects)
                {
                    if (ReferenceEquals(o, c) || IsDescendantOf(o, c))
                        return true;
                }
            }
            return false;
        }

        private static string Signature(List<ITxObject> objs)
        {
            var names = new List<string>();
            foreach (var o in objs)
            {
                try { names.Add(o.Name); }
                catch { names.Add(o.GetType().Name); }
            }
            names.Sort();
            return string.Join("|", names);
        }

        /// <summary>
        /// 构造 TxCollisionQueryParams (反射定位类型)，并尝试把查询范围
        /// 限定为本插件的 pair — 上次运行未找到范围属性, 定型时会dump params成员。
        /// </summary>
        private object TryBuildQueryParams()
        {
            try
            {
                var asm = typeof(TxCollisionPairCreationData).Assembly;
                var pType = asm.GetType("Tecnomatix.Engineering.TxCollisionQueryParams");
                if (pType == null) return null;
                object p = Activator.CreateInstance(pType);

                try
                {
                    var listType = asm.GetType("Tecnomatix.Engineering.TxCollisionPairList");
                    if (listType != null && _pair != null)
                    {
                        object list = Activator.CreateInstance(listType);
                        var add = listType.GetMethod("Add");
                        if (add != null) add.Invoke(list, new[] { _pair });
                        foreach (var prop in pType.GetProperties())
                        {
                            if (prop.CanWrite && prop.PropertyType.IsAssignableFrom(listType))
                            {
                                prop.SetValue(p, list, null);
                                _log("  干涉查询范围限定: params." + prop.Name + " = 本插件pair");
                                break;
                            }
                        }
                    }
                }
                catch { }

                try
                {
                    var stopProp = pType.GetProperty("StopQueryAfterFirstCollision");
                    if (stopProp != null && stopProp.CanWrite)
                        stopProp.SetValue(p, false, null); // 基线扣除需要完整结果集
                }
                catch { }

                if (!_paramsDumped)
                {
                    _paramsDumped = true;
                    _log("  [信息] TxCollisionQueryParams 可写属性: ");
                    foreach (var prop in pType.GetProperties())
                        if (prop.CanWrite)
                            _log("    params.prop " + prop.PropertyType.Name + " " + prop.Name);
                }
                return p;
            }
            catch { return null; }
        }
        private bool _paramsDumped;

        private void DumpCollisionMembers()
        {
            try
            {
                object rootObj = TxApplication.ActiveDocument.CollisionRoot;
                DumpMembersOf(rootObj.GetType(), "root", "Coll");
                if (_pair != null)
                {
                    _log("    pair 类型: " + _pair.GetType().FullName);
                    DumpMembersOf(_pair.GetType(), "pair");
                }
            }
            catch (Exception ex)
            {
                _log("    dump 失败: " + ex.Message);
            }
        }

        private void DumpMembersOf(Type t, string tag, string nameFilter = null)
        {
            try
            {
                foreach (var m in t.GetMethods())
                {
                    if (m.IsSpecialName) continue;
                    if (nameFilter != null &&
                        m.Name.IndexOf(nameFilter, StringComparison.OrdinalIgnoreCase) < 0) continue;
                    var ps = m.GetParameters()
                        .Select(x => x.ParameterType.Name + " " + x.Name);
                    _log("    " + tag + "." + m.Name + "(" + string.Join(", ", ps) + ") : "
                        + m.ReturnType.Name);
                }
                foreach (var p in t.GetProperties())
                {
                    if (nameFilter != null &&
                        p.Name.IndexOf(nameFilter, StringComparison.OrdinalIgnoreCase) < 0) continue;
                    _log("    " + tag + ".prop " + p.PropertyType.Name + " " + p.Name);
                }
            }
            catch { }
        }

        public void Dispose()
        {
            if (_ownedPair && !KeepPairOnDispose)
            {
                try
                {
                    if (_pair != null) _pair.Delete();   // v5.0: 强类型
                }
                catch { }
            }
            // 复用的碰撞对不删除 — 它属于场景原有配置
            // KeepPairOnDispose=true 时, 本次创建的也不删 (持久干涉集)
            _pair = null;
        }
    }
}
