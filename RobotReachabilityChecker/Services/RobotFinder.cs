// ============================================================================
// RobotFinder.cs
//
// 三套接口：
//   FindAssociatedRobot(op, doc, log)        — 6 种策略找操作关联的机器人
//   FindRobotByName(doc, name, log)          — 按名查找单台机器人
//   FindOperationByName(doc, name)           — 递归查找操作，优先 ITxRoboticOperation
//
// 同名机器人诊断：枚举 PhysicalRoot 下所有 TxRobot，按名分组，>1 个发警告
// （避免 .Robot 属性指向错误的副本）
//
// 操作递归查找：当 TxCompoundOperation 和子 TxWeldOperation 同名时，优先返回
// 子 TxWeldOperation（有 Robot 属性），而不是 Compound。
// ============================================================================
using System;
using System.Collections.Generic;
using Tecnomatix.Engineering;
using TxTools.RobotReachabilityChecker.Diagnostics;
using TxTools.RobotReachabilityChecker.Ui;

namespace TxTools.RobotReachabilityChecker.Services
{
    public static class RobotFinder
    {
        // =====================================================================
        // 找操作关联的机器人 — 6 种策略
        // =====================================================================
        public static TxRobot FindAssociatedRobot(ITxObject operation, TxDocument doc, ILogger log = null)
        {
            log = log ?? NullLogger.Instance;
            if (operation == null) return null;

            // 诊断：同名机器人警告
            try
            {
                if (doc != null)
                {
                    var allRobots = doc.PhysicalRoot.GetAllDescendants(new TxTypeFilter(typeof(TxRobot)));
                    var byName = new Dictionary<string, int>();
                    foreach (ITxObject ro in allRobots)
                    {
                        if (ro is TxRobot r)
                        {
                            string n = r.Name ?? "";
                            byName[n] = byName.ContainsKey(n) ? byName[n] + 1 : 1;
                        }
                    }
                    foreach (var kv in byName)
                        if (kv.Value > 1)
                            log.Log($"  [诊断] ⚠ 场景中存在 {kv.Value} 台同名机器人 '{kv.Key}'，可能影响 .Robot 解析", "WARN");
                }
            }
            catch (Exception ex) { log.Log($"  [诊断] 机器人枚举异常: {ex.Message}", "WARN"); }

            // 方式1：.Robot 属性
            try
            {
                dynamic dop = operation;
                var r = dop.Robot as TxRobot;
                if (r != null) { log.Log($"  关联机器人(.Robot): '{r.Name}' (HashCode={r.GetHashCode()})"); return r; }
            }
            catch { }

            // 方式2：.Device 属性
            try { dynamic dop = operation; var r = dop.Device as TxRobot; if (r != null) { log.Log($"  关联机器人(.Device): {r.Name}", "DEBUG"); return r; } } catch { }

            // 方式3：.RobotDevice 属性
            try { dynamic dop = operation; var r = dop.RobotDevice as TxRobot; if (r != null) { log.Log($"  关联机器人(.RobotDevice): {r.Name}", "DEBUG"); return r; } } catch { }

            // 方式4：Parent 链上溯找 TxRobot
            try
            {
                dynamic cur = operation;
                for (int depth = 0; depth < 10; depth++)
                {
                    object parent = null;
                    try { parent = cur.Parent; } catch { break; }
                    if (parent == null) break;
                    if (parent is TxRobot rp) { log.Log($"  关联机器人(Parent链 depth={depth}): {rp.Name}", "DEBUG"); return rp; }
                    cur = parent;
                }
            }
            catch { }

            // 方式5：ParentOperation.Robot
            try
            {
                dynamic dop = operation;
                object parentCompound = dop.ParentOperation;
                if (parentCompound != null)
                {
                    dynamic dc = parentCompound;
                    var r = dc.Robot as TxRobot;
                    if (r != null) { log.Log($"  关联机器人(ParentOperation.Robot): {r.Name}"); return r; }
                }
            }
            catch { }

            // 方式6：复合操作 → 遍历子操作的 .Robot
            if (operation is TxCompoundOperation compound)
            {
                try
                {
                    var children = compound.GetDirectDescendants(new TxTypeFilter(typeof(ITxObject)));
                    if (children != null)
                    {
                        foreach (ITxObject child in children)
                        {
                            if (child == null) continue;
                            try
                            {
                                dynamic dc = child;
                                var r = dc.Robot as TxRobot;
                                if (r != null) { log.Log($"  关联机器人(子操作.Robot [{child.Name}]): {r.Name}"); return r; }
                            }
                            catch { }
                        }
                    }
                }
                catch { }
            }

            log.Log("  所有方式均未能找到关联机器人", "WARN");
            return null;
        }

        /// <summary>FindAssociatedRobot 的静默版本（不打"未找到"警告日志）—— 用于 UI 加载阶段。</summary>
        public static TxRobot FindAssociatedRobotSilent(ITxObject operation)
        {
            if (operation == null) return null;
            try { dynamic dop = operation; var r = dop.Robot as TxRobot; if (r != null) return r; } catch { }
            try { dynamic dop = operation; var r = dop.Device as TxRobot; if (r != null) return r; } catch { }
            try { dynamic dop = operation; var r = dop.RobotDevice as TxRobot; if (r != null) return r; } catch { }
            try
            {
                dynamic cur = operation;
                for (int depth = 0; depth < 10; depth++)
                {
                    object parent = null;
                    try { parent = cur.Parent; } catch { break; }
                    if (parent == null) break;
                    if (parent is TxRobot rp) return rp;
                    cur = parent;
                }
            }
            catch { }
            try
            {
                dynamic dop = operation;
                object parentCompound = dop.ParentOperation;
                if (parentCompound != null)
                {
                    dynamic dc = parentCompound;
                    var r = dc.Robot as TxRobot;
                    if (r != null) return r;
                }
            }
            catch { }
            if (operation is TxCompoundOperation compound)
            {
                try
                {
                    var children = compound.GetDirectDescendants(new TxTypeFilter(typeof(ITxObject)));
                    if (children != null)
                    {
                        foreach (ITxObject child in children)
                        {
                            if (child == null) continue;
                            try { dynamic dc = child; var r = dc.Robot as TxRobot; if (r != null) return r; } catch { }
                        }
                    }
                }
                catch { }
            }
            return null;
        }

        // =====================================================================
        // 按名查找机器人（用于 Grid_AfterSelChange 单击行驱动姿态）
        // =====================================================================
        public static TxRobot FindRobotByName(TxDocument doc, string robotName)
        {
            if (doc == null || string.IsNullOrEmpty(robotName)) return null;
            try
            {
                var all = doc.PhysicalRoot.GetAllDescendants(new TxTypeFilter(typeof(TxRobot)));
                foreach (ITxObject o in all)
                    if (o is TxRobot r && r.Name == robotName) return r;
            }
            catch { }
            return null;
        }

        // =====================================================================
        // 按名递归找操作 — 优先返回 ITxRoboticOperation 而非同名的 TxCompoundOperation
        // =====================================================================
        public static ITxObject FindOperationByName(TxDocument doc, string name)
        {
            if (doc == null || string.IsNullOrEmpty(name)) return null;
            var kids = doc.OperationRoot.GetDirectDescendants(new TxTypeFilter(typeof(ITxObject)));
            return FindOpRecursive(kids, name, 0);
        }

        private static ITxObject FindOpRecursive(TxObjectList nodes, string name, int depth)
        {
            if (nodes == null || depth > 20) return null;
            ITxObject compoundFallback = null;

            foreach (ITxObject obj in nodes)
            {
                if (obj == null) continue;
                bool isOp = obj is ITxRoboticOperation || obj is TxCompoundOperation
                         || obj is ITxOperation || obj.GetType().Name.Contains("Operation");
                if (!isOp) continue;

                if (obj.Name == name)
                {
                    if (obj is ITxRoboticOperation) return obj;
                    if (obj is TxCompoundOperation co)
                    {
                        try
                        {
                            var sub = co.GetDirectDescendants(new TxTypeFilter(typeof(ITxObject)));
                            if (sub != null)
                                foreach (ITxObject child in sub)
                                    if (child is ITxRoboticOperation && child.Name == name)
                                        return child;
                        }
                        catch { }
                        compoundFallback = obj;
                    }
                    else return obj;
                }

                TxObjectList sub2 = null;
                if (obj is TxCompoundOperation co2)
                    try { sub2 = co2.GetDirectDescendants(new TxTypeFilter(typeof(ITxObject))); } catch { }
                var found = FindOpRecursive(sub2, name, depth + 1);
                if (found != null) return found;
            }
            return compoundFallback;
        }

        internal static TxRobot FindAssociatedRobot(ITxObject pickedObj, ReachabilityCheckerForm reachabilityCheckerForm)
        {
            throw new NotImplementedException();
        }
    }
}
