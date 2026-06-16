// ============================================================================
// InterferenceService.cs
//
// 静态干涉检测服务 — 全部基于 PS 真实 API（来自反射探针 v2 的实测）：
//
//   · TxDocument.CollisionRoot                            → TxCollisionRoot
//   · CollisionRoot.PairList                              → TxObjectList of TxCollisionPair
//   · TxCollisionPair.FirstList / SecondList              → TxObjectList
//   · CollisionRoot.HasCollidingObjects(params)           → bool   ★ 核心查询
//   · CollisionRoot.CreateCollisionPair(creationData)     → TxCollisionPair  ★ 现场创建
//   · TxCollisionPairCreationData(name, firstList, secondList)
//
// 业务约定：
//   · 判定"机器人是否有干涉对"：挂载工具出现在任一 Pair 的 FirstList/SecondList 即视为已有
//     （不看 Active；用户临时禁用的对也算"已有"）
//   · 未发现 → 现场调用 CreateCollisionPair，构造"机器人 × 2m 内资源"的 Pair
//   · 查询时调用 HasCollidingObjects(params)，PS 自己处理 Active 逻辑
// ============================================================================
using System;
using System.Collections;
using System.Collections.Generic;
using Tecnomatix.Engineering;
using TxTools.RobotReachabilityChecker.Diagnostics;

namespace TxTools.RobotReachabilityChecker.Services
{
    public static class InterferenceService
    {
        private const double NEARBY_RADIUS_MM = 2000.0;
        private const string AUTO_PAIR_NAME_PREFIX = "Auto_Reachability_";

        // =====================================================================
        // 准备阶段：确保机器人有可用的干涉对
        //   返回 true → 后续可以走 HasCollidingObjects 查询
        //   返回 false → 准备失败（无 CollisionRoot / 无挂载工具 / 自动创建失败）
        // =====================================================================
        public static bool EnsureRobotHasCollisionPair(
            TxRobot robot, TxDocument doc, ILogger log = null)
        {
            log = log ?? NullLogger.Instance;
            if (robot == null || doc == null) return false;

            TxCollisionRoot root = null;
            try { root = doc.CollisionRoot; } catch { }
            if (root == null) { log.Log("  CollisionRoot 不可用，跳过干涉检查", "WARN"); return false; }

            // 取机器人挂载的工具
            TxObjectList mounted = null;
            try { mounted = robot.MountedTools; } catch { }
            if (mounted == null || mounted.Count == 0)
            {
                log.Log("  机器人未挂载工具，跳过干涉检查", "WARN");
                return false;
            }

            // ── 1. 扫现有 Pair：挂工具出现在任一 First/Second 列表 ──
            var toolSet = new HashSet<string>();
            for (int i = 0; i < mounted.Count; i++)
            {
                ITxObject t = mounted[i];
                if (t != null && !string.IsNullOrEmpty(t.Name)) toolSet.Add(t.Name);
            }

            try
            {
                TxObjectList pairs = root.PairList;
                int pairCount = pairs?.Count ?? 0;
                for (int i = 0; i < pairCount; i++)
                {
                    TxCollisionPair pair = pairs[i] as TxCollisionPair;
                    if (pair == null) continue;
                    if (ListContainsAnyByName(pair.FirstList, toolSet) ||
                        ListContainsAnyByName(pair.SecondList, toolSet))
                    {
                        log.Log($"  ✓ 已发现干涉对：'{pair.Name}'（包含机器人工具）", "OK");
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                log.Log($"  扫描现有干涉对异常: {ex.Message}", "WARN");
            }

            // ── 2. 未发现 → 现场创建"机器人×2m内资源"的 Pair ──
            log.Log("  未发现归属干涉对，自动创建…");
            try
            {
                TxObjectList firstList = new TxObjectList();
                firstList.Add(robot);

                TxObjectList secondList = CollectNearbyObjects(robot, doc, log);
                if (secondList == null || secondList.Count == 0)
                {
                    log.Log($"  2m 内未找到其他资源，跳过创建", "WARN");
                    return false;
                }

                string pairName = AUTO_PAIR_NAME_PREFIX + robot.Name + "_" +
                                  DateTime.Now.ToString("HHmmss");
                var cd = new TxCollisionPairCreationData(pairName, firstList, secondList);
                TxCollisionPair created = root.CreateCollisionPair(cd);
                if (created == null)
                {
                    log.Log("  CreateCollisionPair 返回 null", "ERR");
                    return false;
                }

                log.Log($"  ✓ 已自动创建干涉对 '{pairName}'，B 侧资源 {secondList.Count} 个", "OK");
                log.Log("    （可在 PS 干涉编辑器里手动调整成员）", "OK");
                return true;
            }
            catch (Exception ex)
            {
                log.Log($"  自动创建干涉对失败: {ex.GetType().Name} - {ex.Message}", "ERR");
                return false;
            }
        }

        // =====================================================================
        // 静态查询当前姿态是否发生碰撞
        //   PS 内部按 Active 标志过滤 Pair（被用户禁用的 Pair 不查）
        // =====================================================================
        public static bool CheckCollisionAtCurrentPose(
            TxDocument doc, ILogger log = null)
        {
            log = log ?? NullLogger.Instance;
            if (doc == null) return false;

            TxCollisionRoot root = null;
            try { root = doc.CollisionRoot; } catch { }
            if (root == null) return false;

            try
            {
                var queryParams = new TxCollisionQueryParams();
                queryParams.StopQueryAfterFirstCollision = true;  // 性能优化：见到一个就返回
                return root.HasCollidingObjects(queryParams);
            }
            catch (Exception ex)
            {
                log.Log($"  HasCollidingObjects 异常: {ex.Message}", "WARN");
                return false;
            }
        }

        // =====================================================================
        // 辅助：判断 TxObjectList 中是否包含 toolSet 任一同名对象
        // =====================================================================
        private static bool ListContainsAnyByName(TxObjectList list, HashSet<string> nameSet)
        {
            if (list == null || nameSet == null || nameSet.Count == 0) return false;
            int n = list.Count;
            for (int i = 0; i < n; i++)
            {
                ITxObject obj = list[i];
                if (obj == null || string.IsNullOrEmpty(obj.Name)) continue;
                if (nameSet.Contains(obj.Name)) return true;
            }
            return false;
        }

        // =====================================================================
        // 收集机器人 2m 球形范围内的资源对象（用于自动创建 Pair 的 B 侧）
        //
        // 实现：遍历 PhysicalRoot 下 ITxLocatableObject 的直系子节点（避免重复深入
        // 抓机器人/工具内部零件），按 AbsoluteLocation.Translation 距机器人位置筛选。
        //
        // 不排除工件、夹具等 — 用户后续在 PS 干涉编辑器里手动调整。
        // =====================================================================
        private static TxObjectList CollectNearbyObjects(
            TxRobot robot, TxDocument doc, ILogger log)
        {
            var result = new TxObjectList();
            if (robot == null || doc == null) return result;

            TxVector robotPos = null;
            try
            {
                ITxLocatableObject lobj = robot as ITxLocatableObject;
                if (lobj != null && lobj.AbsoluteLocation != null)
                {
                    var tx = lobj.AbsoluteLocation;
                    // TxTransformation 上的 Translation 属性是 TxVector
                    try { robotPos = tx.Translation; } catch { }
                }
            }
            catch { }

            if (robotPos == null)
            {
                log.Log("  无法读取机器人位置，跳过 2m 过滤，包含整个 PhysicalRoot 直系", "WARN");
            }

            // 排除机器人自身和它的挂载工具（这些应当在 FirstList 里）
            var excludeNames = new HashSet<string>();
            if (!string.IsNullOrEmpty(robot.Name)) excludeNames.Add(robot.Name);
            try
            {
                TxObjectList mt = robot.MountedTools;
                if (mt != null)
                {
                    for (int i = 0; i < mt.Count; i++)
                    {
                        var t = mt[i];
                        if (t != null && !string.IsNullOrEmpty(t.Name)) excludeNames.Add(t.Name);
                    }
                }
            }
            catch { }

            try
            {
                // 取 PhysicalRoot 直系子节点（不深入零件），同时挑出可定位的设备
                var direct = doc.PhysicalRoot.GetDirectDescendants(
                    new TxTypeFilter(typeof(ITxLocatableObject)));
                if (direct == null) return result;

                for (int i = 0; i < direct.Count; i++)
                {
                    ITxObject obj = direct[i];
                    if (obj == null) continue;
                    if (excludeNames.Contains(obj.Name)) continue;

                    // 距离过滤
                    if (robotPos != null)
                    {
                        TxVector pos = null;
                        try
                        {
                            ITxLocatableObject lobj = obj as ITxLocatableObject;
                            if (lobj != null && lobj.AbsoluteLocation != null)
                                pos = lobj.AbsoluteLocation.Translation;
                        }
                        catch { }
                        if (pos == null) continue;

                        double dx = pos.X - robotPos.X;
                        double dy = pos.Y - robotPos.Y;
                        double dz = pos.Z - robotPos.Z;
                        double dist = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                        if (dist > NEARBY_RADIUS_MM) continue;
                    }

                    result.Add(obj);
                }
            }
            catch (Exception ex)
            {
                log.Log($"  枚举 PhysicalRoot 异常: {ex.Message}", "WARN");
            }
            return result;
        }
    }
}
