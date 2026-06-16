// ============================================================================
// InterferenceProbe.cs  (第二轮 — 深度反射)
//
// 第一轮发现：
//   · TxApplication 上没有 CollisionSetsToActivate 属性
//   · TxCollisionQueryContext 没有任何公共属性/方法（疑似只是个空容器）
//   · TxCollisionQueryResults 只有一个 ArrayList States
//   · 不知道怎么"执行"查询
//
// 第二轮目标：
//   A. 全程序集扫描所有名字含 Collision 的类型，列出其公开 API
//   B. 全程序集扫描所有 Tx* 类型，找有 Collision* 返回类型的属性/方法 → 入口
//   C. 尝试构造 Context 和 Params，调用候选执行方法，dump States 元素结构
// ============================================================================
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Tecnomatix.Engineering;
using TxTools.RobotReachabilityChecker.Services;

namespace TxTools.RobotReachabilityChecker.Diagnostics
{
    public static class InterferenceProbe
    {
        public static void Run(ITxObject pickedOperation, ILogger log)
        {
            log = log ?? NullLogger.Instance;
            log.Log("════════════════════════════════════════", "OK");
            log.Log("[Interference Probe v2] 深度反射 PS 碰撞 API", "OK");

            var asm = typeof(TxApplication).Assembly;

            // ──────────────────────────────────────────────────────────────
            // A. 全程序集扫描：所有名字含 "Collision" 的类型
            // ──────────────────────────────────────────────────────────────
            log.Log("");
            log.Log("─── A. 全程序集中名字含 Collision 的类型 ───");
            Type[] collTypes = null;
            try
            {
                collTypes = asm.GetTypes()
                    .Where(t => t.Name.IndexOf("Collision", StringComparison.OrdinalIgnoreCase) >= 0)
                    .OrderBy(t => t.Name).ToArray();
                foreach (var t in collTypes)
                {
                    string kind = t.IsEnum ? "enum " : t.IsClass ? "class" : t.IsInterface ? "intf " : "stuct";
                    log.Log($"  [{kind}] {t.FullName}");
                }
                log.Log($"  共 {collTypes.Length} 个类型");
            }
            catch (Exception ex) { log.Log($"  A 异常: {ex.Message}", "WARN"); collTypes = new Type[0]; }

            // ──────────────────────────────────────────────────────────────
            // B. 列出每个非枚举碰撞类型的完整 API（构造/属性/方法）
            //    重点关注它们如何关联到一起
            // ──────────────────────────────────────────────────────────────
            log.Log("");
            log.Log("─── B. 各碰撞类型的完整 API ───");
            foreach (var t in collTypes ?? new Type[0])
            {
                if (t.IsEnum) continue;
                log.Log("");
                log.Log($"  ── {t.Name} ──");
                try
                {
                    foreach (var c in t.GetConstructors(BindingFlags.Public | BindingFlags.Instance))
                    {
                        var ps = string.Join(",", c.GetParameters().Select(x => x.ParameterType.Name + " " + x.Name));
                        log.Log($"    [Ctor] {t.Name}({ps})");
                    }
                    foreach (var p in t.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static))
                    {
                        string sm = p.GetGetMethod()?.IsStatic == true ? "(static)" : "";
                        log.Log($"    [Prop{sm}] {p.PropertyType.Name}  {p.Name}");
                    }
                    foreach (var m in t.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly))
                    {
                        if (m.IsSpecialName) continue;
                        if (m.GetParameters().Length > 4) continue;
                        var ps = string.Join(",", m.GetParameters().Select(x => x.ParameterType.Name + " " + x.Name));
                        string sm = m.IsStatic ? "(static)" : "";
                        log.Log($"    [Mthd{sm}] {m.ReturnType.Name}  {m.Name}({ps})");
                    }
                }
                catch (Exception ex) { log.Log($"    异常: {ex.Message}", "WARN"); }
            }

            // ──────────────────────────────────────────────────────────────
            // C. 全程序集扫描：哪些 Tx* 类型上有"返回 Collision* 类型"的属性/方法
            //    这才是真正的入口
            // ──────────────────────────────────────────────────────────────
            log.Log("");
            log.Log("─── C. 引用 Collision 类型的入口 (返回类型 / 参数类型) ───");
            try
            {
                var collNames = new HashSet<string>(
                    (collTypes ?? new Type[0]).Select(t => t.Name));

                int found = 0;
                foreach (var t in asm.GetTypes())
                {
                    if (!t.IsPublic) continue;
                    if (t.Name.IndexOf("Collision", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                    // 静态属性
                    foreach (var p in t.GetProperties(BindingFlags.Static | BindingFlags.Public | BindingFlags.Instance))
                    {
                        if (collNames.Contains(p.PropertyType.Name))
                        {
                            string sm = p.GetGetMethod()?.IsStatic == true ? "static " : "";
                            log.Log($"  [Prop] {sm}{t.Name}.{p.Name} : {p.PropertyType.Name}");
                            found++;
                        }
                    }
                    // 方法
                    foreach (var m in t.GetMethods(BindingFlags.Static | BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
                    {
                        if (m.IsSpecialName) continue;
                        bool retHit = collNames.Contains(m.ReturnType.Name);
                        bool paramHit = m.GetParameters().Any(pp => collNames.Contains(pp.ParameterType.Name));
                        if (retHit || paramHit)
                        {
                            var ps = string.Join(",", m.GetParameters().Select(x => x.ParameterType.Name + " " + x.Name));
                            string sm = m.IsStatic ? "static " : "";
                            log.Log($"  [Mthd] {sm}{t.Name}.{m.Name}({ps}) : {m.ReturnType.Name}");
                            found++;
                        }
                    }
                    if (found > 80) { log.Log("  ...(超 80 条结果，截断)"); break; }
                }
                log.Log($"  共 {found} 条候选入口");
            }
            catch (Exception ex) { log.Log($"  C 异常: {ex.Message}", "WARN"); }

            // ──────────────────────────────────────────────────────────────
            // D. 尝试空查询：构造 Context + Params + Results，看 States 元素长啥样
            // ──────────────────────────────────────────────────────────────
            log.Log("");
            log.Log("─── D. 尝试构造查询对象，dump States 元素结构 ───");
            try
            {
                Type tCtx = collTypes.FirstOrDefault(t => t.Name == "TxCollisionQueryContext");
                Type tPar = collTypes.FirstOrDefault(t => t.Name == "TxCollisionQueryParams");
                Type tRes = collTypes.FirstOrDefault(t => t.Name == "TxCollisionQueryResults");
                if (tCtx == null || tPar == null || tRes == null)
                {
                    log.Log("  缺少 Context/Params/Results 类型,跳过", "WARN");
                }
                else
                {
                    object ctx = Activator.CreateInstance(tCtx);
                    object par = Activator.CreateInstance(tPar);
                    log.Log($"  ✓ 已构造 Context = {ctx?.GetType().Name}, Params = {par?.GetType().Name}");

                    // 看 Results 怎么构造
                    var resCtor0 = tRes.GetConstructors(BindingFlags.Public | BindingFlags.Instance)
                        .FirstOrDefault(c => c.GetParameters().Length == 0);
                    object res = null;
                    if (resCtor0 != null)
                    {
                        res = Activator.CreateInstance(tRes);
                        log.Log($"  ✓ 已构造 Results = {res?.GetType().Name}");
                    }
                    else
                    {
                        log.Log("  Results 没有无参构造，需要从查询方法返回", "WARN");
                    }

                    // dump Results.States 静态结构（元素类型）
                    var statesProp = tRes.GetProperty("States");
                    if (statesProp != null)
                    {
                        log.Log($"  States 属性类型 = {statesProp.PropertyType.FullName}");
                        // ArrayList 没有泛型参数，但运行时元素类型可能是某个 Tx*State
                        // 试着看程序集里有没有名字带 CollisionState 的类型
                        var stateTypes = asm.GetTypes()
                            .Where(t => t.Name.IndexOf("CollisionState", StringComparison.OrdinalIgnoreCase) >= 0
                                     || t.Name.IndexOf("CollisionPair", StringComparison.OrdinalIgnoreCase) >= 0
                                     || t.Name.IndexOf("CollisionResult", StringComparison.OrdinalIgnoreCase) >= 0)
                            .ToArray();
                        log.Log($"  推测的 State 元素候选类型 ({stateTypes.Length} 个):");
                        foreach (var st in stateTypes)
                        {
                            log.Log($"    · {st.FullName}");
                            foreach (var p in st.GetProperties(BindingFlags.Public | BindingFlags.Instance))
                            {
                                if (p.IsSpecialName) continue;
                                log.Log($"        [Prop] {p.PropertyType.Name}  {p.Name}");
                            }
                            foreach (var m in st.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
                            {
                                if (m.IsSpecialName) continue;
                                if (m.GetParameters().Length > 2) continue;
                                var ps = string.Join(",", m.GetParameters().Select(x => x.ParameterType.Name + " " + x.Name));
                                log.Log($"        [Mthd] {m.ReturnType.Name}  {m.Name}({ps})");
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { log.Log($"  D 异常: {ex.Message}", "WARN"); }

            // ──────────────────────────────────────────────────────────────
            // E. 当前机器人 MountedTools — 留作后续对比用
            // ──────────────────────────────────────────────────────────────
            log.Log("");
            log.Log("─── E. 当前机器人 MountedTools ───");
            try
            {
                if (pickedOperation == null) log.Log("  未提供操作", "WARN");
                else
                {
                    TxRobot robot = RobotFinder.FindAssociatedRobotSilent(pickedOperation);
                    if (robot != null)
                    {
                        log.Log($"  机器人: {robot.Name}");
                        var mounted = robot.MountedTools;
                        int n = mounted?.Count ?? 0;
                        log.Log($"  MountedTools.Count = {n}");
                        for (int i = 0; i < n; i++)
                        {
                            ITxObject t = mounted[i];
                            log.Log($"    [{i}] {t.Name}  ({t.GetType().Name})");
                        }
                    }
                }
            }
            catch (Exception ex) { log.Log($"  E 异常: {ex.Message}", "WARN"); }

            log.Log("");
            log.Log("[Interference Probe v2] 完成", "OK");
            log.Log("════════════════════════════════════════", "OK");
        }
    }
}