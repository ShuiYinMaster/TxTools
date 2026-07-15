// SpotGrouper.cs — C# 7.3
// 读取层（纯读，不写 PS）：
//   1. 在选中节点下枚举所有「焊点位操作」并反查 TxWeldPoint
//   2. 读每个焊点绑定的零件名（AssignedParts 级联 + 反射兜底）
//   3. 按「多重集指纹」分组
//
// 全部 SDK 调用必须在 PS 主线程；调用方负责线程。

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Tecnomatix.Engineering;

namespace TxTools.WeldSpotGrouper
{
    public static class SpotGrouper
    {
        private static void Nop(string s) { }

        // ────────────────────────────────────────────────────────────
        // 1. 枚举 + 分组（对外主入口）
        // ────────────────────────────────────────────────────────────
        public static List<SpotGroup> ScanAndGroup(ITxObject scopeNode, GroupOptions opt,
                                                    GroupReport rep, Action<string> log)
        {
            log = log ?? Nop;
            opt = opt ?? new GroupOptions();
            rep = rep ?? new GroupReport();

            var spots = CollectSpots(scopeNode, log);
            rep.ScannedSpots = spots.Count;
            log("[扫描] 焊点位共 " + spots.Count + " 个");

            // 读零件 + 算指纹
            foreach (var s in spots)
            {
                ReadPartNames(s, log);
                s.Signature = BuildSignature(s.PartNames, opt.IgnoreCase);
                if (s.HasBinding) rep.BoundSpots++;
            }

            // 分组
            var byKey = new Dictionary<string, SpotGroup>(StringComparer.Ordinal);
            foreach (var s in spots)
            {
                if (!s.HasBinding)
                {
                    if (opt.SkipUnbound) { rep.SkippedUnbound++; continue; }
                }
                SpotGroup g;
                if (!byKey.TryGetValue(s.Signature, out g))
                {
                    g = new SpotGroup { Signature = s.Signature };
                    // 展示用零件名：排序去“顺序差异”，但保留重复以体现数量
                    g.SamplePartNames.AddRange(SortedWithDups(s.PartNames, opt.IgnoreCase));
                    byKey[s.Signature] = g;
                }
                g.Spots.Add(s);
            }

            var groups = byKey.Values
                .OrderByDescending(g => g.Spots.Count)
                .ThenBy(g => g.PartLabel, StringComparer.Ordinal)
                .ToList();

            log("[分组] 得到 " + groups.Count + " 组");
            foreach (var g in groups)
                log(string.Format("[分组]   [{0}] {1} 个焊点 ← {2}", groups.IndexOf(g) + 1, g.Spots.Count, g.PartLabel));
            return groups;
        }

        // ────────────────────────────────────────────────────────────
        // 2. 在选中节点下枚举焊点位操作（并反查 TxWeldPoint）
        // ────────────────────────────────────────────────────────────
        private static List<SpotItem> CollectSpots(ITxObject scopeNode, Action<string> log)
        {
            var list = new List<SpotItem>();
            var seenLoc = new HashSet<ITxObject>();
            var seenWp = new HashSet<TxWeldPoint>();

            if (scopeNode == null) { log("[扫描] 范围节点为空"); return list; }

            // 取所有后代（操作树/物理树通吃）。ITxObjectCollection.GetAllDescendants 必须给非空过滤器。
            IEnumerable<ITxObject> descendants = EnumDescendants(scopeNode, log);

            foreach (var obj in descendants)
            {
                if (obj == null) continue;

                // 情况 A：后代本身就是焊点特征
                var asWp = obj as TxWeldPoint;
                if (asWp != null)
                {
                    if (seenWp.Add(asWp))
                    {
                        var locOp = FirstWeldLocationOp(asWp);
                        if (locOp == null || seenLoc.Add(locOp))
                            list.Add(MakeSpot(locOp ?? obj, asWp));
                    }
                    continue;
                }

                // 情况 B：后代是焊点位操作（weld location operation）→ 反查它的焊点
                var wp = WpFromLocOp(obj);
                if (wp != null && seenLoc.Add(obj))
                {
                    seenWp.Add(wp);
                    list.Add(MakeSpot(obj, wp));
                }
            }

            // 兜底：节点本身就是单个焊点
            if (list.Count == 0)
            {
                var asWp = scopeNode as TxWeldPoint;
                if (asWp != null) list.Add(MakeSpot(FirstWeldLocationOp(asWp) ?? scopeNode, asWp));
            }
            return list;
        }

        private static SpotItem MakeSpot(ITxObject locOp, TxWeldPoint wp)
        {
            var s = new SpotItem { LocOp = locOp, Wp = wp };
            s.Name = SafeName(wp) ?? SafeName(locOp) ?? "(未命名)";
            return s;
        }

        private static IEnumerable<ITxObject> EnumDescendants(ITxObject node, Action<string> log)
        {
            // 主路：ITxObjectCollection.GetAllDescendants(TxTypeFilter)
            try
            {
                var coll = node as ITxObjectCollection;
                if (coll != null)
                {
                    var filter = new TxTypeFilter(typeof(ITxObject));
                    TxObjectList all = coll.GetAllDescendants(filter);
                    if (all != null)
                    {
                        var outl = new List<ITxObject>(all.Count);
                        foreach (ITxObject o in all) outl.Add(o);
                        return outl;
                    }
                }
            }
            catch (Exception ex) { log("[扫描] GetAllDescendants 失败：" + ex.Message); }

            // 兜底：dynamic 递归子节点
            var acc = new List<ITxObject>();
            RecurseKids(node, acc, new HashSet<ITxObject>(), 0, log);
            return acc;
        }

        private static void RecurseKids(ITxObject node, List<ITxObject> acc, HashSet<ITxObject> seen, int depth, Action<string> log)
        {
            if (node == null || depth > 40 || !seen.Add(node)) return;
            TxObjectList kids = null;
            foreach (var p in new[] { "Children", "Operations", "DescendantOperations" })
            {
                try
                {
                    var pi = node.GetType().GetProperty(p, BindingFlags.Public | BindingFlags.Instance);
                    if (pi != null) kids = pi.GetValue(node) as TxObjectList;
                }
                catch { kids = null; }
                if (kids != null && kids.Count > 0) break;
            }
            if (kids == null) return;
            foreach (ITxObject k in kids) { acc.Add(k); RecurseKids(k, acc, seen, depth + 1, log); }
        }

        // 从「焊点位操作」反查 TxWeldPoint：d.MfgFeature / d.Feature / d.WeldPoint（与 PsReader.GetWpFromPm 同源）
        internal static TxWeldPoint WpFromLocOp(ITxObject locOp)
        {
            foreach (var prop in new[] { "WeldPoint", "MfgFeature", "Feature", "OperationFeature" })
            {
                try
                {
                    object v = TryGet(locOp, prop);
                    var wp = v as TxWeldPoint;
                    if (wp != null) return wp;
                }
                catch { }
            }
            return null;
        }

        private static ITxObject FirstWeldLocationOp(TxWeldPoint wp)
        {
            try
            {
                TxObjectList ops = wp.WeldLocationOperations;
                if (ops != null && ops.Count > 0)
                    foreach (ITxObject o in ops) return o;
            }
            catch { }
            return null;
        }

        // ────────────────────────────────────────────────────────────
        // 3. 读绑定零件名（焊点的 AssignedParts 级联 + 反射兜底；只要名字）
        // ────────────────────────────────────────────────────────────
        private static void ReadPartNames(SpotItem s, Action<string> log)
        {
            if (s == null || s.Wp == null) return;
            var seen = new HashSet<string>(StringComparer.Ordinal);

            string[] candidates = { "AssignedParts", "WeldedParts", "Parts", "WeldedComponents", "AssignedComponents" };
            foreach (var prop in candidates)
            {
                object raw = TryGet(s.Wp, prop);
                if (raw == null) continue;
                CollectNames(raw, s.PartNames, seen);
                if (s.PartNames.Count > 0) break; // 命中即止
            }

            // 反射兜底：找名字含 Part/Component 且可枚举的属性
            if (s.PartNames.Count == 0)
            {
                try
                {
                    foreach (var pi in s.Wp.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
                    {
                        string pn = pi.Name;
                        if (pn.IndexOf("Part", StringComparison.OrdinalIgnoreCase) < 0 &&
                            pn.IndexOf("Component", StringComparison.OrdinalIgnoreCase) < 0) continue;
                        object v; try { v = pi.GetValue(s.Wp); } catch { continue; }
                        if (v != null) CollectNames(v, s.PartNames, seen);
                        if (s.PartNames.Count > 0) break;
                    }
                }
                catch (Exception ex) { log("[零件] 反射兜底异常 (" + s.Name + ")：" + ex.Message); }
            }
        }

        private static void CollectNames(object raw, List<string> outNames, HashSet<string> seen)
        {
            // raw 可能是单个对象，也可能是 TxObjectList / IEnumerable
            var en = raw as System.Collections.IEnumerable;
            if (en != null && !(raw is string))
            {
                foreach (var o in en) AddName(o, outNames, seen);
            }
            else AddName(raw, outNames, seen);
        }

        private static void AddName(object o, List<string> outNames, HashSet<string> seen)
        {
            if (o == null) return;
            var tx = o as ITxObject;
            string nm = tx != null ? SafeName(tx) : null;
            if (string.IsNullOrEmpty(nm)) { try { dynamic d = o; nm = (string)d.Name; } catch { } }
            if (string.IsNullOrEmpty(nm)) return;
            // 同名零件重复出现要保留（多重集语义），但同一引用不重复 —— 用 "name#序号" 防误去重
            // 这里按「名字」收集；真正的重数由出现次数体现，所以不去重相同名字。
            outNames.Add(nm);
        }

        // ────────────────────────────────────────────────────────────
        // 指纹：排序 + 保留重复 → 多重集相等
        // ────────────────────────────────────────────────────────────
        public static string BuildSignature(List<string> partNames, bool ignoreCase)
        {
            if (partNames == null || partNames.Count == 0) return "";
            var sorted = SortedWithDups(partNames, ignoreCase);
            return string.Join("\u0001", sorted);
        }

        private static List<string> SortedWithDups(List<string> names, bool ignoreCase)
        {
            var cmp = ignoreCase ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
            var copy = new List<string>(names);
            copy.Sort(cmp);
            return copy;
        }

        // ────────────────────────────────────────────────────────────
        // 工具
        // ────────────────────────────────────────────────────────────
        internal static object TryGet(object target, string prop)
        {
            if (target == null) return null;
            try
            {
                var pi = target.GetType().GetProperty(prop, BindingFlags.Public | BindingFlags.Instance);
                if (pi != null) return pi.GetValue(target);
            }
            catch { }
            return null;
        }

        internal static string SafeName(ITxObject o)
        {
            if (o == null) return null;
            try { return o.Name; } catch { return null; }
        }
    }
}
