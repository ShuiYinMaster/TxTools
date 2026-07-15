// GroupWriter.cs — C# 7.3
// 写入层：OperationRoot 根下先建一个父复合操作（汇总容器），每组的焊接操作建在该复合操作
// 之下，再把组内焊点位操作移入对应焊接操作。
//
// 原因：实测 OperationRoot 根下只暴露 CreateCompoundOperation(TxCompoundOperationCreationData)，
// 不能直接建焊接操作；而焊接操作（ITxWeldOperationCreation.CreateWeldOperation）建在复合操作下。
//
// 创建器解析两次：
//   ① 在 opRoot 上解析「建复合操作」→ 建父容器
//   ② 在父容器上解析「建焊接操作」→ 每组建一条；若父容器不暴露建焊接，则降级为在其下建复合子操作
// 解析法：先按名字打分选方法，再用所选方法「自己声明的参数类型」构造 CreationData，杜绝错配。

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Tecnomatix.Engineering;

namespace TxTools.WeldSpotGrouper
{
    public static class GroupWriter
    {
        private static void Nop(string s) { }

        public static GroupReport Apply(List<SpotGroup> groups, GroupOptions opt, GroupReport rep, Action<string> log)
        {
            log = log ?? Nop;
            opt = opt ?? new GroupOptions();
            rep = rep ?? new GroupReport();

            var doc = TxApplication.ActiveDocument;
            if (doc == null) { Fail(rep, "ActiveDocument 为空"); return rep; }
            ITxObject opRoot = doc.OperationRoot as ITxObject;
            if (opRoot == null) { Fail(rep, "OperationRoot 为空"); return rep; }

            // ① 根下「建复合操作」的创建器
            var makeCompound = ResolveCreator(opRoot, ScoreCompound, "[建-父]", log);
            if (makeCompound == null) { Fail(rep, "OperationRoot 根下未找到建复合操作的方法（见日志）"); return rep; }

            object um = OpenUndo("焊点自动分组", log);
            try
            {
                // 建父复合操作（一个汇总容器，承载本次所有分组焊接操作）
                string parentName = (opt.NamePrefix ?? "") + "汇总_" + DateTime.Now.ToString("HHmmss");
                ITxObject parent = null;
                try { parent = makeCompound(opRoot, parentName, log); }
                catch (Exception ex) { Fail(rep, "建父复合操作异常：" + ex.Message); }
                if (parent == null) { Fail(rep, "建父复合操作失败"); AbortUndo(um, log); return rep; }
                log("[写入] 父复合操作 ✓ " + parentName);

                // ② 在父复合操作下「建焊接操作」的创建器
                var makeWeld = ResolveCreator(parent, ScoreWeld, "[建-焊]", log);
                if (makeWeld == null)
                {
                    log("[建-焊] 父容器下未找到建焊接操作的方法，降级为在其下建复合子操作");
                    makeWeld = ResolveCreator(parent, ScoreCompound, "[建-焊降级]", log);
                }
                if (makeWeld == null) { Fail(rep, "父容器下无可用创建方法"); AbortUndo(um, log); return rep; }

                int idx = 1;
                var usedNames = new HashSet<string>(StringComparer.Ordinal);
                foreach (var g in groups)
                {
                    string name = MakeOpName(opt.NamePrefix, idx, g, usedNames);
                    g.TargetOpName = name;

                    ITxObject newOp = null;
                    try { newOp = makeWeld(parent, name, log); }
                    catch (Exception ex) { Fail(rep, "[" + name + "] 建操作异常：" + ex.Message); }
                    if (newOp == null) { Fail(rep, "[" + name + "] 建操作返回 null"); idx++; continue; }

                    g.CreatedOp = newOp;
                    rep.GroupsCreated++;
                    log("[写入] 焊接操作 ✓ " + name + " ← " + g.PartLabel);

                    ITxObject after = null;
                    foreach (var s in g.Spots)
                    {
                        if (s.LocOp == null) { Fail(rep, "[" + s.Name + "] 无可移动的焊点位操作"); continue; }
                        if (MoveLocationInto(newOp, s.LocOp, after, log)) { rep.SpotsMoved++; after = s.LocOp; }
                        else Fail(rep, "[" + s.Name + "] 移动失败");
                    }
                    idx++;
                }
                CommitUndo(um, log);
            }
            catch (Exception ex) { AbortUndo(um, log); Fail(rep, "事务异常：" + ex.Message); }

            log("[写入] " + rep);
            return rep;
        }

        // ════════════════════════════════════════════════════════════
        // 通用创建器解析：在 parent 的接口/自身方法里找「单参 Create*(XxxCreationData)」，
        // 按 scorer 打分选最优，再用所选方法实际参数类型构造数据。
        // ════════════════════════════════════════════════════════════
        private static Func<ITxObject, string, Action<string>, ITxObject> ResolveCreator(
            ITxObject parent, Func<MethodInfo, int> scorer, string tag, Action<string> log)
        {
            var methods = new List<MethodInfo>();
            foreach (var iface in parent.GetType().GetInterfaces())
            {
                try { methods.AddRange(iface.GetMethods()); } catch { }
            }
            try { methods.AddRange(parent.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance)); } catch { }

            var creators = methods
                .Where(m => m.GetParameters().Length == 1
                         && m.Name.IndexOf("Create", StringComparison.OrdinalIgnoreCase) >= 0
                         && m.Name.IndexOf("Can", StringComparison.Ordinal) < 0
                         && m.GetParameters()[0].ParameterType.Name.IndexOf("CreationData", StringComparison.Ordinal) >= 0)
                .GroupBy(m => m.Name + "|" + m.GetParameters()[0].ParameterType.Name)
                .Select(g => g.First())
                .ToList();

            log(tag + " 候选 " + creators.Count + " 个：");
            foreach (var m in creators) log(tag + "   · " + m.Name + "(" + m.GetParameters()[0].ParameterType.Name + ")");

            MethodInfo best = creators.OrderByDescending(scorer).FirstOrDefault();
            if (best == null || scorer(best) <= -200) { log(tag + " 无合适创建方法"); return null; }

            Type pType = best.GetParameters()[0].ParameterType;
            log(tag + " 选用：" + best.Name + "(" + pType.Name + ")");
            var chosen = best;
            return (root, name, lg) =>
            {
                object data = BuildCreationData(pType, name, lg);
                if (data == null) { lg(tag + " 构造 " + pType.Name + " 失败"); return null; }
                return chosen.Invoke(root, new[] { data }) as ITxObject;
            };
        }

        // 偏好焊接操作；焊点位(Location)强排斥；复合操作垫底
        private static int ScoreWeld(MethodInfo m)
        {
            string n = m.Name, p = m.GetParameters()[0].ParameterType.Name;
            int s = 0;
            if (Has(n, "Weld") || Has(p, "Weld")) s += 100;
            if (Has(n, "Location") || Has(p, "Location")) s -= 300;
            if (Has(n, "Compound")) s -= 20;
            if (Has(n, "Continuous") || Has(n, "Generic")) s -= 10;
            if (Has(n, "Operation")) s += 5;
            return s;
        }

        // 偏好复合操作；焊点位强排斥；焊接操作不当父容器
        private static int ScoreCompound(MethodInfo m)
        {
            string n = m.Name, p = m.GetParameters()[0].ParameterType.Name;
            int s = 0;
            if (Has(n, "Compound") || Has(p, "Compound")) s += 100;
            if (Has(n, "Location") || Has(p, "Location")) s -= 300;
            if (Has(n, "Weld")) s -= 50;
            if (Has(n, "Operation")) s += 5;
            return s;
        }

        private static bool Has(string s, string token) => s.IndexOf(token, StringComparison.Ordinal) >= 0;

        private static object BuildCreationData(Type dataType, string name, Action<string> log)
        {
            if (dataType == null) return null;
            try { var c = dataType.GetConstructor(new[] { typeof(string) }); if (c != null) return c.Invoke(new object[] { name }); }
            catch (Exception ex) { log("[建] data ctor(string)×：" + ex.Message); }
            try
            {
                var c0 = dataType.GetConstructor(Type.EmptyTypes);
                if (c0 != null)
                {
                    object d = c0.Invoke(null);
                    var ni = dataType.GetProperty("Name", BindingFlags.Public | BindingFlags.Instance);
                    if (ni != null && ni.CanWrite) ni.SetValue(d, name);
                    return d;
                }
            }
            catch (Exception ex) { log("[建] data ctor()×：" + ex.Message); }
            log("[建] 无法构造 " + dataType.Name);
            return null;
        }

        // ════════════════════════════════════════════════════════════
        // 移动一条 location 进目标操作（验证过的级联）
        // ════════════════════════════════════════════════════════════
        private static bool MoveLocationInto(ITxObject target, ITxObject loc, ITxObject after, Action<string> log)
        {
            if (after != null)
            {
                try { dynamic d = target; d.AddObjectAfter(loc, after); return true; } catch (Exception e) { log("[移动] AddObjectAfter×：" + e.Message); }
                try { dynamic d = target; d.MoveObjectAfter(loc, after); return true; } catch (Exception e) { log("[移动] MoveObjectAfter×：" + e.Message); }
            }
            try { dynamic d = target; d.AddObject(loc); try { if (after != null) d.MoveChildAfter(loc, after); } catch { } return true; } catch (Exception e) { log("[移动] AddObject×：" + e.Message); }
            try { dynamic d = target; var l = new TxObjectList(); l.Add(loc); d.AddObjects(l); return true; } catch (Exception e) { log("[移动] AddObjects×：" + e.Message); }
            return false;
        }

        // ════════════════════════════════════════════════════════════
        // Undo（TxApplication.ActiveUndoManager + OpenUndoTransaction）
        // ════════════════════════════════════════════════════════════
        private static object OpenUndo(string name, Action<string> log)
        {
            try
            {
                dynamic um = TxApplication.ActiveUndoManager;
                if (um == null) return null;
                try { um.OpenUndoTransaction(name); return um; } catch { }
                try { um.OpenTransaction(name); return um; } catch { }
                try { um.StartTransaction(name); return um; } catch { }
                try { um.BeginUndoTransaction(name); return um; } catch { }
            }
            catch { }
            log("[Undo] 未开启事务（PS 仍可 Ctrl+Z）");
            return null;
        }
        private static void CommitUndo(object um, Action<string> log)
        {
            if (um == null) return;
            try { dynamic d = um; try { d.CommitUndoTransaction(); return; } catch { } try { d.CommitTransaction(); return; } catch { } try { d.Commit(); return; } catch { } } catch { }
        }
        private static void AbortUndo(object um, Action<string> log)
        {
            if (um == null) return;
            try { dynamic d = um; try { d.AbortUndoTransaction(); return; } catch { } try { d.AbortTransaction(); return; } catch { } try { d.Rollback(); return; } catch { } } catch { }
            log("[Undo] 已尝试回滚");
        }

        // ════════════════════════════════════════════════════════════
        private static string MakeOpName(string prefix, int idx, SpotGroup g, HashSet<string> used)
        {
            string label = g.SamplePartNames.Count > 0 ? string.Join("+", g.SamplePartNames) : "无绑定";
            if (label.Length > 48) label = label.Substring(0, 48) + "…";
            string name = (prefix ?? "") + idx.ToString("D2") + "_" + label;
            string final = name; int k = 1;
            while (used.Contains(final)) final = name + "(" + (++k) + ")";
            used.Add(final);
            return final;
        }

        private static void Fail(GroupReport rep, string msg) { rep.Failed++; rep.Errors.Add(msg); }
    }
}