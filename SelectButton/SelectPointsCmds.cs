// SelectPointsCmds.cs  —  C# 7.3   (v3: 用确认的 TxSelection 真实 API)
// TxTools 插件：三个一键选点按钮
//   - SelectAllWeldPointsCmd       全选所选操作下的焊点（TxWeldPoint）
//   - SelectAllPathPointsCmd       全选所选操作下的过渡点（Via / Arc / 非焊点非连续点的 Location-like）
//   - SelectAllContinuousPointsCmd 全选所选操作下的连续点（Continuous*）
//
// 工作流程：
//   1. 用户在 PS 中先选中一个或多个操作（焊接操作 / 复合操作 / Compound 等）
//   2. 点击对应按钮
//   3. 插件在所选操作的子树中递归查找匹配类型的点
//   4. 用 TxSelection.SetItems(...) 把 ActiveSelection 替换为这些点
//
// v3 变更（vs v2）:
//   ① 改用确认的官方 API：TxSelection.SetItems(TxObjectList)
//       一步到位，不再 Clear() + AddItems()，不再反射
//   ② 无选中时不再扫 OperationRoot（避免一次选中全场上千个点）
//       改为提示用户"请先选择一个操作"
//   ③ Walk 去重改用 ObjectId（dynamic 取，PsReader 同款），避免 HashCode 碰撞
//   ④ 日志仍保留反射式探测（TxSelection 文档不涉及日志 API）

using System;
using System.Collections.Generic;
using System.Reflection;
using Tecnomatix.Engineering;

namespace MyPlugin.SelectPoints
{
    // ════════════════════════════════════════════════════════════
    //  三个按钮命令
    // ════════════════════════════════════════════════════════════

    public class SelectAllWeldPointsCmd : TxButtonCommand
    {
        public override string Name => "SelectAllWeldPointsCmd";
        public override string Category => "TxTools";
        public override void Execute(object cmdParams)
        {
            SelectPointsCore.Run(PointKind.Weld, "焊点");
        }
    }

    public class SelectAllPathPointsCmd : TxButtonCommand
    {
        public override string Name => "SelectAllPathPointsCmd";
        public override string Category => "TxTools";
        public override void Execute(object cmdParams)
        {
            SelectPointsCore.Run(PointKind.Path, "过渡点");
        }
    }

    public class SelectAllContinuousPointsCmd : TxButtonCommand
    {
        public override string Name => "SelectAllContinuousPointsCmd";
        public override string Category => "TxTools";
        public override void Execute(object cmdParams)
        {
            SelectPointsCore.Run(PointKind.Continuous, "连续点");
        }
    }

    // ════════════════════════════════════════════════════════════
    //  共享实现
    // ════════════════════════════════════════════════════════════

    public enum PointKind { Weld, Path, Continuous }

    public static class SelectPointsCore
    {
        private const string LOG_PREFIX = "[SelectPts] ";

        public static void Run(PointKind kind, string label)
        {
            try
            {
                // 1) 取所选操作作为扫描根
                TxObjectList sel;
                try { sel = TxApplication.ActiveSelection.GetItems(); }
                catch (Exception ex) { Log("读取选中失败: " + ex.Message); return; }

                if (sel == null || sel.Count == 0)
                {
                    Log("未选中任何对象。请先在路径编辑器/对象树中选中一个操作，再点击按钮。");
                    return;
                }

                // 2) 在所选操作子树中递归收集
                var hits = new List<ITxObject>();
                var seen = new HashSet<object>();
                int rootCount = 0;
                foreach (ITxObject r in sel)
                {
                    if (r == null) continue;
                    rootCount++;
                    Walk(r, kind, hits, seen, 0);
                }

                if (hits.Count == 0)
                {
                    Log(string.Format("在所选 {0} 个对象子树中未找到{1}。", rootCount, label));
                    // 无命中时不动当前选中
                    return;
                }

                // 3) 用 SetItems 一步覆盖式设置选中
                bool ok = ApplySelection(hits);

                Log(string.Format("{0}：选中 {1} 个{2}（扫描根 {3} 个）",
                    ok ? "完成" : "尝试", hits.Count, label, rootCount));
            }
            catch (Exception ex)
            {
                Log("顶层异常: " + ex.Message);
            }
        }

        // ── 递归遍历 ──────────────────────────────────────────────

        private static void Walk(ITxObject node, PointKind want,
                                 List<ITxObject> hits, HashSet<object> seen, int depth)
        {
            if (node == null || depth > 64) return;

            // 去重：优先用 ObjectId（PS 内置唯一 id），失败回退到 RuntimeHelpers.GetHashCode
            object id = TryGetObjectId(node);
            if (!seen.Add(id)) return;

            // 命中判断
            if (Matches(node, want))
                hits.Add(node);

            // 子节点
            TxObjectList kids = GetKids(node);
            if (kids == null || kids.Count == 0) return;
            foreach (ITxObject c in kids)
                Walk(c, want, hits, seen, depth + 1);
        }

        private static object TryGetObjectId(ITxObject node)
        {
            try
            {
                dynamic d = node;
                object id = d.ObjectId;
                if (id != null) return id;
            }
            catch { }
            return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(node);
        }

        private static TxObjectList GetKids(ITxObject node)
        {
            // 优先 GetAllChildren（如可用），其次反射 Children 属性
            try
            {
                dynamic d = node;
                object raw = null;
                try { raw = d.GetAllChildren(); } catch { }
                if (raw == null) try { raw = d.Children; } catch { }
                return raw as TxObjectList;
            }
            catch { return null; }
        }

        // ── 类型匹配（与 PsReader.KindOf / IsLocNode 一致口径） ────

        private static bool Matches(ITxObject node, PointKind want)
        {
            // 焊点：TxWeldPoint 是最权威标志
            if (node is TxWeldPoint)
                return want == PointKind.Weld;

            string tn = node.GetType().Name;

            // 只处理 Location-like 节点，避免把操作/容器误选
            bool isLocLike = tn.Contains("Location") || tn.Contains("Via")
                          || tn.Contains("Arc") || tn.Contains("PathPoint")
                          || tn.Contains("WeldLoc") || tn.Contains("Continuous");
            if (!isLocLike) return false;

            // 三类互斥判定
            bool isWeld = tn.Contains("Weld") || tn.Contains("Seam");
            bool isCont = tn.Contains("Continuous");
            bool isPath = !isWeld && !isCont;  // 过渡点 = Location-like 且非焊点非连续点

            switch (want)
            {
                case PointKind.Weld: return isWeld;
                case PointKind.Continuous: return isCont;
                case PointKind.Path: return isPath;
                default: return false;
            }
        }

        // ════════════════════════════════════════════════════════════
        //  写回 ActiveSelection — 使用 TxSelection.SetItems（官方 API）
        // ════════════════════════════════════════════════════════════

        private static bool ApplySelection(List<ITxObject> hits)
        {
            try
            {
                if (hits == null || hits.Count == 0)
                {
                    TxApplication.ActiveSelection.Clear();
                    return true;
                }

                var list = new TxObjectList();
                foreach (var h in hits) list.Add(h);

                // SetItems 是覆盖式：内部会先清除原选中，再设置为给定集合
                TxApplication.ActiveSelection.SetItems(list);
                return true;
            }
            catch (Exception ex)
            {
                Log("SetItems 失败: " + ex.Message);
                return false;
            }
        }

        // ════════════════════════════════════════════════════════════
        //  日志（反射式：尽量输出到 PS，无可用出口时静默到 Debug）
        // ════════════════════════════════════════════════════════════

        private static MethodInfo _logMethod;
        private static bool _logProbed;

        private static void Log(string msg)
        {
            string full = LOG_PREFIX + msg;

            if (!_logProbed)
            {
                _logProbed = true;
                ProbeLogMethod();
            }

            if (_logMethod != null)
            {
                try { _logMethod.Invoke(null, new object[] { full }); return; }
                catch { /* 出错就走 fallback */ }
            }

            try { System.Diagnostics.Debug.WriteLine(full); } catch { }
        }

        private static void ProbeLogMethod()
        {
            string[] candidates = { "WriteToOutput", "WriteMessage", "PrintMessage", "ReportInfo", "Log" };
            try
            {
                Type appT = typeof(TxApplication);
                foreach (var name in candidates)
                {
                    var m = appT.GetMethod(name,
                        BindingFlags.Public | BindingFlags.Static,
                        null, new[] { typeof(string) }, null);
                    if (m != null) { _logMethod = m; return; }
                }
            }
            catch { }
        }
    }
}