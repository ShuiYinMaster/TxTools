// SelectPointsCmds.cs  —  C# 7.3   (v6: GetAllDescendants + 强类型点位提取)
// TxTools 插件：三个一键选点按钮
//   - SelectAllWeldPointsCmd       全选焊点（TxWeldLocationOperation 操作 + TxWeldPoint 点位）
//   - SelectAllPathPointsCmd       全选过渡点（Via/Home/Arc 等非焊非连续的 ITxLocationOperation）
//   - SelectAllContinuousPointsCmd 全选连续点（Continuous/Seam 容器 → 展开仅取 SeamLocation 子操作）
//
// 工作流程：
//   1. 用户在 PS 中选中操作（操作树 / 路径编辑器）或点位对象（3D 视口）
//   2. 点击对应按钮
//   3. 插件遍历选中对象的操作子树（用 GetAllDescendants(TxNoTypeFilter)）
//   4. 从后代操作提取点位对象：
//      - 焊点：选中 TxWeldLocationOperation 操作本身（ITxDisplayableObject，可选中）
//              + 尝试附带 TxWeldPoint（ITxWeldPoint 不可直接选中，但如果 PS 允许则额外高亮）
//      - 过渡点：ITxLocationOperation 且非焊非连续 → 选操作本身
//      - 连续点：Continuous/Seam 容器 → 不选中容器本身，展开仅取 SeamLocation 子操作
//                （如 TxRoboticSeamLocationOperation, TxSeamLocationOperation 等）
//   5. 用 TxSelection.SetItems 设置选区
//
// v5.1 变更（vs v5）:
//   ① 连续点选取：不再选中 Continuous/Seam 容器本身
//      → 改为展开容器，选中其内部子操作（下一层级）
//      → Continuous/Seam 容器是 ITxObjectCollection，可调用 GetAllDescendants
//      → 子操作类型多样（SeamLocation, Location 等），全部选中
//   ② IsTargetPoint(Continuous) 不再匹配 Continuous/Seam 容器
//      → 由 CollectTargets 特殊处理逻辑接管
//   ③ 新增 DrillContinuousContainer() 方法：展开容器仅取 SeamLocation 子操作
//
// v5 变更（vs v4）:
//   ① 废弃 GetEnumerator 手动递归 → 改用 GetAllDescendants(TxNoTypeFilter)
//      （ITxObjectCollection 的官方遍历方法，更可靠）
//   ② 废弃 SimulatedObjects 策略（返回模拟对象如机器人/焊枪，不含点位）
//   ③ 焊点提取改用强类型 TxWeldLocationOperation.WeldPoint（已确认 API）
//      + dynamic WeldPoint 兜底（TxRoboticSeamLocationOperation 等子类）
//   ④ 焊点目标 = 操作本身（TxWeldLocationOperation，ITxDisplayableObject，可选中）
//      过渡点/连续点目标 = 操作本身（无独立点位对象）
//      ※ ITxWeldPoint : ITxObject 不继承 ITxDisplayableObject，SetItems 无法选中
//   ⑤ 新增 GetPlanningItems() 兜底取操作选中

using C1.Util.DX.Direct2D;
using System;
using System.Collections.Generic;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.ModelObjects;

namespace TxTools.SelectPoints
{
    // ════════════════════════════════════════════════════════════
    //  三个按钮命令
    // ════════════════════════════════════════════════════════════

    public class SelectAllWeldPointsCmd : TxButtonCommand
    {
        public override string Name => ".选中所有焊点";
        public override string Category => "TxTools";
        public override string Description => "选中所有焊点";
        public override string Bitmap => "image.weldpoint.bmp";
        public override void Execute(object cmdParams)
        {
            SelectPointsCore.Run(PointKind.Weld, "焊点");
        }
    }

    public class SelectAllPathPointsCmd : TxButtonCommand
    {
        public override string Name => ".选中所有过渡点";
        public override string Category => "TxTools";
        public override string Description => "选中所有过渡点";
        public override string Bitmap => "image.via.bmp";
        public override void Execute(object cmdParams)
        {
            SelectPointsCore.Run(PointKind.Path, "过渡点");
        }
    }

    public class SelectAllContinuousPointsCmd : TxButtonCommand
    {
        public override string Name => ".选中所有连续点";
        public override string Category => "TxTools";
        public override string Description => "选中所有连续点";
        public override string Bitmap => "image.seam.bmp";
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

        // ── 主入口 ──────────────────────────────────────────────────

        public static void Run(PointKind kind, string label)
        {
            try
            {
                // ① 取当前选中（双通道：GetItems + GetPlanningItems）
                TxObjectList sel = GetSelection();
                if (sel == null || sel.Count == 0)
                {
                    DLog("未选中任何对象，请先在操作树或路径编辑器中选中一个操作/操作组");
                    return;
                }

                // ② 从选中对象收集目标点位
                var hits = new List<ITxObject>();
                var hitIds = new HashSet<string>();
                int rootCount = 0;

                foreach (ITxObject r in sel)
                {
                    if (r == null) continue;
                    rootCount++;
                    CollectTargets(r, kind, hits, hitIds);
                }

                // ③ 设置选中
                if (hits.Count == 0)
                {
                    DLog(string.Format("在所选 {0} 个对象中未找到{1}", rootCount, label));
                    return;
                }

                var list = new TxObjectList();
                foreach (var h in hits) list.Add(h);

                try
                {
                    TxApplication.ActiveSelection.SetItems(list);
                    DLog(string.Format("选中 {0} 个{1}", hits.Count, label));
                }
                catch (Exception ex)
                {
                    DLog("SetItems 失败: " + ex.Message + "  命中对象数: " + hits.Count);
                }
            }
            catch (Exception ex)
            {
                DLog("顶层异常: " + ex.Message);
            }
        }

        // ── 获取选中（双通道）──────────────────────────────────────

        private static TxObjectList GetSelection()
        {
            TxObjectList sel = null;

            // 通道 1: GetItems（工程表示 — 3D 视口 + 操作树通用）
            try { sel = TxApplication.ActiveSelection.GetItems(); }
            catch (Exception ex) { DLog("GetItems 异常: " + ex.Message); }

            if (sel != null && sel.Count > 0)
            {
                DLog("GetItems 返回 " + sel.Count + " 项");
                return sel;
            }

            // 通道 2: GetPlanningItems（规划表示 — 操作树选中时可能在此返回）
            try { sel = TxApplication.ActiveSelection.GetPlanningItems(); }
            catch (Exception ex) { DLog("GetPlanningItems 异常: " + ex.Message); }

            if (sel != null)
                DLog("GetPlanningItems 返回 " + sel.Count + " 项");

            return sel;
        }

        // ── 从单个选中对象收集目标 ──────────────────────────────

        private static void CollectTargets(ITxObject obj, PointKind kind,
                                           List<ITxObject> hits, HashSet<string> hitIds)
        {
            if (obj == null) return;
            string tn = obj.GetType().Name;
            string nm = SafeName(obj);
            DLog("处理选中: " + tn + " [" + nm + "]");

            // ── ★ 连续点特殊：Continuous/Seam 容器 → 展开仅取 SeamLocation 子操作 ──
            //    不选中容器本身，仅选中其内部的 SeamLocation 类型操作
            if (kind == PointKind.Continuous && IsContinuousType(tn))
            {
                ITxObjectCollection contContainer = obj as ITxObjectCollection;
                if (contContainer != null)
                {
                    DrillContinuousContainer(contContainer, tn, hits, hitIds);
                    return;
                }
                else if (IsSeamLocation(tn))
                {
                    // SeamLocation 叶操作 → 直接选中
                    DLog("  → SeamLocation 叶操作: " + tn);
                    AddUnique(obj, hits, hitIds);
                    return;
                }
                else
                {
                    DLog("  → 非 SeamLocation 的 Continuous 类型，跳过: " + tn);
                    return;
                }
            }

            // ── 情况 A：obj 本身就是目标点位类型 → 直接加入 ──
            if (IsTargetPoint(obj, kind))
            {
                DLog("  → 直接匹配目标类型，加入命中");
                AddUnique(obj, hits, hitIds);
                // 不再继续遍历此对象的子树（点位对象无操作后代）
                return;
            }

            // ── 情况 B：obj 是容器（TxCompoundOperation / TxOperationRoot）
            //    → 用 GetAllDescendants(TxNoTypeFilter) 获取全部后代 ──
            ITxObjectCollection container = obj as ITxObjectCollection;
            if (container != null)
            {
                DLog("  → 容器类型，调用 GetAllDescendants");
                try
                {
                    TxObjectList descendants = container.GetAllDescendants(new TxNoTypeFilter());
                    if (descendants != null && descendants.Count > 0)
                    {
                        DLog("  → 后代总数: " + descendants.Count);
                        int matched = 0;

                        foreach (ITxObject desc in descendants)
                        {
                            if (desc == null) continue;

                            // ★ 连续点：后代是 Continuous/Seam 容器 → 展开仅取 SeamLocation 子操作
                            if (kind == PointKind.Continuous)
                            {
                                string descTn = desc.GetType().Name;
                                if (IsContinuousType(descTn))
                                {
                                    ITxObjectCollection contSub = desc as ITxObjectCollection;
                                    if (contSub != null)
                                    {
                                        DrillContinuousContainer(contSub, descTn, hits, hitIds);
                                        matched++;
                                        continue;
                                    }
                                    else if (IsSeamLocation(descTn))
                                    {
                                        // SeamLocation 叶操作 → 直接选中
                                        DLog("    → SeamLocation 叶操作: " + descTn);
                                        AddUnique(desc, hits, hitIds);
                                        matched++;
                                        continue;
                                    }
                                    else
                                    {
                                        // 非 SeamLocation → 跳过
                                        continue;
                                    }
                                }
                                // 非 Continuous/Seam 后代 → 跳过（连续点只关心 SeamLocation）
                                continue;
                            }

                            // 后代本身是目标点位类型 → 直接加入
                            if (IsTargetPoint(desc, kind))
                            {
                                AddUnique(desc, hits, hitIds);
                                matched++;
                                continue;
                            }

                            // 后代是操作 → 从操作提取点位对象
                            if (desc is ITxOperation op)
                            {
                                if (ExtractPointFromOp(op, kind, hits, hitIds))
                                    matched++;
                            }
                        }

                        DLog("  → 匹配命中: " + matched);
                    }
                    else
                    {
                        DLog("  → GetAllDescendants 返回空（容器无后代）");
                    }
                }
                catch (Exception ex)
                {
                    DLog("  → GetAllDescendants 异常: " + ex.Message);
                }
                return;
            }

            // ── 情况 C：obj 是叶操作（非容器）→ 直接提取点位 ──
            if (obj is ITxOperation leafOp)
            {
                DLog("  → 叶操作类型，直接提取点位");
                ExtractPointFromOp(leafOp, kind, hits, hitIds);
                return;
            }

            // ── 情况 D：obj 既非点位也非操作也非容器 → 无关联 ──
            DLog("  → 跳过（非点位/操作/容器）: " + tn);
        }

        // ── 从叶操作提取点位对象 ────────────────────────────────

        /// <returns>true 如果提取到至少一个命中</returns>
        private static bool ExtractPointFromOp(ITxOperation op, PointKind kind,
                                                List<ITxObject> hits, HashSet<string> hitIds)
        {
            string tn = op.GetType().Name;
            bool found = false;

            switch (kind)
            {
                // ── 焊点 ──
                // 主目标: 操作本身（TxWeldLocationOperation，ITxDisplayableObject，可选中）
                // 附带目标: TxWeldPoint（ITxWeldPoint 不继承 ITxDisplayableObject，
                //            SetItems 可能无法选中，但尝试附带以防 PS 特殊处理）
                // 判定: 是焊操作 → 加入操作本身 + 尝试附带 WeldPoint
                case PointKind.Weld:
                {
                    bool isWeldOp = op is TxWeldLocationOperation;
                    string tn2 = op.GetType().Name;
                    if (!isWeldOp)
                        isWeldOp = tn2.Contains("Weld") || tn2.Contains("Seam");

                    if (isWeldOp)
                    {
                        // ① 选中焊操作本身（与过渡点行为一致，操作是可选中/可显示的对象）
                        DLog("    [Weld] 选中焊操作: " + tn2);
                        AddUnique(op, hits, hitIds);
                        found = true;

                        // ② 尝试附带 TxWeldPoint（强类型 API）
                        if (op is TxWeldLocationOperation wop)
                        {
                            try
                            {
                                ITxWeldPoint wp = wop.WeldPoint;
                                if (wp != null)
                                {
                                    DLog("    [Weld] 附带 WeldPoint → " + SafeName(wp as ITxObject));
                                    AddUnique(wp as ITxObject, hits, hitIds);
                                }
                            }
                            catch (Exception ex) { DLog("    [Weld] WeldPoint 异常: " + ex.Message); }
                        }

                        // ③ dynamic WeldPoint 兜底（TxRoboticSeamLocationOperation 等子类）
                        try
                        {
                            dynamic dOp = op;
                            object wpObj = dOp.WeldPoint;
                            if (wpObj != null)
                            {
                                DLog("    [Weld] 附带 dynamic WeldPoint → " + wpObj.GetType().Name);
                                AddUnique(wpObj as ITxObject, hits, hitIds);
                            }
                        }
                        catch { /* 属性不存在则跳过 */ }
                    }
                    else
                    {
                        DLog("    [Weld] 操作 " + tn2 + " 不是焊操作，跳过");
                    }

                    break;
                }

                // ── 连续点 ──
                // Continuous/Seam 容器 → 不选中容器本身，展开仅取 SeamLocation 子操作
                // SeamLocation 叶操作 → 直接选中
                case PointKind.Continuous:
                {
                    if (IsContinuousType(tn))
                    {
                        ITxObjectCollection contContainer = op as ITxObjectCollection;
                        if (contContainer != null)
                        {
                            // 容器 → 展开仅取 SeamLocation
                            DLog("    [Continuous] 容器 " + tn + "，展开仅取 SeamLocation");
                            DrillContinuousContainer(contContainer, tn, hits, hitIds);
                            found = true;
                        }
                        else if (IsSeamLocation(tn))
                        {
                            DLog("    [Continuous] SeamLocation 叶操作 → " + tn);
                            AddUnique(op, hits, hitIds);
                            found = true;
                        }
                        else
                        {
                            DLog("    [Continuous] 非 SeamLocation 的 Continuous 类型，跳过: " + tn);
                        }
                    }
                    else if (IsSeamLocation(tn))
                    {
                        // 直接匹配 SeamLocation（独立出现的情况）
                        DLog("    [Continuous] 直接匹配 SeamLocation → " + tn);
                        AddUnique(op, hits, hitIds);
                        found = true;
                    }
                    else
                    {
                        DLog("    [Continuous] 操作 " + tn + " 非 SeamLocation，跳过");
                    }
                    break;
                }

                // ── 过渡点 ──
                // 目标对象: 操作本身（ITxLocationOperation 且非焊非连续）
                // 包括: Via, Home, Arc, GenericRoboticLoc 等类型操作
                case PointKind.Path:
                {
                    bool isLocOp  = op is ITxLocationOperation || tn.Contains("Location");
                    bool isWeld   = op is TxWeldLocationOperation || tn.Contains("Weld") || tn.Contains("Seam");
                    bool isCont   = tn.Contains("Continuous");

                    if (isLocOp && !isWeld && !isCont)
                    {
                        DLog("    [Path] → " + tn);
                        AddUnique(op, hits, hitIds);
                        found = true;
                    }
                    else
                    {
                        DLog("    [Path] 操作 " + tn + " 不符合过渡点条件，跳过");
                    }
                    break;
                }
            }

            return found;
        }

        // ════════════════════════════════════════════════════════════
        //  目标点位类型判定
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 判断 obj 是否本身就是我们要选中的点位对象。
        /// 焊点 → TxWeldLocationOperation 操作（ITxDisplayableObject，可选中）
        ///        + TxWeldPoint 模型对象（尝试附带，但非 ITxDisplayableObject）
        /// 过渡点 → 操作对象本身（无独立点位模型）
        /// 连续点 → 不在此处匹配，由 CollectTargets 特殊逻辑处理（展开容器取子操作）
        /// </summary>
        private static bool IsTargetPoint(ITxObject obj, PointKind kind)
        {
            if (obj == null) return false;

            switch (kind)
            {
                case PointKind.Weld:
                    // 焊点可选中对象 = 焊操作本身（ITxDisplayableObject）
                    //                  + TxWeldPoint 模型对象（尝试附带）
                    return obj is TxWeldPoint || obj is ITxWeldPoint
                        || obj is TxWeldLocationOperation;

                case PointKind.Continuous:
                    // 连续点不在此处匹配：
                    //   Continuous/Seam 容器 → 由 CollectTargets 特殊逻辑展开取子操作
                    //   Continuous/Seam 叶操作 → 由 ExtractPointFromOp 处理
                    return false;

                case PointKind.Path:
                {
                    // 过渡点 = 操作本身（Location-like 且非焊非连续）
                    var tn = obj.GetType().Name;
                    bool isLocLike = tn.Contains("Via") || tn.Contains("Home")
                                  || tn.Contains("Arc") || tn.Contains("Location")
                                  || tn.Contains("PathPoint");
                    bool isWeld = obj is TxWeldPoint || obj is ITxWeldPoint
                              || obj is TxWeldLocationOperation
                              || tn.Contains("Weld") || tn.Contains("Seam");
                    bool isCont = tn.Contains("Continuous");
                    return isLocLike && !isWeld && !isCont;
                }

                default: return false;
            }
        }

        // ════════════════════════════════════════════════════════════
        //  连续点容器展开
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 判断类型名是否是 Continuous/Seam 类型（容器或叶操作）
        /// </summary>
        private static bool IsContinuousType(string typeName)
        {
            return typeName.Contains("Continuous") || typeName.Contains("Seam");
        }

        /// <summary>
        /// 判断类型名是否是 SeamLocation 类型（连续点的真正目标）
        /// 如 TxRoboticSeamLocationOperation, TxSeamLocationOperation 等
        /// </summary>
        private static bool IsSeamLocation(string typeName)
        {
            return typeName.Contains("SeamLocation");
        }

        /// <summary>
        /// 展开 Continuous/Seam 容器，仅选中其内部 SeamLocation 类型的子操作。
        /// 连续点容器是 CompoundOperation，SeamLocation 子操作才是真正的"连续点位"。
        /// 其他子操作（如 Via, Home 等过渡点）不在连续点按钮的选中范围内。
        /// </summary>
        private static void DrillContinuousContainer(ITxObjectCollection container, string tn,
                                                      List<ITxObject> hits, HashSet<string> hitIds)
        {
            DLog("  → 展开 Continuous 容器 " + tn);
            try
            {
                TxObjectList childDescs = container.GetAllDescendants(new TxNoTypeFilter());
                if (childDescs != null && childDescs.Count > 0)
                {
                    DLog("  → 子操作总数: " + childDescs.Count);
                    foreach (ITxObject child in childDescs)
                    {
                        if (child == null) continue;
                        string childTn = child.GetType().Name;
                        // 仅选中 SeamLocation 类型的子操作（如 TxRoboticSeamLocationOperation）
                        if (IsSeamLocation(childTn))
                        {
                            DLog("    → SeamLocation: " + childTn + " [" + SafeName(child) + "]");
                            AddUnique(child, hits, hitIds);
                        }
                    }
                }
                else
                {
                    // 容器无子操作 → 退回选中容器本身
                    DLog("  → " + tn + " 无子操作，退回选中容器本身");
                    AddUnique(container as ITxObject, hits, hitIds);
                }
            }
            catch (Exception ex)
            {
                DLog("  → GetAllDescendants 异常: " + ex.Message + "，退回选中容器本身");
                AddUnique(container as ITxObject, hits, hitIds);
            }
        }

        // ════════════════════════════════════════════════════════════
        //  去重 + 工具
        // ════════════════════════════════════════════════════════════

        private static void AddUnique(ITxObject obj, List<ITxObject> hits, HashSet<string> hitIds)
        {
            if (obj == null) return;
            string id = TryGetId(obj);
            if (id != null && hitIds.Contains(id)) return;
            if (id != null) hitIds.Add(id);
            hits.Add(obj);
        }

        private static string TryGetId(ITxObject obj)
        {
            try { dynamic d = obj; return d.Id as string; }
            catch { return null; }
        }

        private static string SafeName(ITxObject obj)
        {
            if (obj == null) return "(null)";
            try { return obj.Name ?? "(无名)"; }
            catch { return "(无名)"; }
        }

        // ════════════════════════════════════════════════════════════
        //  日志
        // ════════════════════════════════════════════════════════════

        private static void DLog(string msg)
        {
            try { System.Diagnostics.Debug.WriteLine(LOG_PREFIX + msg); } catch { }
        }
    }
}
