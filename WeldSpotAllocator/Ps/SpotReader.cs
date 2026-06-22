// SpotReader.cs  —  C# 7.3
// 从操作树读取焊接操作 → 有序焊点/过渡点，持有底层 TxWeldPoint / via 引用（写入层需要）。
// 复用 PsReader 的静态工具：TxToArr / GetAppearancesFromWeldPoint。
//
// 枚举策略：选中对象中递归找出 TxWeldOperation（含复合操作下层），
//   对每条 weldOp 用 GetDirectDescendants 取其直接子操作，按出现顺序分流：
//     · 类型名含 "WeldLocation" → TxWeldLocationOperation，.WeldPoint 得 TxWeldPoint
//     · 类型名含 "Via"          → TxRoboticViaLocationOperation（过渡点）
//   ⚠ 若 GetDirectDescendants 不保证轨迹顺序，改用 GetChildAt(i) 逐序号遍历（见 EnumChildrenOrdered）。

using System;
using System.Collections;
using System.Collections.Generic;
using Tecnomatix.Engineering;
using MyPlugin.ExportGun;

namespace MyPlugin.WeldSpotAllocator
{
    public static class SpotReader
    {
        private static readonly TxTypeFilter AnyFilter = new TxTypeFilter(typeof(ITxObject));

        // ── 读取一条焊接操作 ────────────────────────────────────────────────
        public static OpData ReadOneWeldOp(ITxObject weldOp, Action<string> log)
        {
            var od = new OpData { Name = SafeName(weldOp), Raw = weldOp };
            foreach (ITxObject child in EnumChildrenOrdered(weldOp))
            {
                string tn = child.GetType().Name;
                try
                {
                    if (tn.Contains("WeldLocation"))           // TxWeldLocationOperation
                    {
                        dynamic d = child;
                        TxWeldPoint wp = null;
                        try { wp = d.WeldPoint as TxWeldPoint; } catch { }
                        if (wp == null) continue;
                        var sd = MakeWeldSpot(wp, child);
                        od.Spots.Add(sd); od.Ordered.Add(sd);
                    }
                    else if (tn.Contains("Via"))               // TxRoboticViaLocationOperation
                    {
                        var sd = MakeViaSpot(child);
                        if (sd != null) { od.Vias.Add(sd); od.Ordered.Add(sd); }
                    }
                }
                catch (Exception ex) { log($"[Reader] [{od.Name}] 子项[{tn}] 解析异常：{ex.Message}"); }
            }
            return od;
        }

        // 优先按序号遍历以保证轨迹顺序；失败回退 GetDirectDescendants
        private static IEnumerable<ITxObject> EnumChildrenOrdered(ITxObject weldOp)
        {
            // 尝试 GetChildCount + GetChildAt(i)
            int count = -1;
            try { dynamic d = weldOp; count = (int)d.GetChildCount(); } catch { count = -1; }
            if (count >= 0)
            {
                for (int i = 0; i < count; i++)
                {
                    ITxObject c = null;
                    try { dynamic d = weldOp; c = d.GetChildAt(i) as ITxObject; } catch { }
                    if (c != null) yield return c;
                }
                yield break;
            }
            var kids = DirectKids(weldOp);
            if (kids != null) foreach (ITxObject k in kids) yield return k;
        }

        // ── 构造 SpotData ───────────────────────────────────────────────────
        private static SpotData MakeWeldSpot(TxWeldPoint wp, ITxObject locOp)
        {
            TxTransformation tx = SafeTx(() => wp.AbsoluteLocation);
            double[] m = PsReader.TxToArr(tx);
            var sd = new SpotData
            {
                Name = SafeName(wp),
                Kind = PointType.WeldPoint,
                Matrix = m,
                Position = new[] { m[3], m[7], m[11] },
                SymCenter = ReadSymCenter(wp, locOp),
                Raw = wp,
                LocOp = locOp
            };
            return sd;
        }

        private static SpotData MakeViaSpot(ITxObject via)
        {
            TxTransformation tx = null;
            try { dynamic d = via; tx = d.AbsoluteLocation as TxTransformation; } catch { }
            if (tx == null) return null;
            double[] m = PsReader.TxToArr(tx);
            return new SpotData
            {
                Name = SafeName(via),
                Kind = PointType.PathPoint,
                Matrix = m,
                Position = new[] { m[3], m[7], m[11] },
                Raw = via,
                LocOp = via
            };
        }

        // ── 小工具 ──────────────────────────────────────────────────────────
        private static double[] ReadSymCenter(TxWeldPoint wp, ITxObject locOp)
        {
            // 首选：焊点位 Attach/绑定的分身车件原点（文档所述对称中心默认值）
            if (locOp != null)
            {
                try
                {
                    dynamic d = locOp;
                    var parent = d.AttachmentParent as ITxObject;
                    if (parent != null)
                    {
                        TxTransformation tx = null;
                        try { dynamic pd = parent; tx = pd.AbsoluteLocation as TxTransformation; } catch { }
                        if (tx != null) return PsReader.TxToArr(tx);
                    }
                }
                catch { }
            }
            // 回退：焊点绑定零件的外观矩阵
            try
            {
                var apps = PsReader.GetAppearancesFromWeldPoint(wp);
                if (apps != null && apps.Count > 0 && apps[0].Matrix != null) return apps[0].Matrix;
            }
            catch { }
            return null;
        }

        private static TxObjectList DirectKids(ITxObject node)
        {
            try { dynamic d = node; return d.GetDirectDescendants(AnyFilter) as TxObjectList; } catch { }
            // 兜底：可枚举
            try { if (node is IEnumerable ie) { var l = new TxObjectList(); foreach (ITxObject c in ie) l.Add(c); return l; } } catch { }
            return null;
        }

        private static TxTransformation SafeTx(Func<TxTransformation> f) { try { return f(); } catch { return null; } }
        private static string SafeName(ITxObject o) { try { dynamic d = o; return (d.Name as string) ?? o.GetType().Name; } catch { return o.GetType().Name; } }
    }
}
