using System;
using System.Collections.Generic;
using System.Reflection;
using Tecnomatix.Engineering;

namespace LineToSolid
{
    /// <summary>
    /// 一个线段（直线段）的几何信息，世界坐标系。
    /// </summary>
    public class LineSegment
    {
        public TxVector Start { get; set; }
        public TxVector End { get; set; }
        public string SourceName { get; set; }
        public int IndexInSource { get; set; }

        public double Length
        {
            get
            {
                double dx = End.X - Start.X, dy = End.Y - Start.Y, dz = End.Z - Start.Z;
                return Math.Sqrt(dx * dx + dy * dy + dz * dz);
            }
        }

        public TxVector Direction
        {
            get
            {
                double L = Length;
                if (L < 1e-9) return new TxVector(1, 0, 0);
                return new TxVector(
                    (End.X - Start.X) / L,
                    (End.Y - Start.Y) / L,
                    (End.Z - Start.Z) / L);
            }
        }

        public TxVector Midpoint
        {
            get
            {
                return new TxVector(
                    (Start.X + End.X) * 0.5,
                    (Start.Y + End.Y) * 0.5,
                    (Start.Z + End.Z) * 0.5);
            }
        }
    }

    public enum FeatureKind
    {
        Unknown,
        Polyline,           // TxPolyline
        OtherCurve          // TxLine / TxCurve / 任意可参数化曲线
    }

    /// <summary>
    /// 从 PS 场景读取曲线特征。
    /// PS 2402 验证过的 API：
    ///   - TxPolyline.GetVertices() : TxParameterizedPoint[]
    ///   - TxPolyline.GetPointByParameter(double) : TxVector  （TxPolyline 类方法，非接口方法）
    ///   - TxPolyline.Length() : double
    /// 策略：
    ///   - TxPolyline 用 GetVertices() 直接取节点
    ///   - 其他曲线（TxLine/TxCurve/...）用反射调 GetPointByParameter，
    ///     按弦高自适应采样
    /// </summary>
    public static class PolylineReader
    {
        public class ReadOptions
        {
            public double MaxSagitta { get; set; } = 0.5;
            public int MinSegmentsPerCurve { get; set; } = 8;
            public int MaxSegmentsPerCurve { get; set; } = 5000;
        }

        public class ReadResult
        {
            public List<LineSegment> Segments { get; set; } = new List<LineSegment>();
            public int FeatureCount { get; set; }
            public List<string> Diagnostics { get; set; } = new List<string>();
        }

        /// <summary>
        /// 从当前选择中取所有支持的曲线特征。
        /// 如果选中的是 Resource/Part/Compound 这样的容器对象（非曲线），
        /// 则递归遍历它的后代，把所有曲线后代也加进来。
        /// </summary>
        public static List<ITxObject> GetSelectedCurveFeatures()
        {
            List<string> _;
            return GetSelectedCurveFeatures(out _);
        }

        public static List<ITxObject> GetSelectedCurveFeatures(out List<string> diagnostics)
        {
            var list = new List<ITxObject>();
            diagnostics = new List<string>();
            try
            {
                var sel = TxApplication.ActiveSelection;
                if (sel == null) { diagnostics.Add("ActiveSelection=null"); return list; }
                var items = sel.GetItems();
                if (items == null) { diagnostics.Add("GetItems()=null"); return list; }

                int total = 0;
                foreach (var obj in items) total++;
                diagnostics.Add(string.Format("选中对象数：{0}", total));

                foreach (var obj in items)
                {
                    var ito = obj as ITxObject;
                    if (ito == null)
                    {
                        diagnostics.Add(string.Format("  [跳过] 非 ITxObject：{0}", obj.GetType().Name));
                        continue;
                    }

                    var kind = ClassifyFeature(obj);
                    if (kind != FeatureKind.Unknown)
                    {
                        if (!list.Contains(ito)) list.Add(ito);
                        diagnostics.Add(string.Format("  [直接加入] {0}（{1}）",
                            SafeName(ito), obj.GetType().Name));
                        continue;
                    }

                    // 非曲线 → 尝试当作容器递归
                    diagnostics.Add(string.Format("  [容器] {0}（{1}）尝试递归遍历...",
                        SafeName(ito), obj.GetType().Name));
                    int before = list.Count;
                    CollectCurvesRecursive(ito, list, diagnostics, 0);
                    diagnostics.Add(string.Format("  [容器] {0}：新增 {1} 个曲线",
                        SafeName(ito), list.Count - before));
                }
            }
            catch (Exception ex)
            {
                diagnostics.Add("异常：" + ex.Message);
            }
            return list;
        }

        private static string SafeName(ITxObject o)
        {
            try { return o.Name; } catch { return "?"; }
        }

        /// <summary>
        /// 公开重载：把 container 下所有曲线后代收集到 list 里（不需要诊断信息）。
        /// </summary>
        public static void CollectCurvesRecursive(object container, List<ITxObject> list)
        {
            CollectCurvesRecursive(container, list, null, 0);
        }

        /// <summary>
        /// 递归遍历容器对象的所有后代，把曲线特征收集到 list 里。
        /// 探测顺序：
        ///   1) GetAllDescendants(null) —— 深度遍历，一次性
        ///   2) GetAllDescendants() 无参 —— 某些版本可能支持
        ///   3) GetDirectDescendants(null) + 手动递归
        ///   4) IEnumerable —— TxComponent 实现了 IEnumerable，直接迭代子项
        /// </summary>
        private static void CollectCurvesRecursive(object container, List<ITxObject> list,
            List<string> diag, int depth)
        {
            if (container == null) return;
            if (depth > 50) return;  // 防无限递归

            string indent = new string(' ', depth * 2 + 4);

            // 1) GetAllDescendants(ITxTypeFilter) 传 null
            var descendants = TryInvokeWithNullArg(container, "GetAllDescendants");
            if (descendants != null)
            {
                int n = 0, curve = 0;
                foreach (var d in descendants)
                {
                    n++;
                    if (d == null) continue;
                    if (ClassifyFeature(d) != FeatureKind.Unknown)
                    {
                        var ito = d as ITxObject;
                        if (ito != null && !list.Contains(ito)) { list.Add(ito); curve++; }
                    }
                }
                if (diag != null)
                    diag.Add(indent + string.Format("GetAllDescendants(null) → {0} 个后代，{1} 个曲线",
                        n, curve));
                return;
            }

            // 2) GetDirectDescendants(null) + 手工递归
            var direct = TryInvokeWithNullArg(container, "GetDirectDescendants");
            if (direct != null)
            {
                int n = 0;
                foreach (var d in direct)
                {
                    n++;
                    if (d == null) continue;
                    if (ClassifyFeature(d) != FeatureKind.Unknown)
                    {
                        var ito = d as ITxObject;
                        if (ito != null && !list.Contains(ito)) list.Add(ito);
                    }
                    else
                    {
                        CollectCurvesRecursive(d, list, diag, depth + 1);
                    }
                }
                if (diag != null)
                    diag.Add(indent + string.Format("GetDirectDescendants(null) → {0} 个直接子",
                        n));
                return;
            }

            // 3) 退路：把容器当 IEnumerable 迭代
            var enumerable = container as System.Collections.IEnumerable;
            if (enumerable != null)
            {
                int n = 0;
                foreach (var d in enumerable)
                {
                    n++;
                    if (d == null) continue;
                    if (ClassifyFeature(d) != FeatureKind.Unknown)
                    {
                        var ito = d as ITxObject;
                        if (ito != null && !list.Contains(ito)) list.Add(ito);
                    }
                    else
                    {
                        CollectCurvesRecursive(d, list, diag, depth + 1);
                    }
                }
                if (diag != null)
                    diag.Add(indent + string.Format("IEnumerable → {0} 个子", n));
                return;
            }

            if (diag != null)
                diag.Add(indent + string.Format("容器 {0} 没有可枚举接口",
                    container.GetType().Name));
        }

        /// <summary>
        /// 在对象上找接受 1 个参数的指定方法，用 null 调用，返回结果 IEnumerable。
        /// </summary>
        private static System.Collections.IEnumerable TryInvokeWithNullArg(object container, string methodName)
        {
            try
            {
                var t = container.GetType();
                foreach (var mi in t.GetMethods())
                {
                    if (mi.Name != methodName) continue;
                    var pars = mi.GetParameters();
                    if (pars.Length != 1) continue;
                    // 接受参数类型为引用类型或可空
                    if (pars[0].ParameterType.IsValueType) continue;
                    var r = mi.Invoke(container, new object[] { null });
                    return r as System.Collections.IEnumerable;
                }
            }
            catch { }
            return null;
        }

        public static FeatureKind ClassifyFeature(object obj)
        {
            if (obj == null) return FeatureKind.Unknown;
            if (obj is TxPolyline) return FeatureKind.Polyline;

            // 其他曲线：只要能找到 GetPointByParameter(double) 方法就接受
            var mi = FindGetPointMethod(obj);
            return mi != null ? FeatureKind.OtherCurve : FeatureKind.Unknown;
        }

        public static ReadResult ExtractAll(IEnumerable<ITxObject> features, ReadOptions opts)
        {
            var result = new ReadResult();
            if (opts == null) opts = new ReadOptions();
            if (features == null) return result;

            foreach (var feat in features)
            {
                if (feat == null) continue;
                result.FeatureCount++;
                string name = GetName(feat);

                try
                {
                    if (feat is TxPolyline)
                    {
                        AppendPolylineSegments((TxPolyline)feat, name, result);
                    }
                    else
                    {
                        var mi = FindGetPointMethod(feat);
                        if (mi != null)
                            AppendParametricCurveSegments(feat, mi, name, opts, result);
                        else
                            result.Diagnostics.Add(string.Format(
                                "[跳过] {0}：未找到 GetPointByParameter（类型 {1}）",
                                name, feat.GetType().Name));
                    }
                }
                catch (Exception ex)
                {
                    result.Diagnostics.Add(string.Format(
                        "[错误] 处理 {0} 时异常：{1}", name, ex.Message));
                }
            }

            return result;
        }

        // ---------- TxPolyline ----------

        private static void AppendPolylineSegments(TxPolyline poly, string srcName, ReadResult result)
        {
            TxParameterizedPoint[] verts = poly.GetVertices();
            if (verts == null || verts.Length < 2)
            {
                result.Diagnostics.Add(string.Format("[跳过] {0}：顶点数 < 2", srcName));
                return;
            }

            var points = new List<TxVector>(verts.Length);
            foreach (var pp in verts)
            {
                var v = ExtractPoint(pp);
                if (v != null) points.Add(v);
            }
            if (points.Count < 2)
            {
                result.Diagnostics.Add(string.Format("[跳过] {0}：能解析的顶点 < 2", srcName));
                return;
            }

            int kept = 0;
            for (int i = 0; i < points.Count - 1; i++)
            {
                var seg = new LineSegment
                {
                    Start = points[i],
                    End = points[i + 1],
                    SourceName = srcName,
                    IndexInSource = kept
                };
                if (seg.Length > 1e-6) { result.Segments.Add(seg); kept++; }
            }
            result.Diagnostics.Add(string.Format(
                "[Polyline] {0}：{1} 顶点 → {2} 段", srcName, verts.Length, kept));
        }

        private static TxVector ExtractPoint(TxParameterizedPoint pp)
        {
            if (pp == null) return null;
            var t = pp.GetType();
            foreach (var name in new[] { "Point", "Position", "Location", "Vertex" })
            {
                try
                {
                    var p = t.GetProperty(name);
                    if (p != null)
                    {
                        var v = p.GetValue(pp, null);
                        if (v is TxVector) return (TxVector)v;
                    }
                }
                catch { }
            }
            try
            {
                dynamic d = pp;
                return new TxVector((double)d.X, (double)d.Y, (double)d.Z);
            }
            catch { }
            return null;
        }

        // ---------- 参数化曲线（TxLine/TxCurve/任意） ----------

        /// <summary>
        /// 在曲线对象上找 GetPointByParameter(double) 方法。
        /// 这个方法不在 ITx1DimensionalGeometry 接口上，是各个具体类自己定义的。
        /// </summary>
        private static MethodInfo FindGetPointMethod(object feat)
        {
            if (feat == null) return null;
            try
            {
                var t = feat.GetType();
                return t.GetMethod("GetPointByParameter", new[] { typeof(double) });
            }
            catch { return null; }
        }

        /// <summary>
        /// 弦高自适应采样：先粗采样，再二分细分至所有相邻参数对的弦高 ≤ MaxSagitta。
        /// </summary>
        private static void AppendParametricCurveSegments(
            object curve, MethodInfo getPoint, string srcName,
            ReadOptions opts, ReadResult result)
        {
            var samples = new SortedDictionary<double, TxVector>();
            int initial = Math.Max(opts.MinSegmentsPerCurve, 2);
            for (int i = 0; i <= initial; i++)
            {
                double t = (double)i / initial;
                samples[t] = SafeInvokeGetPoint(curve, getPoint, t);
            }

            int safety = opts.MaxSegmentsPerCurve;
            bool changed = true;
            while (changed && samples.Count < opts.MaxSegmentsPerCurve && safety > 0)
            {
                changed = false;
                var keys = new List<double>(samples.Keys);
                for (int i = 0; i < keys.Count - 1; i++)
                {
                    if (samples.Count >= opts.MaxSegmentsPerCurve) break;
                    double a = keys[i], b = keys[i + 1];
                    var pa = samples[a]; var pb = samples[b];
                    if (pa == null || pb == null) continue;

                    double tm = 0.5 * (a + b);
                    if (samples.ContainsKey(tm)) continue;
                    var pm = SafeInvokeGetPoint(curve, getPoint, tm);
                    if (pm == null) continue;

                    var chordMid = new TxVector(
                        0.5 * (pa.X + pb.X),
                        0.5 * (pa.Y + pb.Y),
                        0.5 * (pa.Z + pb.Z));
                    double sag = Dist(pm, chordMid);
                    if (sag > opts.MaxSagitta)
                    {
                        samples[tm] = pm;
                        changed = true;
                    }
                    safety--;
                    if (safety <= 0) break;
                }
            }

            var ptsOrdered = new List<TxVector>(samples.Count);
            foreach (var kv in samples) if (kv.Value != null) ptsOrdered.Add(kv.Value);

            int kept = 0;
            for (int i = 0; i < ptsOrdered.Count - 1; i++)
            {
                var seg = new LineSegment
                {
                    Start = ptsOrdered[i],
                    End = ptsOrdered[i + 1],
                    SourceName = srcName,
                    IndexInSource = kept
                };
                if (seg.Length > 1e-6) { result.Segments.Add(seg); kept++; }
            }
            result.Diagnostics.Add(string.Format(
                "[Curve] {0}（{1}）：{2} 采样点 → {3} 段，弦高≤{4}",
                srcName, curve.GetType().Name, ptsOrdered.Count, kept, opts.MaxSagitta));
        }

        private static TxVector SafeInvokeGetPoint(object curve, MethodInfo mi, double t)
        {
            try
            {
                var r = mi.Invoke(curve, new object[] { t });
                return r as TxVector;
            }
            catch { return null; }
        }

        // ---------- 工具 ----------

        private static string GetName(object obj)
        {
            var io = obj as ITxObject;
            if (io != null) { try { return io.Name; } catch { } }
            return obj.GetType().Name;
        }

        private static double Dist(TxVector a, TxVector b)
        {
            double dx = a.X - b.X, dy = a.Y - b.Y, dz = a.Z - b.Z;
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }
    }
}