using System;
using System.Collections.Generic;
using System.Reflection;
using Tecnomatix.Engineering;

namespace TxTools.FenceBuilder
{
    public class BaseSegment
    {
        public TxVector Start;
        public TxVector End;
        public string SourceFeature;
        public int SegmentIndexInFeature;

        public double Length
        {
            get
            {
                double dx = End.X - Start.X, dy = End.Y - Start.Y;
                return Math.Sqrt(dx * dx + dy * dy);
            }
        }

        public TxVector DirectionXY
        {
            get
            {
                double dx = End.X - Start.X, dy = End.Y - Start.Y;
                double len = Math.Sqrt(dx * dx + dy * dy);
                if (len < 1e-6) return new TxVector(1, 0, 0);
                return new TxVector(dx / len, dy / len, 0);
            }
        }
    }

    public class BaseSegmentChain
    {
        public string FeatureName;
        public List<BaseSegment> Segments = new List<BaseSegment>();
        public double GroundZ;
    }

    public enum FeatureKind
    {
        Unknown,
        Polyline,
        OtherCurve
    }

    /// <summary>
    /// 基线读取器,参考 TxTools.LineToSolid 的 PolylineReader 策略:
    ///   - TxPolyline 强类型: GetVertices() → TxParameterizedPoint[]
    ///   - 其他曲线(TxLine/TxCurve/...): 反射调 GetPointByParameter(double)
    ///   - 容器对象(Resource/Part/Compound): 自动递归遍历后代挑曲线
    /// </summary>
    public static class FenceBaselineReader
    {
        public const int CurveSampleCount = 32;

        public static BaseSegmentChain ReadAsChain(ITxObject obj, Action<string> log)
        {
            if (obj == null) return null;
            string objName = TrySafeName(obj);
            string typeName = obj.GetType().Name;
            log?.Invoke("[Reader] 读取: " + objName + " (Type=" + typeName + ")");

            FeatureKind kind = ClassifyFeature(obj);
            if (kind == FeatureKind.Unknown)
            {
                log?.Invoke("[Reader] WARN 非曲线特征,跳过 " + objName);
                return null;
            }

            List<TxVector> pts = null;

            // 路径 1: TxPolyline 强类型
            if (kind == FeatureKind.Polyline)
            {
                pts = TryReadAsPolyline(obj, log);
            }

            // 路径 2: 反射 GetPointByParameter 参数化采样,共线检测
            if (pts == null)
            {
                pts = TryReadAs1DGeometry(obj, log);
            }

            // 路径 3: 反射 StartPoint/EndPoint 直线两端
            if (pts == null)
            {
                pts = TryReadAsLineEndpoints(obj, log);
            }

            if (pts == null || pts.Count < 2)
            {
                log?.Invoke("[Reader] WARN 未能读取顶点序列,跳过 " + objName);
                return null;
            }

            double groundZ = pts[0].Z;
            log?.Invoke("[Reader] 顶点数=" + pts.Count + " 地面 Z=" + groundZ.ToString("F2"));

            BaseSegmentChain chain = new BaseSegmentChain
            {
                FeatureName = objName,
                GroundZ = groundZ
            };

            for (int i = 0; i < pts.Count - 1; i++)
            {
                BaseSegment seg = new BaseSegment
                {
                    Start = new TxVector(pts[i].X, pts[i].Y, groundZ),
                    End = new TxVector(pts[i + 1].X, pts[i + 1].Y, groundZ),
                    SourceFeature = objName,
                    SegmentIndexInFeature = i
                };
                if (seg.Length < 1e-3)
                {
                    log?.Invoke("[Reader] 跳过零长度段 #" + i);
                    continue;
                }
                chain.Segments.Add(seg);
            }
            return chain.Segments.Count == 0 ? null : chain;
        }

        /// <summary>
        /// 沿用 TxTools.LineToSolid 同款判定:
        ///   - TxPolyline → Polyline
        ///   - 有 GetPointByParameter(double) → OtherCurve
        ///   - 有 StartPoint/EndPoint 属性对 → OtherCurve
        ///   - 其他 → Unknown
        /// </summary>
        public static FeatureKind ClassifyFeature(object obj)
        {
            if (obj == null) return FeatureKind.Unknown;
            if (obj is TxPolyline) return FeatureKind.Polyline;
            try
            {
                Type t = obj.GetType();
                if (t.GetMethod("GetPointByParameter", new[] { typeof(double) }) != null)
                    return FeatureKind.OtherCurve;
                // 直线端点对(TxLine 等)
                string[][] pairs = {
                    new[] { "StartPoint", "EndPoint" },
                    new[] { "Start", "End" },
                    new[] { "P1", "P2" }
                };
                foreach (var pair in pairs)
                {
                    if (t.GetProperty(pair[0]) != null && t.GetProperty(pair[1]) != null)
                        return FeatureKind.OtherCurve;
                }
            }
            catch { }
            return FeatureKind.Unknown;
        }

        public static bool IsCurveFeature(object obj)
        {
            return ClassifyFeature(obj) != FeatureKind.Unknown;
        }

        /// <summary>
        /// 把种子列表展开为全部曲线特征。
        /// 种子本身是曲线 → 直接加入;否则当容器递归遍历后代收集曲线。
        /// 沿用 TxTools.LineToSolid PolylineReader.GetSelectedCurveFeatures 的逻辑。
        /// </summary>
        public static List<ITxObject> ExpandToCurveFeatures(
            IEnumerable<ITxObject> seeds, Action<string> log)
        {
            var result = new List<ITxObject>();
            int seedIdx = 0;
            foreach (var seed in seeds)
            {
                seedIdx++;
                if (seed == null) continue;
                string seedName = TrySafeName(seed);
                string seedType = seed.GetType().Name;

                FeatureKind k = ClassifyFeature(seed);
                if (k != FeatureKind.Unknown)
                {
                    if (!result.Contains(seed)) result.Add(seed);
                    log?.Invoke("[Reader] 种子#" + seedIdx + " " + seedName + " (" + seedType + ") → 是曲线特征");
                    continue;
                }

                log?.Invoke("[Reader] 种子#" + seedIdx + " " + seedName + " (" + seedType + ") → 容器,递归遍历...");
                int before = result.Count;
                CollectCurvesRecursive(seed, result, log, 0);
                log?.Invoke("[Reader] 种子#" + seedIdx + " 新增曲线 " + (result.Count - before) + " 个");
            }
            return result;
        }

        /// <summary>
        /// 递归遍历容器,收集所有曲线后代。
        /// 4 路径(沿用 TxTools.LineToSolid):
        ///   1) GetAllDescendants(null) 一次性深度遍历
        ///   2) GetDirectDescendants(null) + 手动递归
        ///   3) Components/Children/SubObjects/Items/Members 属性
        ///   4) IEnumerable 兜底
        /// </summary>
        public static void CollectCurvesRecursive(object container, List<ITxObject> list,
            Action<string> log, int depth)
        {
            if (container == null || depth > 50) return;
            string indent = new string(' ', depth * 2 + 4);

            // 路径 1: GetAllDescendants(null)
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
                log?.Invoke("[Reader]" + indent + "GetAllDescendants(null) → 后代 " + n + " 个,曲线 " + curve);
                return;
            }

            // 路径 2: GetDirectDescendants(null) + 递归
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
                        CollectCurvesRecursive(d, list, log, depth + 1);
                    }
                }
                log?.Invoke("[Reader]" + indent + "GetDirectDescendants(null) → 直接子 " + n);
                return;
            }

            // 路径 3: 反射常见容器属性
            var byProp = TryGetByEnumerableProperty(container, log, depth);
            if (byProp != null)
            {
                int n = 0;
                foreach (var d in byProp)
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
                        CollectCurvesRecursive(d, list, log, depth + 1);
                    }
                }
                log?.Invoke("[Reader]" + indent + "属性遍历 → " + n + " 个子");
                return;
            }

            // 路径 4: IEnumerable
            var en = container as System.Collections.IEnumerable;
            if (en != null)
            {
                int n = 0;
                foreach (var d in en)
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
                        CollectCurvesRecursive(d, list, log, depth + 1);
                    }
                }
                log?.Invoke("[Reader]" + indent + "IEnumerable → " + n + " 个子");
                return;
            }

            log?.Invoke("[Reader]" + indent + "WARN 容器 " + container.GetType().Name + " 无可枚举接口");
        }

        // ============ TxPolyline 路径 ============

        private static List<TxVector> TryReadAsPolyline(ITxObject obj, Action<string> log)
        {
            try
            {
                TxPolyline poly = obj as TxPolyline;
                if (poly == null) return null;
                TxParameterizedPoint[] verts = poly.GetVertices();
                if (verts == null || verts.Length < 2) return null;

                var list = new List<TxVector>();
                foreach (var pp in verts)
                {
                    TxVector v = ExtractPoint(pp);
                    if (v.X != 0 || v.Y != 0 || v.Z != 0 || list.Count == 0)
                        list.Add(v);
                    else list.Add(v); // 包括原点也添加
                }
                if (list.Count >= 2)
                {
                    log?.Invoke("[Reader] 路径1 TxPolyline.GetVertices() OK, 共 " + list.Count + " 顶点");
                    return list;
                }
            }
            catch (Exception ex) { log?.Invoke("[Reader] 路径1 异常: " + ex.Message); }
            return null;
        }

        private static TxVector ExtractPoint(TxParameterizedPoint pp)
        {
            if (pp == null) return new TxVector(0, 0, 0);
            Type t = pp.GetType();
            foreach (string n in new[] { "Point", "Position", "Location", "Vertex" })
            {
                try
                {
                    PropertyInfo p = t.GetProperty(n);
                    if (p != null)
                    {
                        object v = p.GetValue(pp, null);
                        if (v is TxVector) return (TxVector)v;
                    }
                }
                catch { }
            }
            // 反射 X/Y/Z
            try
            {
                PropertyInfo px = t.GetProperty("X");
                PropertyInfo py = t.GetProperty("Y");
                PropertyInfo pz = t.GetProperty("Z");
                if (px != null && py != null && pz != null)
                {
                    double x = Convert.ToDouble(px.GetValue(pp, null));
                    double y = Convert.ToDouble(py.GetValue(pp, null));
                    double z = Convert.ToDouble(pz.GetValue(pp, null));
                    return new TxVector(x, y, z);
                }
            }
            catch { }
            return new TxVector(0, 0, 0);
        }

        // ============ 反射参数化曲线 ============

        private static List<TxVector> TryReadAs1DGeometry(ITxObject obj, Action<string> log)
        {
            try
            {
                Type t = obj.GetType();
                MethodInfo mi = t.GetMethod("GetPointByParameter", new[] { typeof(double) });
                if (mi == null) return null;

                var list = new List<TxVector>();
                for (int i = 0; i < CurveSampleCount; i++)
                {
                    double tt = (double)i / (CurveSampleCount - 1);
                    try
                    {
                        object res = mi.Invoke(obj, new object[] { tt });
                        TxVector v;
                        if (TryGetPointFromAny(res, out v)) list.Add(v);
                    }
                    catch { }
                }
                if (list.Count >= 2)
                {
                    List<TxVector> simplified = SimplifyIfCollinear(list, 1.0);
                    if (simplified.Count < list.Count)
                        log?.Invoke("[Reader] 路径2 GetPointByParameter OK, 采样 " + list.Count + " 点 → 共线简化为 " + simplified.Count + " 点(直线段)");
                    else
                        log?.Invoke("[Reader] 路径2 GetPointByParameter OK, 采样 " + list.Count + " 点(曲线)");
                    return simplified;
                }
            }
            catch (Exception ex) { log?.Invoke("[Reader] 路径2 异常: " + ex.Message); }
            return null;
        }

        private static List<TxVector> SimplifyIfCollinear(List<TxVector> pts, double tolerance)
        {
            if (pts.Count <= 2) return pts;
            TxVector p0 = pts[0];
            TxVector pn = pts[pts.Count - 1];
            double dx = pn.X - p0.X, dy = pn.Y - p0.Y, dz = pn.Z - p0.Z;
            double len = Math.Sqrt(dx * dx + dy * dy + dz * dz);
            if (len < 1e-6) return pts;
            for (int i = 1; i < pts.Count - 1; i++)
            {
                double ax = pts[i].X - p0.X, ay = pts[i].Y - p0.Y, az = pts[i].Z - p0.Z;
                double cx = ay * dz - az * dy;
                double cy = az * dx - ax * dz;
                double cz = ax * dy - ay * dx;
                double d = Math.Sqrt(cx * cx + cy * cy + cz * cz) / len;
                if (d > tolerance) return pts;
            }
            return new List<TxVector> { p0, pn };
        }

        // ============ 直线端点 ============

        private static List<TxVector> TryReadAsLineEndpoints(ITxObject obj, Action<string> log)
        {
            string[][] pairs = {
                new[] { "StartPoint", "EndPoint" },
                new[] { "Start", "End" },
                new[] { "P1", "P2" },
                new[] { "FirstPoint", "LastPoint" }
            };
            Type t = obj.GetType();
            foreach (var pair in pairs)
            {
                try
                {
                    PropertyInfo ps = t.GetProperty(pair[0]);
                    PropertyInfo pe = t.GetProperty(pair[1]);
                    if (ps == null || pe == null) continue;
                    object vs = ps.GetValue(obj, null);
                    object ve = pe.GetValue(obj, null);
                    TxVector s, e;
                    if (TryGetPointFromAny(vs, out s) && TryGetPointFromAny(ve, out e))
                    {
                        log?.Invoke("[Reader] 路径3 端点 OK (" + pair[0] + "/" + pair[1] + ")");
                        return new List<TxVector> { s, e };
                    }
                }
                catch { }
            }
            return null;
        }

        // ============ 点解析 ============

        private static bool TryGetPointFromAny(object item, out TxVector v)
        {
            v = new TxVector(0, 0, 0);
            if (item == null) return false;
            if (item is TxVector) { v = (TxVector)item; return true; }

            Type t = item.GetType();
            string[] subNames = { "Vector", "Position", "Point", "Location", "Value" };
            foreach (string n in subNames)
            {
                try
                {
                    PropertyInfo pi = t.GetProperty(n);
                    if (pi == null) continue;
                    object sub = pi.GetValue(item, null);
                    if (sub == null) continue;
                    if (sub is TxVector) { v = (TxVector)sub; return true; }
                    if (TryReadXYZ(sub, out v)) return true;
                }
                catch { }
            }
            return TryReadXYZ(item, out v);
        }

        private static bool TryReadXYZ(object item, out TxVector v)
        {
            v = new TxVector(0, 0, 0);
            try
            {
                Type t = item.GetType();
                PropertyInfo px = t.GetProperty("X");
                PropertyInfo py = t.GetProperty("Y");
                PropertyInfo pz = t.GetProperty("Z");
                if (px != null && py != null && pz != null)
                {
                    double x = Convert.ToDouble(px.GetValue(item, null));
                    double y = Convert.ToDouble(py.GetValue(item, null));
                    double z = Convert.ToDouble(pz.GetValue(item, null));
                    v = new TxVector(x, y, z);
                    return true;
                }
            }
            catch { }
            return false;
        }

        // ============ 反射工具 ============

        /// <summary>
        /// 在对象上找指定名字、单参引用类型的方法,传 null 调用,返回 IEnumerable。
        /// 沿用 TxTools.LineToSolid 同款实现。
        /// </summary>
        private static System.Collections.IEnumerable TryInvokeWithNullArg(object container, string methodName)
        {
            try
            {
                Type t = container.GetType();
                foreach (var mi in t.GetMethods())
                {
                    if (mi.Name != methodName) continue;
                    var pars = mi.GetParameters();
                    if (pars.Length != 1) continue;
                    if (pars[0].ParameterType.IsValueType) continue;
                    object r = mi.Invoke(container, new object[] { null });
                    return r as System.Collections.IEnumerable;
                }
            }
            catch { }
            return null;
        }

        private static System.Collections.IEnumerable TryGetByEnumerableProperty(object container,
            Action<string> log, int depth)
        {
            Type t = container.GetType();
            string[] propNames = { "Components", "Children", "SubObjects", "Items", "Members", "Descendants" };
            foreach (string name in propNames)
            {
                try
                {
                    PropertyInfo pi = t.GetProperty(name);
                    if (pi == null) continue;
                    var val = pi.GetValue(container, null) as System.Collections.IEnumerable;
                    if (val != null) return val;
                }
                catch { }
            }
            return null;
        }

        private static string TrySafeName(ITxObject obj)
        {
            try { return obj.Name ?? "(unnamed)"; }
            catch { return "(unknown)"; }
        }
    }
}