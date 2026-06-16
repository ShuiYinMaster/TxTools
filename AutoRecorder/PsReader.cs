using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using Tecnomatix.Engineering;

namespace TxTools.AutoRecorder
{
    /// <summary>
    /// 集中所有 PS SDK 调用。其它类不直接接触 Tecnomatix.Engineering 命名空间细节。
    /// 出于 PS 版本兼容考虑，大量使用 dynamic + try/catch 防御。
    /// </summary>
    internal static class PsReader
    {
        // ============================================================
        // 1. 拿主视口
        // ============================================================
        public static TxGraphicViewer GetGraphicViewer()
        {
            try { return TxApplication.ViewersManager.GraphicViewer; }
            catch { return null; }
        }

        // ============================================================
        // 2. 收集操作的相关可视对象（用于聚焦视角 / 估算 bbox）
        //    包含：操作本体 + 所有子位置点 + 各点 AssignedParts + Robot + Tool/Gun
        // ============================================================
        public static List<ITxObject> CollectOperationObjects(ITxObject op)
        {
            var seen = new HashSet<ITxObject>();
            var result = new List<ITxObject>();
            if (op == null) return result;

            AddUnique(result, seen, op);

            // 2.1 操作下的所有后代（weld points / vias / sub-operations）
            var descendants = CollectDescendants(op);
            foreach (var d in descendants) AddUnique(result, seen, d);

            // 2.2 每个子节点的 AssignedParts（焊接零件等）
            foreach (var node in new List<ITxObject>(result))
            {
                foreach (var part in TryGetAssignedParts(node))
                    AddUnique(result, seen, part);
            }

            // 2.3 操作绑定的 Robot
            AddUnique(result, seen, TryGetMember(op, "Robot") as ITxObject);

            // 2.4 操作绑定的 Tool / Gun
            AddUnique(result, seen, TryGetMember(op, "Tool") as ITxObject);
            AddUnique(result, seen, TryGetMember(op, "Gun") as ITxObject);

            return result;
        }

        private static List<ITxObject> CollectDescendants(ITxObject parent)
        {
            var list = new List<ITxObject>();
            if (parent == null) return list;

            // 路径 A：GetAllDescendants(null) —— 最常见
            try
            {
                dynamic dp = parent;
                var all = dp.GetAllDescendants(null) as IEnumerable;
                if (all != null)
                {
                    foreach (var item in all)
                    {
                        var t = item as ITxObject;
                        if (t != null) list.Add(t);
                    }
                    if (list.Count > 0) return list;
                }
            }
            catch { }

            // 路径 B：GetDirectDescendants 递归
            try
            {
                dynamic dp = parent;
                var direct = dp.GetDirectDescendants(null) as IEnumerable;
                if (direct != null)
                {
                    foreach (var item in direct)
                    {
                        var t = item as ITxObject;
                        if (t != null)
                        {
                            list.Add(t);
                            list.AddRange(CollectDescendants(t));
                        }
                    }
                }
            }
            catch { }

            return list;
        }

        private static IEnumerable<ITxObject> TryGetAssignedParts(ITxObject node)
        {
            var result = new List<ITxObject>();
            try
            {
                dynamic d = node;
                var parts = d.AssignedParts as IEnumerable;
                if (parts != null)
                {
                    foreach (var p in parts)
                    {
                        var t = p as ITxObject;
                        if (t != null) result.Add(t);
                    }
                }
            }
            catch { }
            return result;
        }

        private static object TryGetMember(ITxObject obj, string name)
        {
            if (obj == null) return null;
            try
            {
                var prop = obj.GetType().GetProperty(name,
                    BindingFlags.Public | BindingFlags.Instance);
                if (prop != null && prop.CanRead)
                    return prop.GetValue(obj, null);
            }
            catch { }
            return null;
        }

        private static void AddUnique(List<ITxObject> list,
            HashSet<ITxObject> seen, ITxObject obj)
        {
            if (obj == null) return;
            if (seen.Add(obj)) list.Add(obj);
        }

        // ============================================================
        // 3. 聚焦视角（PS 内部 Zoom to Selection 命令）
        // ============================================================
        public static bool ZoomToObjects(IList<ITxObject> objects)
        {
            if (objects == null || objects.Count == 0) return false;

            TxObjectList savedSel = null;
            try { savedSel = TxApplication.ActiveSelection.GetItems(); }
            catch { }

            bool ok = false;
            try
            {
                var list = new TxObjectList(objects.Count);
                foreach (var o in objects)
                    if (o != null) list.Add(o);

                TxApplication.ActiveSelection.SetItems(list);
                TxApplication.CommandsManager.ExecuteCommand("GraphicViewer.ZoomToSelection");
                ok = true;
            }
            catch
            {
                // 兜底：直接 ZoomToFit 全场景
                try
                {
                    var v = GetGraphicViewer();
                    if (v != null) { v.ZoomToFit(); ok = true; }
                }
                catch { }
            }
            finally
            {
                // 恢复原选中
                try
                {
                    if (savedSel != null)
                        TxApplication.ActiveSelection.SetItems(savedSel);
                    else
                        TxApplication.ActiveSelection.SetItems(new TxObjectList(0));
                }
                catch { }
                try { TxApplication.RefreshDisplay(); } catch { }
            }
            return ok;
        }

        // ============================================================
        // 4. 获取系统支持的视频编解码器
        // ============================================================
        public static List<TxVideoCodec> GetSupportedCodecs()
        {
            var result = new List<TxVideoCodec>();
            try
            {
                var supported = TxMovieRecordingSettings.GetSupportedCodecs();
                if (supported != null)
                {
                    foreach (var c in supported) result.Add(c);
                }
            }
            catch { }
            if (result.Count == 0) result.Add(TxVideoCodec.MPEG4);
            return result;
        }

        // ============================================================
        // 5. 创建 TxViewerRecordingSettings
        //    构造签名：TxViewerRecordingSettings(string filePath,
        //                  ITxViewerRecordingSource viewerSource,
        //                  uint width, uint height)
        //    注意：width / height 必须是 4 的倍数且在合法范围内（构造函数会校验）
        // ============================================================
        public static TxViewerRecordingSettings CreateViewerSettings(
            TxGraphicViewer viewer, string filePath, uint width, uint height)
        {
            if (viewer == null) throw new ArgumentNullException("viewer");
            if (string.IsNullOrEmpty(filePath))
                throw new ArgumentNullException("filePath");

            // 提前规整到 4 的倍数，避免构造函数抛 TxArgumentOutOfRangeException
            width = RoundUpTo4(width);
            height = RoundUpTo4(height);

            // TxGraphicViewer 隐式实现了 ITxViewerRecordingSource
            return new TxViewerRecordingSettings(filePath, viewer, width, height);
        }

        private static uint RoundUpTo4(uint v)
        {
            if (v < 4) return 4;
            return (v / 4) * 4;
        }

        // ============================================================
        // 5c. 读取视口当前像素尺寸（用作录像默认分辨率）
        //     防御性走多个候选路径
        // ============================================================
        public static bool TryGetViewerSize(TxGraphicViewer viewer, out int width, out int height)
        {
            width = 0; height = 0;
            if (viewer == null) return false;

            // 候选 1: Bounds.Width / Bounds.Height
            try
            {
                dynamic d = viewer;
                var b = d.Bounds;
                int w = Convert.ToInt32(b.Width);
                int h = Convert.ToInt32(b.Height);
                if (w > 0 && h > 0) { width = w; height = h; return true; }
            }
            catch { }

            // 候选 2: Size.Width / Size.Height
            try
            {
                dynamic d = viewer;
                var s = d.Size;
                int w = Convert.ToInt32(s.Width);
                int h = Convert.ToInt32(s.Height);
                if (w > 0 && h > 0) { width = w; height = h; return true; }
            }
            catch { }

            // 候选 3: Width / Height 直接属性
            try
            {
                dynamic d = viewer;
                int w = Convert.ToInt32(d.Width);
                int h = Convert.ToInt32(d.Height);
                if (w > 0 && h > 0) { width = w; height = h; return true; }
            }
            catch { }

            return false;
        }

        // ============================================================
        // 5b. 创建 TxMovieRecorder —— 反射找构造，失败时列出所有可用签名
        //     （SDK 没暴露文档化的构造签名，先用反射试探）
        // ============================================================
        public static TxMovieRecorder CreateMovieRecorder(TxGraphicViewer viewer)
        {
            var t = typeof(TxMovieRecorder);
            var ctors = t.GetConstructors(
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);

            // 优先 1: 带 TxGraphicViewer 参数的构造
            foreach (var ctor in ctors)
            {
                var prms = ctor.GetParameters();
                if (prms.Length == 1 && viewer != null
                    && prms[0].ParameterType.IsInstanceOfType(viewer))
                {
                    try { return (TxMovieRecorder)ctor.Invoke(new object[] { viewer }); }
                    catch { }
                }
            }

            // 优先 2: 接受 TxApplication 或其它单参的（试一遍）
            foreach (var ctor in ctors)
            {
                var prms = ctor.GetParameters();
                if (prms.Length == 1)
                {
                    try { return (TxMovieRecorder)ctor.Invoke(new object[] { viewer }); }
                    catch { }
                    try { return (TxMovieRecorder)ctor.Invoke(new object[] { null }); }
                    catch { }
                }
            }

            // 优先 3: 无参
            foreach (var ctor in ctors)
            {
                if (ctor.GetParameters().Length == 0)
                {
                    try { return (TxMovieRecorder)ctor.Invoke(null); }
                    catch { }
                }
            }

            // 优先 4: 看是否有静态工厂方法
            foreach (var m in t.GetMethods(BindingFlags.Public | BindingFlags.Static))
            {
                if (m.ReturnType == t && m.GetParameters().Length <= 1)
                {
                    try
                    {
                        var p = m.GetParameters();
                        return (TxMovieRecorder)m.Invoke(null,
                            p.Length == 0 ? null : new object[] { viewer });
                    }
                    catch { }
                }
            }

            // 全部失败 —— 列出所有候选签名给开发者看
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("无法实例化 TxMovieRecorder。可用构造签名：");
            foreach (var ctor in ctors)
            {
                sb.Append("  TxMovieRecorder(");
                var prms = ctor.GetParameters();
                for (int i = 0; i < prms.Length; i++)
                {
                    if (i > 0) sb.Append(", ");
                    sb.Append(prms[i].ParameterType.Name).Append(" ").Append(prms[i].Name);
                }
                sb.AppendLine(")");
            }
            sb.AppendLine("可用静态方法（返回 TxMovieRecorder）：");
            foreach (var m in t.GetMethods(BindingFlags.Public | BindingFlags.Static))
            {
                if (m.ReturnType == t)
                    sb.AppendLine("  static " + m.Name + "(...)");
            }
            throw new InvalidOperationException(sb.ToString());
        }

        // ============================================================
        // 6. 操作持续时间（仅用于 UI 显示估计，可能拿不到）
        // ============================================================
        public static double? GetOperationDuration(ITxObject op)
        {
            if (op == null) return null;
            try
            {
                dynamic d = op;
                return Convert.ToDouble(d.Duration);
            }
            catch { return null; }
        }

        // ============================================================
        // 7. 取对象 Name（防御性，不挑类型）
        // ============================================================
        public static string GetObjectName(ITxObject obj)
        {
            if (obj == null) return "[null]";
            try
            {
                dynamic d = obj;
                var n = d.Name as string;
                if (!string.IsNullOrEmpty(n)) return n;
            }
            catch { }
            try { return obj.ToString(); }
            catch { return "[Unknown]"; }
        }

        // ============================================================
        // 8. 相机 I/O —— 读 / 写当前视口相机
        //    TxGraphicViewer 隐式实现 ITxGraphicDisplayer
        // ============================================================
        public static TxCamera GetCurrentCamera(TxGraphicViewer viewer)
        {
            if (viewer == null) return null;
            try
            {
                var d = (ITxGraphicDisplayer)viewer;
                return d.CurrentCamera;
            }
            catch { }
            // 兜底：dynamic 访问
            try
            {
                dynamic d = viewer;
                return d.CurrentCamera as TxCamera;
            }
            catch { return null; }
        }

        public static bool SetCurrentCamera(TxGraphicViewer viewer, TxCamera camera)
        {
            if (viewer == null || camera == null) return false;
            try
            {
                var d = (ITxGraphicDisplayer)viewer;
                d.CurrentCamera = camera;
                try { TxApplication.RefreshDisplay(); } catch { }
                return true;
            }
            catch { }
            try
            {
                dynamic d = viewer;
                d.CurrentCamera = camera;
                try { TxApplication.RefreshDisplay(); } catch { }
                return true;
            }
            catch { return false; }
        }

        // ============================================================
        // 9. 自动计算操作的"最佳"相机
        //    策略：相机放在工件几何中心的"机器人对面"方向，沿 R→T 延长线退到 bbox*1.8 处，
        //    略抬高（bbox*0.7），以世界 Z 为 up。
        //    返回 null 表示无法自动计算（机器人或焊点位置取不到），调用方应回退到 ZoomToSelection。
        // ============================================================
        public static TxCamera ComputeOptimalCamera(ITxObject op)
        {
            if (op == null) return null;
            try
            {
                // 1. 机器人基座位置
                var robot = TryGetMember(op, "Robot") as ITxObject;
                if (robot == null) return null;
                double Rx, Ry, Rz;
                if (!TryGetAbsolutePosition(robot, out Rx, out Ry, out Rz)) return null;

                // 2. 焊点位置集合 —— 用 GetOperationLocations 兼容焊接操作的 IEnumerable 路径
                var locItems = GetOperationLocations(op);
                var positions = new List<double[]>();
                foreach (var d in locItems)
                {
                    double x, y, z;
                    if (TryGetLocationPosition(d, out x, out y, out z))
                        positions.Add(new[] { x, y, z });
                }
                if (positions.Count == 0) return null;

                // 几何中心 + bbox
                double Tx = 0, Ty = 0, Tz = 0;
                double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
                double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;
                foreach (var p in positions)
                {
                    Tx += p[0]; Ty += p[1]; Tz += p[2];
                    if (p[0] < minX) minX = p[0]; if (p[0] > maxX) maxX = p[0];
                    if (p[1] < minY) minY = p[1]; if (p[1] > maxY) maxY = p[1];
                    if (p[2] < minZ) minZ = p[2]; if (p[2] > maxZ) maxZ = p[2];
                }
                Tx /= positions.Count;
                Ty /= positions.Count;
                Tz /= positions.Count;

                double bboxDiag = Math.Sqrt(
                    (maxX - minX) * (maxX - minX) +
                    (maxY - minY) * (maxY - minY) +
                    (maxZ - minZ) * (maxZ - minZ));
                if (bboxDiag < 500) bboxDiag = 500;  // 单点焊或紧凑布置，保底 500mm

                // 3. R → T 在水平面上的方向
                double dx = Tx - Rx;
                double dy = Ty - Ry;
                double horizDist = Math.Sqrt(dx * dx + dy * dy);
                if (horizDist < 1.0)
                {
                    // 机器人正好在工件正上 / 正下方，退化用 +X 方向
                    dx = 1; dy = 0; horizDist = 1;
                }
                double ux = dx / horizDist;
                double uy = dy / horizDist;

                // 4. 相机位置：工件后方（远离机器人）+ 抬高
                double dist = bboxDiag * 1.8;
                double height = bboxDiag * 0.7;

                var refPoint = new TxVector(Tx, Ty, Tz);
                var camPos = new TxVector(Tx + ux * dist, Ty + uy * dist, Tz + height);
                var upVec = new TxVector(0, 0, 1);
                return new TxCamera(refPoint, camPos, upVec);
            }
            catch { return null; }
        }
        // ============================================================
        // 10. 几何辅助：取对象的世界坐标位置
        // ============================================================
        private static bool TryGetAbsolutePosition(ITxObject obj,
            out double x, out double y, out double z)
        {
            x = y = z = 0;
            if (obj == null) return false;
            try
            {
                dynamic d = obj;
                var loc = d.AbsoluteLocation;   // TxTransformation

                // 路径 1：loc.Translation 返回 TxVector
                try
                {
                    var t = loc.Translation;
                    x = Convert.ToDouble(t.X);
                    y = Convert.ToDouble(t.Y);
                    z = Convert.ToDouble(t.Z);
                    return true;
                }
                catch { }

                // 路径 2：loc.X / loc.Y / loc.Z 直接挂在 transformation 上
                try
                {
                    x = Convert.ToDouble(loc.X);
                    y = Convert.ToDouble(loc.Y);
                    z = Convert.ToDouble(loc.Z);
                    return true;
                }
                catch { }
            }
            catch { }
            return false;
        }

        // ============================================================
        // 11. 取操作的 location 列表（按执行顺序）
        //     参考 ExportGun PsReader 的三层路径：
        //     L1: op 自身 IEnumerable（焊接操作直接含焊点 PM items）
        //     L2: op.Locations / op.LocationList / op.RoboticLocations（路径型操作）
        //     L3: 扫子孙节点（兜底）
        // ============================================================
        public static List<object> GetOperationLocations(ITxObject op)
        {
            var result = new List<object>();
            if (op == null) return result;

            // L1：op 自身可枚举（TxWeldOperation 和很多 robotic op 都实现了 IEnumerable）
            try
            {
                var src = op as System.Collections.IEnumerable;
                if (src != null)
                {
                    foreach (var item in src)
                    {
                        if (item == null) continue;
                        // 必须能取到位置坐标，过滤掉无关 PM 子项
                        double x, y, z;
                        if (TryGetLocationPosition(item, out x, out y, out z))
                            result.Add(item);
                    }
                    if (result.Count > 0) return result;
                }
            }
            catch { }

            // L2：尝试已知的 location 集合属性
            foreach (var prop in new[] { "Locations", "LocationList", "RoboticLocations" })
            {
                try
                {
                    dynamic d = op;
                    object raw = d.GetType().GetProperty(prop)?.GetValue(d);
                    var seq = raw as System.Collections.IEnumerable;
                    if (seq == null) continue;
                    foreach (var l in seq)
                    {
                        if (l == null) continue;
                        double x, y, z;
                        if (TryGetLocationPosition(l, out x, out y, out z))
                            result.Add(l);
                    }
                    if (result.Count > 0) return result;
                }
                catch { }
            }

            // L3：兜底 —— 扫子孙节点，按"有坐标 + 不是 operation"过滤
            var descendants = CollectDescendants(op);
            foreach (var d in descendants)
            {
                if (d == null) continue;
                if (ReferenceEquals(d, op)) continue;
                string tn = ""; try { tn = d.GetType().Name; } catch { }
                if (tn.IndexOf("Operation", StringComparison.OrdinalIgnoreCase) >= 0)
                    continue;
                double x, y, z;
                if (!TryGetLocationPosition(d, out x, out y, out z)) continue;
                result.Add(d);
            }
            return result;
        }

        /// <summary>
        /// 从任意 location-like 对象上提取友好显示名（支持 PM items / ITxObject）
        /// </summary>
        public static string GetLocationName(object item)
        {
            if (item == null) return "[null]";
            try
            {
                dynamic d = item;
                var n = d.Name as string;
                if (!string.IsNullOrEmpty(n)) return n;
            }
            catch { }
            try
            {
                dynamic d = item;
                var n = d.DisplayName as string;
                if (!string.IsNullOrEmpty(n)) return n;
            }
            catch { }
            // 焊点 PM item 可能要再翻一层
            try
            {
                dynamic d = item;
                object mfg = d.MfgFeature;
                if (mfg != null)
                {
                    dynamic dm = mfg;
                    var n = dm.Name as string;
                    if (!string.IsNullOrEmpty(n)) return n;
                }
            }
            catch { }
            try { return item.ToString(); }
            catch { return "[Unknown]"; }
        }

        /// <summary>
        /// 从任意 location-like 对象（PM item / TxWeldPoint / ITxLocatableObject）取世界坐标
        /// 路径参考 ExportGun PsReader.GetTxFromPm
        /// </summary>
        private static bool TryGetLocationPosition(object item, out double x, out double y, out double z)
        {
            x = y = z = 0;
            if (item == null) return false;

            // 路径 A：item.AbsoluteLocation
            try
            {
                dynamic d = item;
                var loc = d.AbsoluteLocation;
                if (loc != null && ExtractTranslation(loc, out x, out y, out z)) return true;
            }
            catch { }
            // 路径 B：item.Location
            try
            {
                dynamic d = item;
                var loc = d.Location;
                if (loc != null && ExtractTranslation(loc, out x, out y, out z)) return true;
            }
            catch { }
            // 路径 C：item.LocationData.Frame
            try
            {
                dynamic d = item;
                var ld = d.LocationData;
                if (ld != null)
                {
                    var fr = ld.Frame;
                    if (fr != null && ExtractTranslation(fr, out x, out y, out z)) return true;
                }
            }
            catch { }
            // 路径 D：item.MfgFeature.AbsoluteLocation（PM item → 焊点）
            try
            {
                dynamic d = item;
                var mfg = d.MfgFeature;
                if (mfg != null)
                {
                    dynamic dm = mfg;
                    var loc = dm.AbsoluteLocation;
                    if (loc != null && ExtractTranslation(loc, out x, out y, out z)) return true;
                }
            }
            catch { }
            return false;
        }

        private static bool ExtractTranslation(object loc, out double x, out double y, out double z)
        {
            x = y = z = 0;
            try
            {
                dynamic d = loc;
                try
                {
                    var t = d.Translation;
                    x = Convert.ToDouble(t.X);
                    y = Convert.ToDouble(t.Y);
                    z = Convert.ToDouble(t.Z);
                    return true;
                }
                catch { }
                try
                {
                    x = Convert.ToDouble(d.X);
                    y = Convert.ToDouble(d.Y);
                    z = Convert.ToDouble(d.Z);
                    return true;
                }
                catch { }
            }
            catch { }
            return false;
        }
    }
}