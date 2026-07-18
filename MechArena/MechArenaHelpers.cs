using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.DataTypes;

namespace TxTools.MechArena
{
    // =========================================================================
    //  几何工厂 —— 支持透明度
    // =========================================================================
    public static class MechArenaGeometry
    {
        /// <summary>
        /// 建立方块组件。
        /// </summary>
        /// <param name="transparency">0.0 = 完全不透明，1.0 = 完全透明。默认 0。</param>
        public static TxComponent CreateBox(
            string name,
            double sx, double sy, double sz,
            TxColor color = null,
            double transparency = 0.0)
        {
            var doc = TxApplication.ActiveDocument;
            if (doc == null) throw new Exception("MechArenaGeometry: 无 ActiveDocument");

            var res = doc.PhysicalRoot.CreateResource(new TxResourceCreationData(name));
            var comp = res as TxComponent;
            if (comp == null)
                throw new Exception("CreateResource 未返回 TxComponent: " + name);
            if (!comp.CanOpenForModeling)
                throw new Exception("组件无法进入建模: " + name);

            comp.SetModelingScope();
            try
            {
                var absLoc = new TxTransformation();
                absLoc.Translation = new TxVector(0, 0, -sz / 2.0);
                var edgeSizes = new TxVector(sx, sy, sz);
                var offset = new TxVector(0, 0, 0);
                var box = new TxBoxCreationData(name, absLoc, edgeSizes, offset);
                comp.CreateSolidBox(box);
            }
            finally
            {
                try
                {
                    var mi = comp.GetType().GetMethod("EndModelingScope",
                        BindingFlags.Public | BindingFlags.Instance);
                    mi?.Invoke(comp, null);
                }
                catch { }
            }

            if (color != null) TrySetColor(comp, color);
            if (transparency > 0.001) TrySetTransparency(comp, transparency);
            return comp;
        }

        private static void TrySetColor(TxComponent comp, TxColor c)
        {
            try
            {
                var pi = comp.GetType().GetProperty("Color",
                    BindingFlags.Public | BindingFlags.Instance);
                if (pi != null && pi.CanWrite) { pi.SetValue(comp, c, null); return; }
            }
            catch { }
            try
            {
                var mi = comp.GetType().GetMethod("SetColor",
                    BindingFlags.Public | BindingFlags.Instance,
                    null, new[] { typeof(TxColor) }, null);
                if (mi != null) mi.Invoke(comp, new object[] { c });
            }
            catch { }
        }

        /// <summary>
        /// 透明度 API 在不同 PS 版本命名/类型都可能不同，做多路径反射。
        /// t: 0..1，0 完全不透明，1 完全透明。
        /// </summary>
        private static void TrySetTransparency(TxComponent comp, double t)
        {
            if (t < 0) t = 0;
            if (t > 1) t = 1;

            var candidates = new (string name, bool invert)[] {
                ("Transparency", false),
                ("Opacity", true),
            };
            foreach (var cand in candidates)
            {
                try
                {
                    var pi = comp.GetType().GetProperty(cand.name,
                        BindingFlags.Public | BindingFlags.Instance);
                    if (pi != null && pi.CanWrite)
                    {
                        double v = cand.invert ? 1.0 - t : t;
                        var pt = pi.PropertyType;
                        object arg = null;
                        if (pt == typeof(double)) arg = v;
                        else if (pt == typeof(float)) arg = (float)v;
                        else if (pt == typeof(int)) arg = (int)Math.Round(v * 100);
                        if (arg != null)
                        {
                            pi.SetValue(comp, arg, null);
                            return;
                        }
                    }
                }
                catch { }
            }

            try
            {
                var mi = comp.GetType().GetMethod("SetTransparency",
                    BindingFlags.Public | BindingFlags.Instance);
                if (mi != null)
                {
                    var pt = mi.GetParameters()[0].ParameterType;
                    object arg;
                    if (pt == typeof(double)) arg = t;
                    else if (pt == typeof(float)) arg = (float)t;
                    else if (pt == typeof(int)) arg = (int)Math.Round(t * 100);
                    else arg = t;
                    mi.Invoke(comp, new[] { arg });
                    return;
                }
            }
            catch { }

            try { ((dynamic)comp).Transparency = t; return; } catch { }
            try { ((dynamic)comp).Transparency = (int)Math.Round(t * 100); return; } catch { }
        }
    }

    // =========================================================================
    //  机器人关节赋值 —— 三重防御
    // =========================================================================
    public static class MechArenaRobotHelper
    {
        public static void SetJointValues(TxRobot robot, double[] values)
        {
            if (robot == null || values == null || values.Length == 0) return;

            // 优先：TxPoseData 路径（CollisionWorld.ApplyJoints 已验证，最可靠）
            if (TrySetViaPoseData(robot, values)) return;

            // 回退1：逐个关节写 CurrentValue（需拿到关节对象）
            IList joints = null;
            try { joints = robot.Joints; } catch { }
            if (joints == null)
            {
                try { joints = ((dynamic)robot).Joints as IList; } catch { }
            }
            if (joints != null && joints.Count > 0)
            {
                int n = Math.Min(values.Length, joints.Count);
                if (TrySetPerJoint(joints, values, n)) return;
            }

            // 回退2：dynamic CurrentJointValues = ArrayList
            TrySetViaArrayList(robot, values, values.Length);
        }

        private static bool TrySetPerJoint(IList joints, double[] values, int n)
        {
            try
            {
                for (int i = 0; i < n; i++)
                {
                    dynamic j = joints[i];
                    j.CurrentValue = values[i];
                }
                return true;
            }
            catch { return false; }
        }

        private static bool TrySetViaArrayList(TxRobot robot, double[] values, int n)
        {
            try
            {
                var arr = new ArrayList();
                for (int i = 0; i < n; i++) arr.Add(values[i]);
                dynamic r = robot;
                r.CurrentJointValues = arr;
                return true;
            }
            catch { return false; }
        }

        private static bool TrySetViaPoseData(TxRobot robot, double[] values)
        {
            try
            {
                var al = new ArrayList();
                foreach (double v in values) al.Add(v);
                var pd = new TxPoseData();
                pd.JointValues = al;
                robot.CurrentPose = pd;
                return true;
            }
            catch { }
            return false;
        }
    }

    // =========================================================================
    //  可见性 helper
    // =========================================================================
    public static class MechArenaVisibility
    {
        public static void TrySetVisible(ITxObject obj, bool visible)
        {
            if (obj == null) return;

            var disp = obj as ITxDisplayableObject;
            if (disp != null)
            {
                try
                {
                    if (visible) disp.Display();
                    else disp.Blank();
                    return;
                }
                catch { }
            }

            try
            {
                var mi = obj.GetType().GetMethod("SetVisibility",
                    BindingFlags.Public | BindingFlags.Instance,
                    null, new[] { typeof(bool) }, null);
                if (mi != null) { mi.Invoke(obj, new object[] { visible }); return; }
            }
            catch { }

            try { ((dynamic)obj).Visible = visible; return; } catch { }
            try { ((dynamic)obj).IsVisible = visible; } catch { }
        }
    }

    // =========================================================================
    //  相机控制 — 参考 AutoRecorder/PsReader（ITxGraphicDisplayer.CurrentCamera）
    // =========================================================================
    public static class MechArenaCamera
    {
        /// <summary>游戏开始前的相机快照，Dispose 时还原。</summary>
        private static TxCamera _savedCamera;

        /// <summary>获取主视口。返回 null 表示 PS 尚未完全初始化。</summary>
        public static TxGraphicViewer GetViewer()
        {
            try { return TxApplication.ViewersManager.GraphicViewer; }
            catch { return null; }
        }

        /// <summary>保存当前相机（游戏开始时调用一次）。</summary>
        public static void SaveCurrentCamera()
        {
            try
            {
                var viewer = GetViewer();
                if (viewer == null) { _savedCamera = null; return; }
                _savedCamera = ((ITxGraphicDisplayer)viewer).CurrentCamera;
            }
            catch { _savedCamera = null; }
        }

        /// <summary>还原 SaveCurrentCamera 保存的相机（游戏结束/Dispose 时调用）。</summary>
        public static void RestoreSavedCamera()
        {
            try
            {
                if (_savedCamera == null) return;
                var viewer = GetViewer();
                if (viewer == null) return;
                ((ITxGraphicDisplayer)viewer).CurrentCamera = _savedCamera;
                try { TxApplication.RefreshDisplay(); } catch { }
            }
            catch { }
            finally { _savedCamera = null; }
        }

        /// <summary>
        /// 任意视点相机：lookAt = 焦点（玩家），camPos = 相机位置。
        /// </summary>
        public static bool SetLookAtCamera(TxVector lookAt, TxVector camPos, bool refresh = true)
        {
            try
            {
                var viewer = GetViewer();
                if (viewer == null) return false;

                var upVec = new TxVector(0, 0, 1);
                var camera = new TxCamera(lookAt, camPos, upVec);
                ((ITxGraphicDisplayer)viewer).CurrentCamera = camera;
                if (refresh)
                {
                    try { TxApplication.RefreshDisplay(); } catch { }
                }
                return true;
            }
            catch { return false; }
        }

        /// <summary>ZoomToFit 全局场景（兜底方案）。</summary>
        public static bool ZoomToFit()
        {
            try
            {
                var viewer = GetViewer();
                if (viewer == null) return false;
                viewer.ZoomToFit();
                try { TxApplication.RefreshDisplay(); } catch { }
                return true;
            }
            catch { return false; }
        }
    }

    // =========================================================================
    //  干涉集碰撞服务
    //  创建路径为已验证的强类型 API（SnakeGame 确认）：
    //    TxApplication.ActiveDocument.CollisionRoot : TxCollisionRoot
    //    root.CreateCollisionPair(new TxCollisionPairCreationData{FirstList,SecondList})
    //    pair.FirstList/SecondList.Add/Remove、pair.Delete() 强类型直接可用
    //    root.CheckCollisions 手动置 true，清理时恢复原值
    //  查询路径（GetCollidingObjects / QueryParams / States）签名未验证过，
    //  用反射自适应；失败时 dump 成员清单到 %TEMP%\MechArena_CollisionDump.txt
    //  并置 QueryUsable=false，引擎自动退回 AABB 数学检测。
    // =========================================================================
    public class MechArenaCollisionService : IDisposable
    {
        public struct CollidingPair
        {
            public ITxObject A;
            public ITxObject B;
            public CollidingPair(ITxObject a, ITxObject b) { A = a; B = b; }
        }

        public bool Ready { get; private set; }
        public bool QueryUsable { get { return Ready && !_queryBroken; } }
        public string StatusText { get; private set; } = "未初始化";

        private TxCollisionRoot _root;
        private TxCollisionPair _playerPair;   // First: 玩家方块  Second: 全部机器人
        private TxCollisionPair _bulletPair;   // First: 子弹们    Second: 全部机器人（懒创建）
        private readonly List<TxRobot> _robots = new List<TxRobot>();
        private TxComponent _playerBody;

        private bool _origCheckCollisions;
        private bool _origSaved;

        // 查询反射机
        private bool _queryInit;
        private bool _queryBroken;
        private MethodInfo _queryMethod;
        private object _queryParams;

        // =====================================================================
        public bool Initialize(TxComponent playerBody, IEnumerable<TxRobot> robots)
        {
            try
            {
                _playerBody = playerBody;
                _robots.Clear();
                foreach (var r in robots) if (r != null) _robots.Add(r);

                _root = TxApplication.ActiveDocument.CollisionRoot;
                if (_root == null) { StatusText = "无 CollisionRoot"; return false; }

                try { _origCheckCollisions = _root.CheckCollisions; _origSaved = true; }
                catch { _origSaved = false; }

                if (_robots.Count > 0 && _playerBody != null)
                {
                    var fl = new TxObjectList();
                    fl.Add(_playerBody);
                    var sl = new TxObjectList();
                    foreach (var r in _robots) sl.Add(r);

                    _playerPair = _root.CreateCollisionPair(new TxCollisionPairCreationData
                    {
                        FirstList = fl,
                        SecondList = sl
                    });
                }

                try { _root.CheckCollisions = true; } catch { }

                Ready = true;
                StatusText = "干涉集OK";
                return true;
            }
            catch (Exception ex)
            {
                Ready = false;
                StatusText = "初始化失败: " + ex.Message;
                return false;
            }
        }

        /// <summary>子弹加入干涉检测（子弹干涉对懒创建）。</summary>
        public void AddBullet(TxComponent bullet)
        {
            if (!Ready || bullet == null) return;
            try
            {
                if (_bulletPair == null)
                {
                    var fl = new TxObjectList();
                    fl.Add(bullet);
                    var sl = new TxObjectList();
                    foreach (var r in _robots) sl.Add(r);
                    _bulletPair = _root.CreateCollisionPair(new TxCollisionPairCreationData
                    {
                        FirstList = fl,
                        SecondList = sl
                    });
                }
                else
                {
                    _bulletPair.FirstList.Add(bullet);
                }
            }
            catch { }
        }

        /// <summary>子弹销毁前从干涉对移除。</summary>
        public void RemoveBullet(TxComponent bullet)
        {
            if (_bulletPair == null || bullet == null) return;
            try { _bulletPair.FirstList.Remove(bullet); } catch { }
        }

        /// <summary>机器人死亡后从两个干涉对移除。</summary>
        public void RemoveRobot(TxRobot robot)
        {
            if (robot == null) return;
            try { _playerPair?.SecondList.Remove(robot); } catch { }
            try { _bulletPair?.SecondList.Remove(robot); } catch { }
        }

        // =====================================================================
        //  查询（反射自适应）
        // =====================================================================
        public List<CollidingPair> Query()
        {
            var outList = new List<CollidingPair>();
            if (!Ready) return outList;
            EnsureQueryMachinery();
            if (_queryBroken) return outList;

            try
            {
                object results = _queryMethod.Invoke(_root,
                    _queryMethod.GetParameters().Length == 1 ? new[] { _queryParams } : null);
                if (results == null) return outList;

                var states = (GetProp(results, "States") as IEnumerable)
                          ?? (GetProp(results, "CollisionStates") as IEnumerable)
                          ?? (results as IEnumerable);
                if (states == null)
                {
                    Fail("查询结果无 States 集合: " + results.GetType().FullName, results);
                    return outList;
                }

                foreach (var st in states)
                {
                    if (st == null) continue;

                    // 有 Type/State 属性且值名含 Clearance 的跳过（那是安全距离告警不是碰撞）
                    if (IsClearanceState(st)) continue;

                    var a = ExtractObj(st, "Object1", "FirstObject", "ObjectA");
                    var b = ExtractObj(st, "Object2", "SecondObject", "ObjectB");
                    if (a == null || b == null)
                    {
                        // 兜底：CollidingObjects 集合取前两个
                        var co = GetProp(st, "CollidingObjects") as IEnumerable;
                        if (co != null)
                        {
                            var two = co.Cast<object>().OfType<ITxObject>().Take(2).ToList();
                            if (two.Count == 2) { a = two[0]; b = two[1]; }
                        }
                    }
                    if (a != null && b != null)
                        outList.Add(new CollidingPair(a, b));
                }
            }
            catch (Exception ex)
            {
                Fail("查询异常: " + (ex.InnerException?.Message ?? ex.Message), null);
            }
            return outList;
        }

        private void EnsureQueryMachinery()
        {
            if (_queryInit) return;
            _queryInit = true;
            try
            {
                // 找 GetCollidingObjects（优先 1 参，其次 0 参）
                var methods = _root.GetType()
                    .GetMethods(BindingFlags.Public | BindingFlags.Instance)
                    .Where(m => m.Name == "GetCollidingObjects")
                    .OrderBy(m => Math.Abs(m.GetParameters().Length - 1))
                    .ToList();
                if (methods.Count == 0)
                {
                    Fail("TxCollisionRoot 无 GetCollidingObjects 方法", null);
                    return;
                }
                _queryMethod = methods[0];

                if (_queryMethod.GetParameters().Length == 1)
                {
                    var pType = _queryMethod.GetParameters()[0].ParameterType;
                    _queryParams = CreateQueryParams(pType);
                    if (_queryParams == null)
                    {
                        Fail("无法构造查询参数 " + pType.FullName, null);
                        return;
                    }
                }
                StatusText = "干涉集OK(查询:" + _queryMethod.GetParameters().Length + "参)";
            }
            catch (Exception ex)
            {
                Fail("查询机构建异常: " + ex.Message, null);
            }
        }

        private object CreateQueryParams(Type t)
        {
            object inst = null;
            try { inst = Activator.CreateInstance(t); } catch { }
            if (inst == null)
            {
                // 带参构造：全部给默认值（枚举→第一个值，值类型→default，引用→null）
                foreach (var ctor in t.GetConstructors().OrderBy(c => c.GetParameters().Length))
                {
                    try
                    {
                        var ps = ctor.GetParameters();
                        var args = new object[ps.Length];
                        for (int i = 0; i < ps.Length; i++)
                        {
                            var pt = ps[i].ParameterType;
                            if (pt.IsEnum)
                            {
                                var vals = Enum.GetValues(pt);
                                args[i] = vals.Length > 0 ? vals.GetValue(0) : null;
                            }
                            else if (pt.IsValueType) args[i] = Activator.CreateInstance(pt);
                            else args[i] = null;
                        }
                        inst = ctor.Invoke(args);
                        if (inst != null) break;
                    }
                    catch { }
                }
            }
            if (inst != null) TrySetQueryType(inst);
            return inst;
        }

        /// <summary>QueryType 枚举优先挑名字含 Colliding 的值。</summary>
        private static void TrySetQueryType(object qp)
        {
            try
            {
                var pi = qp.GetType().GetProperty("QueryType");
                if (pi == null || !pi.CanWrite || !pi.PropertyType.IsEnum) return;
                var names = Enum.GetNames(pi.PropertyType);
                string pick =
                    names.FirstOrDefault(n => n.IndexOf("Colliding", StringComparison.OrdinalIgnoreCase) >= 0) ??
                    names.FirstOrDefault(n => n.IndexOf("Collision", StringComparison.OrdinalIgnoreCase) >= 0) ??
                    names.FirstOrDefault(n => n.IndexOf("All", StringComparison.OrdinalIgnoreCase) >= 0) ??
                    names.FirstOrDefault();
                if (pick != null)
                    pi.SetValue(qp, Enum.Parse(pi.PropertyType, pick), null);
            }
            catch { }
        }

        private static bool IsClearanceState(object st)
        {
            foreach (var name in new[] { "Type", "State", "CollisionType" })
            {
                try
                {
                    var v = GetProp(st, name);
                    if (v != null && v.ToString()
                        .IndexOf("Clearance", StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
                catch { }
            }
            return false;
        }

        private static ITxObject ExtractObj(object st, params string[] propNames)
        {
            foreach (var n in propNames)
            {
                try
                {
                    var v = GetProp(st, n) as ITxObject;
                    if (v != null) return v;
                }
                catch { }
            }
            return null;
        }

        private static object GetProp(object o, string name)
        {
            if (o == null) return null;
            try
            {
                var pi = o.GetType().GetProperty(name,
                    BindingFlags.Public | BindingFlags.Instance);
                return pi?.GetValue(o, null);
            }
            catch { return null; }
        }

        // =====================================================================
        //  失败诊断：dump 成员清单到临时文件，供下一轮迭代粘回
        // =====================================================================
        private void Fail(string reason, object extraObj)
        {
            _queryBroken = true;
            string path = "";
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("MechArena 干涉集查询诊断  " + DateTime.Now);
                sb.AppendLine("原因: " + reason);
                sb.AppendLine();
                DumpMembers(sb, _root, "TxCollisionRoot 实例");
                DumpMembers(sb, _playerPair, "TxCollisionPair 实例");
                if (extraObj != null) DumpMembers(sb, extraObj, "查询结果对象");
                sb.AppendLine();
                sb.AppendLine("=== 程序集内含 Collision 的类型 ===");
                try
                {
                    foreach (var t in typeof(TxCollisionRoot).Assembly.GetTypes()
                        .Where(t => t.Name.IndexOf("Collision", StringComparison.OrdinalIgnoreCase) >= 0)
                        .OrderBy(t => t.Name))
                        sb.AppendLine("  " + t.FullName);
                }
                catch { }

                path = Path.Combine(Path.GetTempPath(), "MechArena_CollisionDump.txt");
                File.WriteAllText(path, sb.ToString());
            }
            catch { }
            StatusText = "查询失效→AABB兜底 (诊断:" +
                (string.IsNullOrEmpty(path) ? "写入失败" : path) + ")";
        }

        private static void DumpMembers(StringBuilder sb, object obj, string title)
        {
            sb.AppendLine("=== " + title + " ===");
            if (obj == null) { sb.AppendLine("  (null)"); return; }
            var t = obj.GetType();
            sb.AppendLine("  类型: " + t.FullName);
            try
            {
                foreach (var p in t.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                    .OrderBy(p => p.Name))
                    sb.AppendLine("  P " + p.PropertyType.Name + " " + p.Name);
                foreach (var m in t.GetMethods(BindingFlags.Public | BindingFlags.Instance)
                    .Where(m => !m.IsSpecialName).OrderBy(m => m.Name))
                    sb.AppendLine("  M " + m.ReturnType.Name + " " + m.Name + "(" +
                        string.Join(", ", m.GetParameters()
                            .Select(pp => pp.ParameterType.Name)) + ")");
            }
            catch { }
            sb.AppendLine();
        }

        // =====================================================================
        public void Dispose()
        {
            try { _bulletPair?.Delete(); } catch { }
            try { _playerPair?.Delete(); } catch { }
            _bulletPair = null;
            _playerPair = null;

            // 恢复 CheckCollisions 原值
            try { if (_root != null && _origSaved) _root.CheckCollisions = _origCheckCollisions; }
            catch { }

            Ready = false;
            StatusText = "已清理";
        }
    }
}