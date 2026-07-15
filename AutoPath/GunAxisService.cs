using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Tecnomatix.Engineering;

namespace TxTools.AutoPathPlanner
{
    /// <summary>
    /// 焊钳外部轴服务 (v6.0)。
    ///
    /// ═══════════════════════════════════════════════════════════════
    /// API 事实 (Tecnomatix.NET 文档 + 编译器实证):
    ///
    ///   TxRobotExternalAxisData —— **只有无参构造** (CS1729 证实无 3 参重载)
    ///     ITxDevice Device    {get;set;}
    ///     TxJoint   Joint     {get;set;}
    ///     double    JointValue{get;set;}
    ///   用法: new TxRobotExternalAxisData() 然后逐个赋属性。
    ///
    ///   ITxRoboticLocationOperation:
    ///     TxRobotExternalAxisData[] RobotExternalAxesData          {get;set;}  ← 到达值
    ///     TxRobotExternalAxisData[] RobotDepartureExternalAxesData {get;set;}  ← 离开值
    ///
    ///   焊钳开口无独立 Open/Close API, 只能:
    ///     ① 外部轴路径: 写 RobotExternalAxesData 里枪轴的 JointValue  ← 本类主用
    ///     ② 设备路径: ITxDevice.CurrentPose / GetPoseByName("OPEN"/"CLOSE")  ← 降级
    ///
    ///   ⚠ robot.ExternalAxes 返回的条目可能 **Joint 为 null** (实测),
    ///     必须从 Device 反查驱动关节 (ResolveDrivingJoint), 否则:
    ///       - 限位读出 [0,0] → 开口全被钳成 0 (焊钳闭合!)
    ///       - 数组里塞 null Joint → PS 抛 NullReferenceException
    ///
    ///   ⚠⚠ 开口关节 **不保证正向开启**! 常见约定:
    ///       CLOSE=0, OPEN=-60   ← 负向开启 (很常见)
    ///       CLOSE=0, OPEN=+60   ← 正向开启
    ///       CLOSE=5, OPEN=+65   ← 闭合值非零
    ///     UI 填的"开口 30mm"是**幅值**。不判符号直接写 +30, 负向开启的枪
    ///     会被往闭合方向压 30 → 夹住钣金 / 超限位 / 焊钳不可用。
    ///     本类通过 ProbeSignConvention() 读 OPEN/CLOSE 位姿探测方向,
    ///     OpeningToJointValue(幅值) 负责换算成带符号的实际关节值。
    /// ═══════════════════════════════════════════════════════════════
    ///
    /// 三个必须处理的现实:
    ///
    ///   【多焊钳】一台机器人可能挂 4 把 TxServoGun (实测), 但一个焊接操作只用一把。
    ///            必须从 op.Gun 解析"活动枪", 只写它的轴, 其余枪不动。
    ///
    ///   【多外部轴】机器人可能有 枪轴 + 第7轴导轨 + 变位机。
    ///              RobotExternalAxesData 是**全量数组** —— 写的时候必须把所有轴
    ///              都填上, 漏一个 PS 会用默认值把导轨归零 (灾难)。
    ///              策略: 枪轴写我们算的开口; 非枪轴从相邻焊点继承。
    ///
    ///   【枪未注册为外部轴】此时无法逐点写开口, 只能全局设 CurrentPose。
    ///                     明确告知用户, 不静默降级。
    /// </summary>
    public sealed class GunAxisService
    {
        private readonly Action<string> _log;
        private TxRobot _robot;

        /// <summary>机器人的全部外部轴 (原始快照, 含枪轴与非枪轴)</summary>
        private readonly List<AxisInfo> _axes = new List<AxisInfo>();

        /// <summary>活动枪对应的外部轴索引 (-1 = 枪未注册为外部轴)</summary>
        private int _gunAxisIndex = -1;

        public GunAxisService(Action<string> log)
        {
            _log = log ?? delegate { };
        }

        /// <summary>枪是否注册为外部轴 (决定能否逐点写开口)</summary>
        public bool HasGunAxis { get { return _gunAxisIndex >= 0; } }

        /// <summary>活动焊枪 (op.Gun 解析)</summary>
        public ITxObject ActiveGun { get; private set; }

        /// <summary>外部轴总数</summary>
        public int AxisCount { get { return _axes.Count; } }

        /// <summary>枪轴关节下限/上限 (开口范围, mm 或 °)</summary>
        public double GunMin { get; private set; }
        public double GunMax { get; private set; }

        /// <summary>枪轴当前值</summary>
        public double GunCurrent
        {
            get { return _gunAxisIndex >= 0 ? _axes[_gunAxisIndex].Value : 0; }
        }

        private sealed class AxisInfo
        {
            public ITxDevice Device;
            public TxJoint Joint;
            public double Value;        // 探测时的当前值 (非枪轴的继承默认)
            public string DeviceName;
            public bool IsGun;          // 该轴的 Device 是焊枪类型
            public bool IsActiveGun;    // 且是本操作的活动枪
        }

        // ════════════════════════════════════════════════════════════
        //  探测
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 探测机器人外部轴, 并解析本操作的活动焊枪。
        /// weldOp 可为 null (此时只探测轴, 不认活动枪)。
        /// 返回 true = 至少探到一根外部轴。
        /// </summary>
        public bool Probe(TxRobot robot, ITxObject weldOp)
        {
            _robot = robot;
            _axes.Clear();
            _gunAxisIndex = -1;
            ActiveGun = null;

            if (robot == null) return false;

            // ---- 1. 解析活动焊枪 (op.Gun) ----
            ActiveGun = ResolveActiveGun(weldOp, robot);

            // ---- 2. 读 robot.ExternalAxes ----
            IEnumerable extAxes = null;
            try { extAxes = ((dynamic)robot).ExternalAxes as IEnumerable; }
            catch { }

            if (extAxes == null)
            {
                _log("  [外部轴] robot.ExternalAxes 不可读 — 焊钳开口将走设备降级路径");
                return false;
            }

            foreach (object ax in extAxes)
            {
                if (ax == null) continue;
                var info = BuildAxisInfo(ax);
                if (info == null) continue;
                _axes.Add(info);
            }

            if (_axes.Count == 0)
            {
                _log("  [外部轴] 机器人无外部轴 — 焊钳开口将走设备降级路径");
                return false;
            }

            // ---- 3. 定位活动枪的轴 ----
            for (int i = 0; i < _axes.Count; i++)
            {
                if (_axes[i].IsActiveGun) { _gunAxisIndex = i; break; }
            }
            // 活动枪没匹配上, 但只有一根枪轴 → 就用它
            if (_gunAxisIndex < 0)
            {
                var gunIdx = new List<int>();
                for (int i = 0; i < _axes.Count; i++)
                    if (_axes[i].IsGun) gunIdx.Add(i);
                if (gunIdx.Count == 1)
                {
                    _gunAxisIndex = gunIdx[0];
                    _log("  [外部轴] 活动枪未从 op.Gun 匹配, 但只有 1 根枪轴 → 采用");
                }
                else if (gunIdx.Count > 1)
                {
                    _log(string.Format(
                        "  [外部轴] 有 {0} 根枪轴但无法确定活动枪 (op.Gun 解析失败) — 开口不写",
                        gunIdx.Count));
                }
            }

            // ---- 4. 读枪轴限位 + 探测开口符号约定 ----
            if (_gunAxisIndex >= 0)
            {
                ReadGunLimits();
                ProbeSignConvention();   // v6.2: 必须在限位之后
            }

            LogSummary();
            return true;
        }

        /// <summary>从 TxRobotExternalAxisData (或等价对象) 构造轴信息</summary>
        private AxisInfo BuildAxisInfo(object ax)
        {
            try
            {
                dynamic d = ax;

                ITxDevice dev = null;
                TxJoint joint = null;
                double val = 0;

                try { dev = d.Device as ITxDevice; } catch { }
                try { joint = d.Joint as TxJoint; } catch { }
                try { val = (double)d.JointValue; } catch { }

                // v6.1 关键修复: robot.ExternalAxes 返回的对象可能**只有 Device 没有 Joint**
                // (实测: 限位读出 [0,0], 写入 NRE)。此时从 Device 反查它的驱动关节。
                if (joint == null && dev != null)
                {
                    joint = ResolveDrivingJoint(dev);
                    if (joint != null)
                    {
                        // 关节值也一并从设备当前姿态读
                        double jv;
                        if (TryReadJointValue(dev, joint, out jv)) val = jv;
                    }
                }

                if (dev == null && joint == null) return null;

                string devName = "?";
                try { if (dev != null) devName = ((ITxObject)dev).Name; } catch { }

                bool isGun = IsGunDevice(dev);
                bool isActive = isGun && ActiveGun != null && SameObject(dev, ActiveGun);

                return new AxisInfo
                {
                    Device = dev,
                    Joint = joint,
                    Value = val,
                    DeviceName = devName,
                    IsGun = isGun,
                    IsActiveGun = isActive
                };
            }
            catch { return null; }
        }

        /// <summary>
        /// 从设备反查它的驱动关节 (外部轴的实际可动轴)。
        /// 伺服焊枪的"开口关节"就是它的驱动关节之一 (通常唯一)。
        ///
        /// v6.5: 强类型化。关节类型属性是 TxJoint.Type (返回嵌套枚举
        /// TxJoint.TxJointType), 不是 .TxJointType —— 后者是类型名不是属性名。
        /// </summary>
        private static TxJoint ResolveDrivingJoint(ITxDevice dev)
        {
            if (dev == null) return null;

            // ① DrivingJoints (最准 —— 就是外部轴要驱动的那根)
            try
            {
                var dj = ((dynamic)dev).DrivingJoints as IEnumerable;
                if (dj != null)
                {
                    foreach (object j in dj)
                    {
                        var tj = j as TxJoint;
                        if (tj != null) return tj;
                    }
                }
            }
            catch { }

            // ② Joints 里第一个 Revolute/Prismatic (跳过固定关节)
            // 注意: Joints 不在 ITxDevice 接口上 (CS1061), 只在具体设备类上 → dynamic
            foreach (var tj in EnumJoints(dev))
            {
                try
                {
                    // TxJoint.Type → TxJoint.TxJointType 枚举 (Revolute / Prismatic)
                    if (tj.Type == TxJoint.TxJointType.Revolute
                        || tj.Type == TxJoint.TxJointType.Prismatic)
                        return tj;
                }
                catch
                {
                    return tj;   // 类型读不到就用第一个
                }
            }

            return null;
        }

        /// <summary>
        /// 枚举设备的关节。
        /// v6.5.1: ITxDevice 接口上**没有** Joints (CS1061) —— 它在具体设备类
        /// (TxServoGun / TxDevice 等) 上。这是一处必要的 dynamic。
        /// </summary>
        private static List<TxJoint> EnumJoints(ITxDevice dev)
        {
            var result = new List<TxJoint>();
            if (dev == null) return result;
            try
            {
                var joints = ((dynamic)dev).Joints as IEnumerable;
                if (joints == null) return result;
                foreach (object o in joints)
                {
                    var tj = o as TxJoint;
                    if (tj != null) result.Add(tj);
                }
            }
            catch { }
            return result;
        }

        /// <summary>从设备当前姿态读某个关节的值 (v6.5 强类型)</summary>
        private static bool TryReadJointValue(ITxDevice dev, TxJoint joint, out double val)
        {
            val = 0;
            try
            {
                int idx = FindJointIndex(dev, joint);
                if (idx < 0) return false;

                TxPoseData pose = dev.CurrentPose;
                if (pose == null) return false;

                ArrayList jv = pose.JointValues;
                if (jv == null || idx >= jv.Count) return false;

                val = Convert.ToDouble(jv[idx]);
                return true;
            }
            catch { return false; }
        }

        /// <summary>
        /// 解析本操作的活动焊枪。
        /// v6.5: ITxWeldOperation.Gun 返回 ITxGun (文档确认), 强类型直取,
        ///       不再用 dynamic 猜属性。
        /// </summary>
        private ITxObject ResolveActiveGun(ITxObject weldOp, TxRobot robot)
        {
            // ① 操作自身的 Gun (最可靠)
            var wop = weldOp as ITxWeldOperation;
            if (wop != null)
            {
                try
                {
                    ITxGun g = wop.Gun;
                    if (g != null) return g as ITxObject;
                }
                catch { }
            }

            // ② 首个焊点子项的 Gun
            var coll = weldOp as ITxObjectCollection;
            if (coll != null)
            {
                try
                {
                    var filter = new TxTypeFilter(typeof(TxWeldLocationOperation));
                    TxObjectList kids = coll.GetAllDescendants(filter);
                    foreach (ITxObject k in kids)
                    {
                        var kw = k as ITxWeldOperation;
                        if (kw == null) continue;
                        try
                        {
                            ITxGun g = kw.Gun;
                            if (g != null) return g as ITxObject;
                        }
                        catch { }
                    }
                }
                catch { }
            }

            // ③ 机器人只挂一把枪 → 就是它
            try
            {
                var guns = CollisionSetService.CollectMountedTools(robot)
                    .Where(IsGunObject).ToList();
                if (guns.Count == 1) return guns[0];
            }
            catch { }

            return null;
        }

        /// <summary>
        /// 读枪轴关节限位。
        ///
        /// v6.5 根因修复: 之前拿 LowerLimit/MinLimit/LowLimit 三个**猜的**属性名硬试,
        /// 全落空 → 限位读出 [0,0] → _limitsUnknown=true → 限位推断走不通 →
        /// 符号约定兜底成"正向" → 负向开启的枪被写入正值。
        ///
        /// 真实 API (文档确认): TxJoint.LowerSoftLimit / UpperSoftLimit
        /// 限位读对之后, 负向开启就能靠 ③限位推断 自动识别出来。
        /// </summary>
        private void ReadGunLimits()
        {
            GunMin = 0;
            GunMax = 0;
            _limitsUnknown = false;

            var j = _axes[_gunAxisIndex].Joint;
            if (j == null) { _limitsUnknown = true; return; }

            try
            {
                GunMin = j.LowerSoftLimit;
                GunMax = j.UpperSoftLimit;
            }
            catch (Exception ex)
            {
                _log("  [外部轴] 关节限位读取失败: " + ex.Message);
                _limitsUnknown = true;
            }

            // 限位无效 (两端相等) → 给保守默认, 否则开口会被钳成 0 (焊钳闭合)
            if (Math.Abs(GunMax - GunMin) < 1e-6)
            {
                GunMin = 0;
                GunMax = DefaultGunMaxOpening;
                _limitsUnknown = true;
            }
        }

        /// <summary>限位读不到时的默认开口上限 (mm)</summary>
        public double DefaultGunMaxOpening = 200.0;
        private bool _limitsUnknown;

        // ════════════════════════════════════════════════════════════
        //  v6.2 开口符号约定 —— 关键!
        //
        //  伺服焊枪的开口关节**不保证正向开启**。常见约定:
        //    CLOSE = 0,  OPEN = -60     ← 负向开启 (很常见!)
        //    CLOSE = 0,  OPEN = +60     ← 正向开启
        //    CLOSE = 5,  OPEN = 65      ← 闭合值非零
        //
        //  UI 上填的"开口 30mm"是**幅值**。如果不判符号直接写 +30:
        //    负向开启的枪 → 往闭合方向再压 30 → 夹住钣金 / 超限位 / 焊钳不可用
        //
        //  探测: 读 CLOSE / OPEN 预定义位姿的枪轴关节值 → 求方向
        //  兜底: 用限位推断 (哪侧离闭合位远, 哪边就是开启方向)
        // ════════════════════════════════════════════════════════════

        /// <summary>闭合位的关节值 (开口幅值 0 对应的关节值)</summary>
        public double ClosedValue { get; private set; }

        /// <summary>全开位的关节值</summary>
        public double FullOpenValue { get; private set; }

        /// <summary>开启方向: +1 = 正向开启, -1 = 负向开启</summary>
        public int OpenDirection { get; private set; }

        /// <summary>
        /// v6.4 开口方向手动覆盖 (自动探测不可靠时的兜底):
        ///    0 = 自动探测 (默认)
        ///   +1 = 强制正向 (CLOSE=0, OPEN=+N)
        ///   -1 = 强制负向 (CLOSE=0, OPEN=-N)
        /// </summary>
        public int OpenDirectionOverride = 0;

        /// <summary>v6.4 手动指定最大开口幅值 (mm); &lt;=0 = 用探测值/默认值</summary>
        public double MaxOpeningOverride = 0;

        /// <summary>最大开口幅值 (|FullOpenValue - ClosedValue|)</summary>
        public double MaxOpeningMagnitude
        {
            get
            {
                if (MaxOpeningOverride > 0) return MaxOpeningOverride;
                return Math.Abs(FullOpenValue - ClosedValue);
            }
        }

        private string _signSource = "?";

        /// <summary>
        /// 探测开口符号约定。
        ///
        /// v6.4 优先级:
        ///   ⓪ OpenDirectionOverride ≠ 0 → 用手动指定的方向 (最高优先级)
        ///   ① CLOSE/CLOSED + OPEN 预定义位姿的枪轴关节值 (最可靠)
        ///   ② 只有 OPEN 位姿 → 闭合位取限位中离 OPEN 远的那端
        ///   ③ 关节限位推断: 绝对值小的一端为闭合, 另一端为全开
        ///   ④ 全失败 → dump 关节/设备成员, 让下一轮能对症下药
        /// </summary>
        private void ProbeSignConvention()
        {
            OpenDirection = 1;
            ClosedValue = 0;
            FullOpenValue = _limitsUnknown ? DefaultGunMaxOpening : GunMax;
            _signSource = "默认(正向)";

            if (_gunAxisIndex < 0) return;
            var a = _axes[_gunAxisIndex];
            if (a.Device == null || a.Joint == null) return;

            // ---- ⓪ 手动覆盖 (最高优先级) ----
            if (OpenDirectionOverride != 0)
            {
                OpenDirection = OpenDirectionOverride > 0 ? 1 : -1;
                ClosedValue = 0;
                double mag = MaxOpeningOverride > 0 ? MaxOpeningOverride : DefaultGunMaxOpening;
                FullOpenValue = ClosedValue + OpenDirection * mag;
                _signSource = "手动指定";
                return;
            }

            // ---- ① 从预定义位姿读 (多种命名) ----
            double closeV, openV;
            bool hasClose = TryReadPoseJointValue(a.Device, a.Joint, CloseNames, out closeV);
            bool hasOpen = TryReadPoseJointValue(a.Device, a.Joint, OpenNames, out openV);

            if (hasClose && hasOpen && Math.Abs(openV - closeV) > 1e-6)
            {
                ClosedValue = closeV;
                FullOpenValue = openV;
                OpenDirection = openV > closeV ? 1 : -1;
                _signSource = "OPEN/CLOSE 位姿";
                return;
            }

            // ---- ② 只有 OPEN: 闭合位取限位中离 OPEN 远的那端 ----
            if (hasOpen && !_limitsUnknown)
            {
                double cLo = Math.Abs(openV - GunMin);
                double cHi = Math.Abs(openV - GunMax);
                ClosedValue = cLo > cHi ? GunMin : GunMax;
                FullOpenValue = openV;
                OpenDirection = openV > ClosedValue ? 1 : -1;
                _signSource = "OPEN 位姿 + 限位";
                return;
            }

            // ---- ③ 纯限位推断 ----
            if (!_limitsUnknown && Math.Abs(GunMax - GunMin) > 1e-6)
            {
                if (Math.Abs(GunMin) <= Math.Abs(GunMax))
                {
                    ClosedValue = GunMin;
                    FullOpenValue = GunMax;
                    OpenDirection = GunMax > GunMin ? 1 : -1;
                }
                else
                {
                    ClosedValue = GunMax;
                    FullOpenValue = GunMin;
                    OpenDirection = GunMin > GunMax ? 1 : -1;   // 通常是 -1
                }
                _signSource = "限位推断";
                return;
            }

            // ---- ④ 全失败: dump 真实 API, 供下一轮对症下药 ----
            DumpJointAndDeviceMembers(a);
        }

        private static readonly string[] OpenNames =
        {
            "OPEN", "Open", "open", "OPENED", "OPEN_POS", "OPENPOS",
            "SEMIOPEN", "SEMI_OPEN", "HALFOPEN", "开", "张开"
        };

        private static readonly string[] CloseNames =
        {
            "CLOSE", "Close", "close", "CLOSED", "CLOSE_POS", "CLOSEPOS",
            "SHUT", "HOME", "关", "闭合"
        };

        /// <summary>
        /// v6.4 探测全失败时: dump 枪轴关节与设备的真实成员。
        /// 目的是把"我在猜 API"变成"我知道 API" —— 下一轮照着 dump 出来的
        /// 属性名直接读, 而不是继续拿候选名硬试。
        /// </summary>
        private void DumpJointAndDeviceMembers(AxisInfo a)
        {
            _log("  [外部轴诊断] 符号约定探测全失败 — dump 真实成员:");

            // 关节属性
            try
            {
                _log(string.Format("    枪轴关节类型: {0}", a.Joint.GetType().FullName));
                foreach (var pi in a.Joint.GetType().GetProperties())
                {
                    if (pi.GetIndexParameters().Length > 0) continue;
                    string vs;
                    try
                    {
                        object v = pi.GetValue(a.Joint, null);
                        vs = v == null ? "null" : v.ToString();
                        if (vs.Length > 40) vs = vs.Substring(0, 40) + "...";
                    }
                    catch (Exception ex) { vs = "<" + ex.GetType().Name + ">"; }
                    _log(string.Format("      joint.{0} ({1}) = {2}",
                        pi.Name, pi.PropertyType.Name, vs));
                }
            }
            catch (Exception ex) { _log("    关节 dump 失败: " + ex.Message); }

            // 设备的位姿列表 (看看到底有哪些位姿名)
            try
            {
                _log(string.Format("    枪设备类型: {0}", a.Device.GetType().FullName));

                TxObjectList poses = a.Device.PoseList;   // v6.5 强类型
                if (poses == null || poses.Count == 0)
                {
                    _log("    设备位姿列表为空 — 该枪没有定义 OPEN/CLOSE 位姿");
                    _log("    建议: 在 PS 里给焊钳定义 OPEN/CLOSE 位姿, 符号约定就能自动探到");
                }
                else
                {
                    var names = new List<string>();
                    foreach (ITxObject p in poses)
                    {
                        try { names.Add(p.Name); }
                        catch { names.Add("?"); }
                    }
                    _log(string.Format("    设备位姿名 ({0} 个): {1}",
                        names.Count, string.Join(", ", names.ToArray())));
                }
            }
            catch (Exception ex) { _log("    设备 dump 失败: " + ex.Message); }

            _log("  [外部轴诊断] → 请把上面的 dump 发我, 或在界面手动指定开口方向");
        }

        /// <summary>
        /// 从设备的某个具名位姿里读指定关节的值。
        /// v6.5: ITxDevice.PoseList 强类型 (TxObjectList)。
        /// 条目可能是 TxPose (含 PoseData 属性) 或直接是 TxPoseData —— 两种都试。
        /// </summary>
        private static bool TryReadPoseJointValue(ITxDevice dev, TxJoint joint,
            string[] poseNames, out double val)
        {
            val = 0;
            if (dev == null || joint == null) return false;

            int idx = FindJointIndex(dev, joint);
            if (idx < 0) return false;

            try
            {
                TxObjectList poses = dev.PoseList;
                if (poses == null) return false;

                foreach (ITxObject p in poses)
                {
                    string n;
                    try { n = p.Name; }
                    catch { continue; }
                    if (string.IsNullOrEmpty(n)) continue;

                    bool match = false;
                    foreach (string want in poseNames)
                    {
                        if (string.Equals(n, want, StringComparison.OrdinalIgnoreCase))
                        { match = true; break; }
                    }
                    if (!match) continue;

                    ArrayList jv = ExtractPoseJointValues(p);
                    if (jv == null || idx >= jv.Count) continue;

                    val = Convert.ToDouble(jv[idx]);
                    return true;
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 从 PoseList 条目里取关节值。
        /// 条目类型未在文档中确认 —— 可能是 TxPose (需 .PoseData) 或 TxPoseData 本身。
        /// 这里保留一处必要的 dynamic 兜底。
        /// </summary>
        private static ArrayList ExtractPoseJointValues(object poseItem)
        {
            if (poseItem == null) return null;

            // ① 条目本身就是 TxPoseData
            // (参数类型必须是 object —— TxPoseData 不实现 ITxObject, 用 ITxObject
            //  做参数会导致 CS0039: 无法转换)
            var direct = poseItem as TxPoseData;
            if (direct != null)
            {
                try { return direct.JointValues; }
                catch { }
            }

            // ② 条目是 TxPose 之类, 带 PoseData 属性 (必要 dynamic 兜底)
            try
            {
                var pd = ((dynamic)poseItem).PoseData as TxPoseData;
                if (pd != null) return pd.JointValues;
            }
            catch { }

            return null;
        }

        /// <summary>
        /// 开口幅值 (mm, 恒非负) → 实际关节值 (可能是负数!)。
        ///
        ///     jointValue = ClosedValue + OpenDirection × |幅值|
        ///
        /// 这是避免"负向开启的枪被写入正值"的关键换算。
        /// 结果再钳到限位内。
        /// </summary>
        public double OpeningToJointValue(double openingMagnitude)
        {
            double m = Math.Abs(openingMagnitude);

            // 幅值不能超过物理全开
            double maxM = MaxOpeningMagnitude;
            if (maxM > 1e-6 && m > maxM) m = maxM;

            double jv = ClosedValue + OpenDirection * m;
            return ClampToLimits(jv);
        }

        /// <summary>实际关节值 → 开口幅值 (诊断用)</summary>
        public double JointValueToOpening(double jointValue)
        {
            return Math.Abs(jointValue - ClosedValue);
        }

        /// <summary>把关节值钳到限位内 (注意 GunMin/GunMax 未必是 闭合/全开 的顺序)</summary>
        private double ClampToLimits(double v)
        {
            if (_limitsUnknown) return v;
            double lo = Math.Min(GunMin, GunMax);
            double hi = Math.Max(GunMin, GunMax);
            return Math.Max(lo, Math.Min(hi, v));
        }

        private void LogSummary()
        {
            _log(string.Format("  [外部轴] 探到 {0} 根:", _axes.Count));
            for (int i = 0; i < _axes.Count; i++)
            {
                var a = _axes[i];
                string tag = a.IsActiveGun ? " ★活动枪"
                           : a.IsGun ? " (焊枪, 非活动)"
                           : " (导轨/变位机)";
                string jn = "关节=?";
                try { if (a.Joint != null) jn = "关节=" + a.Joint.Name; }
                catch { }
                if (a.Joint == null) jn = "关节=null ⚠";

                _log(string.Format("    [{0}] {1} = {2:F1}  {3}{4}",
                    i, a.DeviceName, a.Value, jn, tag));
            }

            if (_gunAxisIndex >= 0)
            {
                if (_axes[_gunAxisIndex].Joint == null)
                {
                    _log("  [外部轴] ⚠ 枪轴关节为 null — 开口写不进去");
                    _log("           robot.ExternalAxes 未提供 Joint, 且从设备反查 DrivingJoints/Joints 也失败");
                    return;
                }

                string src = _limitsUnknown ? " (限位读不到, 用默认上限)" : "";
                _log(string.Format("  [外部轴] 枪轴可写 — 关节限位 [{0:F1}, {1:F1}]{2}, 当前 {3:F1}",
                    GunMin, GunMax, src, GunCurrent));

                // v6.2/6.4: 开口符号约定 —— 必须让人肉眼确认
                string dirTxt = OpenDirection > 0 ? "正向 (+)" : "负向 (−)";
                _log(string.Format(
                    "  [外部轴] 开口约定: 闭合={0:F1} 全开={1:F1} → 开启方向 {2}, 最大幅值 {3:F1}  [来源: {4}]",
                    ClosedValue, FullOpenValue, dirTxt, MaxOpeningMagnitude, _signSource));

                if (_signSource == "手动指定")
                {
                    _log(string.Format("           ✔ 采用界面手动指定 — 开口 30mm 将写成关节值 {0:F1}",
                        OpeningToJointValue(30)));
                }
                else if (OpenDirection < 0)
                {
                    _log("           ⚠ 该枪为**负向开启** — 开口幅值将写成负关节值 (若直接写正值会夹紧钣金)");
                }

                if (_signSource == "默认(正向)")
                {
                    _log("           ⚠⚠ 符号约定未探到 — 按正向处理, 极可能是错的!");
                    _log("               请在界面 [节拍 & 焊钳] 页手动指定「开口方向」");
                }
            }
            else
            {
                _log("  [外部轴] 枪未注册为外部轴 — 无法逐点写开口");
                _log("           建议: 在机器人属性中把焊钳添加为外部轴, 才能按点控制开合");
            }
        }

        // ════════════════════════════════════════════════════════════
        //  写值
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 把开口值写入位置操作的外部轴数组 (到达值 + 可选离开值)。
        ///
        /// 全量重建 TxRobotExternalAxisData[]:
        ///   枪轴   → opening (钳制到限位内)
        ///   非枪轴 → 从 inheritFrom 继承; 继承不到则用探测时的当前值
        ///            (绝不能漏填 —— 漏了 PS 会把导轨归零)
        ///
        /// inheritFrom: 通常传相邻的焊点位置, 从它身上读导轨/变位机的工艺值。
        /// </summary>
        public bool WriteOpening(ITxObject location, double opening,
            ITxObject inheritFrom, bool alsoDeparture = true)
        {
            if (location == null || _gunAxisIndex < 0 || _axes.Count == 0) return false;

            // v6.1: 枪轴的 Joint 必须有 —— 没有就写不了 (之前这里直接 NRE 刷屏)
            if (_axes[_gunAxisIndex].Joint == null)
            {
                if (!_warnedNoJoint)
                {
                    _warnedNoJoint = true;
                    _log("  [外部轴] 枪轴关节解析失败 (Joint=null) — 开口无法写入");
                    _log("           robot.ExternalAxes 只给了 Device 没给 Joint, 且从设备反查驱动关节也失败");
                }
                return false;
            }

            try
            {
                Dictionary<string, double> inherited = ReadAxisValues(inheritFrom);

                // v6.2 关键: opening 是**幅值**, 必须按符号约定换算成实际关节值。
                // 负向开启的枪 (CLOSE=0, OPEN=-60), 直接写 +30 会往闭合方向压 ——
                // 夹住钣金 / 超限位 / 焊钳不可用。
                double gunJointValue = OpeningToJointValue(opening);

                // 只收 Joint 非 null 的轴 (null 的塞进数组会让 PS 抛 NRE)
                var list = new List<TxRobotExternalAxisData>(_axes.Count);
                for (int i = 0; i < _axes.Count; i++)
                {
                    var a = _axes[i];
                    if (a.Device == null || a.Joint == null) continue;

                    double v;
                    if (i == _gunAxisIndex)
                        v = gunJointValue;                 // 枪轴: 符号换算后的关节值
                    else if (inherited != null && inherited.ContainsKey(AxisKey(a)))
                        v = inherited[AxisKey(a)];         // 非枪轴: 从邻近焊点继承
                    else
                        v = a.Value;                       // 兜底: 探测时的当前值

                    // TxRobotExternalAxisData 只有无参构造 (CS1729: 无 3 参重载),
                    // Device/Joint/JointValue 是 get/set 属性 —— 逐个赋值。
                    var d = new TxRobotExternalAxisData();
                    d.Device = a.Device;
                    d.Joint = a.Joint;
                    d.JointValue = v;
                    list.Add(d);
                }

                if (list.Count == 0) return false;

                // v6.5: ITxRoboticLocationOperation 强类型 (原来是 dynamic loc)
                var loc = location as ITxRoboticLocationOperation;
                if (loc == null)
                {
                    if (!_warnedNoJoint)
                    {
                        _warnedNoJoint = true;
                        _log("  [外部轴] 目标不是 ITxRoboticLocationOperation — 开口写不进去");
                    }
                    return false;
                }

                loc.RobotExternalAxesData = list.ToArray();
                if (alsoDeparture)
                {
                    try { loc.RobotDepartureExternalAxesData = list.ToArray(); }
                    catch { }
                }
                return true;
            }
            catch (Exception ex)
            {
                if (!_warnedWriteFail)
                {
                    _warnedWriteFail = true;   // 只报一次, 不刷屏
                    _log("  [外部轴] 写入失败: " + ex.Message + " (后续同类错误不再重复报告)");
                }
                return false;
            }
        }

        private bool _warnedNoJoint;
        private bool _warnedWriteFail;

        /// <summary>
        /// 读取某个位置操作上已有的外部轴值 (键 = 设备名|关节名)。
        /// v6.5: ITxRoboticLocationOperation.RobotExternalAxesData 强类型。
        /// </summary>
        private Dictionary<string, double> ReadAxisValues(ITxObject location)
        {
            var loc = location as ITxRoboticLocationOperation;
            if (loc == null) return null;

            try
            {
                TxRobotExternalAxisData[] arr = loc.RobotExternalAxesData;
                if (arr == null || arr.Length == 0) return null;

                var map = new Dictionary<string, double>();
                foreach (var d in arr)
                {
                    if (d == null) continue;
                    string k = AxisKeyOf(d.Device, d.Joint);
                    if (k != null && !map.ContainsKey(k)) map[k] = d.JointValue;
                }
                return map.Count > 0 ? map : null;
            }
            catch { return null; }
        }

        private string AxisKey(AxisInfo a) { return AxisKeyOf(a.Device, a.Joint); }

        private static string AxisKeyOf(ITxDevice dev, TxJoint joint)
        {
            string d = "?", j = "?";
            try { if (dev != null) d = ((ITxObject)dev).Name; } catch { }
            try { if (joint != null) j = joint.Name; } catch { }
            if (d == "?" && j == "?") return null;
            return d + "|" + j;
        }

        /// <summary>
        /// 把开口**幅值**钳到 [0, 最大幅值]。
        /// v6.2: 这里只管幅值 (恒非负), 符号由 OpeningToJointValue 负责。
        /// </summary>
        public double ClampOpening(double openingMagnitude)
        {
            double m = Math.Abs(openingMagnitude);
            double maxM = MaxOpeningMagnitude;
            if (maxM > 1e-6 && m > maxM) m = maxM;
            return m;
        }

        // ════════════════════════════════════════════════════════════
        //  按需最小开口 (节拍优化核心)
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 按需最小开口搜索: 从小到大试开口, 第一个不干涉的就用。
        ///
        /// 为什么不无脑全开:
        ///   ① 开口大 = 伺服行程长 = 慢 (直接吃节拍)
        ///   ② 开口大 = 枪臂张开 = 包络更大 = 反而更容易撞
        ///
        /// applyOpening: 把开口值实际作用到设备上 (供碰撞查询看见)
        /// isFree:       在当前开口下查询该点是否无干涉
        /// 返回: 找到的最小安全开口; 全部失败返回 double.NaN
        /// </summary>
        public double FindMinSafeOpening(
            Action<double> applyOpening,
            Func<bool> isFree,
            double[] ladder)
        {
            if (applyOpening == null || isFree == null) return double.NaN;
            if (ladder == null || ladder.Length == 0) return double.NaN;

            foreach (double o in ladder)
            {
                double c = ClampOpening(o);
                try
                {
                    applyOpening(c);
                    if (isFree()) return c;
                }
                catch { }
            }
            return double.NaN;
        }

        /// <summary>
        /// 把开口**幅值**作用到活动枪设备上 (碰撞查询前用)。
        ///
        /// v6.5 修复: TxPoseData **没有 SetJointValue 方法** (文档确认) ——
        /// 之前那行 pose.SetJointValue(idx, v) 走 dynamic, 运行时抛异常被 catch 吞掉,
        /// 于是开口从来没真正作用到设备上, "按需最小开口"搜索的结果全是错的。
        /// 正确路径: 读 JointValues (ArrayList) → 改指定下标 → 写回。
        ///
        /// 开口值走 OpeningToJointValue 换算 —— 负向开启的枪会得到负数。
        /// </summary>
        public bool ApplyOpeningToDevice(double openingMagnitude)
        {
            if (_gunAxisIndex < 0) return false;
            var a = _axes[_gunAxisIndex];
            if (a.Device == null || a.Joint == null) return false;

            try
            {
                int idx = FindJointIndex(a.Device, a.Joint);
                if (idx < 0) return false;

                TxPoseData cur = a.Device.CurrentPose;
                if (cur == null) return false;

                // TxPoseData.JointValues 是 ArrayList (强类型确认)
                ArrayList jv = cur.JointValues;
                if (jv == null || idx >= jv.Count) return false;

                var newJv = new ArrayList(jv);      // 复制, 不改原对象
                newJv[idx] = OpeningToJointValue(openingMagnitude);

                var pd = new TxPoseData();
                pd.JointValues = newJv;
                a.Device.CurrentPose = pd;

                // v6.3: 开口变了 = 枪的碰撞包络变了 → 位姿缓存必须失效
                if (OnGeometryChanged != null) OnGeometryChanged();
                return true;
            }
            catch (Exception ex)
            {
                if (!_warnedApplyFail)
                {
                    _warnedApplyFail = true;
                    _log("  [外部轴] 开口作用到设备失败: " + ex.Message
                        + " — 按需最小开口搜索将不准 (后续不再重复报告)");
                }
                return false;
            }
        }

        private bool _warnedApplyFail;

        /// <summary>
        /// v6.3 几何变更通知 —— 焊钳开口改变时触发。
        /// CollisionWorld 订阅它来失效位姿缓存 (开口不同 = 碰撞包络不同,
        /// 缓存结果不能跨开口复用)。
        /// </summary>
        public Action OnGeometryChanged;

        /// <summary>关节在设备关节表里的序号</summary>
        private static int FindJointIndex(ITxDevice dev, TxJoint target)
        {
            if (dev == null || target == null) return -1;

            var joints = EnumJoints(dev);   // ITxDevice 上没有 Joints, 走 dynamic
            for (int i = 0; i < joints.Count; i++)
            {
                var tj = joints[i];
                if (ReferenceEquals(tj, target)) return i;
                try { if (tj.Name == target.Name) return i; }   // 同名兜底
                catch { }
            }
            return -1;
        }

        // ════════════════════════════════════════════════════════════
        //  降级路径: 设备 Pose (枪未注册为外部轴时)
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 按名字应用预定义位姿 (OPEN / CLOSE)。
        /// v6.5: ITxDevice **没有 JumpToPose 方法** (文档确认) —— 已删除该分支。
        /// 唯一路径: 取位姿的 TxPoseData → 赋给 device.CurrentPose。
        /// </summary>
        public bool ApplyNamedPose(ITxObject gun, string poseName)
        {
            var dev = gun as ITxDevice;
            if (dev == null) return false;

            try
            {
                TxObjectList poses = dev.PoseList;
                if (poses == null) return false;

                foreach (ITxObject p in poses)
                {
                    string n;
                    try { n = p.Name; }
                    catch { continue; }
                    if (!string.Equals(n, poseName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    ArrayList jv = ExtractPoseJointValues(p);
                    if (jv == null) continue;

                    var pd = new TxPoseData();
                    pd.JointValues = jv;
                    dev.CurrentPose = pd;

                    if (OnGeometryChanged != null) OnGeometryChanged();
                    return true;
                }
            }
            catch { }
            return false;
        }

        // ════════════════════════════════════════════════════════════
        //  类型判定
        // ════════════════════════════════════════════════════════════

        private static bool IsGunDevice(ITxDevice dev)
        {
            return dev != null && IsGunObject(dev as ITxObject);
        }

        /// <summary>类型名含 "Gun" (含继承链) — 覆盖 TxServoGun / TxWeldGun / TxPneumaticGun 等</summary>
        public static bool IsGunObject(ITxObject o)
        {
            if (o == null) return false;
            var t = o.GetType();
            while (t != null && t != typeof(object))
            {
                if (t.Name.IndexOf("Gun", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
                t = t.BaseType;
            }
            return false;
        }

        private static bool SameObject(ITxDevice dev, ITxObject other)
        {
            if (dev == null || other == null) return false;
            var d = dev as ITxObject;
            if (d == null) return false;
            if (ReferenceEquals(d, other)) return true;
            // PS 可能返回不同代理 → 名字兜底
            try { return d.Name == other.Name; }
            catch { return false; }
        }
    }
}
