using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using Tecnomatix.Engineering;

namespace TxTools.RobotBaseChecker
{
    // ============================================================
    // 线程上下文：所有 PS SDK 调用必须回到主线程
    // ============================================================
    internal static class PsContext
    {
        private static SynchronizationContext _ctx;

        public static void Capture()
        {
            _ctx = SynchronizationContext.Current;
        }

        public static void Run(Action action)
        {
            if (_ctx == null || _ctx == SynchronizationContext.Current)
            {
                action();
                return;
            }
            Exception captured = null;
            _ctx.Send(_ =>
            {
                try { action(); }
                catch (Exception e) { captured = e; }
            }, null);
            if (captured != null) throw captured;
        }
    }

    // ============================================================
    // 位姿（统一用 XYZ + RPY 表达；RPY 单位未定，见下方说明）
    // 自身坐标与 BASE0 都走同一个提取函数，保证单位一致，
    // 因此 ΔRot 差值即便单位待定，"是否一致"的判断依然有效。
    // ============================================================
    internal sealed class Pose6
    {
        public double X, Y, Z;
        public double RX, RY, RZ;
        public bool PosValid;
        public bool RotValid;
        public string Source = "";

        public string PosText => PosValid ? $"{X:F3}, {Y:F3}, {Z:F3}" : "—";
        public string RotText => RotValid ? $"{RX:F4}, {RY:F4}, {RZ:F4}" : "—";

        public override string ToString() => $"[{Source}] T({PosText})  R({RotText})";
    }

    // ============================================================
    // 单台机器人的分析结果
    // ============================================================
    internal sealed class RobotBaseResult
    {
        public string RobotName = "";
        public TxRobot RobotRef;               // 机器人实例引用（同名场景下精确同步）
        public string Brand = "";              // 识别出的品牌（FANUC / Generic）
        public Pose6 Self = new Pose6();       // 机器人自身坐标（安装位姿）
        public Pose6 Base0 = new Pose6();      // 当前存储的 BASE0（系统坐标系首个）
        public Pose6 Expected = new Pose6();   // 品牌感知的期望 BASE0（对比/同步目标）

        public bool Comparable;     // 当前 BASE0 与期望 BASE0 都可读
        public double DeltaPos;     // 平移欧氏距离差 (mm) —— 当前 BASE0 vs 期望
        public double DeltaRot;     // 旋转分量最大绝对差
        public string Verdict = "";
        public string Detail = "";  // 该机器人的完整探测日志
    }

    // ============================================================
    // 核心读取器
    // ============================================================
    internal static class RobotBaseReader
    {
        // 入口：在 PS 主线程执行，返回所有机器人的分析结果
        public static List<RobotBaseResult> Analyze(double posTolMm, double rotTol, BrandMode brandMode)
        {
            var results = new List<RobotBaseResult>();

            PsContext.Run(() =>
            {
                var doc = TxApplication.ActiveDocument;
                if (doc == null)
                    throw new InvalidOperationException("没有打开的研究 (ActiveDocument 为 null)。");

                var robots = CollectRobots(doc);
                foreach (var robot in robots)
                {
                    var log = new StringBuilder();
                    var r = new RobotBaseResult { RobotName = SafeName(robot), RobotRef = robot };
                    log.AppendLine("机器人: " + r.RobotName);

                    // --- 自身坐标（安装位姿） ---
                    r.Self = ReadSelf(robot, log);
                    TxTransformation selfTx = SafeTx(() => robot.AbsoluteLocation);
                    log.AppendLine("自身坐标: " + r.Self);

                    // --- 当前存储 BASE0 ---
                    r.Base0 = ReadBase0(robot, log);
                    log.AppendLine("当前BASE0: " + r.Base0);

                    // --- 品牌识别 + 期望 BASE0（方向一：FANUC 用 J1∩J2） ---
                    r.Brand = RobotKinematics.DetectBrand(robot, brandMode, log);
                    log.AppendLine($"品牌识别: {r.Brand}  (模式: {brandMode})");
                    r.Expected = RobotKinematics.ComputeExpectedPose(robot, selfTx, r.Self, r.Brand, log);
                    log.AppendLine("期望BASE0: " + r.Expected);

                    // --- 对比：当前 BASE0 vs 期望 BASE0 ---
                    Compare(r, posTolMm, rotTol);
                    log.AppendLine($"结论: {r.Verdict}  (ΔPos={FmtDelta(r.DeltaPos)} mm, ΔRot={FmtDelta(r.DeltaRot)})");

                    r.Detail = log.ToString();
                    results.Add(r);
                }
            });

            return results;
        }

        // ----------------------------------------------------------------
        // 1) 遍历场景内所有机器人
        // ----------------------------------------------------------------
        private static List<TxRobot> CollectRobots(object doc)
        {
            var seen = new HashSet<TxRobot>();
            var list = new List<TxRobot>();
            var filter = new TxTypeFilter(typeof(TxRobot));

            foreach (var root in GetRoots(doc))
            {
                if (root == null) continue;
                try
                {
                    var found = ((dynamic)root).GetAllDescendants(filter);
                    foreach (var o in found)
                    {
                        var rb = o as TxRobot;
                        if (rb != null && seen.Add(rb)) list.Add(rb);
                    }
                }
                catch { /* 该根不支持 GetAllDescendants，跳过 */ }
            }
            return list;
        }

        // 物理树多个候选根：机器人通常在 ResourceRoot 下，但 PhysicalRoot 更全
        private static IEnumerable<object> GetRoots(object doc)
        {
            dynamic d = doc;
            foreach (var name in new[] { "PhysicalRoot", "ResourceRoot", "ComponentRoot" })
            {
                object root = null;
                try { root = TryGet(d, name); } catch { }
                if (root != null) yield return root;
            }
        }

        // ----------------------------------------------------------------
        // 2) 自身坐标（机器人在 PS 场景中的安装位姿，相对 world）
        // ----------------------------------------------------------------
        private static Pose6 ReadSelf(TxRobot robot, StringBuilder log)
        {
            dynamic dr = robot;
            TxTransformation t = null;

            foreach (var member in new[] { "AbsoluteLocation", "LocationRelativeToWorkingFrame" })
            {
                try
                {
                    object o = TryGet(dr, member);
                    t = AsTx(o);
                    if (t != null)
                    {
                        var p = FromTx(t, "Self." + member);
                        log.AppendLine("  自身坐标来源: " + member);
                        return p;
                    }
                }
                catch (Exception e) { log.AppendLine($"  自身坐标 {member} 失败: {e.Message}"); }
            }
            return new Pose6 { Source = "Self(未读取)" };
        }

        // ----------------------------------------------------------------
        // 3) RRS BASE0（控制器内部基坐标）
        //    主路径：robot.GetAllSystemFrames()
        //    该方法返回机器人树下 <robot>.<frame> 的全部系统坐标系（工具/用户/基），
        //    BASE0 即其中的基坐标系（与 TCP 枚举一致的已验证取法）。
        // ----------------------------------------------------------------
        // ----------------------------------------------------------------
        // 3) RRS BASE0（控制器内部基坐标）
        //    改为：直接读取系统坐标系列表中的第一个坐标系
        // ----------------------------------------------------------------
        private static Pose6 ReadBase0(TxRobot robot, StringBuilder log)
        {
            System.Collections.IEnumerable frames = null;
            try
            {
                dynamic dr = robot;
                frames = dr.GetAllSystemFrames() as System.Collections.IEnumerable;
            }
            catch (Exception e) { log.AppendLine("  GetAllSystemFrames 异常: " + e.Message); }

            if (frames != null)
            {
                // 先统计数量，便于诊断
                int frameCount = 0;
                object firstFrame = null;
                foreach (object f in frames)
                {
                    frameCount++;
                    if (firstFrame == null) firstFrame = f;
                }

                log.AppendLine("  GetAllSystemFrames 返回 " + frameCount + " 项");

                if (firstFrame != null)
                {
                    string firstType = firstFrame.GetType().Name;
                    string pickName = FrameName(firstFrame);
                    log.AppendLine("  首帧真实类型: " + firstType + "  名称: " + pickName);

                    TxTransformation t = GetFrameWorldTx(firstFrame);
                    if (t != null)
                    {
                        return FromTx(t, "FirstFrame(" + pickName + ")");
                    }

                    log.AppendLine("  首帧 [" + pickName + "] (" + firstType + ") 无法读取世界矩阵");
                }

                log.AppendLine("  系统坐标系中未找到有效的坐标");
            }

            // 兜底诊断：dump 实例参数，便于定位
            DumpInstanceParams(robot, log);
            return new Pose6 { Source = "BASE0(未读取)" };
        }

        // 系统坐标系名（可能带机器人前缀）
        private static string FrameName(object f)
        {
            try { var n = ((dynamic)f).Name as string; if (!string.IsNullOrEmpty(n)) return n; } catch { }
            return SafeName(f);
        }

        // 从坐标系对象提取世界系变换（不挑类型，全部走 dynamic 多属性兜底）
        // ★ 不再优先用 (frame as TxFrame).AbsoluteLocation ——
        //   default 控制器返回的对象不一定强类型为 TxFrame，
        //   用 as TxFrame 筛会导致首帧被丢掉 → "未找到有效坐标"
        private static TxTransformation GetFrameWorldTx(object frame)
        {
            if (frame == null) return null;

            // 优先路径：强类型 TxFrame 仍尝试（如果确实是 TxFrame，取其 AbsoluteLocation 更可靠）
            var fr = frame as TxFrame;
            if (fr != null)
            {
                var t = SafeTx(() => fr.AbsoluteLocation);
                if (t != null) return t;
            }

            // 动态多属性兜底（非 TxFrame 或强类型返回 null 时）
            foreach (var m in new[] { "AbsoluteLocation", "LocationRelativeToWorld", "Location", "Transformation" })
            {
                var tx = AsTx(InvokeOrGetSafe(frame, m));
                if (tx != null) return tx;
            }
            try
            {
                var inner = ((dynamic)frame).Frame as TxFrame;
                if (inner != null) return SafeTx(() => inner.AbsoluteLocation);
            }
            catch { }
            return null;
        }

        // 诊断：dump 机器人实例参数名（BASE0 未命中时定位用）
        private static void DumpInstanceParams(TxRobot robot, StringBuilder log)
        {
            try
            {
                dynamic dr = robot;
                object allp = null;
                try { allp = dr.GetAllInstanceParameters(); } catch { }
                var ie = allp as System.Collections.IEnumerable;
                if (ie == null) { log.AppendLine("  [诊断] GetAllInstanceParameters 不可用"); return; }

                var names = new List<string>();
                foreach (object pr in ie)
                {
                    if (pr == null) continue;
                    string pn = null;
                    try { dynamic dp = pr; pn = (dp.Type as string) ?? (dp.Name as string); } catch { }
                    if (string.IsNullOrEmpty(pn)) pn = SafeName(pr);
                    if (!string.IsNullOrEmpty(pn)) names.Add(pn);
                }
                log.AppendLine("  [诊断] 实例参数: " + (names.Count > 0 ? string.Join(", ", names) : "(空)"));
            }
            catch (Exception e) { log.AppendLine("  [诊断] 异常: " + e.Message); }
        }

        // ----------------------------------------------------------------
        // 4) 对比
        // ----------------------------------------------------------------
        private static void Compare(RobotBaseResult r, double posTolMm, double rotTol)
        {
            // 对比基准：当前存储 BASE0 vs 品牌感知的期望 BASE0
            if (!r.Expected.PosValid)
            {
                r.Comparable = false;
                r.Verdict = "无法对比（缺少坐标）";
                r.DeltaPos = double.NaN;
                r.DeltaRot = double.NaN;
                return;
            }
            if (!r.Base0.PosValid)
            {
                // 无控制器 / 无 BASE0 帧：当前值不存在，但期望已算出，仅作提示（不参与同步）
                r.Comparable = false;
                r.Verdict = "无当前BASE0";
                r.DeltaPos = double.NaN;
                r.DeltaRot = double.NaN;
                return;
            }

            r.Comparable = true;
            double dx = r.Base0.X - r.Expected.X;
            double dy = r.Base0.Y - r.Expected.Y;
            double dz = r.Base0.Z - r.Expected.Z;
            r.DeltaPos = Math.Sqrt(dx * dx + dy * dy + dz * dz);

            if (r.Base0.RotValid && r.Expected.RotValid)
            {
                double drx = Math.Abs(NormDelta(r.Base0.RX - r.Expected.RX));
                double dry = Math.Abs(NormDelta(r.Base0.RY - r.Expected.RY));
                double drz = Math.Abs(NormDelta(r.Base0.RZ - r.Expected.RZ));
                r.DeltaRot = Math.Max(drx, Math.Max(dry, drz));
            }
            else r.DeltaRot = double.NaN;

            bool posOk = r.DeltaPos <= posTolMm;
            bool rotOk = double.IsNaN(r.DeltaRot) || r.DeltaRot <= rotTol;

            r.Verdict = (posOk && rotOk) ? "一致" : "存在偏差";
        }

        // ----------------------------------------------------------------
        // 工具方法
        // ----------------------------------------------------------------
        private static Pose6 FromTx(TxTransformation t, string source)
        {
            var p = new Pose6 { Source = source };
            if (t == null) return p;
            dynamic dt = t;

            // 平移
            try
            {
                dynamic v = dt.Translation;
                p.X = (double)v.X; p.Y = (double)v.Y; p.Z = (double)v.Z;
                p.PosValid = true;
            }
            catch { }

            // 旋转（RPY），多策略；self 与 base0 用同一函数，单位保持一致
            foreach (var m in new[] { "RotationRPY_ZYX", "RotationRPY", "GetRotationRPY" })
            {
                try
                {
                    dynamic rv = InvokeOrGet(t, m);
                    if (rv != null)
                    {
                        p.RX = (double)rv.X; p.RY = (double)rv.Y; p.RZ = (double)rv.Z;
                        p.RotValid = true;
                        break;
                    }
                }
                catch { }
            }
            return p;
        }

        private static TxTransformation AsTx(object o)
        {
            return o as TxTransformation;
        }

        private static object InvokeOrGetSafe(object obj, string name)
        {
            try { return InvokeOrGet(obj, name); } catch { return null; }
        }

        private static TxTransformation SafeTx(Func<TxTransformation> f)
        {
            try { return f(); } catch { return null; }
        }

        // 取属性（无方法调用）
        private static object TryGet(object obj, string name)
        {
            if (obj == null) return null;
            var pi = obj.GetType().GetProperty(name);
            return pi != null ? pi.GetValue(obj, null) : null;
        }

        // 先当属性、再当无参方法
        private static object InvokeOrGet(object obj, string name)
        {
            if (obj == null) return null;
            var t = obj.GetType();
            var pi = t.GetProperty(name);
            if (pi != null) return pi.GetValue(obj, null);
            var mi = t.GetMethod(name, Type.EmptyTypes);
            if (mi != null) return mi.Invoke(obj, null);
            return null;
        }

        private static string SafeName(object o)
        {
            try { var n = TryGet(o, "Name") as string; if (!string.IsNullOrEmpty(n)) return n; } catch { }
            try { return o.ToString(); } catch { return "(unnamed)"; }
        }

        private static double NormDelta(double d)
        {
            // 若 RPY 为角度，归一化到 [-180,180]；若为弧度该归一化无害（小角不受影响）
            while (d > 180) d -= 360;
            while (d < -180) d += 360;
            return d;
        }

        private static string FmtDelta(double d)
        {
            return double.IsNaN(d) ? "N/A" : d.ToString("F4");
        }
    }
}