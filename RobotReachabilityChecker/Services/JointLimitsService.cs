// ============================================================================
// JointLimitsService.cs
//
// 关节限位获取（度）。两套接口：
//   GetJointLimits(robot)                    — 基线一次性读取
//   ReadLimitsAtPose(robot, j1..j6, fb)      — 逐点位读取（解决 FANUC J2/J3 联动问题）
//
// 基线读取的多源回退策略：
//   源1: robot.Joints（运动学关节，最权威）
//   源2: robot.DrivingJoints + 反射枚举属性
//   源3: robot.GetParameter / GetAllInstanceParameters
//   源4: 兜底 [-360, 360]
//
// 单位策略：所有数值 |v| <= 2π+ε 视为弧度，统一乘以 180/π 转度。
// ============================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tecnomatix.Engineering;
using TxTools.RobotReachabilityChecker.Diagnostics;

namespace TxTools.RobotReachabilityChecker.Services
{
    public static class JointLimitsService
    {
        // =====================================================================
        // 基线限位读取
        // =====================================================================
        public static List<(double lo, double hi)> GetJointLimits(TxRobot robot, ILogger log = null)
        {
            log = log ?? NullLogger.Instance;
            var limits = new List<(double lo, double hi)>();
            if (robot == null) return DefaultLimits();

            int djCount = 0;
            try { djCount = robot.DrivingJoints?.Count ?? 6; } catch { djCount = 6; }
            if (djCount <= 0) djCount = 6;

            // ── 源1：robot.Joints ─────────────────────────────────────────
            try
            {
                var kinJoints = robot.Joints;
                if (kinJoints != null && kinJoints.Count > 0)
                {
                    log.Log($"  尝试 robot.Joints: 共 {kinJoints.Count} 个关节, 取前 {djCount} 个驱动轴");
                    int jIdx = 0;
                    foreach (object j in kinJoints)
                    {
                        if (jIdx >= djCount) break;
                        var lim = TryReadJointLimit(j, jIdx);
                        if (lim.HasValue) limits.Add(lim.Value);
                        jIdx++;
                    }
                    if (limits.Count > 0 && limits.Any(l => Math.Abs(l.hi - l.lo) < 719))
                    {
                        log.Log($"  从 robot.Joints 获取限位成功: {limits.Count} 轴", "OK");
                        EnsureDegrees(limits, log);
                        return limits;
                    }
                    limits.Clear();
                }
            }
            catch (Exception ex) { log.Log($"  robot.Joints 异常: {ex.Message}", "WARN"); }

            // ── 源2：DrivingJoints + 反射 ────────────────────────────────
            try
            {
                TxObjectList dj = robot.DrivingJoints;
                if (dj != null && dj.Count > 0)
                {
                    int jIdx = 0;
                    foreach (object j in dj)
                    {
                        var lim = TryReadJointLimit(j, jIdx);
                        if (lim.HasValue) limits.Add(lim.Value);
                        jIdx++;
                    }
                    if (limits.Count > 0 && limits.Any(l => Math.Abs(l.hi - l.lo) < 719))
                    {
                        log.Log($"  从 DrivingJoints 获取限位成功: {limits.Count} 轴", "OK");
                        EnsureDegrees(limits, log);
                        return limits;
                    }
                    limits.Clear();
                }
            }
            catch (Exception ex) { log.Log($"  DrivingJoints 反射异常: {ex.Message}", "WARN"); }

            // ── 源3：robot.GetParameter ──────────────────────────────────
            try
            {
                for (int i = 1; i <= 6; i++)
                {
                    double lo = -360, hi = 360;
                    string[] loNames = { $"J{i}_Min", $"j{i}_min", $"Joint{i}Min", $"A{i}_Min", $"joint{i}LowerLimit" };
                    string[] hiNames = { $"J{i}_Max", $"j{i}_max", $"Joint{i}Max", $"A{i}_Max", $"joint{i}UpperLimit" };
                    foreach (string n in loNames)
                        try { dynamic v = robot.GetParameter(n); lo = Convert.ToDouble(v); break; } catch { }
                    foreach (string n in hiNames)
                        try { dynamic v = robot.GetParameter(n); hi = Convert.ToDouble(v); break; } catch { }
                    limits.Add((lo, hi));
                }
                if (limits.Any(l => Math.Abs(l.hi - l.lo) < 719))
                {
                    log.Log("  从 GetParameter 获取限位成功", "OK");
                    EnsureDegrees(limits, log);
                    return limits;
                }
                limits.Clear();
            }
            catch (Exception ex) { log.Log($"  GetParameter 异常: {ex.Message}", "WARN"); }

            log.Log("  ⚠ 所有方式均无法获取关节限位，使用默认值 [-360, 360]", "WARN");
            return DefaultLimits();
        }

        // =====================================================================
        // 逐点位读取限位（处理 FANUC J2/J3 软限位联动）
        //
        // 流程：保存 robot.Joints CurrentValue → 切到目标姿态 → 读
        // LowerSoftLimit/UpperSoftLimit → finally 恢复
        // =====================================================================
        public static List<(double lo, double hi)> ReadLimitsAtPose(
            TxRobot robot, double j1Deg, double j2Deg, double j3Deg,
            double j4Deg, double j5Deg, double j6Deg,
            List<(double lo, double hi)> fallbackLimits, ILogger log = null)
        {
            log = log ?? NullLogger.Instance;
            if (robot == null) return fallbackLimits;

            var joints = robot.Joints;
            if (joints == null || joints.Count == 0) return fallbackLimits;

            int djCount = 0;
            try { djCount = robot.DrivingJoints?.Count ?? 6; } catch { djCount = 6; }
            if (djCount <= 0) djCount = 6;
            int n = Math.Min(djCount, joints.Count);

            // 保存原值
            var savedVals = new double[joints.Count];
            var savedOk = new bool[joints.Count];
            for (int i = 0; i < joints.Count; i++)
            {
                try { dynamic dj = joints[i]; savedVals[i] = (double)dj.CurrentValue; savedOk[i] = true; }
                catch { savedOk[i] = false; }
            }

            try
            {
                // 切到 IK 解姿态
                double[] targetDeg = { j1Deg, j2Deg, j3Deg, j4Deg, j5Deg, j6Deg };
                for (int i = 0; i < n && i < targetDeg.Length; i++)
                {
                    try
                    {
                        bool isRad = savedOk[i] && Math.Abs(savedVals[i]) <= 2 * Math.PI + 0.05;
                        double valToWrite = isRad ? targetDeg[i] * Math.PI / 180.0 : targetDeg[i];
                        dynamic jt = joints[i];
                        jt.CurrentValue = valToWrite;
                    }
                    catch { }
                }

                // 读该姿态下的软限位
                var perPoint = new List<(double lo, double hi)>();
                for (int i = 0; i < n; i++)
                {
                    try
                    {
                        dynamic jt = joints[i];
                        double lo = 0, hi = 0;
                        bool gotLo = false, gotHi = false;
                        try { lo = (double)jt.LowerSoftLimit; gotLo = true; } catch { }
                        try { hi = (double)jt.UpperSoftLimit; gotHi = true; } catch { }
                        if (gotLo && gotHi) perPoint.Add((lo, hi));
                        else perPoint.Add(GetFallback(fallbackLimits, i));
                    }
                    catch { perPoint.Add(GetFallback(fallbackLimits, i)); }
                }

                EnsureDegrees(perPoint, log);
                return perPoint;
            }
            finally
            {
                // 严格恢复机器人原姿态
                for (int i = 0; i < joints.Count; i++)
                {
                    if (!savedOk[i]) continue;
                    try { dynamic jt = joints[i]; jt.CurrentValue = savedVals[i]; } catch { }
                }
            }
        }

        // =====================================================================
        // 辅助
        // =====================================================================
        public static (double lo, double hi)? TryReadJointLimit(object joint, int index)
        {
            if (joint == null) return null;
            // 优先标准属性
            string[] loProps = { "LowerSoftLimit", "LowLimit", "MinValue", "LowerLimit", "Min" };
            string[] hiProps = { "UpperSoftLimit", "HighLimit", "MaxValue", "UpperLimit", "Max" };
            double lo = 0, hi = 0;
            bool gotLo = false, gotHi = false;

            foreach (string p in loProps)
            {
                try
                {
                    var pi = joint.GetType().GetProperty(p);
                    if (pi != null)
                    {
                        object v = pi.GetValue(joint);
                        if (v != null) { lo = Convert.ToDouble(v); gotLo = true; break; }
                    }
                }
                catch { }
            }
            foreach (string p in hiProps)
            {
                try
                {
                    var pi = joint.GetType().GetProperty(p);
                    if (pi != null)
                    {
                        object v = pi.GetValue(joint);
                        if (v != null) { hi = Convert.ToDouble(v); gotHi = true; break; }
                    }
                }
                catch { }
            }
            if (gotLo && gotHi) return (lo, hi);
            return null;
        }

        public static (double lo, double hi) GetFallback(List<(double lo, double hi)> fb, int i)
        {
            if (fb != null && i < fb.Count) return fb[i];
            return (-360, 360);
        }

        public static void EnsureDegrees(List<(double lo, double hi)> limits, ILogger log = null)
        {
            log = log ?? NullLogger.Instance;
            bool allSmall = limits.Count > 0 && limits.All(
                l => Math.Abs(l.lo) <= 2 * Math.PI + 0.5 && Math.Abs(l.hi) <= 2 * Math.PI + 0.5);
            if (allSmall)
            {
                for (int i = 0; i < limits.Count; i++)
                    limits[i] = (limits[i].lo * 180.0 / Math.PI, limits[i].hi * 180.0 / Math.PI);
            }
        }

        private static List<(double lo, double hi)> DefaultLimits()
        {
            var l = new List<(double lo, double hi)>();
            for (int i = 0; i < 6; i++) l.Add((-360, 360));
            return l;
        }
    }
}
