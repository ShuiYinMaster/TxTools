// ============================================================================
// TcpMarginChecker.cs
//
// TCP XYZ 余量检测：沿 ±X / ±Y / ±Z 六个方向偏移 TCP，使用 GetPoseAtLocation
// 探测每个方向的最大可移动距离（mm），取最小值即为该点 TCP 余量。
//
// 性能优化：
//   1) 阈值短路：一旦确认"余量 ≥ 阈值"，直接返回 marginMm（无需精确数值）
//   2) 二分搜索代替线性扫描：30 步 → ~5 步，IK 调用次数下降 ~6 倍
//   3) joint save/restore 提到 CheckTcpXyzMargin 外层做一次（之前每次 TestIkAtOffset
//      都做，开销 30+ 倍）
// ============================================================================
using System;
using Tecnomatix.Engineering;
using TxTools.RobotReachabilityChecker.Diagnostics;

namespace TxTools.RobotReachabilityChecker.Services
{
    public static class TcpMarginChecker
    {
        public static string CheckTcpXyzMargin(
            TxRobot robot, ITxRoboticLocationOperation loc, double marginMm, ILogger log = null)
        {
            log = log ?? NullLogger.Instance;

            // 性能优化：joint save/restore 提到这一层（每点 1 次 vs 36+ 次）
            double[] savedJointVals = null;
            try
            {
                var joints = robot.Joints;
                if (joints != null && joints.Count > 0)
                {
                    savedJointVals = new double[joints.Count];
                    for (int i = 0; i < joints.Count; i++)
                    {
                        try { dynamic jt = joints[i]; savedJointVals[i] = (double)jt.CurrentValue; } catch { }
                    }
                }
            }
            catch { }

            try
            {
                TxTransformation baseTx = LocationGeometry.GetLocationTransform(loc, log);
                if (baseTx == null) return "";

                double[] basePos = LocationGeometry.ExtractTranslation(baseTx);
                if (basePos == null) return "";

                double[][] directions =
                {
                    new double[] { 1, 0, 0 }, new double[] { -1, 0, 0 },
                    new double[] { 0, 1, 0 }, new double[] {  0,-1, 0 },
                    new double[] { 0, 0, 1 }, new double[] {  0, 0,-1 }
                };
                string[] dirNames = { "+X", "-X", "+Y", "-Y", "+Z", "-Z" };

                double minMargin = double.MaxValue;
                string minDir = "";
                for (int d = 0; d < 6; d++)
                {
                    double maxDist = ProbeDirectionMargin(robot, loc, baseTx, basePos, directions[d], marginMm);
                    if (maxDist < minMargin) { minMargin = maxDist; minDir = dirNames[d]; }
                }

                if (minMargin < double.MaxValue && minMargin < marginMm)
                    return $"TCP余量 {minMargin:F0}mm ({minDir}方向) < 阈值 {marginMm:F0}mm";
            }
            catch (Exception ex)
            {
                log.Log($"    TCP余量检查异常: {ex.Message}", "WARN");
            }
            finally
            {
                if (savedJointVals != null)
                {
                    try
                    {
                        var joints = robot.Joints;
                        if (joints != null)
                        {
                            int n = Math.Min(joints.Count, savedJointVals.Length);
                            for (int i = 0; i < n; i++)
                            {
                                try { dynamic jt = joints[i]; jt.CurrentValue = savedJointVals[i]; } catch { }
                            }
                        }
                    }
                    catch { }
                }
            }
            return "";
        }

        // =====================================================================
        // 沿指定方向探测 TCP 的最大可移动距离（mm）
        // =====================================================================
        private static double ProbeDirectionMargin(TxRobot robot, ITxRoboticLocationOperation loc,
            TxTransformation baseTx, double[] basePos, double[] dir, double marginMm)
        {
            // 步骤 1：阈值距离测试 — 通过即返回（无需精确）
            if (TestIkAtOffset(robot, loc, baseTx, basePos, dir, marginMm))
                return marginMm;

            // 步骤 2：二分搜索 [0, marginMm)
            double lo = 0, hi = marginMm;
            while (hi - lo > 5)
            {
                double mid = (lo + hi) / 2;
                if (TestIkAtOffset(robot, loc, baseTx, basePos, dir, mid)) lo = mid;
                else hi = mid;
            }

            // 步骤 3：精化到 1mm
            double fineEnd = Math.Min(lo + 5, marginMm);
            double bestOk = lo;
            for (double dist = lo + 1; dist <= fineEnd; dist += 1)
            {
                if (TestIkAtOffset(robot, loc, baseTx, basePos, dir, dist)) bestOk = dist;
                else break;
            }
            return bestOk;
        }

        // =====================================================================
        // IK 测试：临时改写 location.AbsoluteLocation 为偏移矩阵，调
        // GetPoseAtLocation 测试 IK，恢复 location。
        // 支持普通场景和 RTCP 场景（PS 内部已封装 RTCP 几何转换）。
        // =====================================================================
        private static bool TestIkAtOffset(TxRobot robot, ITxRoboticLocationOperation loc,
            TxTransformation baseTx, double[] basePos, double[] dir, double dist)
        {
            if (loc == null || baseTx == null || basePos == null) return false;
            ITxLocatableObject locatable = loc as ITxLocatableObject;
            if (locatable == null) return false;

            TxTransformation savedAbs = null;
            bool written = false;
            try
            {
                try { savedAbs = locatable.AbsoluteLocation; } catch { return false; }
                if (savedAbs == null) return false;

                // 构造偏移矩阵
                TxTransformation offsetTx = LocationGeometry.CloneTransformation(baseTx);
                if (offsetTx == null) return false;

                dynamic dt = offsetTx;
                double newX = basePos[0] + dir[0] * dist;
                double newY = basePos[1] + dir[1] * dist;
                double newZ = basePos[2] + dir[2] * dist;

                bool wrote = false;
                try { dt[0, 3] = newX; dt[1, 3] = newY; dt[2, 3] = newZ; wrote = true; } catch { }
                if (!wrote) try { dt.X = newX; dt.Y = newY; dt.Z = newZ; wrote = true; } catch { }
                if (!wrote) return false;

                try { locatable.AbsoluteLocation = offsetTx; written = true; } catch { return false; }

                try
                {
                    TxPoseData pd = robot.GetPoseAtLocation(loc);
                    return pd != null;
                }
                catch { return false; }
            }
            finally
            {
                if (written && savedAbs != null)
                    try { locatable.AbsoluteLocation = savedAbs; } catch { }
            }
        }
    }
}
