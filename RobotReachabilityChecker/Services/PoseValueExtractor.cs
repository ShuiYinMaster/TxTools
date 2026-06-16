// ============================================================================
// PoseValueExtractor.cs
//
// 从 TxPoseData 提取关节值数组。
//
// TxPoseData 在不同 PS 版本暴露的接口不同，提供 4 种策略链式回退：
//   策略1：pd.JointValues  (ArrayList, 官方文档属性)
//   策略2：pd.Values        (double[], 部分版本暴露)
//   策略3：pd[i] 索引器
//   策略4：pd.GetValues()
//
// 单位约定：旋转关节 PS 内部用弧度，由调用方判定/转换。
//
// 另提供 ReadDrivingJoints(robot) — 从 robot.DrivingJoints 读取当前值，
// 用于方式 B/C2 中"驱动到目标姿态后读关节"的场景。
// ============================================================================
using System;
using System.Collections;
using System.Collections.Generic;
using Tecnomatix.Engineering;

namespace TxTools.RobotReachabilityChecker.Services
{
    public static class PoseValueExtractor
    {
        /// <summary>
        /// 从 TxPoseData 提取关节值数组（原始单位 — 多为弧度，调用方按需转度）。
        /// 失败返回 null。expectedCount &lt;= 0 时取全部。
        /// </summary>
        public static double[] TryExtractPoseValues(TxPoseData pd, int expectedCount)
        {
            if (pd == null) return null;

            // 策略1：JointValues (官方文档：ArrayList of double)
            try
            {
                ArrayList raw = pd.JointValues;
                if (raw != null && raw.Count > 0)
                {
                    return ConvertArrayList(raw, expectedCount);
                }
            }
            catch { }

            // 策略2：Values (部分版本暴露 double[] 或 IEnumerable)
            try
            {
                dynamic dpd = pd;
                object vals = dpd.Values;
                double[] arr = ExtractDoubleArray(vals, expectedCount);
                if (arr != null && arr.Length > 0) return arr;
            }
            catch { }

            // 策略3：索引器 pd[i]
            try
            {
                dynamic dpd = pd;
                int n = expectedCount > 0 ? expectedCount : 6;
                var list = new List<double>();
                for (int i = 0; i < n; i++)
                {
                    try { list.Add(Convert.ToDouble(dpd[i])); }
                    catch { break; }
                }
                if (list.Count > 0) return list.ToArray();
            }
            catch { }

            // 策略4：GetValues()
            try
            {
                dynamic dpd = pd;
                object vals = dpd.GetValues();
                double[] arr = ExtractDoubleArray(vals, expectedCount);
                if (arr != null && arr.Length > 0) return arr;
            }
            catch { }

            return null;
        }

        /// <summary>从 robot.DrivingJoints 当前值读取关节数组。</summary>
        public static double[] ReadDrivingJoints(TxRobot robot)
        {
            if (robot == null) return null;
            try
            {
                TxObjectList dj = robot.DrivingJoints;
                if (dj == null || dj.Count == 0) return null;
                var list = new List<double>();
                foreach (ITxObject j in dj)
                {
                    double v = 0;
                    try { dynamic dj2 = j; v = Convert.ToDouble(dj2.Value); }
                    catch
                    {
                        try { dynamic dj2 = j; v = Convert.ToDouble(dj2.CurrentValue); } catch { }
                    }
                    list.Add(v);
                }
                return list.ToArray();
            }
            catch { return null; }
        }

        // =====================================================================
        // 辅助：从 ArrayList unbox 到 double[]，按 expectedCount 截断
        // =====================================================================
        private static double[] ConvertArrayList(ArrayList raw, int expectedCount)
        {
            int n = (expectedCount > 0 && expectedCount <= raw.Count) ? expectedCount : raw.Count;
            double[] vals = new double[n];
            for (int i = 0; i < n; i++)
            {
                object o = raw[i];
                if (o == null) { vals[i] = 0; continue; }
                vals[i] = (o is double) ? (double)o : Convert.ToDouble(o);
            }
            return vals;
        }

        // =====================================================================
        // 辅助：从任意对象（double[] / IEnumerable）提取到 double[]
        // =====================================================================
        public static double[] ExtractDoubleArray(object jv, int expectedCount = -1)
        {
            if (jv == null) return null;
            if (jv is double[] dArr)
            {
                if (expectedCount <= 0 || expectedCount >= dArr.Length) return dArr;
                double[] r = new double[expectedCount];
                Array.Copy(dArr, r, expectedCount);
                return r;
            }
            try
            {
                var list = new List<double>();
                foreach (object o in (IEnumerable)jv)
                {
                    list.Add(Convert.ToDouble(o));
                    if (expectedCount > 0 && list.Count >= expectedCount) break;
                }
                return list.Count > 0 ? list.ToArray() : null;
            }
            catch { return null; }
        }
    }
}
