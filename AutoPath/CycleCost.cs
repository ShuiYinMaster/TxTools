using System;
using System.Collections;
using System.Collections.Generic;
using Tecnomatix.Engineering;

namespace TxTools.AutoPathPlanner
{
    /// <summary>
    /// 关节空间节拍代价 (v6.0)。
    ///
    /// ═══════════════════════════════════════════════════════════════
    /// 认知纠正 —— 之前所有"距离"判断都用 Vec3.Distance (笛卡尔 mm), 这是错的。
    ///
    /// 机器人 PTP (TxMotionType.Joint) 的运动时间 ≈
    ///
    ///     t = max_i ( |Δθi| / vi )        ← 由**最慢的那根轴**决定
    ///
    /// 跟 TCP 走了多少毫米几乎无关。举例:
    ///   路径 A: TCP 走 500mm, 但只是 J1 转 10°        → 快
    ///   路径 B: TCP 走 200mm, 但 J5 翻腕 90°          → 慢得多
    /// 用笛卡尔距离评价, 会选错 B。
    ///
    /// 因此: shortcut 是否更优、焊点如何排序、L1 深度如何取舍 —— 全部应该
    /// 以这个代价为准。
    /// ═══════════════════════════════════════════════════════════════
    /// </summary>
    public sealed class CycleCost
    {
        private readonly Action<string> _log;
        private double[] _jointSpeed;   // °/s (或 mm/s for prismatic)
        private int _n;

        /// <summary>关节速度读不到时的兜底值 (典型 6 轴点焊机器人)</summary>
        private static readonly double[] FallbackSpeed =
        {
            120,  // J1 底座回转
            110,  // J2 大臂
            120,  // J3 小臂
            190,  // J4 腕转
            190,  // J5 腕摆
            280   // J6 腕回转 (最快)
        };

        public CycleCost(TxRobot robot, Action<string> log)
        {
            _log = log ?? delegate { };
            ReadJointSpeeds(robot);
        }

        public int JointCount { get { return _n; } }

        /// <summary>
        /// 读取各关节最大速度。读不到就用典型值 ——
        /// 即使速度值不精确, 相对比例对了, 排序/取舍的结论依然正确。
        /// </summary>
        private void ReadJointSpeeds(TxRobot robot)
        {
            var speeds = new List<double>();
            try
            {
                var joints = ((dynamic)robot).Joints as IEnumerable;
                if (joints != null)
                {
                    foreach (object j in joints)
                    {
                        double v = 0;
                        foreach (string prop in new[] { "MaxSpeed", "MaximumSpeed", "Speed", "VelocityLimit" })
                        {
                            try
                            {
                                var pi = j.GetType().GetProperty(prop);
                                if (pi == null) continue;
                                object val = pi.GetValue(j, null);
                                if (val == null) continue;
                                v = Convert.ToDouble(val);
                                if (v > 0) break;
                            }
                            catch { }
                        }
                        speeds.Add(v);
                    }
                }
            }
            catch { }

            _n = speeds.Count;
            if (_n == 0)
            {
                _jointSpeed = (double[])FallbackSpeed.Clone();
                _n = _jointSpeed.Length;
                _log("  [节拍] 关节表不可读 — 采用典型 6 轴速度模型");
                return;
            }

            _jointSpeed = new double[_n];
            int missing = 0;
            for (int i = 0; i < _n; i++)
            {
                if (speeds[i] > 0)
                {
                    _jointSpeed[i] = speeds[i];
                }
                else
                {
                    _jointSpeed[i] = i < FallbackSpeed.Length ? FallbackSpeed[i] : 150.0;
                    missing++;
                }
            }

            if (missing == 0)
                _log(string.Format("  [节拍] 关节速度已读取 ({0} 轴)", _n));
            else
                _log(string.Format("  [节拍] {0}/{1} 轴速度不可读, 该部分用典型值", missing, _n));
        }

        /// <summary>
        /// PTP 运动时间估计 (秒): max_i(|Δθi| / vi)。
        /// 关节读不出来时返回 -1 (调用方应回退到笛卡尔度量)。
        /// </summary>
        public double PtpTime(TxPoseData a, TxPoseData b)
        {
            double[] ja = ReadJoints(a);
            double[] jb = ReadJoints(b);
            if (ja == null || jb == null) return -1;

            int n = Math.Min(Math.Min(ja.Length, jb.Length), _jointSpeed.Length);
            if (n == 0) return -1;

            double t = 0;
            for (int i = 0; i < n; i++)
            {
                double v = _jointSpeed[i];
                if (v <= 1e-6) continue;
                t = Math.Max(t, Math.Abs(jb[i] - ja[i]) / v);
            }
            return t;
        }

        /// <summary>关节最大跨度 (°) — 构型突变判定 / 粗略代价</summary>
        public static double MaxJointDelta(TxPoseData a, TxPoseData b)
        {
            double[] ja = ReadJoints(a);
            double[] jb = ReadJoints(b);
            if (ja == null || jb == null) return -1;

            int n = Math.Min(ja.Length, jb.Length);
            double m = 0;
            for (int i = 0; i < n; i++)
                m = Math.Max(m, Math.Abs(jb[i] - ja[i]));
            return m;
        }

        /// <summary>
        /// 链路总节拍 (秒)。任一段拿不到关节值则该段按笛卡尔距离折算
        /// (fallbackSpeedMmPerSec, 默认 1000mm/s) 兜底。
        /// </summary>
        public double ChainTime(IList<TxPoseData> poses, IList<Vec3> positions,
            double fallbackSpeedMmPerSec = 1000.0)
        {
            if (poses == null || poses.Count < 2) return 0;
            double total = 0;
            for (int i = 0; i < poses.Count - 1; i++)
            {
                double t = PtpTime(poses[i], poses[i + 1]);
                if (t < 0)
                {
                    if (positions != null && i + 1 < positions.Count)
                        t = Vec3.Distance(positions[i], positions[i + 1])
                            / Math.Max(1.0, fallbackSpeedMmPerSec);
                    else
                        t = 0;
                }
                total += t;
            }
            return total;
        }

        /// <summary>
        /// TxPoseData → double[]。
        /// v6.5: TxPoseData.JointValues 确认为 **ArrayList** (文档 + 实证)。
        /// </summary>
        public static double[] ReadJoints(TxPoseData p)
        {
            if (p == null) return null;
            try
            {
                ArrayList jv = p.JointValues;
                if (jv == null || jv.Count == 0) return null;

                var arr = new double[jv.Count];
                for (int i = 0; i < jv.Count; i++)
                    arr[i] = Convert.ToDouble(jv[i]);
                return arr;
            }
            catch { return null; }
        }
    }
}
