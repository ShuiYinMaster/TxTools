// ============================================================================
// AxisAnalyzer.cs
//
// 综合状态判定：超限 / 奇异 / 近极限 / 临界 / 可达
//
// 严重性顺序：Unreachable > Singular > NearLimit > Critical > Reachable
// 每个轴单独打 AxisFlag，但点位整体 Status 取所有轴中最严重的那个。
//
// 临界带（KUKA / ABB / FANUC / Other）：来自用户提供的临界点规则图
// 奇异点：J5 ∈ [-10°, +10°] 所有品牌通用
// ============================================================================
using System;
using System.Collections.Generic;
using System.Text;
using TxTools.RobotReachabilityChecker.Models;

namespace TxTools.RobotReachabilityChecker.Services
{
    public static class AxisAnalyzer
    {
        private static readonly Dictionary<RobotBrand, List<(double lo, double hi)>> CriticalBands
            = BuildCriticalBands();

        private static Dictionary<RobotBrand, List<(double lo, double hi)>> BuildCriticalBands()
        {
            var d = new Dictionary<RobotBrand, List<(double lo, double hi)>>();
            d.Add(RobotBrand.KUKA, new List<(double lo, double hi)> { (-5.0, 5.0) });
            d.Add(RobotBrand.ABB, new List<(double lo, double hi)>
            {
                (-275, -265), (-185, -175), (-95, -85), (-5, 5),
                (85, 95), (175, 185), (265, 275)
            });
            d.Add(RobotBrand.FANUC, new List<(double lo, double hi)>
            {
                (-185, -175), (175, 185)
            });
            d.Add(RobotBrand.Other, new List<(double lo, double hi)> { (-5, 5) });
            return d;
        }

        /// <summary>1/3/5 轴的临界带判定。</summary>
        public static bool IsInCriticalBand(double valDeg, int axisIdx, RobotBrand brand)
        {
            if (axisIdx != 0 && axisIdx != 2 && axisIdx != 4) return false;
            if (!CriticalBands.TryGetValue(brand, out var bands)) return false;
            for (int i = 0; i < bands.Count; i++)
                if (valDeg >= bands[i].lo && valDeg <= bands[i].hi) return true;
            return false;
        }

        public static bool IsJ5Singular(double j5Deg) => Math.Abs(j5Deg) <= 10.0;

        /// <summary>综合分析单个点位：填充 AxisFlags 并返回最终 Status。</summary>
        public static ReachabilityStatus AnalyzePoint(PathPointResult res,
            List<(double lo, double hi)> jointLimits, double nearThresh,
            bool jointCheckEnabled, RobotBrand brand, out string axisDetail)
        {
            axisDetail = "";
            double[] vals = { res.J1, res.J2, res.J3, res.J4, res.J5, res.J6 };
            res.AxisFlags = new AxisFlag[6];

            int worstLevel = 0;  // 0=Reachable 1=Critical 2=NearLimit 3=Singular 4=Unreachable
            string firstNote = "";

            for (int i = 0; i < 6; i++)
            {
                AxisFlag f = AxisFlag.None;
                double v = vals[i];

                // (1) 轴超限 / 近极限
                if (i < jointLimits.Count)
                {
                    var (lo, hi) = jointLimits[i];
                    if (lo != hi)
                    {
                        if (v < lo - 0.001 || v > hi + 0.001)
                        {
                            f |= AxisFlag.OverLimit;
                            if (worstLevel < 4) { worstLevel = 4; firstNote = $"J{i + 1}超限"; }
                        }
                        else if (jointCheckEnabled)
                        {
                            double margin = Math.Min(v - lo, hi - v);
                            if (margin < nearThresh)
                            {
                                f |= AxisFlag.NearLimit;
                                if (worstLevel < 2)
                                {
                                    worstLevel = 2;
                                    firstNote = $"J{i + 1}近极限({margin:F0}°)";
                                }
                            }
                        }
                    }
                }

                // (2) J5 奇异（优先级高于临界）
                if (i == 4 && IsJ5Singular(v))
                {
                    f |= AxisFlag.Singular;
                    if (worstLevel < 3) { worstLevel = 3; firstNote = "J5奇异"; }
                }
                else if (IsInCriticalBand(v, i, brand))
                {
                    f |= AxisFlag.Critical;
                    if (worstLevel < 1) { worstLevel = 1; firstNote = $"J{i + 1}临界"; }
                }

                res.AxisFlags[i] = f;
            }

            axisDetail = firstNote;
            switch (worstLevel)
            {
                case 4: return ReachabilityStatus.Unreachable;
                case 3: return ReachabilityStatus.Singular;
                case 2: return ReachabilityStatus.NearLimit;
                case 1: return ReachabilityStatus.Critical;
                default: return ReachabilityStatus.Reachable;
            }
        }

        /// <summary>计算各轴最小余量（度）。返回 (最小余量, 轴索引 1-6, 明细)。</summary>
        public static (double minMargin, int minAxis, string detail) CalcJointMargins(
            List<(double lo, double hi)> limits, double threshDeg, params double[] angles)
        {
            double minMargin = double.MaxValue;
            int minAxis = 0;
            var sb = new StringBuilder();
            int n = Math.Min(limits.Count, angles.Length);
            for (int i = 0; i < n; i++)
            {
                var (lo, hi) = limits[i];
                if (lo == hi) continue;
                double v = angles[i];
                double mLo = v - lo;
                double mHi = hi - v;
                double m = Math.Min(mLo, mHi);
                if (m < minMargin) { minMargin = m; minAxis = i + 1; }
                if (m < threshDeg)
                {
                    bool nearLo = mLo < mHi;
                    sb.Append($"J{i + 1}={v:F1}°(接近{(nearLo ? "下限" : "上限")} {(nearLo ? lo : hi):F1}°,余量{m:F1}°)  ");
                }
            }
            if (minMargin == double.MaxValue) minMargin = 999;
            return (minMargin, minAxis, sb.ToString().TrimEnd());
        }

        public static bool IsNearLimit(List<(double lo, double hi)> limits, params double[] angles)
        {
            int n = Math.Min(limits.Count, angles.Length);
            for (int i = 0; i < n; i++)
            {
                var (lo, hi) = limits[i];
                if (lo == hi) continue;
                double v = angles[i];
                if (Math.Min(v - lo, hi - v) < 10.0) return true;
            }
            return false;
        }
    }
}
