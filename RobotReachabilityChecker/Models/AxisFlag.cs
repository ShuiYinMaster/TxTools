// ============================================================================
// AxisFlag.cs
//
// 轴级问题标志（与 ReachabilityStatus 相互独立，每个轴单独标记）。
//
// 单个 PathPointResult 包含 AxisFlags[6]，对应 J1..J6 各自的问题标记。
// 在表格渲染时根据每个轴的 Flag 进行单元格级染色（参考 ReachabilityCheckerForm.Grid.cs）。
// ============================================================================
using System;

namespace TxTools.RobotReachabilityChecker.Models
{
    [Flags]
    public enum AxisFlag
    {
        None      = 0,
        /// <summary>该轴落在临界带（仅 1/3/5 轴可能拥有）</summary>
        Critical  = 1 << 0,
        /// <summary>该轴距软限位 ≤ 阈值</summary>
        NearLimit = 1 << 1,
        /// <summary>仅 J5 可能拥有此标志（腕部奇异）</summary>
        Singular  = 1 << 2,
        /// <summary>该轴超出软限位</summary>
        OverLimit = 1 << 3,
    }
}
