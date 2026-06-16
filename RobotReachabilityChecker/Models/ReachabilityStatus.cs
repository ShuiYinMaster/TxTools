// ============================================================================
// ReachabilityStatus.cs
//
// 点位可达性整体状态枚举（每个点位最终归类到其中一个）。
//
// 严重性顺序（从重到轻，数值越大越严重）：
//   Unreachable > Singular > NearLimit > Critical > Reachable > NotChecked
//
// 在 AxisAnalyzer.AnalyzePoint() 中按此顺序聚合各轴的 AxisFlag 得到最终状态。
// ============================================================================
namespace TxTools.RobotReachabilityChecker.Models
{
    public enum ReachabilityStatus
    {
        /// <summary>正常，所有轴均在安全范围内</summary>
        Reachable,

        /// <summary>临界点（低风险信息）— 1/3/5 轴落入品牌特定的 config 翻转临界带</summary>
        Critical,

        /// <summary>接近软限位 — 至少一个轴距软限位 ≤ 阈值</summary>
        NearLimit,

        /// <summary>奇异点（J5 ∈ ±10°，腕部奇异）</summary>
        Singular,

        /// <summary>不可达 / 超限 — IK无解或至少一个轴超出软限位</summary>
        Unreachable,

        /// <summary>未检查（默认状态，仅在预览阶段使用）</summary>
        NotChecked
    }
}
