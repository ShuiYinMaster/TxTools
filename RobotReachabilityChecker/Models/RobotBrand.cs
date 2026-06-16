// ============================================================================
// RobotBrand.cs
//
// 机器人品牌枚举。BrandResolver 服务通过控制器名 + 机器人名前缀识别品牌，
// 不同品牌使用不同的临界带定义（见 AxisAnalyzer.CriticalBands）。
// ============================================================================
namespace TxTools.RobotReachabilityChecker.Models
{
    public enum RobotBrand
    {
        Auto,
        KUKA,
        ABB,
        FANUC,
        Other
    }
}
