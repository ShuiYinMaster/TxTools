// ============================================================================
// RobotPathCheckTask.cs
//
// 一次完整路径检查任务的聚合容器，包含路径标识、时间戳和所有点位结果。
//
// 派生统计字段（ReachableCount / NearLimitCount 等）用于汇总卡片显示，
// 临界点视为可达，计入 ReachabilityRate。
// ============================================================================
using System;
using System.Collections.Generic;
using System.Linq;

namespace TxTools.RobotReachabilityChecker.Models
{
    public class RobotPathCheckTask
    {
        public string RobotName { get; set; }
        public string PathName { get; set; }
        public DateTime CheckTime { get; set; }
        public List<PathPointResult> Results { get; set; } = new List<PathPointResult>();

        public int TotalPoints => Results.Count;
        public int ReachableCount   => Results.Count(r => r.Status == ReachabilityStatus.Reachable);
        public int UnreachableCount => Results.Count(r => r.Status == ReachabilityStatus.Unreachable);
        public int NearLimitCount   => Results.Count(r => r.Status == ReachabilityStatus.NearLimit);
        public int SingularCount    => Results.Count(r => r.Status == ReachabilityStatus.Singular);
        public int CriticalCount    => Results.Count(r => r.Status == ReachabilityStatus.Critical);

        /// <summary>
        /// 可达率：可达 + 接近极限 + 临界 三类视作可达。
        /// 不可达 / 奇异点 不计入分子。
        /// </summary>
        public double ReachabilityRate => TotalPoints > 0
            ? (double)(ReachableCount + NearLimitCount + CriticalCount) / TotalPoints * 100
            : 0;
    }
}
