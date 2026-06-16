// ============================================================================
// PathPointResult.cs
//
// 单个路径点位检查结果的数据载体，对应 UI 表格中的一行。
//
// 字段约定：
//   J1..J6        — 各轴角度（度），未检查时全为 0
//   JointMargin   — 各轴中最小余量（度），999 表示未计算
//   AxisFlags     — 长度 6 的轴级问题标记数组
//   Status        — 整体可达性状态（按 AxisFlags 中最严重项归类）
//   ErrorMessage  — 简短描述（如 "J5奇异" / "J6近极限(8°)"）
// ============================================================================
namespace TxTools.RobotReachabilityChecker.Models
{
    public class PathPointResult
    {
        public int Index { get; set; }
        public string PointName { get; set; }
        public string OperationName { get; set; }
        public string RobotName { get; set; } = "";
        public string PointType { get; set; } = "";
        public ReachabilityStatus Status { get; set; }

        public double J1 { get; set; }
        public double J2 { get; set; }
        public double J3 { get; set; }
        public double J4 { get; set; }
        public double J5 { get; set; }
        public double J6 { get; set; }

        public double JointMargin { get; set; } = 999;
        public string ErrorMessage { get; set; } = "";

        /// <summary>每个轴的问题标志（J1..J6 对应索引 0..5）</summary>
        public AxisFlag[] AxisFlags { get; set; } = new AxisFlag[6];

        /// <summary>
        /// 该点位是否发生干涉（PS CollisionRoot.HasCollidingObjects 结果）。
        /// 仅在用户启用"静态干涉检查"时计算，否则保持默认 false。
        /// 检测到 true 时，Status 会从 Reachable/Critical 升级为 NearLimit（警告级）。
        /// </summary>
        public bool HasCollision { get; set; } = false;

        /// <summary>
        /// 检查时缓存的 PS 姿态对象（实际类型 Tecnomatix.Engineering.TxPoseData）。
        /// 用 object 避免 Models 层依赖 PS API。
        ///
        /// 用于双击表格行驱动机器人时，一次性 robot.CurrentPose = pd 写入，
        /// 比逐 joint 写 CurrentValue 快 6 倍以上（PS 内部批量更新场景）。
        ///
        /// 内存代价：每个点位约 200~500 字节；1000 点位约 0.5MB，可接受。
        /// </summary>
        public object PoseDataRef { get; set; }
    }
}
