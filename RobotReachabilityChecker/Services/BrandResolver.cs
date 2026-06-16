// ============================================================================
// BrandResolver.cs (v2 — 简化版)
//
// 机器人品牌识别。识别策略：
//   1) 用户在 UI 上手动选择 — 最高优先级
//   2) 从机器人名前缀猜测
//
// v2 变更：
//   - 删除 dynamic Controller 探测（PS 大多场景控制器都是 Default，
//     用 dynamic 访问即报异常即兜底，徒增开销，意义不大）
//   - 调用方不再需要传入 TxRobot 实例
// ============================================================================
using TxTools.RobotReachabilityChecker.Models;

namespace TxTools.RobotReachabilityChecker.Services
{
    public static class BrandResolver
    {
        /// <summary>解析机器人品牌。</summary>
        public static RobotBrand Resolve(string robotName, RobotBrand userSelectedBrand = RobotBrand.Auto)
        {
            if (userSelectedBrand != RobotBrand.Auto) return userSelectedBrand;

            string n = (robotName ?? "").ToUpper();
            if (n.Contains("KR")      || n.Contains("KUKA"))  return RobotBrand.KUKA;
            if (n.Contains("IRB")     || n.Contains("ABB"))   return RobotBrand.ABB;
            if (n.Contains("FANUC")   || n.Contains("R200"))  return RobotBrand.FANUC;
            return RobotBrand.Other;
        }

        /// <summary>从机器人名前缀粗略推断品牌（仅用于表格 "品牌" 列显示）</summary>
        public static string GuessBrandShort(string robotName)
        {
            string n = (robotName ?? "").ToUpper();
            if (n.Contains("KR")      || n.Contains("KUKA"))    return "KUKA";
            if (n.Contains("IRB")     || n.Contains("ABB"))     return "ABB";
            if (n.Contains("FANUC")   || n.Contains("R200"))    return "FANUC";
            if (n.Contains("YASKAWA") || n.Contains("MH"))      return "YASKAWA";
            if (n.Contains("BA")      || n.Contains("OTC"))     return "OTC";
            return "—";
        }
    }
}
