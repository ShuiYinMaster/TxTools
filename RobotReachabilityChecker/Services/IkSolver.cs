// ============================================================================
// IkSolver.cs
//
// 多 Utool 路径核心：TCPF 切换 + InverseFullReach IK + 选解 + 恢复。
//
// 选解算法 v3 — 解决"两次检查不一致 + 路径跳变"：
//
//   首点（anchorDeg == null）:
//     1. 严格匹配 location.RobotConfigurationData (OverheadState + 各轴 State)
//     2. 同组内 Turn 与 location.Turn 最匹配
//     3. 仍有多个 → L1 距离 robot.Joints 当前值最近 (确定性锚点)
//
//   后续点（anchorDeg != null，即上一个成功解）：
//     1. 严格匹配 location.RobotConfigurationData
//     2. 同组内 Turn 与 anchor Turn 一致的优先 (避免 ±360° 跳变)
//     3. 同组内 L1 距离 anchorDeg 最近 (路径姿态连续)
//
// 确定性保证：除首点用 robot.Joints 当前值锚定一次，整条路径不依赖机器人状态。
// 同一路径两次检查必产生同一组关节值。
// ============================================================================
using System;
using System.Collections;
using System.Collections.Generic;
using Tecnomatix.Engineering;
using TxTools.RobotReachabilityChecker.Diagnostics;

namespace TxTools.RobotReachabilityChecker.Services
{
    public static class IkSolver
    {
        /// <summary>
        /// TCPF 切换 + IK 求解 + 路径连续性选解。失败返回 null。
        /// </summary>
        /// <param name="anchorDeg">
        /// 路径上一个成功解的关节值（度）。首点传 null，由 robot.Joints 当前值兜底。
        /// </param>
        public static double[] TryIKWithTcpfSwitch(
            TxRobot robot, ITxObject loc, ITxRoboticLocationOperation locOp,
            double[] anchorDeg, int djCount, out string errMsg, ILogger log = null)
        {
            log = log ?? NullLogger.Instance;
            errMsg = "";
            if (robot == null || loc == null || locOp == null) return null;

            // 1) 保存 TCPF（必恢复）
            TxTransformation savedTCPF = null;
            bool tcpfChanged = false;
            try { savedTCPF = robot.TCPF.AbsoluteLocation; }
            catch (Exception ex) { errMsg = $"读 TCPF 失败: {ex.Message}"; return null; }

            try
            {
                // 2) 切换 TCPF (若 location 有 RRS_TOOL_FRAME)
                TxFrame locTool = ToolFrameReader.ReadLocationToolFrame(loc);
                if (locTool != null)
                {
                    try
                    {
                        var locToolAbs = locTool.AbsoluteLocation;
                        if (locToolAbs != null)
                        {
                            robot.TCPF.AbsoluteLocation = locToolAbs;
                            tcpfChanged = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        log.Log($"  TCPF切换异常({locTool.Name}): {ex.Message}", "DEBUG");
                    }
                }

                // 3) IK 调用
                ArrayList solutions = null;
                try
                {
                    var inv = new TxRobotInverseData();
                    inv.Destination = ((ITxLocatableObject)locOp).AbsoluteLocation;
                    inv.InverseType = TxRobotInverseData.TxInverseType.InverseFullReach;
                    solutions = robot.CalcInverseSolutions(inv);
                }
                catch (Exception ex)
                {
                    errMsg = $"CalcInverseSolutions异常: {ex.Message}";
                    return null;
                }

                if (solutions == null || solutions.Count == 0)
                {
                    errMsg = "IK无解";
                    return null;
                }

                // 4) 取 location 自带 config（决定性参考）
                TxRobotConfigurationData targetCfg = null;
                try { dynamic d = locOp; targetCfg = d.RobotConfigurationData as TxRobotConfigurationData; }
                catch { }

                // 5) 路径连续性选解
                //    首点：anchorDeg == null → 用 robot.Joints 当前值
                //    后续点：anchorDeg 即上一个成功解
                double[] effectiveAnchor = anchorDeg;
                if (effectiveAnchor == null)
                {
                    effectiveAnchor = ReadCurrentJointsDeg(robot, djCount);
                }

                TxPoseData chosenPose = PickWithContinuity(
                    solutions, targetCfg, effectiveAnchor, djCount, robot);
                if (chosenPose == null) { errMsg = "选解失败"; return null; }

                // 6) 提取关节值并归一化到度
                double[] vals = PoseValueExtractor.TryExtractPoseValues(chosenPose, djCount);
                return NormalizeToDegrees(vals);
            }
            finally
            {
                if (tcpfChanged && savedTCPF != null)
                {
                    try { robot.TCPF.AbsoluteLocation = savedTCPF; } catch { }
                }
            }
        }

        // =====================================================================
        // 路径连续性选解：3 层过滤 → 最终评分
        // =====================================================================
        public static TxPoseData PickWithContinuity(
            ArrayList solutions,
            TxRobotConfigurationData targetCfg,
            double[] anchorDeg,
            int djCount, TxRobot robot)
        {
            if (solutions == null || solutions.Count == 0) return null;

            // === 阶段1：按 OverheadState + 各轴 State 过滤 ============
            var stageA = new List<TxPoseData>();

            if (targetCfg != null && robot != null)
            {
                string[] targetStates = ExtractStates(targetCfg, djCount);
                int bestMatched = -1;
                var byMatchCount = new Dictionary<int, List<TxPoseData>>();

                foreach (object o in solutions)
                {
                    var pd = o as TxPoseData;
                    if (pd == null) continue;

                    int matched = 0;
                    bool ohMatch = false;
                    try
                    {
                        var cfg = robot.GetPoseConfiguration(pd);
                        if (cfg != null)
                        {
                            ohMatch = cfg.OverheadState == targetCfg.OverheadState;
                            string[] candStates = ExtractStates(cfg, djCount);
                            int n = Math.Min(targetStates.Length, candStates.Length);
                            for (int i = 0; i < n; i++)
                            {
                                if (string.IsNullOrEmpty(targetStates[i]) || string.IsNullOrEmpty(candStates[i])) continue;
                                if (targetStates[i] == candStates[i]) matched++;
                            }
                        }
                    }
                    catch { }

                    // OverheadState 失配的解直接降为 -1（最低优先级）
                    int score = ohMatch ? matched : -1;
                    if (!byMatchCount.ContainsKey(score)) byMatchCount[score] = new List<TxPoseData>();
                    byMatchCount[score].Add(pd);
                    if (score > bestMatched) bestMatched = score;
                }

                if (bestMatched >= 0 && byMatchCount.ContainsKey(bestMatched))
                    stageA = byMatchCount[bestMatched];
            }

            // 没有 config 信息或全部失配 → 用所有解
            if (stageA.Count == 0)
            {
                foreach (object o in solutions)
                {
                    var pd = o as TxPoseData;
                    if (pd != null) stageA.Add(pd);
                }
            }
            if (stageA.Count == 0) return null;
            if (stageA.Count == 1) return stageA[0];

            // === 阶段2：Turn 圈数匹配（avoid ±360° 跳变） ===========
            //   首点：与 location 自带 Turn 匹配
            //   后续：与 anchor 的等效圈数匹配（Turn = round((deg - 解中deg) / 360)）
            List<TxPoseData> stageB = stageA;

            if (targetCfg != null && robot != null)
            {
                int[] targetTurns = ExtractTurns(targetCfg, djCount);
                if (HasAnyTurn(targetTurns))
                {
                    stageB = FilterByTurns(stageA, targetTurns, djCount, robot);
                }
            }

            if (stageB.Count == 0) stageB = stageA;
            if (stageB.Count == 1) return stageB[0];

            // === 阶段3：L1 距离 anchorDeg 最近 ==========================
            if (anchorDeg == null || anchorDeg.Length == 0) return stageB[0];

            TxPoseData best = stageB[0];
            double bestDist = double.MaxValue;
            foreach (var pd in stageB)
            {
                double[] cand = PoseValueExtractor.TryExtractPoseValues(pd, djCount);
                cand = NormalizeToDegrees(cand);
                if (cand == null) continue;
                double d = L1Distance(cand, anchorDeg);
                if (d < bestDist) { bestDist = d; best = pd; }
            }
            return best;
        }

        // =====================================================================
        // 保留旧接口（向后兼容；内部调用 PickWithContinuity）
        // =====================================================================
        public static TxPoseData PickByConfig(ArrayList solutions,
            TxRobotConfigurationData targetCfg, double[] currentDeg, int djCount, TxRobot robot)
            => PickWithContinuity(solutions, targetCfg, currentDeg, djCount, robot);

        public static TxPoseData PickClosestSolution(ArrayList solutions,
            double[] currentDeg, int djCount)
            => PickWithContinuity(solutions, null, currentDeg, djCount, null);

        // =====================================================================
        // 按 Turn 匹配过滤
        // =====================================================================
        private static List<TxPoseData> FilterByTurns(
            List<TxPoseData> input, int[] targetTurns, int djCount, TxRobot robot)
        {
            var bestList = new List<TxPoseData>();
            int bestMatch = -1;

            foreach (var pd in input)
            {
                int matched = 0;
                try
                {
                    var cfg = robot.GetPoseConfiguration(pd);
                    if (cfg != null)
                    {
                        int[] candTurns = ExtractTurns(cfg, djCount);
                        int n = Math.Min(targetTurns.Length, candTurns.Length);
                        for (int i = 0; i < n; i++)
                        {
                            if (targetTurns[i] == -999 || candTurns[i] == -999) continue;
                            if (targetTurns[i] == candTurns[i]) matched++;
                        }
                    }
                }
                catch { }

                if (matched > bestMatch) { bestMatch = matched; bestList.Clear(); bestList.Add(pd); }
                else if (matched == bestMatch) bestList.Add(pd);
            }
            return bestList;
        }

        private static bool HasAnyTurn(int[] turns)
        {
            if (turns == null) return false;
            for (int i = 0; i < turns.Length; i++) if (turns[i] != -999) return true;
            return false;
        }

        // =====================================================================
        // 读取机器人当前关节值（度）
        // =====================================================================
        public static double[] ReadCurrentJointsDeg(TxRobot robot, int djCount)
        {
            if (robot == null) return null;
            try
            {
                var joints = robot.Joints;
                if (joints == null || joints.Count == 0) return null;
                int n = djCount > 0 ? Math.Min(djCount, joints.Count) : joints.Count;

                double[] vals = new double[n];
                for (int i = 0; i < n; i++)
                {
                    try { dynamic jt = joints[i]; vals[i] = (double)jt.CurrentValue; }
                    catch { vals[i] = 0; }
                }
                return NormalizeToDegrees(vals);
            }
            catch { return null; }
        }

        // =====================================================================
        // 度/弧度归一化
        // =====================================================================
        public static double[] NormalizeToDegrees(double[] vals)
        {
            if (vals == null || vals.Length == 0) return vals;
            double maxAbs = 0;
            for (int i = 0; i < vals.Length; i++)
                if (Math.Abs(vals[i]) > maxAbs) maxAbs = Math.Abs(vals[i]);
            bool isRad = maxAbs > 0 && maxAbs <= 2 * Math.PI + 0.05;
            if (isRad)
            {
                double k = 180.0 / Math.PI;
                double[] r = new double[vals.Length];
                for (int i = 0; i < vals.Length; i++) r[i] = vals[i] * k;
                return r;
            }
            return vals;
        }

        private static double L1Distance(double[] a, double[] b)
        {
            int n = Math.Min(a.Length, b.Length);
            double s = 0;
            for (int i = 0; i < n; i++) s += Math.Abs(a[i] - b[i]);
            return s;
        }

        // =====================================================================
        // 从 TxRobotConfigurationData 提取各轴 State
        // =====================================================================
        private static string[] ExtractStates(TxRobotConfigurationData cfg, int djCount)
        {
            var result = new string[djCount];
            if (cfg == null) return result;
            try
            {
                var jcs = cfg.JointsConfigurations;
                if (jcs == null) return result;
                int idx = 0;
                foreach (object jc in jcs)
                {
                    if (idx >= djCount) break;
                    try
                    {
                        dynamic d = jc;
                        var st = d.State;
                        if (st != null) result[idx] = st.ToString();
                    }
                    catch { }
                    idx++;
                }
            }
            catch { }
            return result;
        }

        // =====================================================================
        // 从 TxRobotConfigurationData 提取各轴 Turn
        // =====================================================================
        private static int[] ExtractTurns(TxRobotConfigurationData cfg, int djCount)
        {
            var result = new int[djCount];
            for (int i = 0; i < djCount; i++) result[i] = -999;
            if (cfg == null) return result;
            try
            {
                var jcs = cfg.JointsConfigurations;
                if (jcs == null) return result;
                int idx = 0;
                foreach (object jc in jcs)
                {
                    if (idx >= djCount) break;
                    try
                    {
                        dynamic d = jc;
                        var turn = d.Turn;
                        if (turn != null && turn.HasValue)
                            result[idx] = (int)turn.Value;
                    }
                    catch { }
                    idx++;
                }
            }
            catch { }
            return result;
        }
    }
}