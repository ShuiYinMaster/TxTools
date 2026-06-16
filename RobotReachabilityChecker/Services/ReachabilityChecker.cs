// ============================================================================
// ReachabilityChecker.cs
//
// 核心可达性检查编排服务 — 对一个操作下的所有点位执行
// "取轴值 → 状态判定 → 余量校验" 流程。
//
// 取轴值策略（v4 — GetPoseAtLocation 为主路径）：
//
//   主路径：robot.GetPoseAtLocation(loc)
//     · PS 内部已用 location.RobotConfigurationData 计算出正确解
//     · 与 PS 手动显示完全一致（这是唯一可靠的真值来源）
//     · 多 Utool 场景下需要先切 TCPF 再调用，否则会取到错误工具上的解
//
//   兜底：IkSolver.TryIKWithTcpfSwitch（仅 GetPoseAtLocation 返回 null 时）
//     · 用于 PS 没缓存解的极个别情况
//
// 双重姿态备份：robot.CurrentPose + 每轴 CurrentValue，确保 RTCP 路径下
// 检查完毕后机器人回到初始姿态，不污染场景。
// ============================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using TxTools.RobotReachabilityChecker.Diagnostics;
using TxTools.RobotReachabilityChecker.Models;

namespace TxTools.RobotReachabilityChecker.Services
{
    /// <summary>检查参数（避免长参数列表）。</summary>
    public class CheckOptions
    {
        public bool   JointMarginCheckEnabled { get; set; } = true;
        public double JointMarginThreshDeg    { get; set; } = 10.0;
        public bool   TcpCheckEnabled         { get; set; } = false;
        public double TcpMarginMm             { get; set; } = 200.0;
        public bool   InterferenceEnabled     { get; set; } = false;
        public RobotBrand UserSelectedBrand   { get; set; } = RobotBrand.Auto;
    }

    public static class ReachabilityChecker
    {
        public static List<PathPointResult> Check(
            ITxObject preferredOp,
            CheckOptions options,
            Action<int, int> progress = null,
            ILogger log = null)
        {
            log = log ?? NullLogger.Instance;
            options = options ?? new CheckOptions();
            var results = new List<PathPointResult>();

            if (preferredOp == null)
            {
                log.Log("Check: 未提供 operation", "ERR");
                return results;
            }
            string operationName = preferredOp.Name ?? "(未命名)";
            log.Log($"开始检查：[{operationName}]");

            // 姿态备份（外层声明确保 catch 可访问）
            TxRobot robotRef = null;
            TxPoseData savedPose = null;
            double[] savedJointValues = null;

            try
            {
                TxDocument doc = TxApplication.ActiveDocument;
                if (doc == null) throw new InvalidOperationException("ActiveDocument 为 null");

                // ── 1. 解析机器人 ─────────────────────────────────────────
                TxRobot robot = RobotFinder.FindAssociatedRobot(preferredOp, doc, log);
                if (robot == null)
                    throw new InvalidOperationException(
                        $"无法从操作 [{operationName}] 找到关联机器人，请确认操作已分配到机器人");
                log.Log($"机器人: {robot.Name}", "DEBUG");

                int djCount = 0;
                try { djCount = robot.DrivingJoints?.Count ?? 0; } catch { }

                // ── 2. 枚举点位 ───────────────────────────────────────────
                var locs = LocationEnumerator.EnumerateLocations(preferredOp, log);
                if (locs.Count == 0)
                    throw new InvalidOperationException($"操作 [{operationName}] 下未找到路径点");

                // ── 3. 基线关节限位 ───────────────────────────────────────
                var jointLimits = JointLimitsService.GetJointLimits(robot, log);

                // ── 4. 保存初始姿态（双重保险）────────────────────────────
                robotRef = robot;
                try { savedPose = robot.CurrentPose; }
                catch (Exception ex) { log.Log($"保存初始姿态失败（非致命）: {ex.Message}", "WARN"); }
                try
                {
                    var joints = robot.Joints;
                    if (joints != null && joints.Count > 0)
                    {
                        savedJointValues = new double[joints.Count];
                        for (int i = 0; i < joints.Count; i++)
                        {
                            try { dynamic jt = joints[i]; savedJointValues[i] = (double)jt.CurrentValue; } catch { }
                        }
                    }
                }
                catch { }

                // 路径连续性锚点：首点 null（IkSolver 内部用 robot.Joints 兜底），
                // 每点检查成功后更新为该点解，供下一点选解使用
                double[] anchorDeg = null;

                RobotBrand brand = BrandResolver.Resolve(robot.Name, options.UserSelectedBrand);

                // 干涉检查准备：用户启用 + 找到或自动建立干涉对 → 后续才查询
                bool ifReady = false;
                if (options.InterferenceEnabled)
                {
                    log.Log("准备干涉检查…");
                    ifReady = InterferenceService.EnsureRobotHasCollisionPair(robot, doc, log);
                    if (!ifReady) log.Log("  干涉检查准备失败，本次跳过干涉判定", "WARN");
                }

                int idx = 1;
                int okA = 0, fail = 0;
                int collisionCount = 0;

                foreach (ITxRoboticLocationOperation loc in locs)
                {
                    // 检测点类型 (Weld / Via)
                    string ptType = "Via";
                    try
                    {
                        string tn = loc.GetType().Name;
                        if (tn.Contains("Weld") || tn.Contains("weld")) ptType = "Weld";
                        else
                        {
                            dynamic dl = loc;
                            string lt = dl.LocationType?.ToString() ?? "";
                            if (lt.Contains("Weld")) ptType = "Weld";
                        }
                    }
                    catch { }

                    var res = new PathPointResult
                    {
                        Index = idx++,
                        PointName = string.IsNullOrEmpty(loc.Name) ? $"P{idx - 1}" : loc.Name,
                        OperationName = operationName,
                        RobotName = robot.Name,
                        PointType = ptType
                    };

                    bool gotJoints = false;
                    double[] joints2 = null;
                    string errMsg = "";

                    // ── 主路径：GetPoseAtLocation
                    //   PS 内部已用 location 的 RobotConfigurationData 计算好正确解，
                    //   直接读最准确、与 PS 手动显示一致。
                    //   注意：多 Utool 场景下需要先把 TCPF 切到 location 的工具坐标，
                    //         GetPoseAtLocation 才能拿到正确结果。
                    Tecnomatix.Engineering.TxTransformation savedTCPF = null;
                    bool tcpfChanged = false;
                    try
                    {
                        try { savedTCPF = robot.TCPF.AbsoluteLocation; } catch { }

                        // 切 TCPF（若 location 绑定了 RRS_TOOL_FRAME）
                        try
                        {
                            var locTool = ToolFrameReader.ReadLocationToolFrame(loc);
                            if (locTool != null && savedTCPF != null)
                            {
                                var locToolAbs = locTool.AbsoluteLocation;
                                if (locToolAbs != null)
                                {
                                    robot.TCPF.AbsoluteLocation = locToolAbs;
                                    tcpfChanged = true;
                                }
                            }
                        }
                        catch (Exception exTcp)
                        {
                            log.Log($"  [{res.PointName}] TCPF切换异常（忽略）: {exTcp.Message}", "DEBUG");
                        }

                        // 取 PS 算好的姿态
                        try
                        {
                            TxPoseData pd = robot.GetPoseAtLocation(loc);
                            if (pd != null)
                            {
                                double[] extracted = PoseValueExtractor.TryExtractPoseValues(pd, djCount);
                                if (extracted != null && extracted.Length > 0)
                                {
                                    joints2 = IkSolver.NormalizeToDegrees(extracted);
                                    gotJoints = true;
                                    okA++;
                                    // 缓存原始 TxPoseData：双击驱动姿态时 robot.CurrentPose = pd
                                    // 一次性写入比逐 joint 写 CurrentValue 快 6 倍以上
                                    res.PoseDataRef = pd;
                                }
                            }
                        }
                        catch (Exception exA)
                        {
                            log.Log($"  [{res.PointName}] GetPoseAtLocation 异常: {exA.Message}", "DEBUG");
                            errMsg = exA.Message;
                        }
                    }
                    finally
                    {
                        // 恢复 TCPF
                        if (tcpfChanged && savedTCPF != null)
                        {
                            try { robot.TCPF.AbsoluteLocation = savedTCPF; } catch { }
                        }
                    }

                    // ── 兜底：IK 选解（仅当 GetPoseAtLocation 失败时）
                    //   理论上 GetPoseAtLocation 应该总返回 PS 内部的解，
                    //   但万一遇到 PS 没缓存的点（罕见）才走 IK 兜底。
                    if (!gotJoints)
                    {
                        try
                        {
                            ITxObject locObj = loc as ITxObject;
                            double[] vals0 = IkSolver.TryIKWithTcpfSwitch(
                                robot, locObj, loc, anchorDeg, djCount, out string err0, log);
                            if (vals0 != null && vals0.Length > 0)
                            {
                                joints2 = vals0; gotJoints = true; okA++;
                                log.Log($"  [{res.PointName}] IK 兜底成功", "DEBUG");
                            }
                            else
                            {
                                errMsg = string.IsNullOrEmpty(err0) ? "IK失败" : err0;
                            }
                        }
                        catch (Exception ex0)
                        {
                            log.Log($"  [{res.PointName}] IK 兜底异常: {ex0.Message}", "WARN");
                            errMsg = $"IK 兜底异常: {ex0.Message}";
                        }
                    }

                    // ── 填写结果
                    if (gotJoints && joints2 != null)
                    {
                        // IkSolver 已归一化到度，这里只取整
                        res.J1 = joints2.Length > 0 ? Math.Round(joints2[0], 2) : 0;
                        res.J2 = joints2.Length > 1 ? Math.Round(joints2[1], 2) : 0;
                        res.J3 = joints2.Length > 2 ? Math.Round(joints2[2], 2) : 0;
                        res.J4 = joints2.Length > 3 ? Math.Round(joints2[3], 2) : 0;
                        res.J5 = joints2.Length > 4 ? Math.Round(joints2[4], 2) : 0;
                        res.J6 = joints2.Length > 5 ? Math.Round(joints2[5], 2) : 0;

                        // 关键：更新 anchorDeg 为本点的解，下一点用它做连续性选解
                        anchorDeg = new double[] { res.J1, res.J2, res.J3, res.J4, res.J5, res.J6 };

                        // 逐点位限位读取（FANUC J2/J3 联动场景）
                        var perPointLimits = JointLimitsService.ReadLimitsAtPose(
                            robot, res.J1, res.J2, res.J3, res.J4, res.J5, res.J6, jointLimits, log);

                        // 计算最小余量
                        var marginRes = AxisAnalyzer.CalcJointMargins(perPointLimits, options.JointMarginThreshDeg,
                            res.J1, res.J2, res.J3, res.J4, res.J5, res.J6);
                        res.JointMargin = Math.Round(marginRes.minMargin, 1);

                        res.Status = AxisAnalyzer.AnalyzePoint(res, perPointLimits, options.JointMarginThreshDeg,
                            options.JointMarginCheckEnabled, brand, out string axisNote);
                        res.ErrorMessage = axisNote;

                        // TCP 余量
                        if (options.TcpCheckEnabled)
                        {
                            string tcpWarn = TcpMarginChecker.CheckTcpXyzMargin(robot, loc, options.TcpMarginMm, log);
                            if (!string.IsNullOrEmpty(tcpWarn))
                            {
                                if (res.Status == ReachabilityStatus.Reachable
                                    || res.Status == ReachabilityStatus.Critical)
                                    res.Status = ReachabilityStatus.NearLimit;
                                res.ErrorMessage = string.IsNullOrEmpty(res.ErrorMessage)
                                    ? tcpWarn
                                    : res.ErrorMessage + "; " + tcpWarn;
                            }
                        }

                        // 静态干涉检查 — 驱动机器人到该姿态再查询
                        // 完整路径检查结束后由 RestoreRobotPose 统一恢复
                        if (ifReady && res.PoseDataRef is TxPoseData pdForIf)
                        {
                            try
                            {
                                robot.CurrentPose = pdForIf;
                                bool hit = InterferenceService.CheckCollisionAtCurrentPose(doc, log);
                                res.HasCollision = hit;
                                if (hit)
                                {
                                    collisionCount++;
                                    if (res.Status == ReachabilityStatus.Reachable
                                        || res.Status == ReachabilityStatus.Critical)
                                        res.Status = ReachabilityStatus.NearLimit;
                                    string note = "存在干涉";
                                    res.ErrorMessage = string.IsNullOrEmpty(res.ErrorMessage)
                                        ? note
                                        : res.ErrorMessage + "; " + note;
                                }
                            }
                            catch (Exception exIf)
                            {
                                log.Log($"  [{res.PointName}] 干涉查询异常: {exIf.Message}", "WARN");
                            }
                        }

                        if (res.Status == ReachabilityStatus.Unreachable) fail++;
                    }
                    else
                    {
                        res.Status = ReachabilityStatus.Unreachable;
                        res.ErrorMessage = string.IsNullOrEmpty(errMsg) ? "IK 无解" : errMsg;
                        fail++;
                        log.Log($"  [{res.PointName}] IK 失败: {res.ErrorMessage}", "ERR");
                        // 注：anchorDeg 不更新 — 跳过失败点，下一个点仍跟随最近一次成功解
                    }

                    results.Add(res);
                    try { progress?.Invoke(results.Count, locs.Count); } catch { }
                }

                // 恢复初始姿态
                RestoreRobotPose(robot, savedPose, savedJointValues);

                log.Log($"检查完成: 成功={okA} 失败={fail}"
                       + (options.InterferenceEnabled ? $" 干涉={collisionCount}" : ""), "OK");

                // 污染检测：所有点失败 → 弹窗指引手动恢复
                if (locs.Count > 0 && fail == locs.Count)
                {
                    try
                    {
                        MessageBox.Show(
                            "检测到所有点位均失败，可能是 PS 内部状态异常导致。\n\n" +
                            "请按以下步骤手动恢复后重试：\n" +
                            "  1. 切换到之前检查成功的任意路径\n" +
                            "  2. 在结果表格中，双击任意一个状态为「正常」或「接近极限」的行\n" +
                            "  3. 在弹出的 Robot Jog 窗口中点击「Close」关闭\n" +
                            "  4. 重新拾取当前路径并点击开始检查",
                            "需要手动恢复",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                log.Log($"检查异常: {ex.Message}", "ERR");
                log.Log($"  StackTrace: {ex.StackTrace?.Split('\n').FirstOrDefault()}", "ERR");
                try { RestoreRobotPose(robotRef, savedPose, savedJointValues); } catch { }
            }
            return results;
        }

        // =====================================================================
        // 恢复机器人姿态
        //
        // 主路径：直接写 joint.CurrentValue —— 与 Grid_AfterSelChange 单击驱动相同
        //         的低层 API，同步、立即生效，已被实操证实可靠。
        // 补充路径：robot.CurrentPose = savedPose —— 高层接口作为额外保险。
        // =====================================================================
        public static void RestoreRobotPose(TxRobot robot, TxPoseData savedPose, double[] savedJointValues)
        {
            if (robot == null) return;

            // 主路径：写每轴 CurrentValue
            try
            {
                if (savedJointValues != null)
                {
                    var joints = robot.Joints;
                    if (joints != null)
                    {
                        int n = Math.Min(joints.Count, savedJointValues.Length);
                        for (int i = 0; i < n; i++)
                        {
                            try { dynamic jt = joints[i]; jt.CurrentValue = savedJointValues[i]; } catch { }
                        }
                    }
                }
            }
            catch { }

            // 补充：CurrentPose 高层接口
            try { if (savedPose != null) robot.CurrentPose = savedPose; } catch { }
        }
    }
}
