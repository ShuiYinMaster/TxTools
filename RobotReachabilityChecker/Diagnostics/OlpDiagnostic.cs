// ============================================================================
// OlpDiagnostic.cs (v2 — 强类型版)
//
// 调试用：打印机器人的工具配置（Toolframe / MountedTools / Toolbox）。
//
// v2 变更：
//   - 全部使用 PS 强类型 API (TxRobot.Toolframe / MountedTools / Toolbox 等)
//   - 去除所有 dynamic / 反射
//   - 不再尝试探测 Location 上的未知 Tool/Frame 属性（v1 的反射部分整段删除）
// ============================================================================
using System;
using Tecnomatix.Engineering;
using TxTools.RobotReachabilityChecker.Services;

namespace TxTools.RobotReachabilityChecker.Diagnostics
{
    public static class OlpDiagnostic
    {
        public static void Run(ITxObject pickedOperation, ILogger log)
        {
            log = log ?? NullLogger.Instance;

            TxRobot robot = RobotFinder.FindAssociatedRobot(pickedOperation, (Ui.ReachabilityCheckerForm)log);

            // ── 已挂载工具列表 ────────────────────────────────────────
            try
            {
                TxObjectList mounted = robot.MountedTools;
                int n = mounted == null ? 0 : mounted.Count;
                log.Log($"[Tool配置] MountedTools.Count = {n}");
                for (int i = 0; i < n; i++)
                {
                    ITxObject t = mounted[i];
                    log.Log($"  [{i}] {t.Name}  ({t.GetType().Name})");
                }
            }
            catch (Exception ex)
            {
                log.Log($"[Tool配置] 读取 MountedTools 异常: {ex.Message}", "WARN");
            }

            // ── 工具箱 ────────────────────────────────────────────────
            try
            {
                // 1. 输出机器人基本信息
                log.Log($"[Tool配置] 机器人: {robot.Name}");
                log.Log($"   Baseframe  : {robot.Baseframe?.Name ?? "null"}");
                log.Log($"   Referenceframe: {robot.Referenceframe?.Name ?? "null"}");
                log.Log($"   Toolframe  : {robot.Toolframe?.Name ?? "null"}");
                log.Log($"   TCPF       : {robot.TCPF?.Name ?? "null"}");

                // 2. 获取当前安装的工具 (MountedTools)
                var mountedTools = robot.MountedTools;   // 类型通常为 TxObjectList
                if (mountedTools == null || mountedTools.Count == 0)
                {
                    log.Log("[Tool配置] 当前机器人未安装任何工具 (MountedTools 为空)");
                }
                else
                {
                    log.Log($"[Tool配置] 已安装工具数量: {mountedTools.Count}");
                    for (int i = 0; i < mountedTools.Count; i++)
                    {
                        ITxObject tool = mountedTools[i];
                        log.Log($"    工具[{i}] {tool.Name} ({tool.GetType().Name})");

                        // 如果工具是 TxRobotWorkTool，可以进一步获取其工作 frames
                        if (tool is TxRobotWorkTool workTool)
                        {
                            // TxRobotWorkTool 可能有 Frames 属性，需根据实际 API 调整
                            // 此处仅示例，若 workTool 无frames可直接跳过
                            log.Log($"        类型: TxRobotWorkTool");
                            // 例如: workTool.Frames?...
                        }
                    }
                }

                // 3. 如果原代码意图是获取某种“frames 集合”，可能是 robot 的某些子对象
                // 这里补充说明：MountedTools 就是当前安装的所有工具，每个工具可能有自己的 frame。
                // 若需要获取每个工具的详细frame，需进一步查询该工具对象的成员。
            }
            catch (Exception ex)
            {
                log.Log($"[Tool配置] 读取机器人工具信息异常: {ex.Message}", "WARN");
            }

            // ── 点位数 ────────────────────────────────────────────────
            try
            {
                var locs = LocationEnumerator.EnumerateLocations(pickedOperation, log);
                log.Log($"[OLP诊断] 操作下点位数 = {locs.Count}");
            }
            catch { }

            log.Log("════════════════════════════════════════", "OK");
        }
    }
}
