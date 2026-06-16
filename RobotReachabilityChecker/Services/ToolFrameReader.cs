// ============================================================================
// ToolFrameReader.cs
//
// 从 location.Parameters 读取该点位绑定的工具坐标系（RRS_TOOL_FRAME）。
//
// Parameters 是 ArrayList，元素是 TxRoboticTxObjectParam（不强依赖基类接口）。
// 工具坐标项的 Type=="RRS_TOOL_FRAME"，Value 是 TxFrame。
//
// 注：无 RRS_TOOL_FRAME 时返回 null，调用方应不切 TCPF 用机器人当前默认工具。
// ============================================================================
using System.Collections;
using Tecnomatix.Engineering;

namespace TxTools.RobotReachabilityChecker.Services
{
    public static class ToolFrameReader
    {
        public const string TYPE_RRS_TOOL_FRAME = "RRS_TOOL_FRAME";

        public static TxFrame ReadLocationToolFrame(ITxObject loc)
        {
            if (loc == null) return null;

            // location 上的 Parameters 通过 dynamic 访问，兼容不同 location 子类型
            ArrayList paramList = null;
            try
            {
                dynamic d = loc;
                paramList = d.Parameters as ArrayList;
            }
            catch { }

            if (paramList == null || paramList.Count == 0) return null;

            for (int i = 0; i < paramList.Count; i++)
            {
                // 工具坐标项的具体类型是 TxRoboticTxObjectParam
                TxRoboticTxObjectParam objParam = paramList[i] as TxRoboticTxObjectParam;
                if (objParam == null) continue;
                if (objParam.Type != TYPE_RRS_TOOL_FRAME) continue;
                return objParam.Value as TxFrame;
            }
            return null;
        }
    }
}
