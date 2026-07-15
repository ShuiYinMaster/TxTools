using System;
using System.Windows.Forms;
using Tecnomatix.Engineering;

namespace TxTools.AutoPathPlanner
{
    /// <summary>TxTools 功能区注册入口 — 机器人组</summary>
    public class AutoPathPlannerCommand : TxButtonCommand
    {
        public override string Category { get { return "TxTools"; } }
        public override string Name { get { return "自动路径规划(RRT)"; } }

        public override void Execute(object cmdParams)
        {
            try
            {
                AutoPathPlannerForm.ShowSingleton();
            }
            catch (Exception ex)
            {
                MessageBox.Show("启动失败: " + ex.Message, "自动路径规划器",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
