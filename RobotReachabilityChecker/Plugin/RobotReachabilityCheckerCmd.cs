// ============================================================================
// RobotReachabilityCheckerCmd.cs
//
// PS 插件入口。在 PS 启动时由插件加载器实例化，菜单 / 工具栏点击触发 Execute。
// ============================================================================
using System;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using TxTools.RobotReachabilityChecker.Ui;

namespace TxTools.RobotReachabilityChecker.Plugin
{
    public class RobotReachabilityCheckerCmd : TxButtonCommand
    {
        public override string Category    => "TxTools";
        public override string Name        => ".可达性检查";
        public override string Description => "机器人路径可达性检查工具";
        public override string Bitmap => base.Bitmap;
        public override string LargeBitmap => "image.RobotReachabilityChecker.png";

        public override void Execute(object cmdParams)
        {
            try
            {
                var form = new ReachabilityCheckerForm();
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
