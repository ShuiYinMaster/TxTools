using System;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;

namespace TxTools.MechArena
{
    /// <summary>
    /// PDPS Ribbon 入口命令：启动 MechArena 机器人竞技场。
    /// TxToolsInstaller 会通过 AssemblyInspector 反射扫描并注入 Ribbon。
    /// </summary>
    public class MechArenaStartCommand : TxButtonCommand
    {        // 必须实现抽象属性
        public override string Name => ".机器人竞技场";
        public override string Category => "TxTools";

        public override void Execute(object cmdParams)
        {
            try
            {
                var form = MechArenaForm.Instance;
                // 非模态：不阻塞 PS 主 UI
                try { form.SemiModal = false; } catch { }
                if (!form.Visible) form.Show();
                form.BringToFront();
                form.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "启动 MechArena 失败:\r\n" + ex.Message + "\r\n\r\n" + ex.StackTrace,
                    "MechArena",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
