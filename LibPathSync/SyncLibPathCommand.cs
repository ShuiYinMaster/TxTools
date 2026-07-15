using System;
using System.Windows.Forms;
using Tecnomatix.Engineering;

namespace TxTools.LibPathSync
{
    /// <summary>
    /// 将所有目标对话框的“上次目录”同步为 System Root Path。
    /// 目标键集中维护在 LibPathTargets.Targets。
    /// </summary>
    public class SyncLibPathCommand : TxButtonCommand
    {
        public override string Name => ".路径同步";
        public override string Category => "TxTools";
        public override string Description => "将类型设置、插入组件、导入MFG等目录路径同步为 System Root Path";

        public override void Execute(object cmdParams)
        {
            try
            {
                string systemRoot = LibPathTargets.ReadValue(
                    LibPathTargets.SystemRootKey, LibPathTargets.SystemRootValue);

                if (string.IsNullOrEmpty(systemRoot))
                {
                    MessageBox.Show(
                        "未找到 System Root Path 或其值为空。\n\n注册表路径：\nHKCU\\" +
                        LibPathTargets.SystemRootKey,
                        "同步库路径", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                LibPathTargets.WriteAll(systemRoot);

                MessageBox.Show(
                    "已同步到 System Root：\n" + systemRoot + "\n\n" +
                    LibPathTargets.ReadBackSummary(),
                    "同步完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("同步失败：\n" + ex.Message,
                    "同步库路径", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}