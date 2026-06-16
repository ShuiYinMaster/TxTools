using System;
using System.Threading;
using System.Windows.Forms;
using Tecnomatix.Engineering;

namespace TxTools.AutoRecorder
{
    public class AutoRecorderCmd : TxButtonCommand
    {
        public override string Name { get { return ".操作录像"; } }
        public override string Category { get { return "TxTools"; } }
        public override string Tooltip { get { return "录制选定操作的仿真过程"; } }
        public override string Description { get { return "自动录制选定操作的仿真过程为视频文件"; } }
        public override string Bitmap { get { return base.Bitmap; } }
        public override string LargeBitmap { get { return "image.AutoRecorder.png"; } }

        public override void Execute(object cmdParams)
        {
            try
            {
                // 在 PS 主线程捕获 SynchronizationContext
                // TxSimulationPlayer 的事件回调线程不确定，需要靠 ctx 把回调切回主线程
                SynchronizationContext psCtx = SynchronizationContext.Current
                                              ?? new SynchronizationContext();

                // TxForm 在 PS 主线程创建和显示，由 PS 框架托管窗口层级
                AutoRecorderForm form = new AutoRecorderForm(psCtx);
                form.Show();   // 非模态

                try { TxApplication.StatusBarMessage = "操作录像插件已启动"; } catch { }
            }
            catch (Exception ex)
            {
                MessageBox.Show("启动失败：" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}