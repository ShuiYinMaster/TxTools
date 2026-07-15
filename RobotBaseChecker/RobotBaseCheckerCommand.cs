using System;
using System.Threading;
using Tecnomatix.Engineering;

namespace TxTools.RobotBaseChecker
{
    // 注册要求：public 类，继承 TxButtonCommand，实现 Name / Category / Execute
    public class RobotBaseCheckerCommand : TxButtonCommand
    {
        public override string Name => ".BASE0 一致性检查校正";

        public override string Category => "TxTools";
        public override string Description => "检查BASE0是否符合标准，并一键校正";
        public override string LargeBitmap => "image.base0.png";

        public override void Execute(object cmdParams)
        {
            // 捕获 PS 主线程上下文，供后续 SDK 调用
            if (SynchronizationContext.Current == null)
                SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
            PsContext.Capture();

            try
            {
                if (RobotBaseCheckerForm.Instance == null || RobotBaseCheckerForm.Instance.IsDisposed)
                {
                    RobotBaseCheckerForm.Instance = new RobotBaseCheckerForm();
                    RobotBaseCheckerForm.Instance.FormClosed += (s, e) => RobotBaseCheckerForm.Instance = null;
                }
                RobotBaseCheckerForm.Instance.Show();   // 非模态，无 owner
                RobotBaseCheckerForm.Instance.Activate();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString(), "启动失败");
            }
        }
    }
}
