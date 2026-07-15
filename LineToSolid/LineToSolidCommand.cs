using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Tecnomatix.Engineering;

namespace TxTools.LineToSolid
{
    /// <summary>
    /// TxTools.LineToSolid 插件入口命令。
    /// 由 PS 通过 TxButtonCommand 加载，菜单点击时弹出窗体。
    /// </summary>
    [Guid("9A2B7F1C-3D88-4E2E-9A77-7E2B45C901AA")]
    public class LineToSolidCommand : TxButtonCommand
    {
        public override string Category
        {
            get { return "TxTools"; }
        }
        public override string Name
        {
            get { return ".创建线槽、水管"; }
        }

        public override string Description
        {
            get { return "以多段线/直线/圆弧特征为基线创建长方体/圆柱体"; }
        }

        public override string LargeBitmap { get { return "image.linetosoild.png"; } }

        public override void Execute(object cmdParams)
        {
            try
            {
                // 取当前 PS 进程的主窗口句柄作为 owner，确保 Z 序正确
                IntPtr psHwnd = IntPtr.Zero;
                try
                {
                    psHwnd = Process.GetCurrentProcess().MainWindowHandle;
                }
                catch { }

                var form = new LineToSolidForm();
                if (psHwnd != IntPtr.Zero)
                    form.Show(new WindowWrapper(psHwnd));
                else
                    form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("TxTools.LineToSolid 启动失败：" + ex.Message,
                    "TxTools.LineToSolid", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 把原生 HWND 包装成 IWin32Window，便于 form.Show(owner)。
        /// </summary>
        private class WindowWrapper : IWin32Window
        {
            private readonly IntPtr _h;
            public WindowWrapper(IntPtr h) { _h = h; }
            public IntPtr Handle { get { return _h; } }
        }
    }
}