using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;

namespace TxTools.FenceBuilder
{
    /// <summary>
    /// 入口命令。窗体非模态打开,允许用户在 PS 中继续选择对象。
    /// 
    /// TxButtonCommand 真实抽象成员:
    ///   - Name (命令内部标识,不是 InternalName)
    ///   - Category (菜单分组)
    ///   - Execute(object) (单个参数,不是 object[])
    /// 没有 DisplayName / InternalName / Tooltip(可能存在为 virtual,但不强求 override)
    /// </summary>
    public class FenceBuilderCommand : TxButtonCommand
    {
        public override string Name { get { return ".创建围栏"; } }
        public override string Description { get { return "创建围栏"; } }
        public override string Category { get { return "TxTools"; } }
        public override string LargeBitmap => "image.AutoFance.png";

        private static FenceBuilderForm _instance;

        public override void Execute(object cmdParams)
        {
            try
            {
                // 单例:已打开则前置,避免开多窗口
                if (_instance != null && !_instance.IsDisposed)
                {
                    _instance.BringToFront();
                    _instance.Activate();
                    return;
                }

                _instance = new FenceBuilderForm();
                _instance.FormClosed += (s, e) => { _instance = null; };

                // 把 PS 主窗口设为 owner,避免被 PS 主窗口遮挡;非模态 Show
                IntPtr psMainHwnd = IntPtr.Zero;
                try { psMainHwnd = Process.GetCurrentProcess().MainWindowHandle; } catch { }

                if (psMainHwnd != IntPtr.Zero)
                {
                    _instance.Show(new WindowWrapper(psMainHwnd));
                }
                else
                {
                    _instance.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "围栏生成器启动失败:\r\n" + ex.Message + "\r\n\r\n" + ex.StackTrace,
                    "FenceBuilder",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>把 HWND 包装成 IWin32Window 用于 Show(owner)</summary>
        private class WindowWrapper : IWin32Window
        {
            private readonly IntPtr _h;
            public WindowWrapper(IntPtr h) { _h = h; }
            public IntPtr Handle { get { return _h; } }
        }
    }
}
