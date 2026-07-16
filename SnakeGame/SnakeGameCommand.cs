using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Tecnomatix.Engineering;

namespace TxTools.SnakeGame
{
    /// <summary>
    /// TxTools.SnakeGame 插件入口命令。
    /// 由 PS 通过 TxButtonCommand 加载，菜单点击时弹出贪吃蛇窗体。
    /// </summary>
    [Guid("7C3E9A54-2B18-4F6D-8E33-9A1D42B7C605")]
    public class SnakeGameCommand : TxButtonCommand
    {
        public override string Category
        {
            get { return "TxTools"; }
        }

        public override string Name
        {
            get { return ".贪吃蛇小游戏"; }
        }

        public override string Description
        {
            get { return "在 Process Simulate 场景内玩一局贪吃蛇（几何体+干涉集）"; }
        }

        public override string LargeBitmap { get { return "image.snakegame.png"; } }

        public override void Execute(object cmdParams)
        {
            try
            {
                if (TxApplication.ActiveDocument == null)
                {
                    MessageBox.Show(
                        "请先打开一个 Process Simulate study。",
                        "TxTools · 贪吃蛇",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // 取当前 PS 进程的主窗口句柄作为 owner，确保 Z 序正确
                IntPtr psHwnd = IntPtr.Zero;
                try
                {
                    psHwnd = Process.GetCurrentProcess().MainWindowHandle;
                }
                catch { }

                var form = new SnakeGameForm();
                if (psHwnd != IntPtr.Zero)
                    form.Show(new WindowWrapper(psHwnd));
                else
                    form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "TxTools.SnakeGame 启动失败：" + ex.Message,
                    "TxTools.SnakeGame",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
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