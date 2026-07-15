using System;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;

namespace TxTools.WeldAnnotator
{
    /// <summary>焊点标注截图插件入口。</summary>
    public class WeldPointAnnotatorCmd : TxButtonCommand
    {
        public override string Name        => ".焊点标注截图";
        public override string Category    => "My Plugins";
        public override string Tooltip     => "焊点标注截图 → Excel";
        public override string Description => "将PS视口截图与焊点标注导出到活动Excel";
        public override string Bitmap      => "image.WeldAnnotator.bmp";
        public override string LargeBitmap => "image.WeldAnnotator.png";

        public override void Execute(object param)
        {
            try
            {
                var form = new WeldAnnotatorForm();
                form.Show();
                // 打开后立即把焦点还给 PS 主窗口
                try
                {
                    var ps = System.Diagnostics.Process.GetCurrentProcess();
                    if (ps != null && ps.MainWindowHandle != IntPtr.Zero)
                        SetForegroundWindow(ps.MainWindowHandle);
                }
                catch { }
            }
            catch (Exception ex)
            {
                TxMessageBox.ShowModal("插件启动失败：" + ex.Message, "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}
