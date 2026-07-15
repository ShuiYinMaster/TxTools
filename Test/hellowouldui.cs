using System.Drawing;
using System.Windows.Forms;
using Tecnomatix.Engineering.Ui;

namespace TxTools.HelloMulti
{
    internal class HelloTxForm : TxForm
    {
        public HelloTxForm(string title, int index)
        {
            // 关键:显式关闭半模态。默认很可能是 true,
            // 之前只是删掉了设 true 的行,从没真正设成 false。
            SemiModal = false;

            Text = title;
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96f, 96f);
            ClientSize = new Size(260, 120);
            MaximizeBox = false;
            Name = "TxTools.HelloMulti.HelloTxForm." + index;

            StartPosition = FormStartPosition.Manual;
            int off = (index % 12) * 28;
            Location = new Point(700 + off, 300 + off);

            var lbl = new Label
            {
                Text = title,
                Dock = DockStyle.Top,
                Height = 64,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 12f, FontStyle.Bold)
            };
            var btn = new Button { Text = "关闭", Dock = DockStyle.Bottom, Height = 36 };
            btn.Click += (s, e) => Close();

            Controls.Add(lbl);
            Controls.Add(btn);
        }
    }
}