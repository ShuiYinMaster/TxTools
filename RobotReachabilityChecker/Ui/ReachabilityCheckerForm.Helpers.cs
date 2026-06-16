// ============================================================================
// ReachabilityCheckerForm.Helpers.cs
//
// 小型 UI 控件工厂：MkLabel / MkButton / MkFuncButton + 自适应宽度辅助。
// 这些方法不持有状态，纯粹为 Layout 文件减负。
// ============================================================================
using System.Drawing;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using TxTools.Common;
using static TxTools.RobotReachabilityChecker.Ui.Theme;

namespace TxTools.RobotReachabilityChecker.Ui
{
    public partial class ReachabilityCheckerForm
    {
        // =====================================================================
        // 控件工厂
        // =====================================================================
        private Label MkLabel(string text) => new Label
        {
            Text = text,
            AutoSize = true,
            Font = SystemFonts.DefaultFont,
            ForeColor = ClrMuted,
            Margin = new Padding(0, 7, 4, 0)
        };

        private Button MkButton(string text, int width)
        {
            return new FormUiKit.FlatColorButton
            {
                Text = text,
                Width = width,
                Height = 26,
                FlatStyle = FlatStyle.Flat,
                BgColor = SystemColors.ControlLight,
                ForeColor = SystemColors.ControlText,
                BorderColor = SystemColors.ControlDark,
                Font = SystemFonts.DefaultFont,
                Margin = new Padding(0, 2, 4, 2)
            };
        }

        /// <summary>功能区按钮：自适应文本宽度、单行、带背景色</summary>
        private Button MkFuncButton(string text, Color bgColor)
        {
            return new FormUiKit.FlatColorButton
            {
                Text = text,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Height = 26,
                FlatStyle = FlatStyle.Flat,
                Font = SystemFonts.DefaultFont,
                BgColor = bgColor,
                ForeColor = Color.White,
                BorderColor = bgColor,
                Margin = new Padding(0, 2, 4, 2),
                Padding = new Padding(8, 2, 8, 2),
                Cursor = Cursors.Hand
            };
        }

        // =====================================================================
        // 自适应宽度
        // =====================================================================

        /// <summary>根据 ComboBox 最长项文本自动设定宽度</summary>
        private void AutoFitComboBoxWidth(ComboBox cb)
        {
            if (cb == null || cb.Items.Count == 0) return;
            using (var g = CreateGraphics())
            {
                float maxW = 0;
                foreach (var item in cb.Items)
                {
                    float w = g.MeasureString(item.ToString(), cb.Font).Width;
                    if (w > maxW) maxW = w;
                }
                // 加上下拉箭头宽度(20) + 边距(8)
                cb.Width = (int)maxW + 28;
            }
        }

        /// <summary>根据 NumericUpDown 的 Maximum 位数自动设定宽度</summary>
        private void AutoFitNumericWidth(NumericUpDown nud)
        {
            if (nud == null) return;
            string maxText = nud.Maximum.ToString("F" + nud.DecimalPlaces);
            using (var g = CreateGraphics())
            {
                float w = g.MeasureString(maxText, nud.Font).Width;
                // 加上上下箭头宽度(18) + 边距(8)
                nud.Width = (int)w + 26;
            }
        }
    }
}
