using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using TxTools.Common;

namespace TxTools.WeldAnnotator
{
    /// <summary>标注样式设置对话框。</summary>
    public class AnnotationStyleForm : Form
    {
        public AnnotationStyle Result { get; private set; }

        private static readonly string[] CATEGORY_OPTIONS_UI = {
            "二层板点焊", "二层板补焊",
            "三层板及以上点焊", "三层板及以上补焊",
            "焊缝", "CO2焊点",
            "螺母", "螺钉螺栓",
            "胶",
            "强度校验点", "重要特性", "关键特性"
        };

        private static readonly Tuple<string, int>[] SHAPE_OPTIONS = {
            Tuple.Create("圆形",     9),   // msoShapeOval
            Tuple.Create("矩形",     1),   // msoShapeRectangle
            Tuple.Create("三角形",   7),   // msoShapeIsoscelesTriangle
            Tuple.Create("菱形",     4),   // msoShapeDiamond
            Tuple.Create("圆角矩形", 5),   // msoShapeRoundedRectangle
            Tuple.Create("五角星",  12),   // msoShape5pointStar
        };

        public AnnotationStyleForm(AnnotationStyle src)
        {
            var catCopy = new Dictionary<string, int>();
            if (src.CategoryShapes != null)
                foreach (var kv in src.CategoryShapes) catCopy[kv.Key] = kv.Value;

            Result = new AnnotationStyle
            {
                DotColor       = src.DotColor,
                DotRadius      = src.DotRadius,
                LineColor      = src.LineColor,
                LineWidth      = src.LineWidth,
                BoxBorderColor = src.BoxBorderColor,
                BoxFillColor   = src.BoxFillColor,
                TextColor      = src.TextColor,
                TextFont       = src.TextFont,
                BoxPadding     = src.BoxPadding,
                OffsetX        = src.OffsetX,
                OffsetY        = src.OffsetY,
                CategoryShapes = catCopy
            };
            Text               = "标注样式设置";
            FormBorderStyle    = FormBorderStyle.FixedDialog;
            MaximizeBox        = false;
            MinimizeBox        = false;
            StartPosition      = FormStartPosition.CenterParent;
            Font               = new Font("Microsoft YaHei UI", 9f);
            AutoSize           = true;
            AutoSizeMode       = AutoSizeMode.GrowAndShrink;
            MinimumSize        = new Size(420, 0);
            Width              = 420;
            Build();
        }

        private void Build()
        {
            var t = new TableLayoutPanel
            {
                Dock = DockStyle.Top, ColumnCount = 2,
                AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(8, 6, 8, 6)
            };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            void Row(string lbl, Control ctrl)
            {
                t.Controls.Add(new Label
                {
                    Text = lbl, AutoSize = true, Anchor = AnchorStyles.Left,
                    Margin = new Padding(0, 6, 8, 6)
                });
                ctrl.Margin = new Padding(0, 3, 0, 3);
                t.Controls.Add(ctrl);
            }

            void Section(string title)
            {
                var header = new Label
                {
                    Text = title, AutoSize = true,
                    Font = new Font("Microsoft YaHei UI", 9.5f, FontStyle.Bold),
                    ForeColor = Color.FromArgb(0, 70, 127), Margin = new Padding(0, 10, 0, 4)
                };
                int rowIdx = t.RowCount;
                t.SetColumnSpan(header, 2);
                t.Controls.Add(header, 0, rowIdx);
                t.RowCount = rowIdx + 1;
            }

            Button Clr(Color init, Action<Color> set)
            {
                var b = new Button { Width = 80, Height = 22, BackColor = init };
                b.Click += (s, e) =>
                {
                    using (var cd = new ColorDialog { Color = b.BackColor })
                        if (cd.ShowDialog() == DialogResult.OK)
                        { b.BackColor = cd.Color; set(cd.Color); }
                };
                return b;
            }

            NumericUpDown Nud(decimal v, decimal mn, decimal mx, decimal inc, int dp, Action<decimal> fn)
            {
                var n = new NumericUpDown
                {
                    Width = 75, Minimum = mn, Maximum = mx,
                    Value = v, Increment = inc, DecimalPlaces = dp
                };
                n.ValueChanged += (s, e) => fn(n.Value);
                return n;
            }

            Button FontBtn(Font init, Action<Font> set)
            {
                var b = new Button
                {
                    Text = $"{init.Name}, {init.SizeInPoints}pt",
                    AutoSize = true, Height = 22,
                    Font = new Font("Microsoft YaHei UI", 8f), FlatStyle = FlatStyle.Flat
                };
                b.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
                b.Click += (s, e) =>
                {
                    using (var fd = new FontDialog { Font = init })
                        if (fd.ShowDialog() == DialogResult.OK)
                        { set(fd.Font); b.Text = $"{fd.Font.Name}, {fd.Font.SizeInPoints}pt"; }
                };
                return b;
            }

            Section("基础样式");
            Row("圆点颜色：",     Clr(Result.DotColor,       c => Result.DotColor = c));
            Row("圆点半径(px)：",  Nud(Result.DotRadius,       2, 30,   1,   0, v => Result.DotRadius = (int)v));
            Row("引线颜色：",     Clr(Result.LineColor,       c => Result.LineColor = c));
            Row("引线宽度(px)：",  Nud((decimal)Result.LineWidth, 1, 10, 0.5m, 1, v => Result.LineWidth = (float)v));
            Row("框边框颜色：",   Clr(Result.BoxBorderColor,  c => Result.BoxBorderColor = c));
            Row("框填充颜色：",   Clr(Result.BoxFillColor,    c => Result.BoxFillColor = c));
            Row("编号颜色：",     Clr(Result.TextColor,       c => Result.TextColor = c));
            Row("编号字体：",     FontBtn(Result.TextFont,    f => Result.TextFont = f));
            Row("框内边距(px)：", Nud(Result.BoxPadding,      1,  20,  1,   0, v => Result.BoxPadding = (int)v));
            Row("水平偏移(px)：", Nud(Result.OffsetX,         10, 300, 5,   0, v => Result.OffsetX = (int)v));
            Row("垂直偏移(px)：", Nud(Result.OffsetY,         10, 300, 5,   0, v => Result.OffsetY = (int)v));

            Section("分类形状（Excel 标记形状）");
            foreach (string cat in CATEGORY_OPTIONS_UI)
            {
                string catLocal = cat;
                var cmb = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList, Width = 130
                };
                foreach (var opt in SHAPE_OPTIONS) cmb.Items.Add(opt.Item1);

                if (!Result.CategoryShapes.TryGetValue(catLocal, out int currentId))
                    currentId = 9;
                int selIdx = 0;
                for (int i = 0; i < SHAPE_OPTIONS.Length; i++)
                    if (SHAPE_OPTIONS[i].Item2 == currentId) { selIdx = i; break; }
                cmb.SelectedIndex = selIdx;

                cmb.SelectedIndexChanged += (s, e) =>
                {
                    int idx = cmb.SelectedIndex;
                    if (idx >= 0 && idx < SHAPE_OPTIONS.Length)
                        Result.CategoryShapes[catLocal] = SHAPE_OPTIONS[idx].Item2;
                };
                Row(catLocal + "：", cmb);
            }

            Controls.Add(t);

            var bp = new Panel { Dock = DockStyle.Bottom, Height = 42 };
            var ok = new FormUiKit.FlatColorButton
            {
                Text = "确定", Width = 80, Height = 26, DialogResult = DialogResult.OK,
                Anchor = AnchorStyles.Right,
                BgColor = Color.FromArgb(0, 100, 167), ForeColor = Color.White,
                BorderColor = Color.FromArgb(0, 100, 167), FlatStyle = FlatStyle.Flat
            };
            var can = new FormUiKit.FlatColorButton
            {
                Text = "取消", Width = 80, Height = 26, DialogResult = DialogResult.Cancel,
                Anchor = AnchorStyles.Right,
                BgColor = Color.FromArgb(120, 124, 135), ForeColor = Color.White,
                BorderColor = Color.FromArgb(120, 124, 135), FlatStyle = FlatStyle.Flat
            };
            bp.Resize += (s, e) =>
            {
                can.Left = bp.ClientSize.Width - can.Width - 12;
                can.Top  = 8;
                ok.Left  = can.Left - ok.Width - 8;
                ok.Top   = 8;
            };
            ok.Click  += (s, e) => Close();
            can.Click += (s, e) => Close();
            bp.Controls.Add(ok);
            bp.Controls.Add(can);
            Controls.Add(bp);
            AcceptButton = ok;
            CancelButton  = can;
        }
    }
}
