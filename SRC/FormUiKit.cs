// ============================================================================
// FormUiKit.cs   —   TxTools 套件统一 GUI 规范（以 ExportGun 界面为准）
//
// 目的：
//   1) 五个插件窗体共用同一套"窗体初始化 + 卡片布局 + 配色控件"写法，
//      不再各写各的。
//   2) 彻底解决"先开 A 窗体、A 的尺寸会影响后开的 B 窗体"问题。
//
// 解耦原理（重点）：
//   PS 宿主只是 system-DPI-aware（非 PerMonitorV2），WinForms 的缩放基准在
//   进程内"第一个显示的窗口"握柄创建时锁定。若各窗体 AutoScaleMode/字体不一致，
//   或在 WinForms 自动缩放之外再手动乘系数，先后打开的窗体就会互相影响尺寸。
//   统一做法：
//     · 所有窗体都用 AutoScaleMode.None + OnLoad 手动 Scale（ExportGun 同款）
//     · 字体统一用 SystemFonts.DefaultFont
//     · 代码里一律写 96-DPI 裸像素，OnLoad 时按 DPI 系数一次性放大
//     · 严禁任何 ×_dpiScale / ApplyDpiFix / ScaleControlsRecursive 二次缩放
// ============================================================================
using System;
using System.Drawing;
using System.Windows.Forms;
using Tecnomatix.Engineering.Ui;

namespace TxTools.Common
{
    /// <summary>
    /// TxTools 所有 TxForm 窗体的统一外观/尺寸入口。
    /// 用法：在构造函数里、添加任何子控件之前，调用 InitStandardForm 一次即可。
    /// </summary>
    public static class FormUiKit
    {
        // ── 统一字体（OS 已 DPI 缩放，不要再自己乘系数）─────────────────────
        public static readonly Font BaseFont = SystemFonts.DefaultFont;
        public static readonly Font BoldFont =
            new Font(SystemFonts.DefaultFont, FontStyle.Bold);
        public static readonly Font TitleFont =
            new Font(SystemFonts.DefaultFont.FontFamily, 9f, FontStyle.Bold);

        // ── 统一配色（卡片标题用宿主主题蓝以外的稳定色，避免被 flat 皮肤吃掉）──
        public static readonly Color CardBack = Color.FromArgb(250, 251, 253);
        public static readonly Color CardBorder = Color.FromArgb(214, 219, 228);
        public static readonly Color TitleBack = Color.FromArgb(60, 90, 150);
        public static readonly Color TitleFore = Color.White;
        public static readonly Color PrimaryBack = Color.FromArgb(60, 90, 150);
        public static readonly Color PrimaryFore = Color.White;

        /// <summary>
        /// 窗体基础设置。必须在 BuildUI / 添加控件之前调用。
        /// 注意 size / minSize 全部写 96-DPI 下的裸像素，缩放交给 WinForms。
        ///
        /// 【串扰修复 —— ExportGun 同款方案】
        /// PS 的 TxForm 框架会把窗口尺寸/位置持久化到 PS 用户配置，
        /// 存储键就是窗体的 Control.Name。纯手写、从不设 Name 的窗体，Name 一直是
        /// 空字符串 ""，于是多个插件窗体全挤在同一个 "" 键上 —— 一个拉宽、其它全跟着变。
        /// ExportGunForm 的 InitializeComponent 里有一句 Name = "ExportGunForm"，
        /// 所以它有唯一键、独立记忆、互不干扰。这里给每个窗体设 Name = 类名，
        /// 等价地让每个窗体获得唯一持久化键：既消除串扰，又能各自记住上次尺寸。
        /// </summary>
        public static void InitStandardForm(TxForm form, string title,
                                            Size size, Size minSize,
                                            bool sizable = false)
        {
            form.SuspendLayout();

            // —— 唯一持久化键：消除 TxForm 跨插件几何串扰（核心）——
            // 用类型全名而非短名，避免不同命名空间下出现重名窗体时再次撞键。
            form.Name = form.GetType().FullName;

            // —— 缩放：AutoScale 关掉，由 OnLoad 手动 DPI 放大（ExportGun 同款）——
            // PS 宿主不会触发 WinForms 的 PerMonitorV2 缩放，AutoScaleMode.Dpi
            // 在进程内第一个窗体句柄创建时就被锁死，后续打开其他窗体不再生效。
            // 方案：设 None，在 OnLoad 中用 Graphics.DpiX / 96f 手动 Scale 一次，
            // 每次打开都是确定性放大，不与持久化尺寸耦合。
            form.AutoScaleDimensions = new SizeF(96F, 96F);
            form.AutoScaleMode = AutoScaleMode.None;
            form.Font = BaseFont;

            form.Text = title;
            form.BackColor = SystemColors.Control;
            form.StartPosition = FormStartPosition.CenterScreen;
            // 默认固定大小（ExportGun 同款 FixedDialog）
            form.FormBorderStyle = sizable ? FormBorderStyle.Sizable
                                           : FormBorderStyle.FixedDialog;
            form.MaximizeBox = sizable;
            form.MinimizeBox = true;
            form.MinimumSize = minSize;     // 必须先设 MinimumSize
            form.Size = size;        // 再设 Size，否则被 MinimumSize 钳制

            TrySetFlatStyleDisabled(form);   // 让自定义配色生效
            form.ResumeLayout(false);
        }

        /// <summary>
        /// 关闭 Siemens flat style 皮肤，否则它会重绘子控件、吃掉自定义配色。
        /// 用反射兼容不同 PS/SDK 版本（属性可能不存在）。
        /// </summary>
        public static void TrySetFlatStyleDisabled(TxForm form)
        {
            try
            {
                var p = form.GetType().GetProperty("FlatStyleEnabled");
                if (p != null && p.CanWrite) p.SetValue(form, false, null);
            }
            catch { /* 静默忽略，确保插件继续运行 */ }
        }

        // ====================================================================
        // DPI 缩放（ExportGun 同款：AutoScaleMode.None + OnLoad 手动 Scale）
        // ====================================================================

        /// <summary>当前 DPI 相对 96 的缩放系数。在 OnLoad 中调用一次，用于确定性缩放。</summary>
        public static float GetDpiScale(Control c)
        {
            try { using (var g = c.CreateGraphics()) return g.DpiX / 96f; }
            catch { return 1f; }
        }

        /// <summary>
        /// 在 OnLoad 中调用一次，完成 DPI 放大。已应用过则跳过（幂等）。
        ///
        /// 【防反复放大】base.OnLoad 会从 PS 持久化配置恢复上次的窗口尺寸，
        /// 若上次已是 DPI 放大后的尺寸，再次 Scale 会导致逐次变大。
        /// 因此在 Scale 前先将 Size 重置为 96-DPI 设计尺寸，确保每次都是
        /// 从同一基准出发做确定性放大。
        ///
        /// onScaled 回调传入系数，方便各窗体在缩放后微调列宽等。
        /// </summary>
        public static void ApplyDpiScaling(Form form, ref bool applied,
                                            Size designSize, Action<float> onScaled = null)
        {
            if (applied) return;
            applied = true;
            try
            {
                // 重置为设计尺寸，防止 TxForm 持久化尺寸叠加放大
                form.Size = designSize;
                float sc = GetDpiScale(form);
                if (sc < 1f) sc = 1f;
                if (sc > 1.01f) form.Scale(new SizeF(sc, sc));
                onScaled?.Invoke(sc);
            }
            catch { }
        }

        // ====================================================================
        // 三行根布局：标题栏(28) + 主体(100%) + 底栏(46)   —— ExportGun 同款
        // ====================================================================
        public static TableLayoutPanel BuildRoot(Control header, Control body, Control bottom)
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Margin = Padding.Empty,
                Padding = new Padding(6, 4, 6, 4),
                BackColor = SystemColors.Control
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));   // 标题栏
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));   // 主体
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 46F));   // 底栏
            if (header != null) root.Controls.Add(header, 0, 0);
            if (body != null) root.Controls.Add(body, 0, 1);
            if (bottom != null) root.Controls.Add(bottom, 0, 2);
            return root;
        }

        // ====================================================================
        // 卡片：左侧控制栏装多张卡片用 FlowLayoutPanel TopDown，绝不用 Dock=Top
        // ====================================================================
        /// <summary>纵向卡片容器（带滚动），固定宽度。</summary>
        public static FlowLayoutPanel BuildCardColumn(int width)
        {
            return new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Width = width,
                Padding = new Padding(4),
                BackColor = SystemColors.Control
            };
        }

        /// <summary>单张卡片：标题条 + 内容区。内容控件用 FlowDirection.TopDown 排。</summary>
        public static Panel MkCard(string title, int width, int contentHeight,
                                   out FlowLayoutPanel content)
        {
            var card = new Panel
            {
                Width = width,
                Height = contentHeight + 30,
                Margin = new Padding(2, 2, 2, 8),
                BackColor = CardBack,
                BorderStyle = BorderStyle.None
            };
            card.Paint += (s, e) =>
            {
                using (var pen = new Pen(CardBorder))
                    e.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1);
            };

            var titleBar = new Label
            {
                Text = title,
                Dock = DockStyle.Top,
                Height = 26,
                Font = TitleFont,
                BackColor = TitleBack,
                ForeColor = TitleFore,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0)
            };

            content = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = false,
                Padding = new Padding(8, 6, 8, 6),
                BackColor = CardBack
            };

            card.Controls.Add(content);
            card.Controls.Add(titleBar);
            return card;
        }

        // ====================================================================
        // 配色控件 —— PS 宿主会覆盖普通 Button/Label 的 BackColor，必须自绘
        // ====================================================================
        public static Button MkButton(string text, bool primary, int width = 0, int height = 30)
        {
            var b = new FlatColorButton
            {
                Text = text,
                Font = primary ? BoldFont : BaseFont,
                AutoSize = width <= 0,
                Height = height,
                Margin = new Padding(3),
                FlatStyle = FlatStyle.Flat,
                BackColor = primary ? PrimaryBack : SystemColors.ControlLight,
                ForeColor = primary ? PrimaryFore : SystemColors.ControlText
            };
            if (width > 0) b.Width = width;
            b.FlatAppearance.BorderColor = CardBorder;
            return b;
        }

        public static Label MkLabel(string text, bool bold = false)
        {
            return new FlatColorLabel
            {
                Text = text,
                AutoSize = true,
                Font = bold ? BoldFont : BaseFont,
                Margin = new Padding(3, 4, 3, 2),
                BackColor = Color.Transparent,
                ForeColor = SystemColors.ControlText
            };
        }

        // —— 自绘控件：用 UserPaint 绕过 PS flat 皮肤的重绘 ——
        public sealed class FlatColorButton : Button
        {
            /// <summary>快捷属性：等同 BackColor，语义更清晰。</summary>
            public Color BgColor { get => BackColor; set => BackColor = value; }
            /// <summary>快捷属性：等同 FlatAppearance.BorderColor。</summary>
            public Color BorderColor { get => FlatAppearance.BorderColor; set => FlatAppearance.BorderColor = value; }

            private bool _hover;
            private bool _pressed;

            public FlatColorButton()
            {
                SetStyle(ControlStyles.UserPaint
                       | ControlStyles.AllPaintingInWmPaint
                       | ControlStyles.OptimizedDoubleBuffer
                       | ControlStyles.ResizeRedraw, true);
                FlatAppearance.BorderColor = SystemColors.ControlDark;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                var rect = ClientRectangle;
                Color fill = BackColor;
                if (!Enabled) { /* 保持原色 */ }
                else if (_pressed)
                    fill = ControlPaint.Dark(BackColor, 0.15f);
                else if (_hover)
                    fill = ControlPaint.Light(BackColor, 0.20f);

                using (var bg = new SolidBrush(fill))
                    e.Graphics.FillRectangle(bg, rect);
                using (var pen = new Pen(FlatAppearance.BorderColor))
                    e.Graphics.DrawRectangle(pen, 0, 0, rect.Width - 1, rect.Height - 1);
                TextRenderer.DrawText(e.Graphics, Text, Font, rect, ForeColor,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter
                  | TextFormatFlags.EndEllipsis);
            }

            protected override void OnMouseEnter(EventArgs e) { _hover = true; Invalidate(); base.OnMouseEnter(e); }
            protected override void OnMouseLeave(EventArgs e) { _hover = false; _pressed = false; Invalidate(); base.OnMouseLeave(e); }
            protected override void OnMouseDown(MouseEventArgs e) { _pressed = true; Invalidate(); base.OnMouseDown(e); }
            protected override void OnMouseUp(MouseEventArgs e) { _pressed = false; Invalidate(); base.OnMouseUp(e); }
        }

        public sealed class FlatColorLabel : Label
        {
            public FlatColorLabel()
            {
                SetStyle(ControlStyles.UserPaint | ControlStyles.SupportsTransparentBackColor, true);
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                if (BackColor != Color.Transparent)
                    using (var bg = new SolidBrush(BackColor))
                        e.Graphics.FillRectangle(bg, ClientRectangle);
                TextRenderer.DrawText(e.Graphics, Text, Font, ClientRectangle, ForeColor,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter
                  | TextFormatFlags.WordBreak);
            }
        }

        public sealed class ColoredGroupBox : GroupBox
        {
            public Color HeaderColor { get; set; } = TitleBack;
            public ColoredGroupBox()
            {
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint
                       | ControlStyles.OptimizedDoubleBuffer, true);
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics;
                using (var bg = new SolidBrush(BackColor))
                    g.FillRectangle(bg, ClientRectangle);
                int top = Font.Height / 2;
                using (var pen = new Pen(CardBorder))
                    g.DrawRectangle(pen, 0, top, Width - 1, Height - top - 1);
                var sz = TextRenderer.MeasureText(Text, Font);
                using (var bg = new SolidBrush(BackColor))
                    g.FillRectangle(bg, 8, 0, sz.Width + 4, Font.Height);
                TextRenderer.DrawText(g, Text, Font, new Point(10, 0), HeaderColor);
            }
        }
    }
}