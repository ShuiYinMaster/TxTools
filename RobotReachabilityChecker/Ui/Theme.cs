// ============================================================================
// Theme.cs
//
// 所有 UI 颜色常量集中定义。基于 PS SDK 的 TxColor，并提供 .Color 转换为
// System.Drawing.Color 供 WinForms 控件使用。
//
// 命名约定：
//   TxClrXxx (TxColor)       — SDK 类型，给 PS 原生控件用
//   ClrXxx   (System.Drawing.Color) — 给 WinForms 控件用
// ============================================================================
using System.Drawing;
using Tecnomatix.Engineering;

namespace TxTools.RobotReachabilityChecker.Ui
{
    internal static class Theme
    {
        // ── 主色调 ─────────────────────────────────────────────
        public static readonly TxColor TxClrAccent  = new TxColor(0,   70,  127);
        public static readonly TxColor TxClrSuccess = new TxColor(0,   128, 0);
        public static readonly TxColor TxClrDanger  = new TxColor(192, 0,   0);
        public static readonly TxColor TxClrWarning = new TxColor(160, 100, 0);

        // ── 功能区按钮色 ────────────────────────────────────────
        public static readonly TxColor TxClrBtnCheck  = new TxColor(0,   100, 167);
        public static readonly TxColor TxClrBtnAll    = new TxColor(0,   120, 90);
        public static readonly TxColor TxClrBtnExport = new TxColor(80,  80,  130);
        public static readonly TxColor TxClrBtnReset  = new TxColor(130, 100, 40);
        public static readonly TxColor TxClrBtnClose  = new TxColor(130, 50,  50);

        // ── 表格色 ─────────────────────────────────────────────
        public static readonly TxColor TxClrGridHeader     = new TxColor(218, 227, 243);
        public static readonly TxColor TxClrGridHeaderText = new TxColor(20,  20,  60);
        public static readonly TxColor TxClrGridAlt        = new TxColor(242, 244, 248);
        public static readonly TxColor TxClrGridHighlight  = new TxColor(189, 215, 238);
        public static readonly TxColor TxClrRowOk          = new TxColor(198, 239, 206);
        public static readonly TxColor TxClrRowFail        = new TxColor(255, 199, 206);
        public static readonly TxColor TxClrRowWarn        = new TxColor(255, 235, 156);
        public static readonly TxColor TxClrRowSingular    = new TxColor(248, 187, 208);  // 浅紫红
        public static readonly TxColor TxClrRowCritical    = new TxColor(207, 216, 220);  // 浅蓝灰

        // ── 单元格级（轴级）问题高亮色 ────────────────────────
        public static readonly TxColor TxClrCellOver     = new TxColor(198, 40,  40);   // 深红 — 轴超限（白字）
        public static readonly TxColor TxClrCellNear     = new TxColor(255, 179, 0);    // 橙黄 — 轴近极限
        public static readonly TxColor TxClrCellSingular = new TxColor(233, 30,  99);   // 紫红 — J5奇异（白字）
        public static readonly TxColor TxClrCellCritical = new TxColor(144, 164, 174);  // 蓝灰 — 临界

        // ── 日志面板 ───────────────────────────────────────────
        public static readonly TxColor TxClrLogBg   = new TxColor(30,  30,  30);
        public static readonly TxColor TxClrLogText = new TxColor(204, 204, 204);
        public static readonly TxColor TxClrLogErr  = new TxColor(255, 100, 100);
        public static readonly TxColor TxClrLogWarn = new TxColor(255, 200, 80);
        public static readonly TxColor TxClrLogOk   = new TxColor(80,  220, 120);

        // ── 点位编辑头栏 ────────────────────────────────────────
        public static readonly TxColor TxClrEditHeader = new TxColor(235, 241, 250);

        // ── WinForms 快捷引用（.Color 转换） ────────────────────
        public static readonly Color ClrAccent  = TxClrAccent.Color;
        public static readonly Color ClrSuccess = TxClrSuccess.Color;
        public static readonly Color ClrDanger  = TxClrDanger.Color;
        public static readonly Color ClrWarning = TxClrWarning.Color;
        public static readonly Color ClrMuted   = SystemColors.GrayText;
        public static readonly Color ClrText    = SystemColors.WindowText;
        public static readonly Color ClrBg      = SystemColors.Control;
    }
}
