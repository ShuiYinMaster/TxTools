// ============================================================================
// ReachabilityCheckerForm.Grid.cs
//
// TxFlexGrid 相关：构建、自适应列宽、数据刷新、过滤。
//
// 列结构：
//   # | 品牌 | 机器人名 | 操作名 | 点名 | 点类型 | J1 | J2 | J3 | J4 | J5 | J6 | 检查结果 | 备注
//
// 染色策略：
//   1) 行级背景色：按整体 Status 着色
//   2) 单元格级染色：J1..J6 按 AxisFlags 单独染色（优先级：超限 > 奇异 > 近极限 > 临界）
// ============================================================================
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Tecnomatix.Engineering.Ui;
using Tecnomatix.Engineering.Ui.WPF;
using TxTools.RobotReachabilityChecker.Models;
using TxTools.RobotReachabilityChecker.Services;
using static TxTools.RobotReachabilityChecker.Ui.Theme;

namespace TxTools.RobotReachabilityChecker.Ui
{
    public partial class ReachabilityCheckerForm
    {
        // =====================================================================
        // 构建表格
        // =====================================================================
        private void BuildGrid()
        {
            _grid = new TxFlexGrid
            {
                Dock = DockStyle.Fill,
                AllowMerging = C1.Win.C1FlexGrid.AllowMergingEnum.None,
                SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row,
                AllowEditing = false,
                AllowSorting = C1.Win.C1FlexGrid.AllowSortingEnum.SingleColumn,
                ShowCursor = true,
                Font = SystemFonts.DefaultFont
            };

            _grid.Rows.Fixed = 1;
            _grid.Cols.Count = 14;

            string[] captions = { "#", "品牌", "机器人名", "操作名", "点名", "点类型",
                                  "J1(°)", "J2(°)", "J3(°)", "J4(°)", "J5(°)", "J6(°)",
                                  "检查结果", "备注" };

            for (int i = 0; i < 14; i++)
            {
                _grid.Cols[i].Caption = captions[i];
                _grid.Cols[i].AllowSorting = true;
            }

            // 用 Graphics 测量表头，确保列宽足以显示完整表头
            using (var g = CreateGraphics())
            {
                var hdrFont = new Font(SystemFonts.DefaultFont, FontStyle.Bold);
                for (int i = 0; i < 14; i++)
                {
                    int textW = (int)g.MeasureString(captions[i], hdrFont).Width + 16;
                    if (i == COL_OP) _grid.Cols[i].Width = 140;
                    else if (i == COL_PT) _grid.Cols[i].Width = 140;
                    else if (i == COL_NOTE) _grid.Cols[i].Width = 120;  // 在 Resize 中填充剩余
                    else _grid.Cols[i].Width = Math.Max(textW, 42);
                }
            }

            // 表头样式
            var hdrStyle = _grid.Styles[C1.Win.C1FlexGrid.CellStyleEnum.Fixed];
            hdrStyle.BackColor = TxClrGridHeader.Color;
            hdrStyle.ForeColor = TxClrGridHeaderText.Color;
            hdrStyle.Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold);

            _grid.Styles[C1.Win.C1FlexGrid.CellStyleEnum.Normal].BackColor = SystemColors.Window;
            _grid.Styles[C1.Win.C1FlexGrid.CellStyleEnum.Alternate].BackColor = TxClrGridAlt.Color;
            _grid.Styles[C1.Win.C1FlexGrid.CellStyleEnum.Highlight].BackColor = TxClrGridHighlight.Color;
            _grid.Styles[C1.Win.C1FlexGrid.CellStyleEnum.Highlight].ForeColor = SystemColors.WindowText;

            // 单击行 → 只更新状态栏（不再驱动姿态，避免与双击冲突）
            //          + 在 _tripleClickWindowMs 内连击双击行 → 触发三击（Robot Jog）
            _grid.AfterSelChange += Grid_AfterSelChange;
            _grid.MouseClick += Grid_SingleClickForTriple;

            // 双击行 → 驱动机器人到该点位姿态
            _grid.MouseDoubleClick += Grid_DblClick_Mouse;

            this.Resize += (s, e) => ResizeGridCols();
            Controls.Add(_grid);

            BuildRowStylePool();
        }

        // =====================================================================
        // 行/单元格样式池：每个状态预创建一个 Style 对象，刷新表格时直接复用，
        // 避免 RefreshGrid 中每行/每单元格都 _grid.Styles.Add 造成大量临时对象
        // =====================================================================
        private C1.Win.C1FlexGrid.CellStyle _rsOk, _rsFail, _rsWarn, _rsSing, _rsCrit, _rsAlt0, _rsAlt1;
        private C1.Win.C1FlexGrid.CellStyle _csOver, _csSing, _csNear, _csCrit;

        private void BuildRowStylePool()
        {
            _rsOk = _grid.Styles.Add("rsOk"); _rsOk.BackColor = TxClrRowOk.Color;
            _rsFail = _grid.Styles.Add("rsFail"); _rsFail.BackColor = TxClrRowFail.Color;
            _rsWarn = _grid.Styles.Add("rsWarn"); _rsWarn.BackColor = TxClrRowWarn.Color;
            _rsSing = _grid.Styles.Add("rsSing"); _rsSing.BackColor = TxClrRowSingular.Color;
            _rsCrit = _grid.Styles.Add("rsCrit"); _rsCrit.BackColor = TxClrRowCritical.Color;
            _rsAlt0 = _grid.Styles.Add("rsAlt0"); _rsAlt0.BackColor = SystemColors.Window;
            _rsAlt1 = _grid.Styles.Add("rsAlt1"); _rsAlt1.BackColor = TxClrGridAlt.Color;

            _csOver = _grid.Styles.Add("csOver"); _csOver.BackColor = TxClrCellOver.Color; _csOver.ForeColor = Color.White;
            _csSing = _grid.Styles.Add("csSing"); _csSing.BackColor = TxClrCellSingular.Color; _csSing.ForeColor = Color.White;
            _csNear = _grid.Styles.Add("csNear"); _csNear.BackColor = TxClrCellNear.Color;
            _csCrit = _grid.Styles.Add("csCrit"); _csCrit.BackColor = TxClrCellCritical.Color;
        }

        private C1.Win.C1FlexGrid.CellStyle PickRowStyle(ReachabilityStatus s, int rowIdx)
        {
            switch (s)
            {
                case ReachabilityStatus.Reachable: return _rsOk;
                case ReachabilityStatus.Unreachable: return _rsFail;
                case ReachabilityStatus.NearLimit: return _rsWarn;
                case ReachabilityStatus.Singular: return _rsSing;
                case ReachabilityStatus.Critical: return _rsCrit;
                default: return rowIdx % 2 == 0 ? _rsAlt0 : _rsAlt1;
            }
        }

        // =====================================================================
        // 自适应列宽
        // =====================================================================
        internal void ResizeGridCols()
        {
            if (_grid == null || _grid.Cols.Count < 14) return;
            try
            {
                int[] autoSizeCols = { COL_IDX, COL_BRAND, COL_ROBOT, COL_TYPE,
                                       COL_J1, COL_J2, COL_J3, COL_J4, COL_J5, COL_J6,
                                       COL_RESULT };
                foreach (int ci in autoSizeCols)
                {
                    try
                    {
                        _grid.AutoSizeCol(ci);
                        using (var g = CreateGraphics())
                        {
                            var hdrFont = new Font(SystemFonts.DefaultFont, FontStyle.Bold);
                            int hdrW = (int)g.MeasureString(_grid.Cols[ci].Caption, hdrFont).Width + 16;
                            if (_grid.Cols[ci].Width < hdrW) _grid.Cols[ci].Width = hdrW;
                        }
                    }
                    catch { }
                }

                int fixedTotal = 0;
                for (int i = 0; i < 13; i++) fixedTotal += _grid.Cols[i].Width;
                int remaining = _grid.ClientSize.Width - fixedTotal
                              - (SystemInformation.VerticalScrollBarWidth + 2);
                if (remaining > 60) _grid.Cols[COL_NOTE].Width = remaining;
            }
            catch { }
        }

        // =====================================================================
        // 刷新表格
        // =====================================================================
        internal void RefreshGrid(List<PathPointResult> results)
        {
            if (_grid == null) return;
            var filtered = ApplyFilter(results);

            // 性能优化：Redraw=false 包围批量赋值
            bool savedRedraw = _grid.Redraw;
            _grid.Redraw = false;
            try
            {
                _grid.Rows.Count = _grid.Rows.Fixed + filtered.Count;
                _rowToResult.Clear();

                for (int i = 0; i < filtered.Count; i++)
                {
                    var r = filtered[i];
                    int row = i + _grid.Rows.Fixed;
                    _rowToResult[row] = r;

                    string jStr(double v) => r.Status == ReachabilityStatus.NotChecked ? "" : v.ToString("F1");
                    string statusText = StatusToText(r.Status);

                    _grid[row, COL_IDX] = r.Index.ToString();
                    _grid[row, COL_BRAND] = BrandResolver.GuessBrandShort(r.RobotName);
                    _grid[row, COL_ROBOT] = r.RobotName;
                    _grid[row, COL_OP] = r.OperationName;
                    _grid[row, COL_PT] = r.PointName;
                    _grid[row, COL_TYPE] = r.PointType;
                    _grid[row, COL_J1] = jStr(r.J1);
                    _grid[row, COL_J2] = jStr(r.J2);
                    _grid[row, COL_J3] = jStr(r.J3);
                    _grid[row, COL_J4] = jStr(r.J4);
                    _grid[row, COL_J5] = jStr(r.J5);
                    _grid[row, COL_J6] = jStr(r.J6);
                    _grid[row, COL_RESULT] = statusText;
                    _grid[row, COL_NOTE] = r.ErrorMessage;

                    // 行背景色：从预制 Style 池中取，避免每行创建新 Style
                    _grid.Rows[row].Style = PickRowStyle(r.Status, i);

                    // 单元格级染色
                    ApplyCellAxisFlags(r, row);
                }
            }
            finally
            {
                _grid.Redraw = savedRedraw;
                _grid.Refresh();
            }

            ResizeGridCols();
        }

        private void ApplyCellAxisFlags(PathPointResult r, int row)
        {
            if (r.AxisFlags == null) return;
            int[] jCols = { COL_J1, COL_J2, COL_J3, COL_J4, COL_J5, COL_J6 };
            for (int ax = 0; ax < 6 && ax < r.AxisFlags.Length; ax++)
            {
                AxisFlag af = r.AxisFlags[ax];
                if (af == AxisFlag.None) continue;

                C1.Win.C1FlexGrid.CellStyle cs;
                if ((af & AxisFlag.OverLimit) != 0) cs = _csOver;
                else if ((af & AxisFlag.Singular) != 0) cs = _csSing;
                else if ((af & AxisFlag.NearLimit) != 0) cs = _csNear;
                else if ((af & AxisFlag.Critical) != 0) cs = _csCrit;
                else continue;

                _grid.SetCellStyle(row, jCols[ax], cs);
            }
        }

        private string StatusToText(ReachabilityStatus s)
        {
            switch (s)
            {
                case ReachabilityStatus.Reachable: return "正常";
                case ReachabilityStatus.Unreachable: return "不可达";
                case ReachabilityStatus.NearLimit: return "接近极限";
                case ReachabilityStatus.Singular: return "奇异";
                case ReachabilityStatus.Critical: return "临界";
                default: return "未检查";
            }
        }

        // =====================================================================
        // 过滤
        // =====================================================================
        private List<PathPointResult> ApplyFilter(List<PathPointResult> src)
        {
            var q = src.AsEnumerable();

            if (_chkHideNormal != null && _chkHideNormal.Checked)
                q = q.Where(r => r.Status != ReachabilityStatus.Reachable);

            if (_cbPointTypeFilter != null)
            {
                int idx = _cbPointTypeFilter.SelectedIndex;
                if (idx == 1) q = q.Where(r => r.PointType == "Weld");
                else if (idx == 2) q = q.Where(r => r.PointType == "Via");
            }

            return q.OrderBy(r => r.Index).ToList();
        }

        internal void ApplyFilterNow()
        {
            if (_currentTask != null) RefreshGrid(_currentTask.Results);
        }

        // 并发锁：防御性保留，避免后续误加调用路径时再次出现重入崩溃
        //         （旧版有 _selTimer 与 Picked 两条调用路径，已删 _selTimer）
        private bool _isPreviewing = false;

        // =====================================================================
        // 预览（拾取 OP 时显示该操作下的所有点位，未检查状态）
        // =====================================================================
        private void PreviewLocations(Tecnomatix.Engineering.ITxRoboticOperation op)
        {
            if (_grid == null || op == null) return;

            // 并发护栏 — 一次只允许一个 Preview 在执行
            if (_isPreviewing) return;
            _isPreviewing = true;

            // UI 线程保护 — 如果不在 UI 线程，转发到 UI 线程
            if (this.InvokeRequired)
            {
                try { this.BeginInvoke(new Action(() => { _isPreviewing = false; PreviewLocations(op); })); }
                catch { _isPreviewing = false; }
                return;
            }

            try
            {
                var locs = LocationEnumerator.EnumerateLocations(op as Tecnomatix.Engineering.ITxObject, this);

                bool savedRedraw = _grid.Redraw;
                _grid.Redraw = false;
                try
                {
                    _grid.Rows.Count = _grid.Rows.Fixed + locs.Count;
                    _rowToResult.Clear();

                    string robotName = "";
                    try
                    {
                        dynamic dop = op;
                        var r = dop.Robot;
                        if (r != null) robotName = r.Name as string ?? "";
                    }
                    catch { }

                    for (int i = 0; i < locs.Count; i++)
                    {
                        int row = i + _grid.Rows.Fixed;
                        var loc = locs[i];
                        string ptType = loc.GetType().Name.Contains("Weld") ? "Weld" : "Via";

                        _grid[row, COL_IDX] = (i + 1).ToString();
                        _grid[row, COL_BRAND] = "";
                        _grid[row, COL_ROBOT] = robotName;
                        _grid[row, COL_OP] = op.Name;
                        _grid[row, COL_PT] = loc.Name ?? $"P{i + 1}";
                        _grid[row, COL_TYPE] = ptType;
                        _grid[row, COL_RESULT] = "未检查";
                    }
                }
                finally
                {
                    _grid.Redraw = savedRedraw;
                    _grid.Refresh();
                }
                SetStatus($"路径 [{op.Name}]，{locs.Count} 个点位");
            }
            catch (Exception ex)
            {
                // 详细记录 — 避免下次再闪退时无线索
                Log($"PreviewLocations 异常: {ex.GetType().Name} - {ex.Message}", "ERR");
            }
            finally
            {
                _isPreviewing = false;
            }
        }

        // =====================================================================
        // 清空
        // =====================================================================
        private void ClearResults()
        {
            if (_grid != null)
            {
                _grid.Rows.Count = _grid.Rows.Fixed;
                _rowToResult.Clear();
            }
            _tasks.Clear();
            _currentTask = null;
            SetStatus("就绪");
        }
    }
}