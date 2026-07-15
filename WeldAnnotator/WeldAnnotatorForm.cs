// =============================================================================
// WeldAnnotatorForm.cs  v4  —  焊点标注截图导出插件（核心逻辑）
//
// 拆分为 5 个文件：
//   WeldPointAnnotatorCmd.cs   — 命令入口
//   AnnotationStyle.cs         — 样式数据模型
//   WeldAnnotatorForm.UI.cs    — UI 构建（partial）
//   AnnotationStyleForm.cs     — 样式设置对话框
//   WeldAnnotatorForm.cs       — 本文件：字段 + 构造 + 事件 + 业务逻辑
// =============================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;
using C1.Win.C1FlexGrid;
using TxTools.ExportGun;
using TxTools.Common;

using D  = System.Drawing;
using DI = System.Drawing.Imaging;
using WF = System.Windows.Forms;

namespace TxTools.WeldAnnotator
{
    public partial class WeldAnnotatorForm : TxForm
    {
        // ════════════════════════════════════════════════════════════════
        //  UI 控件字段
        // ════════════════════════════════════════════════════════════════

        // 卡片①：显示控制
        private WF.Button _btnSnap, _btnRestore, _btnShowOnly, _btnShowAll;
        private WF.Label  _lblSnapStatus;
        // 卡片②/③：导出设置
        private WF.Label    _lblCount;
        private WF.CheckBox _chkNewSheet, _chkWriteList, _chkAutoThick;
        // 卡片④：操作
        private WF.Button _btnExport, _btnClear, _btnStyle;

        private ToolStripProgressBar _progress;

        // 左侧 OP 选择
        private TxObjGridCtrl _opGrid;
        private WF.Label      _lblOpHint;

        // 焊点列表
        private TxFlexGrid _grid;

        // 状态栏 + 日志
        private StatusStrip          _status;
        private ToolStripStatusLabel _lblStatus;
        private WF.Panel             _logPanel;
        private RichTextBox          _logBox;
        private bool                 _logVisible;

        // 标注命名模式
        private WF.ComboBox _cmbLabelMode;
        private WF.TextBox  _txtPrefix, _txtSuffix;
        private WF.Label    _lblPrefix, _lblSuffix;

        // 共享 ToolTip
        private readonly WF.ToolTip _sharedToolTip = new WF.ToolTip();

        // ════════════════════════════════════════════════════════════════
        //  常量
        // ════════════════════════════════════════════════════════════════
        private const int C_IDX = 0, C_NAME = 1, C_OP = 2, C_X = 3, C_Y = 4, C_Z = 5, C_VIS = 6, C_TYPE = 7;

        private static readonly string[] CATEGORY_OPTIONS = {
            "二层板点焊", "二层板补焊",
            "三层板及以上点焊", "三层板及以上补焊",
            "焊缝", "CO2焊点",
            "螺母", "螺钉螺栓",
            "胶",
            "强度校验点", "重要特性", "关键特性"
        };
        private const string DEFAULT_CATEGORY = "二层板点焊";

        // ════════════════════════════════════════════════════════════════
        //  数据
        // ════════════════════════════════════════════════════════════════
        private List<WeldAnnotationPoint> _points         = new List<WeldAnnotationPoint>();
        private string                    _selOpName;
        private ITxObject                 _selOpObj;
        private List<Tuple<ITxObject, bool>> _dispSnapshot;
        private AnnotationStyle           _style          = new AnnotationStyle();
        private Dictionary<int, string>   _categories     = new Dictionary<int, string>();
        private HashSet<int>              _categoriesManuallyEdited = new HashSet<int>();

        private static readonly D.Size _myDefaultSize = new D.Size(1000, 620);
        private static readonly D.Size _myMinimumSize = new D.Size(820, 480);
        private readonly D.Font _font = new D.Font("Microsoft YaHei UI", 9f);
        private bool _dpiApplied;

        // ════════════════════════════════════════════════════════════════
        //  构造 + 生命周期
        // ════════════════════════════════════════════════════════════════
        public WeldAnnotatorForm()
        {
            SemiModal = false;
            FormUiKit.InitStandardForm(this, "焊点标注截图导出",
                _myDefaultSize, _myMinimumSize);

            Font = _font;
            ShowInTaskbar = true;
            PlaceTopRight();

            BuildMainArea();
            BuildLogPanel();
            BuildStatusBar();
            BuildCardPanel();

            this.FormClosing += WeldAnnotatorForm_FormClosing;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            FormUiKit.ApplyDpiScaling(this, ref _dpiApplied, _myDefaultSize);
        }

        public override void OnInitTxForm()
        {
            base.OnInitTxForm();
        }

        private void PlaceTopRight()
        {
            try
            {
                StartPosition = FormStartPosition.Manual;
                var scr = Screen.PrimaryScreen.WorkingArea;
                Location = new D.Point(
                    Math.Max(scr.Left, scr.Right - Width - 20),
                    Math.Max(scr.Top, scr.Top + 60));
            }
            catch { }
        }

        private void WeldAnnotatorForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                bool hasHide = PsReader.HasUnrestoredHide;
                bool hasSnap = _dispSnapshot != null && _dispSnapshot.Count > 0;
                if (!hasHide && !hasSnap) return;

                string msg =
                    "检测到场景中存在本插件产生的未恢复显示状态：\n\n" +
                    (hasHide ? "  • 有对象被'仅显示操作外观'隐藏，未恢复\n" : "") +
                    (hasSnap ? "  • 有已拍摄的快照，未恢复\n" : "") +
                    "\n关闭窗口前是否恢复？\n" +
                    "  [是] 恢复显示后关闭\n" +
                    "  [否] 直接关闭（场景保持当前状态）\n" +
                    "  [取消] 返回窗口";

                var r = MessageBox.Show(this, msg, "未恢复的显示状态",
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                if (r == DialogResult.Cancel) { e.Cancel = true; return; }
                if (r == DialogResult.Yes)
                {
                    if (hasSnap)
                        PsReader.RestoreDisplayStates(_dispSnapshot, s => Log("INFO", s));
                    else
                        PsReader.RestoreFromLastHide(s => Log("INFO", s));
                }
            }
            catch (Exception ex)
            {
                Log("ERR", "关闭前恢复检查异常：" + ex.Message);
            }
        }

        // ════════════════════════════════════════════════════════════════
        //  OP 选择事件
        // ════════════════════════════════════════════════════════════════
        private void OnOpInserted(object sender, TxObjGridCtrl_ObjectInsertedEventArgs e)
        {
            try
            {
                ITxObject txo = _opGrid.GetObject(0);
                if (txo == null) return;

                _selOpName = SafeGetName(txo);
                _selOpObj  = txo;
                _lblOpHint.Text      = _selOpName;
                _lblOpHint.ForeColor = D.Color.DarkGreen;
                SetStatus("已选择操作：" + _selOpName + "  正在读取焊点...");

                bool ok = ReadWeldPointsFromSelectedOp();
                ApplyAutoThicknessIfEnabled();
                RefreshGrid();
                SetStatus(ok
                    ? $"已选择 [{_selOpName}]，读取到 {_points.Count} 个焊点"
                    : $"已选择 [{_selOpName}]，但未读取到焊点（可能该 OP 下无焊点）");
            }
            catch (Exception ex) { Log("WARN", "OnOpInserted：" + ex.Message); }
        }

        private void OnOpDeleted(object sender, TxObjGridCtrl_RowDeletedEventArgs e)
        {
            try
            {
                if (_opGrid.Count > 0)
                {
                    ITxObject txo = _opGrid.GetObject(0);
                    if (txo != null)
                    {
                        _selOpName = SafeGetName(txo);
                        _selOpObj  = txo;
                        _lblOpHint.Text      = _selOpName;
                        _lblOpHint.ForeColor = D.Color.DarkGreen;
                        ReadWeldPointsFromSelectedOp();
                        RefreshGrid();
                        return;
                    }
                }
            }
            catch { }
            _selOpName = null;
            _selOpObj  = null;
            _lblOpHint.Text      = "选OP或视口选中";
            _lblOpHint.ForeColor = D.Color.Gray;
            _points.Clear();
            RefreshGrid();
        }

        // ════════════════════════════════════════════════════════════════
        //  快照控制
        // ════════════════════════════════════════════════════════════════
        private void BtnSnap_Click(object sender, EventArgs e)
        {
            SetStatus("拍摄快照中...");
            try
            {
                _dispSnapshot = PsReader.SnapshotDisplayStates(s => Log("INFO", s));
                int cnt = _dispSnapshot?.Count ?? 0;
                if (cnt == 0)
                {
                    _lblSnapStatus.Text = "快照为空";
                    _lblSnapStatus.ForeColor = D.Color.Gray;
                    _btnRestore.Enabled = false;
                    SetStatus("快照为空");
                    return;
                }
                _lblSnapStatus.Text      = $"已记录 {cnt} 对象";
                _lblSnapStatus.ForeColor = D.Color.DarkGreen;
                _btnRestore.Enabled      = true;
                SetStatus($"快照完成，记录 {cnt} 个对象显示状态");
            }
            catch (Exception ex) { Log("ERR", "拍摄快照失败：" + ex.Message); SetStatus("快照失败：" + ex.Message); }
        }

        private void BtnRestore_Click(object sender, EventArgs e)
        {
            SetStatus("恢复快照中...");
            try
            {
                if (_dispSnapshot != null && _dispSnapshot.Count > 0)
                {
                    PsReader.RestoreDisplayStates(_dispSnapshot, s => Log("INFO", s));
                    _lblSnapStatus.Text = "已恢复快照";
                    _lblSnapStatus.ForeColor = D.Color.DarkGreen;
                    SetStatus("快照已恢复");
                }
                else
                {
                    PsReader.RestoreFromLastHide(s => Log("INFO", s));
                    _lblSnapStatus.Text = "已恢复本次隐藏";
                    _lblSnapStatus.ForeColor = D.Color.Gray;
                    SetStatus("已恢复本次被隐藏的对象");
                }
            }
            catch (Exception ex) { Log("ERR", "恢复失败：" + ex.Message); SetStatus("恢复失败：" + ex.Message); }
        }

        private void BtnShowOnly_Click(object sender, EventArgs e)
        {
            if (_selOpObj == null) { SetStatus("请先在左侧选择操作节点"); return; }
            SetStatus("准备白名单...");

            var op = new OperationInfo
            {
                Name      = _selOpName ?? SafeGetName(_selOpObj),
                TypeLabel = _selOpObj.GetType().Name,
                PsObject  = _selOpObj
            };
            try { PsReader.FillPoints(op, PointType.WeldPoint, false, s => Log("INFO", s)); }
            catch (Exception ex) { Log("WARN", "FillPoints 异常：" + ex.Message); }

            if (op.Points == null || op.Points.Count == 0)
            {
                TxMessageBox.ShowModal($"操作 [{op.Name}] 下未找到任何焊点。", "未找到焊点",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SetStatus("未找到焊点");
                return;
            }

            var whitelist = new List<ITxObject>();
            TxRobot robot = null;
            try { robot = PsReader.FindRobotForOperation(_selOpObj); } catch { }
            if (robot != null) whitelist.Add(robot);
            if (robot != null)
            {
                ITxObject tool = null;
                try
                {
                    dynamic dr = robot;
                    try { tool = dr.ActiveTool as ITxObject; } catch { }
                    if (tool == null) try { tool = dr.CurrentTool as ITxObject; } catch { }
                }
                catch { }
                if (tool != null) whitelist.Add(tool);
            }

            int boundCount = 0;
            foreach (var pi in op.Points)
            {
                if (pi?.AllAppearances == null) continue;
                foreach (var ar in pi.AllAppearances)
                {
                    if (ar?.RawObject != null) { whitelist.Add(ar.RawObject); boundCount++; }
                }
            }

            if (boundCount == 0)
            {
                TxMessageBox.ShowModal(
                    $"已读取 {op.Points.Count} 个焊点，但未能从中获取任何绑定的外观对象。\n\n" +
                    "可能原因：\n• 焊点未绑定零件/外观（PS 中右键焊点 → 检查 Assigned Parts）\n" +
                    "• FillAppearancesForOperation 在此操作类型下未触发\n\n" +
                    "为避免把场景中所有对象都隐藏，已取消操作。",
                    "白名单无外观", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                SetStatus("已取消：未能获取焊点绑定外观");
                return;
            }

            SetStatus("设置仅显示操作外观...");
            try
            {
                PsReader.HideAllExcept(whitelist, s => Log("INFO", s));
                _btnRestore.Enabled      = true;
                _lblSnapStatus.Text       = "已隐藏其他对象";
                _lblSnapStatus.ForeColor  = D.Color.FromArgb(197, 90, 17);
                SetStatus("已仅显示操作绑定外观");
            }
            catch (InvalidOperationException iex) { Log("WARN", iex.Message); TxMessageBox.ShowModal(iex.Message, "无法执行", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            catch (Exception ex) { Log("ERR", "操作失败：" + ex.Message); SetStatus("失败：" + ex.Message); }
        }

        private void BtnShowAll_Click(object sender, EventArgs e)
        {
            try { PsReader.ShowAllDevices(s => Log("INFO", s)); SetStatus("已恢复全部显示"); }
            catch (Exception ex) { Log("ERR", "失败：" + ex.Message); SetStatus("失败：" + ex.Message); }
        }

        // ════════════════════════════════════════════════════════════════
        //  读取焊点 / 清空 / 导出
        // ════════════════════════════════════════════════════════════════
        private bool ReadWeldPointsFromSelectedOp()
        {
            if (_selOpObj == null) { _points.Clear(); return false; }

            var op = new OperationInfo
            {
                Name      = _selOpName ?? SafeGetName(_selOpObj),
                TypeLabel = _selOpObj.GetType().Name,
                PsObject  = _selOpObj
            };
            try { PsReader.FillPoints(op, PointType.WeldPoint, false, s => Log("INFO", s)); }
            catch (Exception ex) { Log("WARN", $"FillPoints [{op.Name}] 异常：{ex.Message}"); }

            var result = new List<WeldAnnotationPoint>();
            int idx = 1;
            if (op.Points != null)
            {
                foreach (var pi in op.Points)
                {
                    if (pi == null || pi.Type != PointType.WeldPoint || pi.Position == null || pi.Position.Length < 3) continue;

                    int boundParts = 0;
                    if (pi.AllAppearances != null && pi.AllAppearances.Count > 0)
                    {
                        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var ar in pi.AllAppearances)
                        {
                            if (ar == null) continue;
                            string key = !string.IsNullOrEmpty(ar.ParentPartName) ? ar.ParentPartName : (ar.Name ?? "");
                            if (!string.IsNullOrEmpty(key)) seen.Add(key);
                        }
                        boundParts = seen.Count;
                    }

                    result.Add(new WeldAnnotationPoint
                    {
                        Index          = idx++,
                        Name           = string.IsNullOrEmpty(pi.Name) ? ("P" + idx) : pi.Name,
                        OpName         = op.Name ?? "(未知)",
                        X = pi.Position[0], Y = pi.Position[1], Z = pi.Position[2],
                        WorldTx        = null,
                        BoundPartsCount = boundParts
                    });
                }
            }

            _points = result;
            foreach (var wp in _points)
                if (!_categories.ContainsKey(wp.Index))
                    _categories[wp.Index] = DEFAULT_CATEGORY;

            Log("INFO", $"[Annotator] 读取 [{op.Name}] 共 {result.Count} 个焊点");
            return result.Count > 0;
        }

        private static string SafeGetName(ITxObject o)
        {
            try { dynamic d = o; return (d.Name as string) ?? o.GetType().Name; }
            catch { return o?.GetType().Name ?? "(null)"; }
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            _points.Clear();
            _categories.Clear();
            _categoriesManuallyEdited.Clear();
            RefreshGrid();
            SetStatus("已清空");
        }

        private void ApplyAutoThicknessIfEnabled()
        {
            if (_chkAutoThick == null || !_chkAutoThick.Checked || _points == null || _points.Count == 0) return;
            int changed = 0;
            foreach (var pt in _points)
            {
                if (_categoriesManuallyEdited.Contains(pt.Index)) continue;
                string cat = CategoryByThickness(pt.BoundPartsCount);
                if (cat == null) continue;
                _categories.TryGetValue(pt.Index, out string old);
                if (old == cat) continue;
                _categories[pt.Index] = cat;
                changed++;
            }
            Log("INFO", $"[自动板厚] 已自动设置 {changed} 个焊点的分类（按绑定零件数）");
            RefreshGrid();
        }

        private static string CategoryByThickness(int parts)
        {
            if (parts <= 1) return null;
            if (parts == 2) return "二层板点焊";
            return "三层板及以上点焊";
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (_selOpObj == null) { SetStatus("请先在左侧选择操作节点（在PS中点击一个OP即可）"); return; }

            SetStatus("读取焊点...");
            if (!ReadWeldPointsFromSelectedOp()) { SetStatus($"OP [{_selOpName}] 下未读取到任何焊点，无法导出。"); return; }
            ApplyAutoThicknessIfEnabled();
            RefreshGrid();

            _btnExport.Enabled  = false;
            _progress.Visible   = true;
            _progress.Value     = 0;
            try
            {
                SetStatus("截图中..."); _progress.Value = 15;
                var bmp = PsReader.CaptureActiveViewer(D.Size.Empty, false, s => Log("INFO", s));
                if (bmp == null) throw new Exception("截图失败（CaptureActiveViewer 返回 null）");

                SetStatus("计算焊点位置..."); _progress.Value = 35;
                PsReader.ProjectPointsToScreen(_points, bmp.Width, bmp.Height, s => Log("INFO", s));
                RefreshGrid();

                SetStatus("写入Excel（含可编辑标注）..."); _progress.Value = 55;
                ExportToActiveExcel(bmp, _style);

                _progress.Value = 100;
                int inVp = _points.Count(p => p.InViewport);
                SetStatus($"导出完成：{_points.Count} 个焊点（视口内 {inVp} 个）");
                Log("INFO", $"已写入活动Excel：{_points.Count} 个焊点，{inVp} 个视口内点已生成可编辑 Shape 标注");
            }
            catch (Exception ex) { Log("ERR", "导出失败：" + ex.Message); SetStatus("导出失败：" + ex.Message); }
            finally { _btnExport.Enabled = true; _progress.Visible = false; }
        }

        private void BtnStyle_Click(object sender, EventArgs e)
        {
            using (var sf = new AnnotationStyleForm(_style))
                if (sf.ShowDialog(this) == DialogResult.OK) { _style = sf.Result; Log("INFO", "样式已更新"); }
        }

        // ════════════════════════════════════════════════════════════════
        //  Grid
        // ════════════════════════════════════════════════════════════════
        private void RefreshGrid()
        {
            _grid.Rows.Count = _points.Count + 1;
            for (int i = 0; i < _points.Count; i++)
            {
                var pt = _points[i];
                int row = i + 1;
                _grid[row, C_IDX]  = pt.Index.ToString();
                _grid[row, C_NAME] = pt.Name;
                _grid[row, C_OP]   = pt.OpName;
                _grid[row, C_X]    = pt.X.ToString("F2");
                _grid[row, C_Y]    = pt.Y.ToString("F2");
                _grid[row, C_Z]    = pt.Z.ToString("F2");
                _grid[row, C_VIS]  = pt.InViewport ? "是" : (pt.ScreenX == 0 && pt.ScreenY == 0 ? "-" : "否");

                if (!_categories.TryGetValue(pt.Index, out string cat) || string.IsNullOrEmpty(cat))
                    _categories[pt.Index] = cat = DEFAULT_CATEGORY;
                _grid[row, C_TYPE] = cat;

                var cs = _grid.Rows[row].Style ?? _grid.Styles.Add("r" + row);
                cs.BackColor = pt.InViewport ? D.Color.FromArgb(198, 239, 206)
                             : (pt.ScreenX != 0 || pt.ScreenY != 0) ? D.Color.FromArgb(255, 235, 156)
                             : D.Color.Empty;
                _grid.Rows[row].Style = cs;
            }

            int inVp = _points.Count(p => p.InViewport);
            _lblCount.Text = _points.Count > 0
                ? $"焊点：{_points.Count} 个  视口内：{inVp} 个"
                : "焊点：0 个";

            try
            {
                int[] minWidths = { 36, 100, 90, 70, 70, 70, 50, 90 };
                for (int c = 0; c < _grid.Cols.Count && c < minWidths.Length; c++)
                {
                    _grid.AutoSizeCol(c);
                    if (_grid.Cols[c].Width < minWidths[c]) _grid.Cols[c].Width = minWidths[c];
                    _grid.Cols[c].Width += 12;
                }
            }
            catch { }
        }

        private void Grid_AfterEdit(object sender, RowColEventArgs e)
        {
            try
            {
                if (e.Col != C_TYPE) return;
                if (e.Row < 1 || e.Row - 1 >= _points.Count) return;
                string newVal = _grid[e.Row, C_TYPE] as string;
                if (string.IsNullOrEmpty(newVal)) newVal = DEFAULT_CATEGORY;
                int key = _points[e.Row - 1].Index;
                _categories[key] = newVal;
                _categoriesManuallyEdited.Add(key);
                Log("INFO", $"[Annotator] 点 [{_points[e.Row - 1].Name}] 类型改为 [{newVal}]");
            }
            catch (Exception ex) { Log("WARN", "Grid_AfterEdit: " + ex.Message); }
        }

        // ════════════════════════════════════════════════════════════════
        //  标注文字
        // ════════════════════════════════════════════════════════════════
        private string GetAnnotationLabel(WeldAnnotationPoint pt, int ordinal)
        {
            var mode = (LabelNamingMode)(_cmbLabelMode?.SelectedIndex ?? 0);
            string name = pt.Name ?? "";
            string pfx  = _txtPrefix?.Text ?? "";
            string sfx  = _txtSuffix?.Text ?? "";
            switch (mode)
            {
                case LabelNamingMode.PointName:  return name;
                case LabelNamingMode.Prefix:     return pfx + ordinal;
                case LabelNamingMode.Suffix:     return ordinal + sfx;
                case LabelNamingMode.SeqAndName: return ordinal + "-" + name;
                default:                         return ordinal.ToString();
            }
        }

        // ════════════════════════════════════════════════════════════════
        //  Excel 导出（dynamic COM）
        // ════════════════════════════════════════════════════════════════
        private void ExportToActiveExcel(D.Bitmap bmp, AnnotationStyle style)
        {
            string tmp = Path.GetTempFileName() + ".png";
            bmp.Save(tmp, DI.ImageFormat.Png);

            dynamic xlApp = null;
            bool savedScreenUpdating = true;
            try
            {
                try { xlApp = Marshal.GetActiveObject("Excel.Application"); }
                catch { throw new Exception("未检测到已打开的 Excel，请先打开工作簿"); }

                dynamic wb = xlApp.ActiveWorkbook;
                if (wb == null) throw new Exception("Excel 中没有活动工作簿");

                try { savedScreenUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }

                dynamic ws;
                if (_chkNewSheet.Checked)
                {
                    dynamic sheets = wb.Sheets;
                    ws = sheets.Add(System.Reflection.Missing.Value, sheets[sheets.Count],
                        System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                    ws.Name = "焊点标注_" + DateTime.Now.ToString("MMddHHmm");
                }
                else { ws = wb.ActiveSheet; }

                double targetLeft = -1, targetTop = -1, targetWidth = -1, targetHeight = -1;
                bool scaleToTarget = false;
                try
                {
                    dynamic sel = xlApp.Selection;
                    int selCellCount = 0; try { selCellCount = (int)sel.Cells.Count; } catch { }
                    bool selMerged = false; try { selMerged = (bool)sel.MergeCells; } catch { }
                    targetLeft = (double)sel.Left;
                    targetTop  = (double)sel.Top;
                    if (selCellCount == 1 && !selMerged) { scaleToTarget = false; }
                    else { targetWidth = (double)sel.Width; targetHeight = (double)sel.Height; scaleToTarget = targetWidth > 4 && targetHeight > 4; }
                }
                catch { targetLeft = -1; targetTop = -1; scaleToTarget = false; }

                const int msoFalsePic = 0;
                dynamic a1 = ws.Cells[1, 1];
                double a1Left = (double)a1.Left, a1Top = (double)a1.Top;
                dynamic pic = ws.Shapes.AddPicture(tmp, msoFalsePic, true, a1Left, a1Top, -1, -1);

                const double maxWidthPt = 800.0;
                if ((double)pic.Width > maxWidthPt)
                {
                    double r = maxWidthPt / (double)pic.Width;
                    pic.Width = maxWidthPt;
                    pic.Height = (double)pic.Height * r;
                }

                string picName = "WPPicture_" + DateTime.Now.ToString("HHmmss");
                try { pic.Name = picName; } catch { }

                var shapeNamesForGroup = new List<string> { picName };
                double picLeft   = (double)pic.Left, picTop = (double)pic.Top;
                double picWidth  = (double)pic.Width, picHeight = (double)pic.Height;
                double sx = picWidth / bmp.Width, sy = picHeight / bmp.Height;

                const int msoShapeOval           = 9;
                const int msoTextOrientationHoriz = 1;
                const int msoFalse               = 0;
                const int msoTrue                = -1;
                const int xlHAlignCenter         = -4108;
                const int xlVAlignCenter         = -4108;
                const int msoConnectorStraight   = 1;
                const int xlFreeFloating         = 3;

                try { pic.Placement = xlFreeFloating; } catch { }

                int oleDot       = D.ColorTranslator.ToOle(style.DotColor);
                int oleLine      = D.ColorTranslator.ToOle(style.LineColor);
                int oleBoxFill   = D.ColorTranslator.ToOle(style.BoxFillColor);
                int oleBoxBorder = D.ColorTranslator.ToOle(style.BoxBorderColor);
                int oleText      = D.ColorTranslator.ToOle(style.TextColor);

                double dotR      = style.DotRadius;
                const double BOX_W_PT = 28.0, BOX_H_PT = 18.0;
                int shapeCount = 0;
                int inVp = _points.Count(p => p.InViewport);
                int processed = 0;

                var placedPts = _points.Where(p => p.InViewport).ToList();
                var labelPos  = new Dictionary<int, Tuple<double, double>>();

                if (placedPts.Count > 0)
                {
                    var pointEdges = placedPts.Select((p, idx) =>
                    {
                        double cx = picLeft + p.ScreenX * sx, cy = picTop + p.ScreenY * sy;
                        double dL = cx - picLeft, dR = picLeft + picWidth - cx;
                        double dT = cy - picTop,   dB = picTop + picHeight - cy;
                        double minD = Math.Min(Math.Min(dL, dR), Math.Min(dT, dB));
                        int side = minD == dT ? 0 : minD == dR ? 1 : minD == dB ? 2 : 3;
                        return new { Pt = p, Cx = cx, Cy = cy, Side = side };
                    }).ToList();

                    foreach (int sideId in new[] { 0, 1, 2, 3 })
                    {
                        var grp = pointEdges.Where(e => e.Side == sideId).ToList();
                        if (grp.Count == 0) continue;
                        if (sideId == 0 || sideId == 2) grp = grp.OrderBy(e => e.Cx).ToList();
                        else grp = grp.OrderBy(e => e.Cy).ToList();

                        int n = grp.Count;
                        for (int i = 0; i < n; i++)
                        {
                            var e = grp[i];
                            double t = n == 1 ? 0.5 : (i + 0.5) / n;
                            double bx, by;
                            if (sideId == 0)      { bx = picLeft + t * (picWidth - BOX_W_PT); by = picTop - BOX_H_PT - 4; }
                            else if (sideId == 2) { bx = picLeft + t * (picWidth - BOX_W_PT); by = picTop + picHeight + 4; }
                            else if (sideId == 3) { bx = picLeft - BOX_W_PT - 4; by = picTop + t * (picHeight - BOX_H_PT); }
                            else                 { bx = picLeft + picWidth + 4;  by = picTop + t * (picHeight - BOX_H_PT); }
                            labelPos[e.Pt.Index] = Tuple.Create(bx, by);
                        }
                    }
                }

                int ordinal = 0;
                foreach (var pt in _points)
                {
                    if (!pt.InViewport) continue;
                    ordinal++;

                    double cx = picLeft + pt.ScreenX * sx, cy = picTop + pt.ScreenY * sy;
                    double bx, by;
                    if (labelPos.TryGetValue(pt.Index, out var bp)) { bx = bp.Item1; by = bp.Item2; }
                    else { bx = cx + style.OffsetX; by = cy - style.OffsetY; }

                    dynamic dotShape = null, boxShape = null;

                    if (!_categories.TryGetValue(pt.Index, out string cat) || string.IsNullOrEmpty(cat))
                        cat = DEFAULT_CATEGORY;
                    int shapeId = msoShapeOval;
                    style.CategoryShapes?.TryGetValue(cat, out shapeId);
                    bool filled = true;
                    style.CategoryFilled?.TryGetValue(cat, out filled);

                    try
                    {
                        dotShape = ws.Shapes.AddShape(shapeId, cx - dotR, cy - dotR, dotR * 2.0, dotR * 2.0);
                        if (filled)
                        {
                            try { dotShape.Fill.ForeColor.RGB = oleDot; dotShape.Fill.Solid(); } catch { }
                            try { dotShape.Line.ForeColor.RGB = oleDot; dotShape.Line.Weight = 1.0; } catch { }
                        }
                        else
                        {
                            try { dotShape.Fill.Visible = msoFalse; } catch { }
                            try { dotShape.Line.ForeColor.RGB = oleDot; dotShape.Line.Weight = 1.75; } catch { }
                        }
                        try { dotShape.Placement = xlFreeFloating; } catch { }
                        try { dotShape.Name = "WPDot_" + pt.Index; shapeNamesForGroup.Add("WPDot_" + pt.Index); } catch { }
                        shapeCount++;
                    }
                    catch (Exception ex) { Log("WARN", "AddShape 失败: " + ex.Message); }

                    try
                    {
                        string label = GetAnnotationLabel(pt, ordinal);
                        double estW = Math.Max(BOX_W_PT, label.Length * 6.5 + 6.0);
                        boxShape = ws.Shapes.AddTextbox(msoTextOrientationHoriz, bx, by, estW, BOX_H_PT);
                        try { boxShape.Fill.ForeColor.RGB = oleBoxFill; boxShape.Fill.Solid(); } catch { }
                        try { boxShape.Line.ForeColor.RGB = oleBoxBorder; boxShape.Line.Weight = 1.0; } catch { }
                        bool setOk = false;
                        try { boxShape.TextFrame.Characters().Text = label; setOk = true; } catch { }
                        if (!setOk) try { boxShape.TextFrame2.TextRange.Text = label; } catch { }
                        try { boxShape.TextFrame.HorizontalAlignment = xlHAlignCenter; } catch { }
                        try { boxShape.TextFrame.VerticalAlignment   = xlVAlignCenter; } catch { }
                        try
                        {
                            double pad = style.BoxPadding;
                            boxShape.TextFrame.MarginLeft = pad; boxShape.TextFrame.MarginRight = pad;
                            boxShape.TextFrame.MarginTop  = pad; boxShape.TextFrame.MarginBottom = pad;
                        }
                        catch { }
                        try
                        {
                            dynamic chars = boxShape.TextFrame.Characters();
                            chars.Font.Name   = style.TextFont.Name;
                            chars.Font.Size   = style.TextFont.SizeInPoints;
                            chars.Font.Bold   = style.TextFont.Bold;
                            chars.Font.Italic = style.TextFont.Italic;
                            chars.Font.Color  = oleText;
                        }
                        catch { }
                        try { boxShape.Placement = xlFreeFloating; } catch { }
                        try { boxShape.Name = "WPBox_" + pt.Index; shapeNamesForGroup.Add("WPBox_" + pt.Index); } catch { }
                        shapeCount++;
                    }
                    catch (Exception ex) { Log("WARN", "AddTextbox 失败: " + ex.Message); }

                    if (dotShape != null && boxShape != null)
                    {
                        try
                        {
                            dynamic conn = ws.Shapes.AddConnector(msoConnectorStraight, cx, cy, bx, by);
                            try { conn.Line.ForeColor.RGB = oleLine; } catch { }
                            try { conn.Line.Weight = (double)style.LineWidth; } catch { }
                            try { conn.Line.Visible = msoTrue; } catch { }
                            try { conn.ConnectorFormat.BeginConnect(dotShape, 1); } catch (Exception ex) { Log("WARN", "BeginConnect: " + ex.Message); }
                            try { conn.ConnectorFormat.EndConnect(boxShape, 1); }   catch (Exception ex) { Log("WARN", "EndConnect: "   + ex.Message); }
                            try { conn.RerouteConnections(); } catch { }
                            try { conn.Line.ForeColor.RGB = oleLine; } catch { }
                            try { conn.Line.Weight = (double)style.LineWidth; } catch { }
                            try { conn.Placement = xlFreeFloating; } catch { }
                            try { conn.Name = "WPConn_" + pt.Index; shapeNamesForGroup.Add("WPConn_" + pt.Index); } catch { }
                            shapeCount++;
                        }
                        catch (Exception ex) { Log("WARN", "AddConnector 失败: " + ex.Message); }
                    }

                    processed++;
                    if (inVp > 0 && _progress != null)
                        try { _progress.Value = 55 + (int)(processed * 40.0 / inVp); } catch { }
                }

                double finalPicWidth = (double)pic.Width, finalPicHeight = (double)pic.Height;
                try
                {
                    if (shapeNamesForGroup.Count >= 2)
                    {
                        object namesArr = shapeNamesForGroup.ToArray();
                        dynamic range   = ws.Shapes.Range(namesArr);
                        dynamic group   = range.Group();
                        try { group.Name = "WPGroup_" + DateTime.Now.ToString("HHmmss"); } catch { }
                        try { group.Placement = xlFreeFloating; } catch { }

                        if (targetLeft >= 0 && targetTop >= 0)
                        {
                            try { group.Left = targetLeft; } catch { }
                            try { group.Top  = targetTop;  } catch { }
                        }
                        if (scaleToTarget && targetWidth > 0 && targetHeight > 0)
                        {
                            double gw = (double)group.Width, gh = (double)group.Height;
                            double scale = Math.Min(targetWidth / gw, targetHeight / gh);
                            try { group.LockAspectRatio = true; } catch { }
                            try { group.Width = gw * scale; } catch { }
                            try { if (Math.Abs((double)group.Height - gh * scale) > 1) group.Height = gh * scale; } catch { }
                        }
                        finalPicWidth  = (double)group.Width;
                        finalPicHeight = (double)group.Height;
                        Log("INFO", $"[Annotator] 已组合 {shapeNamesForGroup.Count} 个 Shape 为单组" +
                            (scaleToTarget ? "，已缩放至选中单元格" : ""));
                    }
                }
                catch (Exception ex) { Log("WARN", "[Annotator] Group/缩放失败: " + ex.Message + "（保留为独立 Shape）"); }

                picHeight = finalPicHeight; picWidth = finalPicWidth;
                if (targetLeft >= 0) picLeft = targetLeft;
                if (targetTop  >= 0) picTop  = targetTop;

                if (_chkWriteList.Checked && _points.Count > 0)
                {
                    int ds = (int)Math.Ceiling(picHeight / 15.0) + 3;
                    string[] hdr = { "#", "焊点名称", "操作名", "X(mm)", "Y(mm)", "Z(mm)", "视口内", "类型" };
                    for (int c = 0; c < hdr.Length; c++)
                    {
                        dynamic cell = ws.Cells[ds, c + 1];
                        cell.Value2 = hdr[c];
                        cell.Font.Bold = true;
                        cell.Interior.Color = D.ColorTranslator.ToOle(D.Color.FromArgb(68, 114, 196));
                        cell.Font.Color     = D.ColorTranslator.ToOle(D.Color.White);
                    }
                    for (int i = 0; i < _points.Count; i++)
                    {
                        var pt = _points[i]; int row = ds + 1 + i;
                        if (!_categories.TryGetValue(pt.Index, out string catRow) || string.IsNullOrEmpty(catRow))
                            catRow = DEFAULT_CATEGORY;
                        ((dynamic)ws.Cells[row, 1]).Value2 = pt.Index;
                        ((dynamic)ws.Cells[row, 2]).Value2 = pt.Name;
                        ((dynamic)ws.Cells[row, 3]).Value2 = pt.OpName;
                        ((dynamic)ws.Cells[row, 4]).Value2 = pt.X;
                        ((dynamic)ws.Cells[row, 5]).Value2 = pt.Y;
                        ((dynamic)ws.Cells[row, 6]).Value2 = pt.Z;
                        ((dynamic)ws.Cells[row, 7]).Value2 = pt.InViewport ? "是" : "否";
                        ((dynamic)ws.Cells[row, 8]).Value2 = catRow;
                        if (i % 2 == 1)
                        {
                            dynamic rng = ws.Range[ws.Cells[row, 1], ws.Cells[row, 8]];
                            rng.Interior.Color = D.ColorTranslator.ToOle(D.Color.FromArgb(242, 242, 242));
                        }
                    }
                    dynamic dr2 = ws.Range[ws.Cells[ds, 1], ws.Cells[ds + _points.Count, 8]];
                    dr2.Columns.AutoFit();
                }

                ws.Activate();
                xlApp.Visible = true;
                Log("INFO", $"[Annotator] Excel 写入完成：{shapeCount} 个 Shape，{inVp} 个视口内点；" +
                    (_chkWriteList.Checked ? "附焊点数据表" : "未附数据表"));
            }
            finally
            {
                if (xlApp != null) try { xlApp.ScreenUpdating = savedScreenUpdating; } catch { }
                try { File.Delete(tmp); } catch { }
            }
        }

        // ════════════════════════════════════════════════════════════════
        //  辅助
        // ════════════════════════════════════════════════════════════════
        private void SetStatus(string msg) { _lblStatus.Text = msg; _status.Refresh(); }

        private void Log(string level, string msg)
        {
            if (_logBox == null) return;
            D.Color c = level == "ERR" ? D.Color.Tomato : level == "WARN" ? D.Color.Yellow : D.Color.LightGray;
            _logBox.SelectionColor = c;
            _logBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {level}: {msg}\r\n");
            _logBox.ScrollToCaret();
            if (level != "INFO" && !_logVisible) _logPanel.Visible = _logVisible = true;
        }
    }
}
