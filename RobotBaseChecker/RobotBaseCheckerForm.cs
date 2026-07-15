using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Tecnomatix.Engineering;      // 必须引用：用于 TxApplication, TxRobot, TxFrame
using Tecnomatix.Engineering.Ui;   // TxForm

// 别名，避免 WPF/WinForms 冲突（按 TxTools 惯例）
using TextBox = System.Windows.Forms.TextBox;
using Button = System.Windows.Forms.Button;
using Label = System.Windows.Forms.Label;

namespace TxTools.RobotBaseChecker
{
    public class RobotBaseCheckerForm : TxForm
    {
        // 静态引用防 GC（非模态窗口）
        internal static RobotBaseCheckerForm Instance;

        private DataGridView _grid;
        private TextBox _logBox;
        private NumericUpDown _posTol;
        private NumericUpDown _rotTol;
        private Button _btnCheck;
        private Button _btnExport;
        private Button _btnSync;       // 同步全部按钮
        private ComboBox _brandMode;   // 品牌处理模式

        private List<RobotBaseResult> _results = new List<RobotBaseResult>();

        private bool _scaled;
        private readonly Size _designSize = new Size(1000, 620);

        public RobotBaseCheckerForm()
        {
            SemiModal = false;            // 非模态：双安全（构造 + OnInitTxForm）
            BuildUi();
            TryInitUiKit();
        }

        public override void OnInitTxForm()
        {
            base.OnInitTxForm();
            SemiModal = false;
        }

        private void TryInitUiKit()
        {
            // 若套件内有 FormUiKit 则统一外观；没有也不影响运行
            try
            {
                var t = Type.GetType("TxTools.Common.FormUiKit, TxTools.Common");
                var mi = t?.GetMethod("InitStandardForm");
                mi?.Invoke(null, new object[] { this });
            }
            catch { }

            // 唯一持久化键，消除跨插件窗口几何串扰
            Name = GetType().FullName;
            try
            {
                var flatStyleProp = this.GetType().GetProperty("FlatStyleEnabled");
                if (flatStyleProp != null && flatStyleProp.CanWrite)
                {
                    flatStyleProp.SetValue(this, false, null);
                }
            }
            catch
            {
                // 反射失败时静默忽略，确保插件继续运行
            }
        }

        private void BuildUi()
        {
            Text = "机器人 BASE0 一致性检查";
            ClientSize = _designSize;
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.None;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(8)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 85));   // 工具栏（容差两行 + GroupBox）
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 70));    // 表格
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 30));    // 日志

            // ---- 工具栏：左卡片(容差) + 右卡片(操作) ----
            var toolbar = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1
            };
            toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            toolbar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            // ── 容差参数卡片 ──
            var tolGroup = new GroupBox { Text = "容差参数", Dock = DockStyle.Fill, Padding = new Padding(6, 2, 6, 2) };
            var tolPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2
            };
            tolPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            tolPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            _posTol = new NumericUpDown { DecimalPlaces = 3, Increment = 0.1M, Minimum = 0, Maximum = 1000, Value = 0.5M, Width = 80 };
            _rotTol = new NumericUpDown { DecimalPlaces = 4, Increment = 0.01M, Minimum = 0, Maximum = 360, Value = 0.01M, Width = 80 };

            tolPanel.Controls.Add(new Label { Text = "位置容差(mm):", AutoSize = true }, 0, 0);
            tolPanel.Controls.Add(_posTol, 1, 0);
            tolPanel.Controls.Add(new Label { Text = "旋转容差:", AutoSize = true }, 0, 1);
            tolPanel.Controls.Add(_rotTol, 1, 1);

            tolGroup.Controls.Add(tolPanel);
            toolbar.Controls.Add(tolGroup, 0, 0);

            // ── 操作卡片 ──
            var opGroup = new GroupBox { Text = "操作", Dock = DockStyle.Fill, Padding = new Padding(6, 2, 6, 2) };
            var opPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, AutoScroll = true };

            opPanel.Controls.Add(new Label { Text = "品牌:", AutoSize = true, Padding = new Padding(0, 6, 0, 0) });
            _brandMode = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 130 };
            _brandMode.Items.AddRange(new object[] { "自动检测", "通用(底座面)", "FANUC(J1∩J2)" });
            _brandMode.SelectedIndex = 0;
            opPanel.Controls.Add(_brandMode);

            _btnCheck = new Button { Text = "开始检查", Width = 100, Height = 28 };
            _btnCheck.Click += (s, e) => RunCheck();
            opPanel.Controls.Add(_btnCheck);

            _btnSync = new Button { Text = "同步全部 BASE0", Width = 130, Height = 28, Enabled = false };
            _btnSync.Click += (s, e) => SyncAll();
            opPanel.Controls.Add(_btnSync);

            _btnExport = new Button { Text = "导出 CSV", Width = 100, Height = 28, Enabled = false };
            _btnExport.Click += (s, e) => ExportCsv();
            opPanel.Controls.Add(_btnExport);

            opGroup.Controls.Add(opPanel);
            toolbar.Controls.Add(opGroup, 1, 0);

            root.Controls.Add(toolbar, 0, 0);

            // ---- 表格 ----
            _grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                MultiSelect = false,
                ColumnHeadersHeight = 48,
                EnableHeadersVisualStyles = false
            };
            _grid.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            _grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            _grid.Columns.Add("col_robot", "机器人");
            _grid.Columns.Add("col_brand", "品牌");
            _grid.Columns.Add("col_self", "自身坐标 (X,Y,Z)");
            _grid.Columns.Add("col_base", "当前BASE0 (X,Y,Z)");
            _grid.Columns.Add("col_expected", "期望BASE0 (X,Y,Z)");
            _grid.Columns.Add("col_dp", "ΔPos(mm)");
            _grid.Columns.Add("col_dr", "ΔRot");
            _grid.Columns.Add("col_verdict", "结论");

            // 单击行 → 日志区显示该行详细信息
            _grid.SelectionChanged += (s, e) => ShowSelectedDetail();

            // 双击行 → 视口跳转至选中机器人的位置
            _grid.CellDoubleClick += (s, e) => FocusRobot(e.RowIndex);

            // 拖拽窗体边缘结束后重新分配列宽
            this.ResizeEnd += (s, e) => FitColumnsToGrid();

            root.Controls.Add(_grid, 0, 1);

            // ---- 日志（统计信息 + 详细信息） ----
            _logBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9F),
                WordWrap = false
            };
            root.Controls.Add(_logBox, 0, 2);

            Controls.Add(root);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            if (!_scaled)
            {
                _scaled = true;
                float sc = CreateGraphics().DpiX / 96f;
                if (sc > 1.01f)
                {
                    Scale(new SizeF(sc, sc));
                    Size = new Size((int)(_designSize.Width * sc) + (Width - ClientSize.Width),
                                    (int)(_designSize.Height * sc) + (Height - ClientSize.Height));
                }
            }
        }

        // ----------------------------------------------------------------
        // 核心修改：同步所有存在偏差的机器人的 BASE0
        // ----------------------------------------------------------------
        private void SyncAll()
        {
            if (_results == null || _results.Count == 0) return;

            if (MessageBox.Show(this, "确定要将所有【存在偏差】的机器人的 BASE0 覆盖为【期望BASE0】吗？\n\n· 通用品牌：期望 = 底座安装面(机器人自身坐标)\n· FANUC：期望 = J1∩J2 运动学交点\n\n注意：此操作将批量修改场景数据。", "同步全部 BASE0", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            {
                return;
            }

            int successCount = 0;
            int skippedCount = 0;
            int failCount = 0;
            var errorLog = new StringBuilder();

            try
            {
                PsContext.Run(() =>
                {
                    foreach (var r in _results)
                    {
                        // 逻辑梳理：跳过已一致或缺乏坐标无法对比的机器人
                        if (!r.Comparable || r.Verdict == "一致")
                        {
                            skippedCount++;
                            continue;
                        }

                        try
                        {
                            // 1. 直接用 Analyze 时保存的机器人实例（避免同名 GetObjectsByName 串扰）
                            TxRobot robot = r.RobotRef;
                            if (robot == null)
                                throw new Exception($"缺少机器人实例引用 [{r.RobotName}]。");

                            // 2. 找到第一个系统坐标系（不再筛 TxFrame 类型，
                            //    default 控制器返回的可能不是 TxFrame）
                            var frames = robot.GetAllSystemFrames() as System.Collections.IEnumerable;
                            object base0Obj = null;   // 首帧对象（不挑类型）
                            if (frames != null)
                            {
                                foreach (object f in frames)
                                {
                                    if (f != null) { base0Obj = f; break; }
                                }
                            }

                            if (base0Obj == null)
                            {
                                // 没有 BASE0 帧（多为未挂控制器）→ 跳过而非失败
                                skippedCount++;
                                continue;
                            }

                            // 3. 计算品牌感知的期望 BASE0（FANUC 用 J1∩J2），写入
                            var sb = new StringBuilder();
                            var target = RobotKinematics.ComputeExpectedTx(robot, r.Brand, sb);
                            if (target == null)
                                throw new Exception("无法构造期望 BASE0 变换。" + sb);

                            // 4. 写入首帧的世界位姿（动态多属性兜底，非仅 TxFrame.AbsoluteLocation）
                            SetFrameWorldTx(base0Obj, target);
                            successCount++;
                        }
                        catch (Exception ex)
                        {
                            failCount++;
                            errorLog.AppendLine($"[{r.RobotName}] 失败: {ex.Message}");
                        }
                    }
                });

                string resultMsg = $"批量同步完成！\n\n成功同步: {successCount} 台\n跳过处理: {skippedCount} 台 (已一致或无法读取)\n失败异常: {failCount} 台";

                if (failCount > 0)
                {
                    resultMsg += $"\n\n失败详情:\n{errorLog}";
                    MessageBox.Show(this, resultMsg, "同步结果 (部分成功)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show(this, resultMsg, "同步结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                // 同步完毕自动触发一次刷新，更新表格显示
                RunCheck();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "批量同步过程发生系统异常: " + ex.Message, "严重错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ----------------------------------------------------------------
        private void RunCheck()
        {
            _btnCheck.Enabled = false;
            _btnSync.Enabled = false;
            _logBox.Text = "正在遍历机器人并读取坐标...\r\n";
            try
            {
                double posTol = (double)_posTol.Value;
                double rotTol = (double)_rotTol.Value;
                BrandMode mode = BrandModeFromUi();
                _results = RobotBaseReader.Analyze(posTol, rotTol, mode);
                FillGrid();

                // 检查完毕后只要有结果就激活全局按钮
                _btnExport.Enabled = _results.Count > 0;
                _btnSync.Enabled = _results.Count > 0;
            }
            catch (Exception ex)
            {
                _logBox.AppendText("出错: " + ex.Message + "\r\n" + ex.StackTrace);
                MessageBox.Show(this, ex.Message, "检查失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _btnCheck.Enabled = true;
            }
        }

        private void FillGrid()
        {
            _grid.Rows.Clear();
            int total = _results.Count, ok = 0, dev = 0, fail = 0, nobase = 0;

            foreach (var r in _results)
            {
                string self = r.Self.PosText;
                string baseTxt = r.Base0.PosText;
                string expTxt = r.Expected.PosText;
                string dp = double.IsNaN(r.DeltaPos) ? "—" : r.DeltaPos.ToString("F3");
                string dr = double.IsNaN(r.DeltaRot) ? "—" : r.DeltaRot.ToString("F4");

                int idx = _grid.Rows.Add(r.RobotName, r.Brand, self, baseTxt, expTxt, dp, dr, r.Verdict);
                var row = _grid.Rows[idx];

                if (r.Verdict == "一致") { row.DefaultCellStyle.BackColor = Color.FromArgb(225, 245, 225); ok++; }
                else if (r.Verdict == "存在偏差") { row.DefaultCellStyle.BackColor = Color.FromArgb(255, 230, 225); dev++; }
                else if (r.Verdict == "无当前BASE0") { row.DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 225); nobase++; }
                else { row.DefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245); fail++; }
            }

            // 统计信息写入日志区（不再使用 _summary 标签）
            var summarySb = new StringBuilder();
            summarySb.AppendLine($"═══ 检查统计 ═══");
            summarySb.AppendLine($"共 {total} 台：一致 {ok}，偏差 {dev}，无BASE0 {nobase}，无法读取 {fail}");
            summarySb.AppendLine();
            foreach (var r in _results)
            {
                summarySb.AppendLine(r.Detail);
                summarySb.AppendLine("────────────────────────────────");
            }
            _logBox.Text = summarySb.ToString();

            AutoFitColumns();
        }

        // 两步法列宽自适应：先按内容算理想宽度，再约束总宽不超表格区域
        private void AutoFitColumns()
        {
            if (_grid.Columns.Count == 0) return;
            _grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            foreach (DataGridViewColumn c in _grid.Columns) c.Width += 12; // 留点余量
            FitColumnsToGrid();
        }

        private void FitColumnsToGrid()
        {
            if (_grid.Columns.Count == 0) return;
            int avail = _grid.ClientSize.Width;
            if (_grid.RowHeadersVisible) avail -= _grid.RowHeadersWidth;
            if (_grid.Controls.OfType<VScrollBar>().Any(sb => sb.Visible))
                avail -= SystemInformation.VerticalScrollBarWidth;
            if (avail <= 0) return;

            int total = 0;
            foreach (DataGridViewColumn c in _grid.Columns) total += c.Width;

            if (total > avail)
            {
                // 内容总宽超出 → 按比例缩小（最小 30px），末列吸收取整误差
                double scale = (double)avail / total;
                int used = 0, n = _grid.Columns.Count;
                for (int i = 0; i < n; i++)
                {
                    int w = (i == n - 1) ? Math.Max(30, avail - used)
                                         : Math.Max(30, (int)(_grid.Columns[i].Width * scale));
                    _grid.Columns[i].Width = w;
                    used += w;
                }
            }
            else
            {
                // 内容总宽不足 → 最后一列吸收剩余空间，填满表格
                int last = _grid.Columns.Count - 1;
                _grid.Columns[last].Width += (avail - total);
            }
        }

        private void ShowSelectedDetail()
        {
            if (_grid.SelectedRows.Count == 0) return;
            int i = _grid.SelectedRows[0].Index;
            if (i < 0 || i >= _results.Count) return;
            _logBox.Text = _results[i].Detail;
        }

        /// <summary>
        /// 双击表格行 → 选中机器人 + ZoomToSelection 跳转视口。
        /// 参考 AutoRecorder 的 ZoomToObjects 实现：
        ///   1. SetItems 选中目标对象
        ///   2. ExecuteCommand("GraphicViewer.ZoomToSelection") 让视口自动聚焦
        ///   3. 兜底：ZoomToFit
        /// </summary>
        private void FocusRobot(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= _results.Count) return;
            var r = _results[rowIndex];
            var robot = r.RobotRef;
            if (robot == null) return;

            try
            {
                PsContext.Run(() =>
                {
                    // 1. 选中机器人
                    var list = new TxObjectList();
                    list.Add(robot);
                    TxApplication.ActiveSelection.SetItems(list);

                    // 2. 视口跳转：使用 PS 内置 ZoomToSelection 命令（与 AutoRecorder 一致）
                    try
                    {
                        TxApplication.CommandsManager.ExecuteCommand("GraphicViewer.ZoomToSelection");
                    }
                    catch
                    {
                        // 兜底：ZoomToFit 全场景
                        try
                        {
                            var gv = TxApplication.ViewersManager?.GraphicViewer;
                            if (gv != null) gv.ZoomToFit();
                        }
                        catch { }
                    }

                    // 3. 刷新显示
                    try { TxApplication.RefreshDisplay(); } catch { }
                });
            }
            catch { }
        }

        private void ExportCsv()
        {
            using (var dlg = new SaveFileDialog { Filter = "CSV 文件|*.csv", FileName = "RobotBase0Check.csv" })
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                var sb = new StringBuilder();
                sb.AppendLine("机器人,品牌,Self_X,Self_Y,Self_Z,Base0_X,Base0_Y,Base0_Z,Expected_X,Expected_Y,Expected_Z,DeltaPos_mm,DeltaRot,结论");
                var c = CultureInfo.InvariantCulture;
                foreach (var r in _results)
                {
                    sb.AppendLine(string.Join(",",
                        Csv(r.RobotName), Csv(r.Brand),
                        N(r.Self.X, r.Self.PosValid), N(r.Self.Y, r.Self.PosValid), N(r.Self.Z, r.Self.PosValid),
                        N(r.Base0.X, r.Base0.PosValid), N(r.Base0.Y, r.Base0.PosValid), N(r.Base0.Z, r.Base0.PosValid),
                        N(r.Expected.X, r.Expected.PosValid), N(r.Expected.Y, r.Expected.PosValid), N(r.Expected.Z, r.Expected.PosValid),
                        double.IsNaN(r.DeltaPos) ? "" : r.DeltaPos.ToString("F4", c),
                        double.IsNaN(r.DeltaRot) ? "" : r.DeltaRot.ToString("F4", c),
                        Csv(r.Verdict)));
                }
                File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
                MessageBox.Show(this, "已导出: " + dlg.FileName, "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private BrandMode BrandModeFromUi()
        {
            switch (_brandMode.SelectedIndex)
            {
                case 1: return BrandMode.Generic;
                case 2: return BrandMode.Fanuc;
                default: return BrandMode.Auto;
            }
        }

        private static string N(double v, bool valid)
            => valid ? v.ToString("F4", CultureInfo.InvariantCulture) : "";

        private static string Csv(string s)
            => s != null && (s.Contains(",") || s.Contains("\"")) ? "\"" + s.Replace("\"", "\"\"") + "\"" : s;

        // ════════════════════════════════════════════════════════════
        //  首帧世界位姿写入（动态多属性兜底，与 GetFrameWorldTx 对称）
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 将期望 BASE0 变换写入首帧的世界位姿属性。
        /// 不挑类型——default 控制器返回的首帧未必是 TxFrame，
        /// 所以优先尝试强类型 TxFrame.AbsoluteLocation，失败后走 dynamic 兜底。
        /// </summary>
        private static void SetFrameWorldTx(object frameObj, TxTransformation target)
        {
            // 优先路径：强类型 TxFrame
            var fr = frameObj as TxFrame;
            if (fr != null)
            {
                try { fr.AbsoluteLocation = target; return; }
                catch { /* 写入失败，走兜底 */ }
            }

            // 动态兜底：尝试多个可写属性
            dynamic dObj = frameObj;
            foreach (var prop in new[] { "AbsoluteLocation", "LocationRelativeToWorld", "Location", "Transformation" })
            {
                try
                {
                    dObj.GetType().GetProperty(prop)?.SetValue(dObj, target, null);
                    return;
                }
                catch { /* 属性不可写或不存在，继续 */ }
            }

            // 最终兜底：直接 dynamic 赋值
            try { dObj.AbsoluteLocation = target; return; }
            catch { }

            throw new Exception(
                "无法写入首帧的世界位姿。首帧类型: " + frameObj.GetType().Name +
                "，所有写入属性均失败。");
        }
    }
}