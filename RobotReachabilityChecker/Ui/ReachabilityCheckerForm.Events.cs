using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using TxTools.RobotReachabilityChecker.Diagnostics;
using TxTools.RobotReachabilityChecker.Models;
using TxTools.RobotReachabilityChecker.Services;

namespace TxTools.RobotReachabilityChecker.Ui
{
    // =========================================================================
    // ReachabilityCheckerForm — 事件处理 partial
    //
    // 包含：
    //   · OnOpNodePicked            — TxObjEditBoxCtrl.Picked 处理（唯一拾取入口）
    //   · Grid_AfterSelChange       — 单击行 → 仅更新状态栏（不再驱动姿态）
    //   · Grid_DblClick_Mouse/Core  — 双击行 → 驱动机器人到记录的关节姿态
    //                                  · 同一行 500ms 内再次单击 → 触发三击（Robot Jog）
    //   · OpenRobotJogDialog        — 通过 CommandsManager 触发 PS 内置 Robot Jog
    //   · BtnCheck_Click / BtnCheckAll_Click / BtnReset_Click / BtnExport_Click / BtnOlpDiag_Click
    //   · StartCheck                — 编排检查任务（异步推进 + 进度回调）
    //   · BuildCheckOptions         — 把 UI 控件值打包为 CheckOptions
    //   · LoadRobotsAndOperations   — 文档校验 + 状态初始化
    //   · UpdateSummaryCards        — 状态栏汇总
    //   · ClearResults              — (实际定义在 Grid.cs)
    //   · OnFormClosing             — (实际定义在 ReachabilityCheckerForm.cs 主 partial)
    // =========================================================================
    public partial class ReachabilityCheckerForm
    {
        // PS 内置 Robot Jog 命令 ID（来自 RibbonConfiguration.xml）
        private const string CMD_ROBOT_JOG =
            "DnProcessSimulateCommands.RobotJog.CApRJRobotJogCmd";

        // =====================================================================
        // OP 节点拾取
        // =====================================================================
        private void OnOpNodePicked(object sender, EventArgs e)
        {
            if (_txtOpNode == null) return;
            ITxObject pickedObj = null;
            try { pickedObj = _txtOpNode.Object as ITxObject; }
            catch { return; }
            if (pickedObj == null) return;

            // 验证拾取对象是否为操作类型
            bool isOp = pickedObj is ITxRoboticOperation
                     || pickedObj is TxCompoundOperation
                     || pickedObj is ITxOperation
                     || pickedObj.GetType().Name.Contains("Operation");
            if (!isOp)
            {
                Log($"拾取对象 [{pickedObj.Name}] 不是操作类型({pickedObj.GetType().Name})，已忽略", "WARN");
                SetStatus($"请选择操作节点，当前对象类型: {pickedObj.GetType().Name}");
                try { _txtOpNode.Object = null; } catch { }
                _pickedOperation = null;
                return;
            }

            // 保存用户拾取到的具体对象实例（避免按名字查找时拿错副本）
            _pickedOperation = pickedObj;

            // 解析关联机器人 — 必须包 try/catch，否则 PS API 抛 SEH 异常会闪退进程
            try
            {
                _currentRobot = RobotFinder.FindAssociatedRobotSilent(pickedObj);
            }
            catch (Exception exFr)
            {
                _currentRobot = null;
                Log($"解析机器人异常（已忽略）: {exFr.Message}", "WARN");
            }

            try
            {
                string opName = pickedObj.Name;
                if (string.IsNullOrEmpty(opName)) return;

                if (pickedObj is ITxRoboticOperation rop)
                {
                    PreviewLocations(rop);
                }

                string robotInfo = _currentRobot != null ? $" → {_currentRobot.Name}" : " (未关联机器人)";
                Log($"OP节点拾取: {opName}{robotInfo}", "OK");
            }
            catch (Exception ex)
            {
                Log($"OP节点拾取处理异常: {ex.Message}", "WARN");
            }
        }

        // =====================================================================
        // 表格交互 — 单击：仅更新状态栏（不再驱动姿态，避免与双击冲突）
        // =====================================================================
        private void Grid_AfterSelChange(object sender, C1.Win.C1FlexGrid.RangeEventArgs e)
        {
            int row = _grid.RowSel;
            if (row < _grid.Rows.Fixed) return;
            if (!_rowToResult.TryGetValue(row, out var res)) return;

            SetStatus($"[{res.PointName}]  {res.RobotName}  " +
                      $"J1={res.J1:F1}° J2={res.J2:F1}° J3={res.J3:F1}°    " +
                      $"双击：驱动姿态 ▌ 三击：Robot Jog");
        }

        // =====================================================================
        // 表格交互 — 三击检测：在双击发生后 _tripleClickWindowMs 内再来一次单击
        //                   → 视为三击，触发 Robot Jog
        // =====================================================================
        private void Grid_SingleClickForTriple(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button != MouseButtons.Left) return;

                var ht = _grid.HitTest(e.X, e.Y);
                int row = ht.Row;
                if (row < _grid.Rows.Fixed) return;

                // 必须是 "双击留下的行" + 还在时间窗内
                if (row != _lastDoubleClickRow) return;
                if ((DateTime.UtcNow - _lastDoubleClickTime).TotalMilliseconds > _tripleClickWindowMs) return;

                // 已确认三击 → 重置窗口，避免连发
                _lastDoubleClickRow = -1;
                _lastDoubleClickTime = DateTime.MinValue;

                Log($">>> 三击触发 (row={row}) → 打开 Robot Jog", "DEBUG");
                Grid_TripleClick_Core(row);
            }
            catch (Exception ex)
            {
                Log($"三击处理异常: {ex.Message}", "ERR");
            }
        }

        // =====================================================================
        // 表格交互 — 双击行 → 驱动机器人到该点位姿态
        //                   并记录 (row, time) 供三击检测使用
        // =====================================================================
        private void Grid_DblClick_Mouse(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button != MouseButtons.Left) return;

                var ht = _grid.HitTest(e.X, e.Y);
                int row = ht.Row;
                if (row < _grid.Rows.Fixed) return;

                _grid.RowSel = row;

                // 记录双击发生 — 供三击检测
                _lastDoubleClickRow = row;
                _lastDoubleClickTime = DateTime.UtcNow;

                Grid_DblClick_DrivePose(row);
            }
            catch (Exception ex)
            {
                Log($"双击处理异常: {ex.Message}", "ERR");
            }
        }

        // 双击：驱动机器人姿态
        // 主路径：res.PoseDataRef 缓存了 TxPoseData → robot.CurrentPose = pd 一次写入（最快）
        // 兜底：逐 joint 写 CurrentValue（缓存丢失时使用，慢但可靠）
        private void Grid_DblClick_DrivePose(int row)
        {
            if (!_rowToResult.TryGetValue(row, out var res)) return;

            // 未检查/不可达 状态下没有有效关节值，跳过驱动
            if (res.Status == ReachabilityStatus.Unreachable
                || (res.J1 == 0 && res.J2 == 0 && res.J3 == 0
                    && res.J4 == 0 && res.J5 == 0 && res.J6 == 0
                    && res.JointMargin >= 999))
            {
                SetStatus($"[{res.PointName}] 无有效姿态可驱动");
                return;
            }

            TxRobot robot = _currentRobot;
            if (robot == null) return;

            // ── 主路径：用缓存的 TxPoseData 一次性写入 ─────────────────────
            //    PS 内部对 CurrentPose 赋值会做批量更新（一次重绘 + 一次运动学求解），
            //    比逐 joint 写 CurrentValue 触发 6 次完整重绘快得多
            if (res.PoseDataRef is TxPoseData pd && pd != null)
            {
                Cursor savedCursor = this.Cursor;
                this.Cursor = Cursors.WaitCursor;
                SetStatus($"▶ 驱动中...");
                try
                {
                    robot.CurrentPose = pd;
                    SetStatus($"▶ 已驱动到 [{res.PointName}]   再次单击同一行可打开 Robot Jog");
                    return;
                }
                catch (Exception ex)
                {
                    Log($"CurrentPose 赋值失败，回退逐 joint 写入: {ex.Message}", "WARN");
                    // 落到下面的兜底
                }
                finally
                {
                    this.Cursor = savedCursor;
                }
            }

            // ── 兜底：逐 joint 写 CurrentValue ───────────────────────────
            Cursor savedCursor2 = this.Cursor;
            this.Cursor = Cursors.WaitCursor;
            SetStatus($"▶ 驱动中...");
            try
            {
                var joints = robot.Joints;
                if (joints == null || joints.Count == 0) return;

                double[] target = { res.J1, res.J2, res.J3, res.J4, res.J5, res.J6 };
                int n = Math.Min(joints.Count, target.Length);

                for (int i = 0; i < n; i++)
                {
                    try
                    {
                        TxJoint joint = joints[i] as TxJoint;
                        if (joint == null) continue;

                        double valDeg = target[i];

                        double curVal = 0;
                        bool sawCur = false;
                        try { curVal = joint.CurrentValue; sawCur = true; } catch { }

                        double valToWrite;
                        if (sawCur && Math.Abs(curVal) <= 2 * Math.PI + 0.05)
                            valToWrite = valDeg * Math.PI / 180.0;
                        else
                            valToWrite = valDeg;

                        joint.CurrentValue = valToWrite;
                    }
                    catch { /* 单关节失败不影响其他 */ }
                }

                SetStatus($"▶ 已驱动到 [{res.PointName}]   再次单击同一行可打开 Robot Jog");
            }
            catch (Exception ex)
            {
                Log($"驱动姿态异常: {ex.Message}", "ERR");
            }
            finally
            {
                this.Cursor = savedCursor2;
            }
        }

        // 三击：选中点位 + 打开 Robot Jog
        private void Grid_TripleClick_Core(int row)
        {
            try
            {
                if (!_rowToResult.TryGetValue(row, out var res))
                {
                    Log($"行 {row} 未在 _rowToResult 字典中找到对应数据 (字典大小={_rowToResult.Count})", "WARN");
                    return;
                }

                string pointName = res.PointName;
                string opName = res.OperationName;
                if (string.IsNullOrEmpty(pointName))
                {
                    Log("当前行没有点位名，无法定位", "WARN");
                    return;
                }
                Log($"三击点位: [{opName}] / [{pointName}]", "DEBUG");

                TxDocument doc = TxApplication.ActiveDocument;
                if (doc == null)
                {
                    Log("ActiveDocument 为 null", "ERR");
                    return;
                }

                ITxObject locObj = LocationEnumerator.FindLocationInDoc(doc, opName, pointName, this);
                if (locObj == null)
                {
                    Log($"未在 PS 文档中找到点位: {pointName}", "ERR");
                    MessageBox.Show($"未在当前 Study 中找到点位 [{pointName}]，\n" +
                                    "可能该点位已被删除或操作已变更。",
                                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 设置 PS ActiveSelection（Robot Jog 会读取它）
                try
                {
                    TxSelection sel = TxApplication.ActiveSelection;
                    sel.Clear();

                    TxObjectList list = new TxObjectList();
                    list.Add(locObj);
                    sel.AddItems(list);

                    Log($"已设置 ActiveSelection: {locObj.Name}", "DEBUG");
                }
                catch (Exception exSel)
                {
                    Log($"设置 ActiveSelection 失败: {exSel.Message}", "ERR");
                    return;
                }

                OpenRobotJogDialog();
            }
            catch (Exception ex)
            {
                Log($"三击处理异常: {ex.Message}", "ERR");
            }
        }

        // =====================================================================
        // 调用 PS 内置 Robot Jog 命令
        //
        // 关键：ExecuteCommand 是同步阻塞的（直到用户关闭 Robot Jog 才返回）。
        // Robot Jog 关闭后清空 ActiveSelection，避免影响后续 OP 节点拾取。
        // =====================================================================
        private void OpenRobotJogDialog()
        {
            TxCommandsManager mgr = null;
            try { mgr = TxApplication.CommandsManager; }
            catch (Exception exMgr)
            {
                Log($"获取 CommandsManager 异常: {exMgr.Message}", "ERR");
                return;
            }

            if (mgr == null)
            {
                Log("CommandsManager 为 null", "ERR");
                return;
            }

            try
            {
                Log($"  执行命令: {CMD_ROBOT_JOG}", "DEBUG");
                mgr.ExecuteCommand(CMD_ROBOT_JOG);
                Log($"  ✓ Robot Jog 已触发", "DEBUG");
            }
            catch (TxCommandIdentifierDoesNotExistException)
            {
                Log($"  命令ID不存在: {CMD_ROBOT_JOG}（PS 版本可能不同）", "ERR");
                MessageBox.Show(
                    $"Robot Jog 命令未注册到当前 PS 实例。\n\n" +
                    $"已使用的命令ID:\n  {CMD_ROBOT_JOG}\n\n" +
                    $"目标点位已选中，可手动点击工具栏的 Robot Jog 按钮。",
                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (TxCannotActivateCommandException exAct)
            {
                Log($"  ✓ Robot Jog 命令已激活: {exAct.Message}", "OK");
            }
            catch (Exception ex)
            {
                Log($"  执行命令异常: {ex.GetType().Name} - {ex.Message}", "ERR");
            }
            finally
            {
                // Robot Jog 关闭后，清空 ActiveSelection 避免影响后续拾取
                try { TxApplication.ActiveSelection.Clear(); } catch { }
            }
        }

        // =====================================================================
        // 按钮事件
        // =====================================================================
        private void BtnCheck_Click(object sender, EventArgs e)
        {
            if (_pickedOperation == null)
            {
                SetStatus("请先在 [OP节点] 中拾取要检查的操作");
                MessageBox.Show("请先在「OP节点」拾取框中点击场景树或 3D 视图中的操作节点。",
                    "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            StartCheck();
        }

        private void BtnCheckAll_Click(object sender, EventArgs e)
        {
            // 弃用：批量按名字查找会拿错同名操作的副本
            SetStatus("当前仅支持单路径检查，请使用「开始检查」");
            MessageBox.Show(
                "为了避免同名操作导致的检查错误，「检查所有路径」功能已暂时改为单路径模式。\n\n" +
                "请在「OP节点」中拾取要检查的具体操作，然后点击「开始检查」。",
                "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            // 合并原"刷新"功能：清空结果 + 清空拾取记忆 + 重新加载文档状态
            ClearResults();
            _pickedOperation = null;
            _currentRobot = null;
            try { if (_txtOpNode != null) _txtOpNode.Object = null; } catch { }
            LoadRobotsAndOperations();
        }

        private void BtnOlpDiag_Click(object sender, EventArgs e)
        {
            try
            {
                if (_pickedOperation == null)
                {
                    MessageBox.Show("请先在 OP节点 拾取一个操作", "提示");
                    return;
                }
                OlpDiagnostic.Run(_pickedOperation, this);
            }
            catch (Exception ex)
            {
                Log($"[OLP诊断] 总体异常: {ex.Message}", "ERR");
            }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (_currentTask == null || _currentTask.Results.Count == 0)
            {
                SetStatus("没有可导出的数据");
                return;
            }

            using (var dlg = new SaveFileDialog
            {
                Filter = "CSV文件|*.csv",
                FileName = $"Reachability_{_currentTask.PathName}_{DateTime.Now:yyyyMMdd_HHmm}",
                Title = "导出可达性报告"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;

                var lines = new List<string>
                {
                    $"# 机器人路径可达性检查报告",
                    $"# 机器人: {_currentTask.RobotName}  路径: {_currentTask.PathName}  时间: {_currentTask.CheckTime:yyyy-MM-dd HH:mm:ss}",
                    $"# 总计:{_currentTask.TotalPoints}  可达:{_currentTask.ReachableCount}  不可达:{_currentTask.UnreachableCount}  近极限:{_currentTask.NearLimitCount}  奇异:{_currentTask.SingularCount}  临界:{_currentTask.CriticalCount}  可达率:{_currentTask.ReachabilityRate:F1}%",
                    "",
                    "序号,机器人名,操作名称,点名,点类型,状态,J1(°),J2(°),J3(°),J4(°),J5(°),J6(°),最小余量(°),备注"
                };
                foreach (var r in _currentTask.Results)
                {
                    string st =
                          r.Status == ReachabilityStatus.Reachable ? "可达"
                        : r.Status == ReachabilityStatus.Unreachable ? "不可达"
                        : r.Status == ReachabilityStatus.NearLimit ? "接近极限"
                        : r.Status == ReachabilityStatus.Singular ? "奇异"
                        : r.Status == ReachabilityStatus.Critical ? "临界"
                        : "未检查";
                    lines.Add(string.Join(",",
                        r.Index, r.RobotName, r.OperationName, r.PointName,
                        r.PointType, st,
                        r.J1.ToString("F1"), r.J2.ToString("F1"), r.J3.ToString("F1"),
                        r.J4.ToString("F1"), r.J5.ToString("F1"), r.J6.ToString("F1"),
                        r.JointMargin.ToString("F1"), $"\"{r.ErrorMessage}\""));
                }
                System.IO.File.WriteAllLines(dlg.FileName, lines, Encoding.UTF8);
                SetStatus($"✓ 已导出: {System.IO.Path.GetFileName(dlg.FileName)}");
            }
        }

        // =====================================================================
        // 检查任务编排
        // =====================================================================
        private CheckOptions BuildCheckOptions()
        {
            var opts = new CheckOptions
            {
                JointMarginCheckEnabled = _chkJointMargin != null && _chkJointMargin.Checked,
                JointMarginThreshDeg = _nudJointMarginDeg != null
                    ? (double)_nudJointMarginDeg.Value : 10.0,
                TcpCheckEnabled = _chkTcpXyz != null && _chkTcpXyz.Checked,
                TcpMarginMm = _nudTcpMargin != null
                    ? (double)_nudTcpMargin.Value : 200.0,
                InterferenceEnabled = _chkStaticInterference != null && _chkStaticInterference.Checked,
                UserSelectedBrand = RobotBrand.Auto
            };

            if (_cbBrand != null && _cbBrand.SelectedIndex > 0)
            {
                switch (_cbBrand.SelectedItem?.ToString())
                {
                    case "KUKA": opts.UserSelectedBrand = RobotBrand.KUKA; break;
                    case "ABB": opts.UserSelectedBrand = RobotBrand.ABB; break;
                    case "FANUC": opts.UserSelectedBrand = RobotBrand.FANUC; break;
                    case "其他": opts.UserSelectedBrand = RobotBrand.Other; break;
                }
            }
            return opts;
        }

        private void StartCheck()
        {
            if (_pickedOperation == null) return;

            string opName = _pickedOperation.Name ?? "(未命名)";
            ITxObject opForCheck = _pickedOperation;

            _tsProgress.Visible = true;
            if (_toolStrip != null) _toolStrip.Visible = true;
            _tsProgress.ProgressBar.Value = 0;
            SetStatus($"正在检查 [{opName}]...");
            Log("========================================");
            Log($"开始检查任务: {opName}");

            CheckOptions opts = BuildCheckOptions();

            // 用 BeginInvoke 让窗体先把"正在检查"状态绘出来再阻塞
            this.BeginInvoke(new Action(() =>
            {
                List<PathPointResult> results = null;
                try
                {
                    results = ReachabilityChecker.Check(opForCheck, opts, (done, total) =>
                    {
                        // 进度回调高频触发，UpdateProgress 自带节流
                        UpdateProgress(opName, done, total);
                    }, this);
                }
                catch (Exception ex)
                {
                    Log($"检查任务异常: {ex.Message}", "ERR");
                    results = new List<PathPointResult>();
                }

                string robotName = (_currentRobot != null) ? _currentRobot.Name
                                 : (results != null && results.Count > 0 ? results[0].RobotName : "未知");
                var task = new RobotPathCheckTask
                {
                    RobotName = robotName,
                    PathName = opName,
                    CheckTime = DateTime.Now,
                    Results = results ?? new List<PathPointResult>()
                };
                _tasks.Add(task);
                _currentTask = task;
                RefreshGrid(task.Results);
                UpdateSummaryCards(task);

                _tsProgress.Visible = false;
                if (_toolStrip != null) _toolStrip.Visible = false;
                SetStatus(
                    $"✓ 完成 | {task.TotalPoints} 点 | " +
                    $"可达 {task.ReachableCount} | 不可达 {task.UnreachableCount} | " +
                    $"近极限 {task.NearLimitCount} | 奇异 {task.SingularCount} | " +
                    $"临界 {task.CriticalCount} | 可达率 {task.ReachabilityRate:F1}%");
            }));
        }

        // 进度更新节流：每 100ms 最多刷一次（高频回调下避免 UI 卡顿）
        private DateTime _lastProgressTime = DateTime.MinValue;
        private void UpdateProgress(string opName, int done, int total)
        {
            DateTime now = DateTime.UtcNow;
            // 首尾必刷，中间按 100ms 节流
            bool isEdge = (done == 0) || (done >= total);
            if (!isEdge && (now - _lastProgressTime).TotalMilliseconds < 100) return;
            _lastProgressTime = now;

            if (_tsProgress?.ProgressBar != null)
            {
                int pct = total > 0 ? Math.Min(done * 100 / total, 100) : 0;
                _tsProgress.ProgressBar.Value = pct;
            }
            SetStatus($"检查中 [{opName}] {done}/{total}");
            Application.DoEvents();
        }

        // =====================================================================
        // 初始化与状态
        // =====================================================================
        private void LoadRobotsAndOperations()
        {
            // 启动期不主动访问 PS API（ActiveDocument 等），减少首次显示延迟
            SetStatus("● 就绪，请通过 OP节点 拾取要检查的操作");
        }

        private void UpdateSummaryCards(RobotPathCheckTask t)
        {
            SetStatus($"共 {t.TotalPoints} 点  可达 {t.ReachableCount}  " +
                      $"不可达 {t.UnreachableCount}  接近极限 {t.NearLimitCount}  " +
                      $"奇异 {t.SingularCount}  临界 {t.CriticalCount}  " +
                      $"可达率 {t.ReachabilityRate:F1}%");
        }

        // 注:ClearResults() 在 ReachabilityCheckerForm.Grid.cs 中实现
        // 注:OnFormClosing() 在 ReachabilityCheckerForm.cs (主 partial) 中实现
    }
}