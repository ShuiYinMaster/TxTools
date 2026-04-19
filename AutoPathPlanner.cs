using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;

namespace PS_AutoPathPlanner
{
    /// <summary>
    /// Process Simulate 自动路径规划器
    /// 
    /// 核心算法流程：
    /// 1. 获取焊点序列 → 生成进/出枪路径点（Z向偏移）
    /// 2. 逐段检测干涉 → X/Y/Z方向逐步搜索无干涉过渡点
    /// 3. 两焊点间插入中间安全点实现快速规划
    /// 4. 全方向干涉时，启用启发式猜测兜底策略
    /// </summary>
    public class AutoPathPlanner
    {
        #region ===== 配置参数 =====

        /// <summary>进/出枪点距焊点的Z向偏移量(mm)</summary>
        public double ApproachRetractDistance { get; set; } = 50.0;

        /// <summary>干涉搜索步长(mm)</summary>
        public double SearchStepSize { get; set; } = 10.0;

        /// <summary>干涉搜索最大偏移量(mm)</summary>
        public double MaxSearchOffset { get; set; } = 200.0;

        /// <summary>两焊点间中间点的数量</summary>
        public int IntermediatePointCount { get; set; } = 1;

        /// <summary>启发式猜测时的安全高度抬升(mm)</summary>
        public double FallbackLiftHeight { get; set; } = 150.0;

        /// <summary>碰撞检测间距(mm) — 用于路径离散化后逐段检测</summary>
        public double CollisionCheckResolution { get; set; } = 20.0;

        #endregion

        #region ===== 内部数据结构 =====

        /// <summary>
        /// 路径点类型枚举
        /// </summary>
        public enum PathPointType
        {
            /// <summary>焊接点（实际焊点）</summary>
            WeldPoint,
            /// <summary>进枪点（焊点前Z向偏移）</summary>
            ApproachPoint,
            /// <summary>出枪点（焊点后Z向偏移）</summary>
            RetractPoint,
            /// <summary>安全过渡点（干涉避让生成）</summary>
            SafeTransition,
            /// <summary>中间插值点</summary>
            IntermediatePoint,
            /// <summary>启发式猜测点（兜底策略）</summary>
            FallbackGuess,
            /// <summary>Home点</summary>
            HomePoint
        }

        /// <summary>
        /// 路径规划点 — 封装位置、姿态、类型等信息
        /// </summary>
        public class PlanPoint
        {
            public double X { get; set; }
            public double Y { get; set; }
            public double Z { get; set; }

            /// <summary>姿态矩阵(3x3旋转部分)，简化为欧拉角存储</summary>
            public double Rx { get; set; }
            public double Ry { get; set; }
            public double Rz { get; set; }

            public PathPointType PointType { get; set; }

            /// <summary>关联的焊点名称（用于溯源）</summary>
            public string SourceWeldName { get; set; } = "";

            /// <summary>该点是否经过干涉验证</summary>
            public bool IsCollisionFree { get; set; } = false;

            /// <summary>干涉搜索偏移记录(调试用)</summary>
            public string OffsetLog { get; set; } = "";

            public PlanPoint Clone()
            {
                return (PlanPoint)this.MemberwiseClone();
            }

            public override string ToString()
            {
                return $"[{PointType}] ({X:F1},{Y:F1},{Z:F1}) R({Rx:F1},{Ry:F1},{Rz:F1}) " +
                       $"Src={SourceWeldName} Free={IsCollisionFree}";
            }
        }

        /// <summary>
        /// 规划结果报告
        /// </summary>
        public class PlanningReport
        {
            public List<PlanPoint> FinalPath { get; set; } = new List<PlanPoint>();
            public int TotalWeldPoints { get; set; }
            public int CollisionFreeSegments { get; set; }
            public int FallbackPointsUsed { get; set; }
            public List<string> Warnings { get; set; } = new List<string>();
            public List<string> Logs { get; set; } = new List<string>();
            public TimeSpan ElapsedTime { get; set; }
        }

        #endregion

        #region ===== 主规划入口 =====

        private TxRobot _robot;
        private TxCollisionQueryCreation _collisionRoot;
        private List<string> _logBuffer = new List<string>();

        /// <summary>
        /// 执行自动路径规划
        /// </summary>
        /// <param name="robot">目标机器人</param>
        /// <param name="weldOps">焊接操作列表(按执行顺序排列)</param>
        /// <returns>规划报告（含完整路径）</returns>
        public PlanningReport ExecutePlanning(TxRobot robot, List<TxWeldOperation> weldOps)
        {
            var report = new PlanningReport();
            var sw = System.Diagnostics.Stopwatch.StartNew();
            _robot = robot;

            try
            {
                Log("========== 自动路径规划开始 ==========");
                Log($"机器人: {GetNameSafe(robot)}");
                Log($"焊点数量: {weldOps.Count}");
                Log($"进出枪距离: {ApproachRetractDistance}mm");
                Log($"搜索步长: {SearchStepSize}mm, 最大偏移: {MaxSearchOffset}mm");

                // ---- Step 0: 参数校验 ----
                if (robot == null || weldOps == null || weldOps.Count == 0)
                {
                    report.Warnings.Add("ERROR: 机器人或焊点列表为空，无法规划");
                    return report;
                }

                // ---- Step 1: 提取焊点位置，生成进出枪点 ----
                Log("\n----- Step 1: 生成进出枪路径点 -----");
                var rawPath = GenerateApproachRetractPoints(weldOps);
                Log($"生成路径点总数: {rawPath.Count} (含进出枪点)");
                report.TotalWeldPoints = weldOps.Count;

                // ---- Step 2: 初始化干涉集 ----
                Log("\n----- Step 2: 初始化干涉检测环境 -----");
                InitializeCollisionSet(robot);

                // ---- Step 3: 逐段干涉检测与避让 ----
                Log("\n----- Step 3: 逐段干涉检测与安全点插入 -----");
                var safePath = ProcessPathWithCollisionAvoidance(rawPath);

                // ---- Step 4: 两焊点间插入中间过渡点 ----
                Log("\n----- Step 4: 插入中间过渡点 -----");
                var enrichedPath = InsertIntermediatePoints(safePath);

                // ---- Step 5: 最终验证 ----
                Log("\n----- Step 5: 最终路径验证 -----");
                var finalPath = FinalValidation(enrichedPath);

                // ---- 汇总 ----
                report.FinalPath = finalPath;
                report.CollisionFreeSegments = finalPath.Count(p => p.IsCollisionFree);
                report.FallbackPointsUsed = finalPath.Count(p => p.PointType == PathPointType.FallbackGuess);

                if (report.FallbackPointsUsed > 0)
                {
                    report.Warnings.Add(
                        $"WARNING: 有 {report.FallbackPointsUsed} 个兜底猜测点，需人工复核");
                }

                Log($"\n========== 规划完成 ==========");
                Log($"最终路径点数: {finalPath.Count}");
                Log($"干涉安全点数: {report.CollisionFreeSegments}/{finalPath.Count}");
                Log($"兜底猜测点数: {report.FallbackPointsUsed}");
            }
            catch (Exception ex)
            {
                report.Warnings.Add($"EXCEPTION: {ex.Message}");
                Log($"[异常] {ex.Message}\n{ex.StackTrace}");
            }
            finally
            {
                sw.Stop();
                report.ElapsedTime = sw.Elapsed;
                report.Logs = new List<string>(_logBuffer);
                Log($"耗时: {sw.Elapsed.TotalSeconds:F2}s");
            }

            return report;
        }

        #endregion

        #region ===== Step 1: 焊点提取 + 进出枪点生成 =====

        /// <summary>
        /// 从焊接操作列表提取焊点位置，并为每个焊点生成进/出枪点
        /// 进枪点 = 焊点位置 + Z向正偏移
        /// 出枪点 = 焊点位置 + Z向正偏移
        /// </summary>
        private List<PlanPoint> GenerateApproachRetractPoints(List<TxWeldOperation> weldOps)
        {
            var path = new List<PlanPoint>();

            for (int i = 0; i < weldOps.Count; i++)
            {
                var op = weldOps[i];
                string opName = GetNameSafe(op);

                // --- 获取焊点位置 ---
                // PS SDK 中焊接操作的位置获取方式因版本不同而异
                // 使用 dynamic + try/catch 防御性获取
                double wx = 0, wy = 0, wz = 0;
                double wrx = 0, wry = 0, wrz = 0;
                bool posOk = TryGetWeldPosition(op, out wx, out wy, out wz,
                                                    out wrx, out wry, out wrz);

                if (!posOk)
                {
                    Log($"  [警告] 焊点 {opName} 位置获取失败，跳过");
                    continue;
                }

                Log($"  焊点[{i}] {opName}: ({wx:F1},{wy:F1},{wz:F1})");

                // --- 生成进枪点 (Approach): Z + offset ---
                var approach = new PlanPoint
                {
                    X = wx, Y = wy, Z = wz + ApproachRetractDistance,
                    Rx = wrx, Ry = wry, Rz = wrz,
                    PointType = PathPointType.ApproachPoint,
                    SourceWeldName = opName
                };

                // --- 焊接点本身 ---
                var weld = new PlanPoint
                {
                    X = wx, Y = wy, Z = wz,
                    Rx = wrx, Ry = wry, Rz = wrz,
                    PointType = PathPointType.WeldPoint,
                    SourceWeldName = opName
                };

                // --- 生成出枪点 (Retract): Z + offset ---
                var retract = new PlanPoint
                {
                    X = wx, Y = wy, Z = wz + ApproachRetractDistance,
                    Rx = wrx, Ry = wry, Rz = wrz,
                    PointType = PathPointType.RetractPoint,
                    SourceWeldName = opName
                };

                path.Add(approach);
                path.Add(weld);
                path.Add(retract);
            }

            return path;
        }

        /// <summary>
        /// 防御性获取焊接操作的TCP位置
        /// 兼容不同PS版本的API差异
        /// </summary>
        private bool TryGetWeldPosition(TxWeldOperation op, 
            out double x, out double y, out double z,
            out double rx, out double ry, out double rz)
        {
            x = y = z = rx = ry = rz = 0;

            try
            {
                // 方式1: 通过 LocationFrame 获取 (常见于PS15+)
                dynamic dynOp = op;
                try
                {
                    var frame = dynOp.LocationFrame;
                    if (frame != null)
                    {
                        dynamic mat = frame.Matrix;
                        // TxMatrix 索引方式: mat[row, col] 或 mat.GetValue(row, col)
                        // 位置在第4列 (index 3), 前3行
                        try
                        {
                            x = (double)mat[0, 3];
                            y = (double)mat[1, 3];
                            z = (double)mat[2, 3];
                        }
                        catch
                        {
                            // 某些版本用 GetValue
                            x = (double)mat.GetValue(0, 3);
                            y = (double)mat.GetValue(1, 3);
                            z = (double)mat.GetValue(2, 3);
                        }

                        // 欧拉角提取(简化 — 从旋转矩阵推导)
                        TryExtractEulerAngles(mat, out rx, out ry, out rz);
                        return true;
                    }
                }
                catch { /* LocationFrame 不可用，尝试下一种方式 */ }

                // 方式2: 通过 WeldPoint 属性获取
                try
                {
                    var wp = dynOp.WeldPoint;
                    if (wp != null)
                    {
                        dynamic wpFrame = wp.LocationFrame ?? wp.AbsoluteLocation;
                        if (wpFrame != null)
                        {
                            dynamic mat = wpFrame.Matrix ?? wpFrame;
                            try
                            {
                                x = (double)mat[0, 3];
                                y = (double)mat[1, 3];
                                z = (double)mat[2, 3];
                            }
                            catch
                            {
                                x = (double)mat.GetValue(0, 3);
                                y = (double)mat.GetValue(1, 3);
                                z = (double)mat.GetValue(2, 3);
                            }
                            TryExtractEulerAngles(mat, out rx, out ry, out rz);
                            return true;
                        }
                    }
                }
                catch { /* WeldPoint 不可用 */ }

                // 方式3: 通过 AbsoluteLocation 获取
                try
                {
                    var absLoc = dynOp.AbsoluteLocation;
                    if (absLoc != null)
                    {
                        x = (double)absLoc.X;
                        y = (double)absLoc.Y;
                        z = (double)absLoc.Z;
                        try
                        {
                            rx = (double)absLoc.Rx;
                            ry = (double)absLoc.Ry;
                            rz = (double)absLoc.Rz;
                        }
                        catch { /* 姿态获取失败，用默认值 */ }
                        return true;
                    }
                }
                catch { /* AbsoluteLocation 不可用 */ }

                // 方式4: 反射遍历查找位置属性
                try
                {
                    var props = op.GetType().GetProperties();
                    foreach (var prop in props)
                    {
                        if (prop.Name.Contains("Location") || prop.Name.Contains("Position"))
                        {
                            var val = prop.GetValue(op);
                            if (val != null)
                            {
                                dynamic dv = val;
                                try
                                {
                                    x = (double)dv.X;
                                    y = (double)dv.Y;
                                    z = (double)dv.Z;
                                    return true;
                                }
                                catch { continue; }
                            }
                        }
                    }
                }
                catch { /* 反射查找失败 */ }
            }
            catch (Exception ex)
            {
                Log($"  [异常] 获取焊点位置: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// 从旋转矩阵提取ZYX欧拉角(度)
        /// </summary>
        private void TryExtractEulerAngles(dynamic mat, 
            out double rx, out double ry, out double rz)
        {
            rx = ry = rz = 0;
            try
            {
                double r00, r01, r02, r10, r11, r12, r20, r21, r22;
                try
                {
                    r00 = (double)mat[0, 0]; r01 = (double)mat[0, 1]; r02 = (double)mat[0, 2];
                    r10 = (double)mat[1, 0]; r11 = (double)mat[1, 1]; r12 = (double)mat[1, 2];
                    r20 = (double)mat[2, 0]; r21 = (double)mat[2, 1]; r22 = (double)mat[2, 2];
                }
                catch
                {
                    r00 = (double)mat.GetValue(0, 0); r01 = (double)mat.GetValue(0, 1); r02 = (double)mat.GetValue(0, 2);
                    r10 = (double)mat.GetValue(1, 0); r11 = (double)mat.GetValue(1, 1); r12 = (double)mat.GetValue(1, 2);
                    r20 = (double)mat.GetValue(2, 0); r21 = (double)mat.GetValue(2, 1); r22 = (double)mat.GetValue(2, 2);
                }

                // ZYX 欧拉角分解
                ry = Math.Asin(-Clamp(r20, -1.0, 1.0));

                if (Math.Abs(Math.Cos(ry)) > 1e-6)
                {
                    rx = Math.Atan2(r21, r22);
                    rz = Math.Atan2(r10, r00);
                }
                else
                {
                    rx = Math.Atan2(-r12, r11);
                    rz = 0;
                }

                // 转为度
                rx = rx * 180.0 / Math.PI;
                ry = ry * 180.0 / Math.PI;
                rz = rz * 180.0 / Math.PI;
            }
            catch { /* 欧拉角提取失败，保持默认0 */ }
        }

        #endregion

        #region ===== Step 2: 干涉集初始化 =====

        /// <summary>
        /// 初始化碰撞检测环境
        /// 获取机器人+末端工具 vs 工件/夹具的干涉对
        /// </summary>
        private void InitializeCollisionSet(TxRobot robot)
        {
            try
            {
                // PS中碰撞检测通常通过 TxCollisionQueryCreation 或
                // TxApplication.ActiveDocument.CollisionRoot 来管理
                // 这里采用防御性方式获取

                Log("  正在获取碰撞检测根...");

                // 尝试通过 TxApplication 获取碰撞查询
                try
                {
                    dynamic app = TxApplication.ActiveDocument;
                    _collisionRoot = app.CollisionRoot as TxCollisionQueryCreation;
                }
                catch
                {
                    Log("  [信息] CollisionRoot 不可用，将使用替代碰撞检测");
                }

                Log("  碰撞检测环境初始化完成");
            }
            catch (Exception ex)
            {
                Log($"  [警告] 碰撞检测初始化异常: {ex.Message}");
            }
        }

        /// <summary>
        /// 核心碰撞检测方法
        /// 将机器人TCP移动到指定位置，检测是否碰撞
        /// </summary>
        /// <param name="point">目标位置</param>
        /// <returns>true=有碰撞, false=无碰撞</returns>
        private bool CheckCollisionAtPoint(PlanPoint point)
        {
            try
            {
                // ---- 移动机器人TCP到目标位置 ----
                if (!MoveRobotToPoint(point))
                {
                    // 移动失败(如超出工作空间)，视为有碰撞
                    return true;
                }

                // ---- 执行碰撞检测 ----
                return PerformCollisionCheck();
            }
            catch (Exception ex)
            {
                Log($"    [碰撞检测异常] {ex.Message}");
                return true; // 异常时保守处理，视为有碰撞
            }
        }

        /// <summary>
        /// 移动机器人TCP到指定点
        /// </summary>
        private bool MoveRobotToPoint(PlanPoint point)
        {
            try
            {
                dynamic robot = _robot;

                // 构造目标位姿矩阵
                // 方式1: 使用 TxTransformation
                try
                {
                    var pose = new TxTransformation();
                    pose.Translation = new TxVector(point.X, point.Y, point.Z);
                    pose.RotationRPY_ZYX = new TxVector(
                        point.Rx * Math.PI / 180.0,
                        point.Ry * Math.PI / 180.0,
                        point.Rz * Math.PI / 180.0);

                    // 尝试逆解并移动
                    try
                    {
                        robot.MoveImmediatelyTo(pose);
                        return true;
                    }
                    catch
                    {
                        // MoveImmediatelyTo 不可用，尝试其他方式
                        try
                        {
                            var tcpPose = robot.TTCP;
                            tcpPose.Translation = pose.Translation;
                            tcpPose.RotationRPY_ZYX = pose.RotationRPY_ZYX;
                            robot.TTCP = tcpPose;
                            return true;
                        }
                        catch { }
                    }
                }
                catch { }

                // 方式2: 使用 MoveTo + TxPoseData
                try
                {
                    dynamic poseData = Activator.CreateInstance(
                        robot.GetType().Assembly.GetType("Tecnomatix.Engineering.TxPoseData"));
                    poseData.X = point.X;
                    poseData.Y = point.Y;
                    poseData.Z = point.Z;
                    poseData.Rx = point.Rx;
                    poseData.Ry = point.Ry;
                    poseData.Rz = point.Rz;
                    robot.MoveTo(poseData);
                    return true;
                }
                catch { }

                // 方式3: ForwardKinematic 逆解
                try
                {
                    var targetFrame = new TxTransformation(
                        new TxVector(point.X, point.Y, point.Z),
                        new TxVector(point.Rx * Math.PI / 180.0,
                                     point.Ry * Math.PI / 180.0,
                                     point.Rz * Math.PI / 180.0));

                    var ikSolutions = robot.InverseKinematics(targetFrame);
                    if (ikSolutions != null)
                    {
                        dynamic firstSol = ((IEnumerable<object>)ikSolutions).FirstOrDefault();
                        if (firstSol != null)
                        {
                            robot.CurrentConfiguration = firstSol;
                            return true;
                        }
                    }
                }
                catch { }

                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 执行一次碰撞检测查询
        /// </summary>
        private bool PerformCollisionCheck()
        {
            try
            {
                // 方式1: 使用全局碰撞检测
                try
                {
                    dynamic doc = TxApplication.ActiveDocument;
                    var collisions = doc.CollisionRoot.GetCollidingPairs();
                    if (collisions != null)
                    {
                        int count = 0;
                        foreach (var pair in (IEnumerable<object>)collisions)
                        {
                            count++;
                        }
                        return count > 0;
                    }
                }
                catch { }

                // 方式2: 使用 TxCollisionQueryCreation
                try
                {
                    if (_collisionRoot != null)
                    {
                        dynamic result = ((dynamic)_collisionRoot).CheckCollisions();
                        if (result != null)
                        {
                            return (bool)result.HasCollisions;
                        }
                    }
                }
                catch { }

                // 方式3: TxPhysicsUtils 碰撞检测
                try
                {
                    var collisionPairs = TxApplication.ActiveDocument
                        .GetType().GetMethod("GetCollisions")
                        ?.Invoke(TxApplication.ActiveDocument, null);
                    if (collisionPairs != null)
                    {
                        dynamic pairs = collisionPairs;
                        return ((IEnumerable<object>)pairs).Any();
                    }
                }
                catch { }

                // 方式4: 使用 TxCollisionChecker 静态方法
                try
                {
                    var checkerType = AppDomain.CurrentDomain.GetAssemblies()
                        .SelectMany(a =>
                        {
                            try { return a.GetTypes(); }
                            catch { return new Type[0]; }
                        })
                        .FirstOrDefault(t => t.Name == "TxCollisionChecker"
                                          || t.Name == "TxCollisionUtils");

                    if (checkerType != null)
                    {
                        var checkMethod = checkerType.GetMethod("CheckCollisions",
                            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public);
                        if (checkMethod != null)
                        {
                            dynamic result = checkMethod.Invoke(null, null);
                            return result != null && (bool)result;
                        }
                    }
                }
                catch { }

                // 所有方式都失败 → 保守返回false(无碰撞)，避免阻塞流程
                Log("    [警告] 碰撞检测API均不可用，假设无碰撞");
                return false;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region ===== Step 3: 逐段干涉检测 + 安全点搜索 =====

        /// <summary>
        /// 对路径中相邻点对进行干涉检测
        /// 干涉时按 X→Y→Z 顺序搜索安全过渡点
        /// </summary>
        private List<PlanPoint> ProcessPathWithCollisionAvoidance(List<PlanPoint> rawPath)
        {
            var safePath = new List<PlanPoint>();
            if (rawPath.Count == 0) return safePath;

            // 第一个点直接检测
            rawPath[0].IsCollisionFree = !CheckCollisionAtPoint(rawPath[0]);
            safePath.Add(rawPath[0]);

            for (int i = 1; i < rawPath.Count; i++)
            {
                var prevPoint = safePath.Last();
                var currPoint = rawPath[i];

                Log($"\n  检测段 [{i - 1}]→[{i}]: {prevPoint.SourceWeldName}({prevPoint.PointType}) → " +
                    $"{currPoint.SourceWeldName}({currPoint.PointType})");

                // 检测当前点本身是否有碰撞
                bool currCollision = CheckCollisionAtPoint(currPoint);
                currPoint.IsCollisionFree = !currCollision;

                if (currCollision && currPoint.PointType == PathPointType.WeldPoint)
                {
                    // 焊点本身碰撞 → 这是工艺问题，记录警告但保留
                    Log($"    !! 焊点 {currPoint.SourceWeldName} 本身存在碰撞!");
                    safePath.Add(currPoint);
                    continue;
                }

                // 检测从上一点到当前点的移动过程中是否有碰撞
                bool segmentCollision = CheckSegmentCollision(prevPoint, currPoint);

                if (!segmentCollision)
                {
                    // 无碰撞，直接添加
                    Log($"    √ 无碰撞");
                    currPoint.IsCollisionFree = true;
                    safePath.Add(currPoint);
                }
                else
                {
                    // 有碰撞 → 搜索安全过渡点
                    Log($"    × 检测到碰撞，启动安全点搜索...");
                    var safePoint = SearchSafeTransitionPoint(prevPoint, currPoint);

                    if (safePoint != null)
                    {
                        Log($"    √ 找到安全过渡点: ({safePoint.X:F1},{safePoint.Y:F1},{safePoint.Z:F1}) " +
                            $"偏移: {safePoint.OffsetLog}");
                        safePath.Add(safePoint);
                    }
                    else
                    {
                        // 所有方向都碰撞 → 启用兜底策略
                        Log($"    !! 全方向碰撞，启用启发式猜测...");
                        var fallbackPoint = GenerateFallbackPoint(prevPoint, currPoint);
                        safePath.Add(fallbackPoint);
                    }

                    safePath.Add(currPoint);
                }
            }

            return safePath;
        }

        /// <summary>
        /// 检测两点之间的路径段是否存在碰撞
        /// 将路径离散化为多个检测点
        /// </summary>
        private bool CheckSegmentCollision(PlanPoint from, PlanPoint to)
        {
            double dist = Distance3D(from, to);
            int steps = Math.Max(2, (int)(dist / CollisionCheckResolution));

            for (int s = 1; s < steps; s++) // 跳过起点(已检测)
            {
                double t = (double)s / steps;
                var midPoint = Interpolate(from, to, t);

                if (CheckCollisionAtPoint(midPoint))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 核心搜索算法: 按X→Y→Z顺序搜索无碰撞的安全过渡点
        /// 
        /// 搜索策略:
        /// 1. 以两点中点为基准
        /// 2. 分别在X+/X-、Y+/Y-、Z+/Z-方向上步进搜索
        /// 3. 找到第一个无碰撞位置即返回
        /// 4. 优先Z正方向(抬起)，因为通常向上移动最安全
        /// </summary>
        private PlanPoint SearchSafeTransitionPoint(PlanPoint from, PlanPoint to)
        {
            // 基准点 = 两点中点
            var basePoint = Interpolate(from, to, 0.5);
            basePoint.PointType = PathPointType.SafeTransition;
            basePoint.SourceWeldName = $"{from.SourceWeldName}→{to.SourceWeldName}";

            // 搜索方向优先级: Z+(抬起) > Z-(下沉) > X+ > X- > Y+ > Y-
            // 但按题目要求先X再Y再Z排列搜索轴，每个轴双向
            var searchAxes = new[]
            {
                // (轴名, dx, dy, dz)
                ("X+", 1.0, 0.0, 0.0),
                ("X-", -1.0, 0.0, 0.0),
                ("Y+", 0.0, 1.0, 0.0),
                ("Y-", 0.0, -1.0, 0.0),
                ("Z+", 0.0, 0.0, 1.0),   // 最常用的安全方向
                ("Z-", 0.0, 0.0, -1.0),
            };

            // 逐步增大偏移量
            for (double offset = SearchStepSize; offset <= MaxSearchOffset; offset += SearchStepSize)
            {
                foreach (var (axisName, dx, dy, dz) in searchAxes)
                {
                    var testPoint = basePoint.Clone();
                    testPoint.X += dx * offset;
                    testPoint.Y += dy * offset;
                    testPoint.Z += dz * offset;

                    if (!CheckCollisionAtPoint(testPoint))
                    {
                        testPoint.IsCollisionFree = true;
                        testPoint.OffsetLog = $"{axisName} {offset:F1}mm";
                        return testPoint;
                    }
                }
            }

            // 尝试组合方向(对角线搜索)
            Log("    尝试组合方向搜索...");
            var comboAxes = new[]
            {
                ("XZ+", 1.0, 0.0, 1.0),
                ("XZ-", -1.0, 0.0, 1.0),
                ("YZ+", 0.0, 1.0, 1.0),
                ("YZ-", 0.0, -1.0, 1.0),
                ("XY+", 1.0, 1.0, 0.0),
                ("XY-", -1.0, -1.0, 0.0),
                ("XYZ+", 1.0, 1.0, 1.0),
                ("XYZ-", -1.0, -1.0, 1.0),
            };

            for (double offset = SearchStepSize; offset <= MaxSearchOffset; offset += SearchStepSize)
            {
                foreach (var (axisName, dx, dy, dz) in comboAxes)
                {
                    var testPoint = basePoint.Clone();
                    double norm = Math.Sqrt(dx * dx + dy * dy + dz * dz);
                    testPoint.X += (dx / norm) * offset;
                    testPoint.Y += (dy / norm) * offset;
                    testPoint.Z += (dz / norm) * offset;

                    if (!CheckCollisionAtPoint(testPoint))
                    {
                        testPoint.IsCollisionFree = true;
                        testPoint.OffsetLog = $"组合{axisName} {offset:F1}mm";
                        return testPoint;
                    }
                }
            }

            return null; // 所有方向均碰撞
        }

        #endregion

        #region ===== Step 4: 中间点插入 =====

        /// <summary>
        /// 在相邻出枪点和下一个进枪点之间插入中间过渡点
        /// 目的: 平滑路径、避免大幅度跳跃
        /// </summary>
        private List<PlanPoint> InsertIntermediatePoints(List<PlanPoint> path)
        {
            if (IntermediatePointCount <= 0) return path;

            var enriched = new List<PlanPoint>();

            for (int i = 0; i < path.Count; i++)
            {
                enriched.Add(path[i]);

                // 在出枪点和下一个进枪点之间插入中间点
                if (i < path.Count - 1
                    && path[i].PointType == PathPointType.RetractPoint
                    && path[i + 1].PointType == PathPointType.ApproachPoint)
                {
                    var from = path[i];
                    var to = path[i + 1];
                    double dist = Distance3D(from, to);

                    Log($"  在 {from.SourceWeldName}(出枪) → {to.SourceWeldName}(进枪) 间插入中间点" +
                        $" (距离={dist:F1}mm)");

                    for (int k = 1; k <= IntermediatePointCount; k++)
                    {
                        double t = (double)k / (IntermediatePointCount + 1);
                        var midPoint = Interpolate(from, to, t);

                        // 中间点默认抬高(安全策略)
                        double liftFactor = Math.Sin(t * Math.PI); // 抛物线抬升
                        midPoint.Z += FallbackLiftHeight * 0.5 * liftFactor;

                        midPoint.PointType = PathPointType.IntermediatePoint;
                        midPoint.SourceWeldName = $"{from.SourceWeldName}→{to.SourceWeldName}";

                        // 检测中间点碰撞
                        bool collision = CheckCollisionAtPoint(midPoint);
                        midPoint.IsCollisionFree = !collision;

                        if (collision)
                        {
                            Log($"    中间点[{k}]碰撞，搜索安全替代...");
                            var safeMid = SearchSafeTransitionPoint(from, to);
                            if (safeMid != null)
                            {
                                safeMid.PointType = PathPointType.IntermediatePoint;
                                enriched.Add(safeMid);
                            }
                            else
                            {
                                // 中间点也无法避开，用抬升替代
                                midPoint.Z += FallbackLiftHeight;
                                midPoint.IsCollisionFree = !CheckCollisionAtPoint(midPoint);
                                midPoint.OffsetLog = "中间点强制抬升";
                                enriched.Add(midPoint);
                            }
                        }
                        else
                        {
                            enriched.Add(midPoint);
                        }
                    }
                }
            }

            Log($"  中间点插入后路径总点数: {enriched.Count}");
            return enriched;
        }

        #endregion

        #region ===== Step 5: 兜底策略 + 最终验证 =====

        /// <summary>
        /// 启发式兜底策略: 当所有方向搜索都碰撞时，猜测一个相对安全的过渡点
        /// 
        /// 策略思路:
        /// 1. 大幅Z向抬升(跳过障碍物上方)
        /// 2. 向机器人基座方向回撤(通常靠近基座更开阔)
        /// 3. 沿两点连线的法向量偏移
        /// 4. 多点组合: 生成一条"抬起→平移→下降"的绕行路径
        /// </summary>
        private PlanPoint GenerateFallbackPoint(PlanPoint from, PlanPoint to)
        {
            // 策略A: 超高抬升
            var liftPoint = Interpolate(from, to, 0.5);
            liftPoint.Z += FallbackLiftHeight * 2;
            liftPoint.PointType = PathPointType.FallbackGuess;
            liftPoint.SourceWeldName = $"兜底:{from.SourceWeldName}→{to.SourceWeldName}";

            bool liftOk = !CheckCollisionAtPoint(liftPoint);
            if (liftOk)
            {
                liftPoint.IsCollisionFree = true;
                liftPoint.OffsetLog = $"兜底-超高抬升 Z+{FallbackLiftHeight * 2:F0}mm";
                Log($"    → 兜底策略A成功: Z+{FallbackLiftHeight * 2:F0}mm");
                return liftPoint;
            }

            // 策略B: 向机器人基座方向回撤
            try
            {
                double robotBaseX = 0, robotBaseY = 0, robotBaseZ = 0;
                try
                {
                    dynamic rob = _robot;
                    dynamic baseLoc = rob.BaseFrame ?? rob.AbsoluteLocation;
                    if (baseLoc != null)
                    {
                        try
                        {
                            dynamic mat = baseLoc.Matrix ?? baseLoc;
                            robotBaseX = (double)mat[0, 3];
                            robotBaseY = (double)mat[1, 3];
                            robotBaseZ = (double)mat[2, 3];
                        }
                        catch
                        {
                            robotBaseX = (double)baseLoc.X;
                            robotBaseY = (double)baseLoc.Y;
                            robotBaseZ = (double)baseLoc.Z;
                        }
                    }
                }
                catch { /* 无法获取基座位置 */ }

                var retractPoint = Interpolate(from, to, 0.5);
                // 向基座方向偏移30%
                retractPoint.X += (robotBaseX - retractPoint.X) * 0.3;
                retractPoint.Y += (robotBaseY - retractPoint.Y) * 0.3;
                retractPoint.Z += FallbackLiftHeight;
                retractPoint.PointType = PathPointType.FallbackGuess;
                retractPoint.SourceWeldName = liftPoint.SourceWeldName;

                if (!CheckCollisionAtPoint(retractPoint))
                {
                    retractPoint.IsCollisionFree = true;
                    retractPoint.OffsetLog = "兜底-向基座回撤+抬升";
                    Log($"    → 兜底策略B成功: 向基座回撤");
                    return retractPoint;
                }
            }
            catch { }

            // 策略C: 法向量偏移(垂直于两点连线方向)
            {
                double dx = to.X - from.X;
                double dy = to.Y - from.Y;
                double dz = to.Z - from.Z;
                double len = Math.Sqrt(dx * dx + dy * dy + dz * dz);

                if (len > 1e-6)
                {
                    // 计算垂直于连线的法向量 (简化: 取与Z轴的叉积)
                    double nx = -dy / len;
                    double ny = dx / len;
                    double nz = 0;

                    var normalPoint = Interpolate(from, to, 0.5);
                    normalPoint.X += nx * MaxSearchOffset;
                    normalPoint.Y += ny * MaxSearchOffset;
                    normalPoint.Z += FallbackLiftHeight;
                    normalPoint.PointType = PathPointType.FallbackGuess;
                    normalPoint.SourceWeldName = liftPoint.SourceWeldName;

                    if (!CheckCollisionAtPoint(normalPoint))
                    {
                        normalPoint.IsCollisionFree = true;
                        normalPoint.OffsetLog = "兜底-法向偏移+抬升";
                        Log($"    → 兜底策略C成功: 法向偏移");
                        return normalPoint;
                    }
                }
            }

            // 所有策略均失败 → 返回超高抬升点，标记为未验证
            Log($"    !! 所有兜底策略均失败，返回未验证超高点 (需人工处理)");
            liftPoint.Z += FallbackLiftHeight; // 再抬高一倍
            liftPoint.IsCollisionFree = false;
            liftPoint.OffsetLog = "兜底-最终猜测(未验证，需人工复核)";
            return liftPoint;
        }

        /// <summary>
        /// 最终路径验证: 逐点确认、清理异常点
        /// </summary>
        private List<PlanPoint> FinalValidation(List<PlanPoint> path)
        {
            Log($"  最终验证: {path.Count} 个点");

            int issueCount = 0;
            for (int i = 0; i < path.Count; i++)
            {
                var p = path[i];

                // 检查NaN/Infinity
                if (double.IsNaN(p.X) || double.IsNaN(p.Y) || double.IsNaN(p.Z) ||
                    double.IsInfinity(p.X) || double.IsInfinity(p.Y) || double.IsInfinity(p.Z))
                {
                    Log($"    [修复] 点[{i}] 包含NaN/Inf，重置为前一点");
                    if (i > 0)
                    {
                        p.X = path[i - 1].X;
                        p.Y = path[i - 1].Y;
                        p.Z = path[i - 1].Z + 50; // 略抬高
                    }
                    issueCount++;
                }

                // 检查相邻点距离是否过大(>1000mm可能有问题)
                if (i > 0)
                {
                    double d = Distance3D(path[i - 1], p);
                    if (d > 1000)
                    {
                        Log($"    [警告] 点[{i-1}]→[{i}] 距离={d:F1}mm 过大");
                        issueCount++;
                    }
                }
            }

            Log($"  验证完成，发现 {issueCount} 个问题");
            return path;
        }

        #endregion

        #region ===== 路径创建: 将规划结果写入PS =====

        /// <summary>
        /// 将规划好的路径点创建为PS中的机器人操作序列
        /// </summary>
        public bool CreatePathInPS(TxRobot robot, PlanningReport report, string pathName = "AutoPlannedPath")
        {
            try
            {
                Log("\n===== 在PS中创建路径 =====");

                var path = report.FinalPath;
                if (path.Count == 0)
                {
                    Log("  路径为空，无法创建");
                    return false;
                }

                // 创建机器人操作
                dynamic dynRobot = robot;

                // 尝试创建 RobotOperation/CompoundOperation
                try
                {
                    // 方式1: 通过 TxRobotOperationCreation
                    var opCreation = new TxRobotOperationCreation();

                    foreach (var point in path)
                    {
                        try
                        {
                            // 构造 Location
                            var pose = new TxTransformation();
                            pose.Translation = new TxVector(point.X, point.Y, point.Z);
                            pose.RotationRPY_ZYX = new TxVector(
                                point.Rx * Math.PI / 180.0,
                                point.Ry * Math.PI / 180.0,
                                point.Rz * Math.PI / 180.0);

                            // 创建路径点
                            string pointName = $"{pathName}_{point.PointType}_{point.SourceWeldName}";

                            // 使用 TxRoboticViaLocationOperation 添加路径点
                            try
                            {
                                dynamic loc = opCreation.CreateRoboticViaLocationOperation(
                                    new TxRoboticViaLocationOperationCreationData
                                    {
                                        Name = pointName,
                                    });

                                if (loc != null)
                                {
                                    loc.AbsolutePosition = pose;
                                    // 设置运动类型
                                    if (point.PointType == PathPointType.WeldPoint)
                                    {
                                        try { loc.MotionType = TxRoboticViaLocationOperationMotionType.MoveL; }
                                        catch { }
                                    }
                                    else
                                    {
                                        try { loc.MotionType = TxRoboticViaLocationOperationMotionType.MoveJ; }
                                        catch { }
                                    }
                                }
                            }
                            catch
                            {
                                // 备选: 直接创建 OLP 命令
                                try
                                {
                                    dynRobot.CreateLocation(pointName, pose);
                                }
                                catch { }
                            }
                        }
                        catch (Exception ex)
                        {
                            Log($"  [警告] 创建路径点失败: {point} - {ex.Message}");
                        }
                    }

                    Log($"  路径 '{pathName}' 创建完成，共 {path.Count} 个点");
                    return true;
                }
                catch (Exception ex)
                {
                    Log($"  [异常] 路径创建失败: {ex.Message}");
                }

                return false;
            }
            catch (Exception ex)
            {
                Log($"  [严重异常] {ex.Message}");
                return false;
            }
        }

        #endregion

        #region ===== 工具方法 =====

        private double Distance3D(PlanPoint a, PlanPoint b)
        {
            double dx = a.X - b.X;
            double dy = a.Y - b.Y;
            double dz = a.Z - b.Z;
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }

        /// <summary>
        /// 线性插值两个PlanPoint
        /// </summary>
        private PlanPoint Interpolate(PlanPoint from, PlanPoint to, double t)
        {
            return new PlanPoint
            {
                X = from.X + (to.X - from.X) * t,
                Y = from.Y + (to.Y - from.Y) * t,
                Z = from.Z + (to.Z - from.Z) * t,
                Rx = from.Rx + (to.Rx - from.Rx) * t,
                Ry = from.Ry + (to.Ry - from.Ry) * t,
                Rz = from.Rz + (to.Rz - from.Rz) * t,
            };
        }

        private double Clamp(double value, double min, double max)
        {
            return Math.Max(min, Math.Min(max, value));
        }

        private string GetNameSafe(object obj)
        {
            try
            {
                dynamic d = obj;
                return (string)d.Name;
            }
            catch
            {
                return obj?.GetType().Name ?? "Unknown";
            }
        }

        private void Log(string message)
        {
            _logBuffer.Add($"[{DateTime.Now:HH:mm:ss.fff}] {message}");
            System.Diagnostics.Debug.WriteLine(message);
        }

        #endregion
    }

    #region ===== 入口Form: 参数配置 + 启动规划 =====

    /// <summary>
    /// 自动路径规划器的交互界面
    /// </summary>
    public class AutoPathPlannerForm : Form
    {
        private NumericUpDown _nudApproachDist;
        private NumericUpDown _nudSearchStep;
        private NumericUpDown _nudMaxOffset;
        private NumericUpDown _nudMidPoints;
        private NumericUpDown _nudLiftHeight;
        private TextBox _txtLog;
        private Button _btnRun;
        private Button _btnClose;
        private ProgressBar _progressBar;
        private Label _lblStatus;

        public AutoPathPlannerForm()
        {
            InitializeUI();
        }

        private void InitializeUI()
        {
            this.Text = "PS 自动路径规划器 v1.0";
            this.Width = 700;
            this.Height = 700;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            int y = 15;
            int labelWidth = 200;
            int inputX = 220;

            // ---- 参数区域 ----
            var grpParams = new GroupBox
            {
                Text = "规划参数",
                Location = new System.Drawing.Point(10, y),
                Size = new System.Drawing.Size(665, 200)
            };

            int gy = 25;
            AddParamRow(grpParams, "进/出枪距离 (mm):", ref gy, out _nudApproachDist, 50, 10, 500);
            AddParamRow(grpParams, "干涉搜索步长 (mm):", ref gy, out _nudSearchStep, 10, 5, 100);
            AddParamRow(grpParams, "最大搜索偏移 (mm):", ref gy, out _nudMaxOffset, 200, 50, 1000);
            AddParamRow(grpParams, "中间过渡点数量:", ref gy, out _nudMidPoints, 1, 0, 5);
            AddParamRow(grpParams, "兜底抬升高度 (mm):", ref gy, out _nudLiftHeight, 150, 50, 500);

            this.Controls.Add(grpParams);
            y += 210;

            // ---- 按钮 ----
            _btnRun = new Button
            {
                Text = "▶ 开始规划",
                Location = new System.Drawing.Point(10, y),
                Size = new System.Drawing.Size(150, 35),
                BackColor = System.Drawing.Color.FromArgb(46, 139, 87),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            _btnRun.Click += OnRunClick;
            this.Controls.Add(_btnRun);

            _btnClose = new Button
            {
                Text = "关闭",
                Location = new System.Drawing.Point(170, y),
                Size = new System.Drawing.Size(100, 35),
                FlatStyle = FlatStyle.Flat
            };
            _btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(_btnClose);

            _lblStatus = new Label
            {
                Text = "就绪",
                Location = new System.Drawing.Point(280, y + 8),
                Size = new System.Drawing.Size(380, 25),
                ForeColor = System.Drawing.Color.DarkBlue
            };
            this.Controls.Add(_lblStatus);
            y += 45;

            // ---- 进度条 ----
            _progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(10, y),
                Size = new System.Drawing.Size(665, 20),
                Style = ProgressBarStyle.Marquee
            };
            this.Controls.Add(_progressBar);
            y += 30;

            // ---- 日志区 ----
            var lblLog = new Label
            {
                Text = "规划日志:",
                Location = new System.Drawing.Point(10, y),
                AutoSize = true
            };
            this.Controls.Add(lblLog);
            y += 20;

            _txtLog = new TextBox
            {
                Location = new System.Drawing.Point(10, y),
                Size = new System.Drawing.Size(665, 330),
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 9),
                WordWrap = false
            };
            this.Controls.Add(_txtLog);
        }

        private void AddParamRow(GroupBox parent, string label, ref int y,
            out NumericUpDown nud, decimal defaultVal, decimal min, decimal max)
        {
            var lbl = new Label
            {
                Text = label,
                Location = new System.Drawing.Point(15, y + 3),
                Size = new System.Drawing.Size(180, 22),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };
            parent.Controls.Add(lbl);

            nud = new NumericUpDown
            {
                Location = new System.Drawing.Point(200, y),
                Size = new System.Drawing.Size(100, 25),
                Minimum = min,
                Maximum = max,
                Value = defaultVal,
                DecimalPlaces = 0
            };
            parent.Controls.Add(nud);

            y += 30;
        }

        private void OnRunClick(object sender, EventArgs e)
        {
            _txtLog.Clear();
            _lblStatus.Text = "正在获取机器人和焊点...";
            _btnRun.Enabled = false;
            _progressBar.Visible = true;

            try
            {
                // ---- 获取当前机器人 ----
                TxRobot robot = GetCurrentRobot();
                if (robot == null)
                {
                    AppendLog("[错误] 未找到机器人，请先在PS中选择一个机器人");
                    _lblStatus.Text = "失败: 未找到机器人";
                    return;
                }
                AppendLog($"找到机器人: {GetNameSafe(robot)}");

                // ---- 获取焊接操作 ----
                var weldOps = GetWeldOperations(robot);
                if (weldOps.Count == 0)
                {
                    AppendLog("[错误] 未找到焊接操作，请确认机器人已分配焊接任务");
                    _lblStatus.Text = "失败: 未找到焊接操作";
                    return;
                }
                AppendLog($"找到焊接操作: {weldOps.Count} 个");

                // ---- 配置规划器 ----
                var planner = new AutoPathPlanner
                {
                    ApproachRetractDistance = (double)_nudApproachDist.Value,
                    SearchStepSize = (double)_nudSearchStep.Value,
                    MaxSearchOffset = (double)_nudMaxOffset.Value,
                    IntermediatePointCount = (int)_nudMidPoints.Value,
                    FallbackLiftHeight = (double)_nudLiftHeight.Value,
                };

                // ---- 执行规划 ----
                _lblStatus.Text = "正在执行路径规划...";
                Application.DoEvents();

                var report = planner.ExecutePlanning(robot, weldOps);

                // ---- 显示结果 ----
                foreach (var log in report.Logs)
                {
                    AppendLog(log);
                }

                AppendLog("\n========================================");
                AppendLog($"  规划结果摘要");
                AppendLog($"  总焊点: {report.TotalWeldPoints}");
                AppendLog($"  最终路径点: {report.FinalPath.Count}");
                AppendLog($"  安全验证通过: {report.CollisionFreeSegments}");
                AppendLog($"  兜底猜测点: {report.FallbackPointsUsed}");
                AppendLog($"  耗时: {report.ElapsedTime.TotalSeconds:F2}s");
                AppendLog("========================================");

                foreach (var w in report.Warnings)
                {
                    AppendLog($"  ⚠ {w}");
                }

                // ---- 询问是否写入PS ----
                if (report.FinalPath.Count > 0)
                {
                    var result = MessageBox.Show(
                        $"规划完成! 共 {report.FinalPath.Count} 个路径点\n" +
                        $"其中 {report.FallbackPointsUsed} 个兜底猜测点需要人工复核\n\n" +
                        "是否将路径写入PS?",
                        "路径规划完成",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        _lblStatus.Text = "正在写入PS...";
                        Application.DoEvents();

                        bool ok = planner.CreatePathInPS(robot, report);
                        _lblStatus.Text = ok ? "路径已成功写入PS" : "路径写入失败";
                    }
                }

                _lblStatus.Text = $"规划完成 - {report.FinalPath.Count}个点, 耗时{report.ElapsedTime.TotalSeconds:F1}s";
            }
            catch (Exception ex)
            {
                AppendLog($"[严重错误] {ex.Message}\n{ex.StackTrace}");
                _lblStatus.Text = "规划出错";
            }
            finally
            {
                _btnRun.Enabled = true;
                _progressBar.Visible = false;
            }
        }

        #region ---- PS对象获取 ----

        private TxRobot GetCurrentRobot()
        {
            try
            {
                // 方式1: 从当前选择获取
                try
                {
                    var sel = TxApplication.ActiveSelection;
                    if (sel != null)
                    {
                        foreach (var item in (IEnumerable<object>)sel.GetItems())
                        {
                            if (item is TxRobot robot) return robot;

                            // 选中的可能是机器人操作，从中找到机器人
                            try
                            {
                                dynamic dynItem = item;
                                var rob = dynItem.Robot as TxRobot;
                                if (rob != null) return rob;
                            }
                            catch { }
                        }
                    }
                }
                catch { }

                // 方式2: 获取场景中所有机器人
                try
                {
                    dynamic doc = TxApplication.ActiveDocument;
                    var objects = doc.GetAllObjects();
                    foreach (var obj in (IEnumerable<object>)objects)
                    {
                        if (obj is TxRobot robot) return robot;
                    }
                }
                catch { }

                // 方式3: 通过 PhysicalRoot 遍历
                try
                {
                    dynamic root = TxApplication.ActiveDocument.PhysicalRoot;
                    return FindRobotRecursive(root);
                }
                catch { }
            }
            catch { }

            return null;
        }

        private TxRobot FindRobotRecursive(dynamic node)
        {
            try
            {
                if (node is TxRobot robot) return robot;

                try
                {
                    var children = node.Children ?? node.GetChildren();
                    foreach (var child in (IEnumerable<object>)children)
                    {
                        var found = FindRobotRecursive(child);
                        if (found != null) return found;
                    }
                }
                catch { }
            }
            catch { }
            return null;
        }

        private List<TxWeldOperation> GetWeldOperations(TxRobot robot)
        {
            var ops = new List<TxWeldOperation>();
            try
            {
                // 方式1: 从机器人的操作列表获取
                try
                {
                    dynamic dynRobot = robot;
                    var operations = dynRobot.Operations ?? dynRobot.GetOperations();
                    foreach (var op in (IEnumerable<object>)operations)
                    {
                        if (op is TxWeldOperation weldOp)
                        {
                            ops.Add(weldOp);
                        }
                    }
                    if (ops.Count > 0) return ops;
                }
                catch { }

                // 方式2: 从操作树遍历
                try
                {
                    dynamic doc = TxApplication.ActiveDocument;
                    var opRoot = doc.OperationRoot;
                    CollectWeldOps(opRoot, ops);
                    if (ops.Count > 0) return ops;
                }
                catch { }

                // 方式3: GetAllDescendants
                try
                {
                    dynamic doc = TxApplication.ActiveDocument;
                    var allOps = doc.OperationRoot.GetAllDescendants(
                        new TxTypeFilter(typeof(TxWeldOperation)));
                    foreach (var op in (IEnumerable<object>)allOps)
                    {
                        if (op is TxWeldOperation weldOp)
                            ops.Add(weldOp);
                    }
                }
                catch { }
            }
            catch { }

            return ops;
        }

        private void CollectWeldOps(dynamic node, List<TxWeldOperation> ops)
        {
            try
            {
                if (node is TxWeldOperation weldOp)
                {
                    ops.Add(weldOp);
                    return;
                }

                try
                {
                    var children = node.Children ?? node.GetChildren();
                    foreach (var child in (IEnumerable<object>)children)
                    {
                        CollectWeldOps(child, ops);
                    }
                }
                catch { }
            }
            catch { }
        }

        #endregion

        private void AppendLog(string msg)
        {
            if (_txtLog.InvokeRequired)
            {
                _txtLog.Invoke(new Action(() => AppendLog(msg)));
                return;
            }
            _txtLog.AppendText(msg + Environment.NewLine);
        }

        private string GetNameSafe(object obj)
        {
            try { return ((dynamic)obj).Name; }
            catch { return obj?.GetType().Name ?? "?"; }
        }
    }

    #endregion

    #region ===== PS Add-in 注册入口 =====

    /// <summary>
    /// TxButtonCommand 注册 — 在PS菜单栏中添加"自动路径规划"按钮
    /// </summary>
    public class AutoPathPlannerCommand : TxButtonCommand
    {
        public override string Category => "自动化工具";
        public override string Name => "自动路径规划器";
        public override string Description => "基于干涉检测的焊接路径自动规划工具";

        public override void Execute(object cmdParams)
        {
            try
            {
                var form = new AutoPathPlannerForm();
                form.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动失败: {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public override bool Connect(object cmdParams) => true;

        public override void When()
        {
            // 当PS中有打开的文档时可用
            try
            {
                this.Enabled = TxApplication.ActiveDocument != null;
            }
            catch
            {
                this.Enabled = false;
            }
        }
    }

    #endregion
}
