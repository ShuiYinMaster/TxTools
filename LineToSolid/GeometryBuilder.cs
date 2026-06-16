using System;
using System.Collections.Generic;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.DataTypes;

namespace LineToSolid
{
    public enum CrossSectionType
    {
        Rectangle,
        Circle
    }

    public enum MultiPipeLayout
    {
        SymmetricCentered,
        OffsetFromStart
    }

    public class GeometryParams
    {
        public CrossSectionType Section { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double Diameter { get; set; }
        public string PartNamePrefix { get; set; } = "LTS_Solid";

        // 局部偏移（截面平面内）
        // X：沿"侧向"（与段方向垂直、近水平的方向）
        // Y：沿"上向"（与段方向、X 都垂直）
        public double OffsetX { get; set; } = 0;
        public double OffsetY { get; set; } = 0;

        // 多根圆柱
        public int PipeCount { get; set; } = 1;
        public double PipeSpacing { get; set; } = 0;
        public MultiPipeLayout PipeLayout { get; set; } = MultiPipeLayout.SymmetricCentered;

        // 圆柱拐角处理
        public bool RoundCornerForCylinder { get; set; } = false;
        /// <summary>
        /// 主线拐弯半径（多根圆柱时的"中心管中心线"半径）。
        /// 默认 0 表示用 PipeSpacing×N（最小避免内侧重叠的值）自动算。
        /// 必须 ≥ 多根管中最大的 |s_i|，否则内侧圆环退化。
        /// </summary>
        public double BendRadius { get; set; } = 0;
    }

    /// <summary>
    /// 已验证 PS 2402 API：
    ///   CreateResource / SetModelingScope / EndModeling
    ///   CreateSolidBox(TxBoxCreationData)
    ///   CreateSolidCylinder(TxCylinderCreationData)
    ///   TxTransformation(TxVector position, TxVector zDirection)
    /// 待验证（用反射 + try/catch 兜底）：
    ///   CreateSolidTorus(TxTorusCreationData) —— SDK 表中存在但属性未确认
    /// </summary>
    public static class GeometryBuilder
    {
        public class BuildResult
        {
            public ITxObject CreatedPart { get; set; }
            public int SuccessCount { get; set; }
            public int FailCount { get; set; }
            public List<string> Messages { get; set; } = new List<string>();
        }

        public static BuildResult BuildForSegments(List<LineSegment> segments, GeometryParams p)
        {
            var result = new BuildResult();
            if (segments == null || segments.Count == 0)
            {
                result.Messages.Add("无可用线段"); return result;
            }
            if (p == null) { result.Messages.Add("参数为空"); return result; }

            var doc = TxApplication.ActiveDocument;
            if (doc == null || doc.PhysicalRoot == null)
            {
                result.Messages.Add("ActiveDocument 或 PhysicalRoot 不可用"); return result;
            }
            var root = doc.PhysicalRoot;

            string ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string resName = string.Format("{0}_{1}", p.PartNamePrefix, ts);
            ITxComponent comp;
            try { comp = root.CreateResource(new TxResourceCreationData(resName)); }
            catch (Exception ex)
            {
                result.Messages.Add("Resource 创建失败：" + ex.Message); return result;
            }
            if (comp == null) { result.Messages.Add("Resource 创建返回 null"); return result; }

            try
            {
                if (!comp.CanOpenForModeling)
                { result.Messages.Add("CanOpenForModeling=false"); return result; }
                comp.SetModelingScope();
            }
            catch (Exception ex)
            {
                result.Messages.Add("SetModelingScope 失败：" + ex.Message); return result;
            }

            var tc = comp as TxComponent;
            int idx = 0;
            int total = segments.Count;
            foreach (var seg in segments)
            {
                idx++;
                string baseName = string.Format("seg_{0:D4}", idx);
                var prev = (idx > 1) ? segments[idx - 2] : null;
                var next = (idx < total) ? segments[idx] : null;
                try
                {
                    int created;
                    if (p.Section == CrossSectionType.Rectangle)
                    {
                        double extra = ComputeBoxExtension(seg, next, p.Width);
                        created = CreateBox(tc, seg, p, baseName, extra) ? 1 : 0;
                    }
                    else
                    {
                        // 圆柱：考虑相邻段做核减
                        created = CreateCylinders(tc, seg, prev, next, p, baseName);

                        // 拐角填补：单根用球，多根用同心环
                        if (p.RoundCornerForCylinder && next != null)
                        {
                            int tCreated;
                            if (p.PipeCount <= 1)
                                tCreated = CreateCornerSpheres(tc, seg, next, p,
                                    string.Format("corner_{0:D4}", idx));
                            else
                                tCreated = CreateCornerTori(tc, seg, next, p,
                                    string.Format("corner_{0:D4}", idx));
                            created += tCreated;
                        }
                    }

                    if (created > 0) result.SuccessCount += created;
                    else { result.FailCount++; result.Messages.Add(string.Format("段{0}：未创建", idx)); }
                }
                catch (Exception ex)
                {
                    result.FailCount++;
                    result.Messages.Add(string.Format("段{0}：异常 {1}", idx, ex.Message));
                }
            }

            // 注意：不调 EndModeling。
            // EndModeling 要求传 .cojt 路径将 Resource 保存到磁盘，
            // PS 2402 上这一步耗时较长（卡顿明显）且会抛 "Error in the application"
            // 警告。实测不调 EndModeling 时，PS 会在后续操作前自动收尾 modeling scope，
            // 不影响生成结果，效率明显更高。

            result.CreatedPart = (ITxObject)comp;
            return result;
        }

        // ---------- 偏移：在段方向的正交平面内 ----------

        /// <summary>
        /// 给定段方向，得到局部 X 轴（与段方向垂直、近水平）和 Y 轴（与段方向、X 都垂直）。
        /// 段方向接近世界 Z 时 X 取世界 X 方向；否则 X 取 cross(worldZ, dir) 归一化。
        /// </summary>
        private static void ComputeLocalAxes(TxVector dir, out TxVector localX, out TxVector localY)
        {
            TxVector worldZ = new TxVector(0, 0, 1);
            TxVector worldX = new TxVector(1, 0, 0);
            TxVector aux = Math.Abs(dir.X * worldZ.X + dir.Y * worldZ.Y + dir.Z * worldZ.Z) > 0.99
                ? worldX : worldZ;
            // X = normalize(aux × dir)
            localX = Normalize(Cross(aux, dir));
            // Y = normalize(dir × X)
            localY = Normalize(Cross(dir, localX));
        }

        /// <summary>
        /// 把"局部 X 偏移 + 局部 Y 偏移"应用到世界点上。
        /// </summary>
        private static TxVector ApplyOffset(TxVector basePoint, TxVector localX, TxVector localY,
            double offsetX, double offsetY)
        {
            return new TxVector(
                basePoint.X + localX.X * offsetX + localY.X * offsetY,
                basePoint.Y + localX.Y * offsetX + localY.Y * offsetY,
                basePoint.Z + localX.Z * offsetX + localY.Z * offsetY);
        }

        // ---------- 长方体 ----------

        private static double ComputeBoxExtension(LineSegment cur, LineSegment next, double width)
        {
            if (next == null || width <= 0) return 0;
            TxVector d1 = cur.Direction, d2 = next.Direction;
            double dot = d1.X * d2.X + d1.Y * d2.Y + d1.Z * d2.Z;
            if (dot >= 0.9999) return 0;
            if (dot <= -0.9999) return width * 0.5;
            double theta = Math.Acos(Math.Max(-1.0, Math.Min(1.0, dot)));
            double extra = (width * 0.5) / Math.Tan(theta * 0.5);
            if (extra < 0) extra = 0;
            if (extra > width * 10) extra = width * 10;
            return extra;
        }

        private static bool CreateBox(
            TxComponent comp, LineSegment seg, GeometryParams p, string name, double extra)
        {
            double length = seg.Length + extra;
            if (length < 1e-6) return false;

            TxVector localX, localY;
            ComputeLocalAxes(seg.Direction, out localX, out localY);
            TxVector startWithOffset = ApplyOffset(seg.Start, localX, localY, p.OffsetX, p.OffsetY);

            var edgeSizes = new TxVector(p.Width, p.Height, length);
            var xform = new TxTransformation(startWithOffset, seg.Direction);
            var offset = new TxVector(0, 0, 0);

            var data = new TxBoxCreationData(name, xform, edgeSizes, offset);
            var r = comp.CreateSolidBox(data);
            return r != null;
        }

        // ---------- 圆柱体 ----------

        /// <summary>
        /// 创建段对应的 N 根圆柱。
        /// 如果勾选了拐角过渡且 PipeCount > 1（多根管），每根管两端按拐弯半径核减：
        ///   核减长度 = R_i / tan(θ/2)
        ///   R_i = MainBendRadius - s_i  （s_i = 该根侧向偏移量，凹侧为正）
        /// 单根管不核减（拐角用球体填缝，不需要让出圆环空间）。
        /// </summary>
        private static int CreateCylinders(
            TxComponent comp, LineSegment seg, LineSegment prev, LineSegment next,
            GeometryParams p, string baseName)
        {
            int n = Math.Max(1, p.PipeCount);
            if (p.Diameter <= 0 || seg.Length < 1e-6) return 0;
            double radius = p.Diameter * 0.5;
            double spacing = Math.Max(0, p.PipeSpacing);

            TxVector localX_world, localY_world;
            ComputeLocalAxes(seg.Direction, out localX_world, out localY_world);
            // 多根 spacing 偏移方向：在 d1-d2 平面内（让圆柱和圆环群对齐）
            TxVector spacingAxis = ChooseSpacingAxis(seg, prev, next);

            // 用户的局部 X/Y 偏移仍按世界相关的 localX/localY（不影响多根 spacing 方向）
            TxVector startBase = ApplyOffset(seg.Start, localX_world, localY_world, p.OffsetX, p.OffsetY);
            TxVector endBase = ApplyOffset(seg.End, localX_world, localY_world, p.OffsetX, p.OffsetY);

            double mainBend = ResolveMainBendRadius(p);

            int ok = 0;
            for (int i = 0; i < n; i++)
            {
                double s_i;
                if (p.PipeLayout == MultiPipeLayout.OffsetFromStart)
                    s_i = i * spacing;
                else
                    s_i = (i - (n - 1) / 2.0) * spacing;

                double cutStart = 0;
                if (p.RoundCornerForCylinder && n > 1 && prev != null)
                    cutStart = ComputeCutLength(prev, seg, mainBend, s_i, isStartSide: true);
                double cutEnd = 0;
                if (p.RoundCornerForCylinder && n > 1 && next != null)
                    cutEnd = ComputeCutLength(seg, next, mainBend, s_i, isStartSide: false);

                // 各根管的起点/终点：先按 spacingAxis 侧向偏移 s_i，再沿段方向核减
                TxVector start = new TxVector(
                    startBase.X + spacingAxis.X * s_i + seg.Direction.X * cutStart,
                    startBase.Y + spacingAxis.Y * s_i + seg.Direction.Y * cutStart,
                    startBase.Z + spacingAxis.Z * s_i + seg.Direction.Z * cutStart);
                TxVector end = new TxVector(
                    endBase.X + spacingAxis.X * s_i - seg.Direction.X * cutEnd,
                    endBase.Y + spacingAxis.Y * s_i - seg.Direction.Y * cutEnd,
                    endBase.Z + spacingAxis.Z * s_i - seg.Direction.Z * cutEnd);

                double dx = end.X - start.X, dy = end.Y - start.Y, dz = end.Z - start.Z;
                if (Math.Sqrt(dx * dx + dy * dy + dz * dz) < 1e-3) continue;

                string name = (n == 1) ? baseName : string.Format("{0}_{1:D2}", baseName, i + 1);
                try
                {
                    var data = new TxCylinderCreationData(name, start, end, radius);
                    var r = comp.CreateSolidCylinder(data);
                    if (r != null) ok++;
                }
                catch { }
            }
            return ok;
        }

        /// <summary>
        /// 在 d1-d2 平面内，返回垂直于 segDir 的单位向量，方向选择"指向 bisector"那侧。
        /// 这是多根管 spacing 偏移的正确轴向 —— 让所有多根管和圆环群共享同一个平面坐标系。
        ///
        /// 当 d1 与 d2 平行时（无拐弯），返回世界相关的 localX（退化为 ComputeLocalAxes）。
        /// </summary>
        private static TxVector ComputeInPlaneSideAxis(TxVector segDir, LineSegment otherSeg)
        {
            if (otherSeg == null)
            {
                TxVector lx, ly;
                ComputeLocalAxes(segDir, out lx, out ly);
                return lx;
            }
            TxVector d2 = otherSeg.Direction;
            double dot = segDir.X * d2.X + segDir.Y * d2.Y + segDir.Z * d2.Z;
            // 共线时退化到世界相关
            if (Math.Abs(dot) >= 0.9999)
            {
                TxVector lx, ly;
                ComputeLocalAxes(segDir, out lx, out ly);
                return lx;
            }
            // d1-d2 平面法线
            TxVector planeNormal = Normalize(Cross(segDir, d2));
            // 在 d1-d2 平面内、垂直 segDir 的方向 = planeNormal × segDir
            TxVector side = Normalize(Cross(planeNormal, segDir));
            // 确认 side 方向 —— 让它指向 bisector（d2 - d1 单位化）
            TxVector bis = Normalize(new TxVector(d2.X - segDir.X, d2.Y - segDir.Y, d2.Z - segDir.Z));
            double sign = side.X * bis.X + side.Y * bis.Y + side.Z * bis.Z;
            if (sign < 0)
                side = new TxVector(-side.X, -side.Y, -side.Z);
            return side;
        }

        /// <summary>
        /// 为单根管在某段上选择 spacing 偏移轴。
        /// 如果该段两端都有拐弯邻段，用前后两个邻段平均的"in-plane side"；
        /// 只有一端有邻段，用那端的；都没邻段，退化到世界相关 localX。
        /// </summary>
        private static TxVector ChooseSpacingAxis(LineSegment seg, LineSegment prev, LineSegment next)
        {
            // 优先用 next 段的拐弯定向（圆柱终点处理是大头）
            if (next != null) return ComputeInPlaneSideAxis(seg.Direction, next);
            if (prev != null)
            {
                // 把 prev 反向作为参考段
                var prevReversed = new LineSegment
                {
                    Start = prev.End,
                    End = prev.Start
                };
                return ComputeInPlaneSideAxis(seg.Direction, prevReversed);
            }
            TxVector lx, ly;
            ComputeLocalAxes(seg.Direction, out lx, out ly);
            return lx;
        }
        private static double ResolveMainBendRadius(GeometryParams p)
        {
            if (p.BendRadius > 0) return p.BendRadius;
            // 自动：要避免内侧管重叠，最小拐弯半径 ≥ 最大 |s_i|
            // 对称中线：max|s_i| = ((N-1)/2) × spacing
            // 起点偏移：max|s_i| = (N-1) × spacing
            int n = Math.Max(1, p.PipeCount);
            double maxAbsS = (p.PipeLayout == MultiPipeLayout.OffsetFromStart)
                ? (n - 1) * p.PipeSpacing
                : ((n - 1) / 2.0) * p.PipeSpacing;
            double radius = p.Diameter * 0.5;
            // 留出 2× 管径的余量，避免太紧
            return Math.Max(maxAbsS + 2 * p.Diameter, p.Diameter * 3);
        }

        /// <summary>
        /// 所有管的缩短长度统一为 T_0 = R / tan(θ/2)，与 s_i 无关。
        /// 这是为了保证平行的多根管长度一致（用户硬约束）。
        /// 环大半径 R_i 不同由 spacing 偏移天然形成同心环。
        /// </summary>
        private static double ComputeCutLength(LineSegment a, LineSegment b,
            double mainBend, double s_i, bool isStartSide)
        {
            TxVector d1 = a.Direction, d2 = b.Direction;
            double dot = d1.X * d2.X + d1.Y * d2.Y + d1.Z * d2.Z;
            if (dot >= 0.9999) return 0;
            if (dot <= -0.9999) return mainBend;

            double theta = Math.Acos(Math.Max(-1.0, Math.Min(1.0, dot)));
            double halfTheta = theta * 0.5;
            double tanHalf = Math.Tan(halfTheta);
            if (tanHalf < 1e-9) return 0;

            // 所有管缩短相同长度：T = R × tan(θ/2)
            return mainBend * tanHalf;
        }

        // ---------- 拐角处理：单根用球体 ----------

        /// <summary>
        /// 单根管拐角处填球体：球心 = apex，半径 = 管径/2。
        /// 球体完全覆盖两段圆柱端面接缝，无论夸角是任何角度都干净。
        /// PS 2402 API（按命名惯例使用强类型，若运行时签名不符会被外层 try/catch 捕获）：
        ///   new TxSphereCreationData(string name, TxVector center, double radius)
        ///   TxComponent.CreateSolidSphere(TxSphereCreationData) → TxSolid
        /// </summary>
        private static int CreateCornerSpheres(
            TxComponent comp, LineSegment cur, LineSegment next, GeometryParams p, string baseName)
        {
            double radius = p.Diameter * 0.5;
            if (radius <= 0) return 0;

            // 共线时无需填缝
            TxVector d1 = cur.Direction, d2 = next.Direction;
            double dot = d1.X * d2.X + d1.Y * d2.Y + d1.Z * d2.Z;
            if (dot >= 0.9999) return 0;

            TxVector localX, localY;
            ComputeLocalAxes(d1, out localX, out localY);
            // 球心 = cur.End + 用户局部偏移（与圆柱体保持一致）
            TxVector apex = ApplyOffset(cur.End, localX, localY, p.OffsetX, p.OffsetY);

            try
            {
                var data = new TxSphereCreationData(baseName, apex, radius);
                var r = comp.CreateSolidSphere(data);
                return r != null ? 1 : 0;
            }
            catch
            {
                return 0;
            }
        }

        // ---------- 拐角圆环 + 裁剪（同心环、共享切除盒） ----------

        /// <summary>
        /// 多根管拐角圆环（按用户确认的几何）：
        ///   1. 缩短长度：所有管统一 T = R / tan(θ/2)
        ///   2. 圆心 C：用中间管(s=0)缩短后端点作 cur 段法线（在 d1-d2 平面内、
        ///      指向凹侧的 spacingAxis 方向），延伸 R 到达
        ///      C = apex - d1×T + spacingAxis×R
        ///   3. 三个同心圆环：圆心同为 C，主轴同为 d1×d2，环大半径不同
        ///      R_i = R - s_i（s_i 沿 spacingAxis 方向，凹侧为正）
        ///   4. 切除盒：端面中心在 C，分别沿 -d1 和 +d2 延伸
        /// </summary>
        private static int CreateCornerTori(
            TxComponent comp, LineSegment cur, LineSegment next, GeometryParams p, string baseName)
        {
            int n = Math.Max(1, p.PipeCount);
            double radius = p.Diameter * 0.5;
            double spacing = Math.Max(0, p.PipeSpacing);

            TxVector d1 = cur.Direction, d2 = next.Direction;
            double dot = d1.X * d2.X + d1.Y * d2.Y + d1.Z * d2.Z;
            if (dot >= 0.9999 || dot <= -0.9999) return 0;

            double theta = Math.Acos(Math.Max(-1.0, Math.Min(1.0, dot)));
            double halfTheta = theta * 0.5;
            double tanHalf = Math.Tan(halfTheta);
            if (tanHalf < 1e-9) return 0;

            double R = ResolveMainBendRadius(p);
            // 缩短距离 T = R × tan(θ/2)（不是 R / tan）
            // 推导：圆心 = P_cur_0 + spacingAxis_cur × R = P_next_0 + spacingAxis_next × R
            // 解此方程得 T = R × tan(θ/2)
            double T = R * tanHalf;

            // 圆环主轴 = d1 × d2 归一化
            TxVector torusAxis = Normalize(Cross(d1, d2));

            // spacingAxis：在 d1-d2 平面内、垂直 cur 段、指向凹侧（next 段相对偏向）
            TxVector spacingAxis = ComputeInPlaneSideAxis(d1, next);

            // 用户偏移
            TxVector localX_world, localY_world;
            ComputeLocalAxes(d1, out localX_world, out localY_world);
            TxVector apex = ApplyOffset(cur.End, localX_world, localY_world, p.OffsetX, p.OffsetY);

            // 圆心 C
            TxVector center = new TxVector(
                apex.X - d1.X * T + spacingAxis.X * R,
                apex.Y - d1.Y * T + spacingAxis.Y * R,
                apex.Z - d1.Z * T + spacingAxis.Z * R);

            // 切除盒尺寸
            // 注意：TxBoxCreationData 的 edgeSizes 是局部 X/Y/Z 三个方向的边长，
            // 但盒子姿态用 TxTransformation(center, zDir) 只规定了 Z 方向，
            // 局部 X、Y 方向是 PS 自动选的，未必对应我想要的"d1-d2 平面内 vs torusAxis"。
            // 为避免 X/Y 方向猜测错误导致切除范围不够，三个边长取 max 用立方体形状。
            // 立方体边长 = max(长, 宽×2, 厚×2) —— 至少 2 倍以确保 X/Y 都够大
            double maxAbsS = (p.PipeLayout == MultiPipeLayout.OffsetFromStart)
                ? (n - 1) * spacing
                : ((n - 1) / 2.0) * spacing;
            double needLen = (R + maxAbsS + radius) * 2.0;            // 沿 -d1/+d2 方向需要的长
            double needWide = (maxAbsS + radius) * 2.0 + radius * 4;  // d1-d2 平面内需要的宽
            double needThick = radius * 4.0;                          // 沿 torusAxis 方向需要的厚
            double cubeSize = Math.Max(needLen, Math.Max(needWide, needThick));
            var boxSizes = new TxVector(cubeSize, cubeSize, cubeSize);
            var zeroOffset = new TxVector(0, 0, 0);

            TxSolid box1 = null, box2 = null;
            try
            {
                // 盒1：端面中心 = C，Z 方向沿 -d1
                var minusD1 = new TxVector(-d1.X, -d1.Y, -d1.Z);
                var box1Xform = new TxTransformation(center, minusD1);
                var box1Data = new TxBoxCreationData(baseName + "_cut1", box1Xform, boxSizes, zeroOffset);
                box1 = comp.CreateSolidBox(box1Data);

                // 盒2：端面中心 = C，Z 方向沿 +d2
                var box2Xform = new TxTransformation(center, d2);
                var box2Data = new TxBoxCreationData(baseName + "_cut2", box2Xform, boxSizes, zeroOffset);
                box2 = comp.CreateSolidBox(box2Data);
            }
            catch { }

            if (box1 == null || box2 == null) return 0;

            // N 个同心圆环
            int ok = 0;
            for (int i = 0; i < n; i++)
            {
                double s_i;
                if (p.PipeLayout == MultiPipeLayout.OffsetFromStart)
                    s_i = i * spacing;
                else
                    s_i = (i - (n - 1) / 2.0) * spacing;

                // R_i：s_i 同向 spacingAxis（凹侧）→ R_i 小（内侧管）
                //      s_i 反向 spacingAxis（凸侧）→ R_i 大（外侧管）
                double R_i = R - s_i;
                if (R_i <= radius * 1.1) continue;  // 内侧管过窄、跳过

                string name = (n == 1) ? baseName : string.Format("{0}_{1:D2}", baseName, i + 1);
                try
                {
                    var torusData = new TxTorusCreationData(
                        name + "_torus", center, R_i, radius, torusAxis);
                    TxSolid torus = comp.CreateSolidTorus(torusData);
                    if (torus == null) continue;

                    var subData = new TxSolidSubtractCreationData();
                    subData.Name = name;
                    subData.SolidToSubtractFrom = torus;
                    subData.DeleteEntities = false;  // 保留盒子供下一根管复用
                    subData.SolidsToSubtract = new TxSolid[] { box1, box2 };
                    var r = comp.CreateSolidBySubtract(subData);
                    if (r != null) ok++;

                    // 手动删除原始完整圆环（DeleteEntities=false 是为了保留盒子复用，
                    // 但原圆环也不会自动删，需要这里清掉）
                    try { ((ITxObject)torus).Delete(); } catch { }
                }
                catch { }
            }

            // 全部环裁剪完毕后清理两个切除盒
            try { ((ITxObject)box1).Delete(); } catch { }
            try { ((ITxObject)box2).Delete(); } catch { }
            return ok;
        }

        // ---------- 向量小工具 ----------

        private static TxVector Cross(TxVector a, TxVector b)
        {
            return new TxVector(
                a.Y * b.Z - a.Z * b.Y,
                a.Z * b.X - a.X * b.Z,
                a.X * b.Y - a.Y * b.X);
        }

        private static TxVector Normalize(TxVector v)
        {
            double L = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
            if (L < 1e-12) return new TxVector(1, 0, 0);
            return new TxVector(v.X / L, v.Y / L, v.Z / L);
        }
    }
}