using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Tecnomatix.Engineering;

namespace TxTools.RobotBaseChecker
{
    // 品牌处理模式（UI 选择）
    internal enum BrandMode { Auto, Generic, Fanuc }

    // 简易三维向量
    internal struct V3
    {
        public double X, Y, Z;
        public V3(double x, double y, double z) { X = x; Y = y; Z = z; }
        public static V3 operator -(V3 a, V3 b) => new V3(a.X - b.X, a.Y - b.Y, a.Z - b.Z);
        public static V3 operator +(V3 a, V3 b) => new V3(a.X + b.X, a.Y + b.Y, a.Z + b.Z);
        public static V3 operator *(V3 a, double s) => new V3(a.X * s, a.Y * s, a.Z * s);
        public double Dot(V3 b) => X * b.X + Y * b.Y + Z * b.Z;
        public double Len => Math.Sqrt(X * X + Y * Y + Z * Z);
        public V3 Norm() { double l = Len; return l < 1e-12 ? this : new V3(X / l, Y / l, Z / l); }
        public override string ToString() => $"({X:F3}, {Y:F3}, {Z:F3})";
    }

    // ============================================================
    //  方向一：基于运动学动态计算 FANUC 的 World / BASE0
    //
    //  FANUC（非倒挂）World 原点定义：
    //    J1 轴（竖直）上、距 J2 轴（水平）最近的点
    //    （即 J2 轴沿 J1 轴“水平交叉”的位置）。
    //  该点对 J1/J2 当前角度不敏感（J1 旋转不改变 J1 上的最近点高度，
    //   J2 旋转不移动 J2 轴本身），故无需先回零位。
    //
    //  其它品牌（KUKA/ABB/Yaskawa 等）：World 原点 ≈ 底座安装面中心，
    //    直接取机器人自身坐标（robot.AbsoluteLocation）。
    // ============================================================
    internal static class RobotKinematics
    {
        // ---------- 品牌识别 ----------
        // 优先级：① 控制器类型串（最可靠，区分所有品牌）
        //         ② 名称反品牌排除（明确是 KUKA/ABB/Yaskawa → 直接 Generic）
        //         ③ FANUC 实例参数特征（收紧：OLP_TOOL_NAME_ 单独不算，需含 FANUC 值）
        //         ④ 名称模糊匹配 FANUC 型号库
        public static string DetectBrand(TxRobot robot, BrandMode mode, StringBuilder log)
        {
            if (mode == BrandMode.Fanuc) { log.AppendLine("  品牌=FANUC (手动指定)"); return "FANUC"; }
            if (mode == BrandMode.Generic) { log.AppendLine("  品牌=Generic (手动指定)"); return "Generic"; }

            // ① 控制器类型（PDPS 中机器人挂的控制器，如 Fanuc-Rj / Kuka-Krc / Abb-Rapid）
            string ctrl = TryGetControllerType(robot);
            if (!string.IsNullOrEmpty(ctrl))
            {
                if (IsNoController(ctrl))
                {
                    // "default"/空 表示未分配真实控制器 → 不能据此判 Generic，继续走特征/名称
                    log.AppendLine("  控制器类型: " + ctrl + " (未分配真实控制器，转名称/特征识别)");
                }
                else
                {
                    log.AppendLine("  控制器类型: " + ctrl);
                    if (ctrl.ToUpperInvariant().Contains("FANUC")) return "FANUC";
                    return "Generic"; // 明确的非发那科控制器 → 底座面
                }
            }

            // ② 名称反品牌排除：如果名字明显是其他品牌，直接 Generic，不再走 FANUC 特征检测
            string name = GetProp(robot, "Name") as string ?? "";
            string otherBrand = LooksLikeOtherBrand(name);
            if (otherBrand != null)
            {
                log.AppendLine("  名称排除: 识别为 " + otherBrand + " (" + name + ")，不走 FANUC 特征");
                return "Generic";
            }

            // ③ FANUC 实例参数特征（收紧：OLP_TOOL_NAME_ 单独不算 FANUC）
            string fanucParam = HasFanucInstanceParam(robot, log);
            if (fanucParam != null) { log.AppendLine("  命中 FANUC 实例参数特征: " + fanucParam); return "FANUC"; }

            // ④ 无控制器：按型号名模糊匹配
            if (LooksLikeFanucName(name)) { log.AppendLine("  名称模糊匹配 FANUC: " + name); return "FANUC"; }

            log.AppendLine("  未识别为 FANUC，按通用(底座面)处理 (name=" + name + ")");
            return "Generic";
        }

        private static bool IsNoController(string ctrl)
        {
            string u = ctrl.Trim().ToUpperInvariant();
            return u == "DEFAULT" || u == "NONE" || u == "UNDEFINED" || u == "N/A" || u == "NULL" || u == "";
        }

        // 名称反品牌排除：如果名字明显是其他品牌（KUKA/ABB/Yaskawa/Denso/Staubli 等），
        // 返回品牌名；否则 null（不排除，继续走 FANUC 检测）
        private static string LooksLikeOtherBrand(string rawName)
        {
            if (string.IsNullOrEmpty(rawName)) return null;
            string n = NormalizeName(rawName);

            // KUKA: kr 前缀（kr16, kr210, kr500, kr150_r2700 等）
            if (n.StartsWith("KR")) return "KUKA";

            // ABB: IRB 前缀（IRB120, IRB6700 等）
            if (n.StartsWith("IRB")) return "ABB";

            // Yaskawa/Motoman: GP/MA/EA/SIA/MPP 前缀 + motoman 关键词
            if (n.StartsWith("GP") || n.StartsWith("MA") || n.StartsWith("SIA")
                || n.StartsWith("MPP") || n.Contains("MOTOMAN") || n.Contains("YASKAWA"))
                return "Yaskawa";

            // Denso: VS/VM 前缀
            if (n.StartsWith("VS") || n.StartsWith("VM")) return "Denso";

            // Staubli: RX/TX/CS 前缀（仅匹配纯型号开头，避免误触 FANUC R-2000i 中的 R/T）
            if (n.StartsWith("RX") || n.StartsWith("TX") || n.StartsWith("CS")) return "Staubli";

            // Kawasaki: RS/FS/BX 前缀 + kawasaki 关键词
            if (n.StartsWith("RS") || n.StartsWith("FS") || n.StartsWith("BX")
                || n.Contains("KAWASAKI")) return "Kawasaki";

            // Hyundai: HH 前缀
            if (n.StartsWith("HH")) return "Hyundai";

            // Nachi: MR/RA 前缀 + nachi 关键词
            if (n.Contains("NACHI")) return "Nachi";

            return null; // 不排除
        }

        // 读取机器人挂载的控制器类型串（多途径兜底）
        private static string TryGetControllerType(TxRobot robot)
        {
            // a) 单参数 getter：robot.GetInstanceParameter("...")
            foreach (var key in new[] { "OLP_CONTROLLER_TYPE", "CONTROLLER_TYPE", "OLP_CONTROLLER_NAME", "OLP_CONTROLLER_VERSION" })
            {
                try
                {
                    object v = InvokeOneArg(robot, "GetInstanceParameter", key);
                    string s = v as string;
                    if (!string.IsNullOrEmpty(s)) return s;
                }
                catch { }
            }
            // b) 直接属性
            foreach (var p in new[] { "ControllerType", "OlpControllerType", "RobotControllerType", "ControllerName" })
            {
                object v = GetProp(robot, p);
                if (v is string s && !string.IsNullOrEmpty(s)) return s;
                if (v != null && !(v is string))
                {
                    string nm = GetProp(v, "Name") as string ?? GetProp(v, "Type") as string;
                    if (!string.IsNullOrEmpty(nm)) return nm;
                }
            }
            // c) Controller 对象 → Name/Type
            object ctrl = GetProp(robot, "Controller") ?? GetProp(robot, "RobotController");
            if (ctrl != null)
            {
                string nm = GetProp(ctrl, "Name") as string
                            ?? GetProp(ctrl, "Type") as string
                            ?? GetProp(ctrl, "ControllerType") as string;
                if (!string.IsNullOrEmpty(nm)) return nm;
            }
            return null;
        }

        // FANUC 实例参数特征检测（收紧版）
        // OLP_TOOL_NAME_x 不是 FANUC 专属（KUKA/ABB 等也会带），仅凭存在不能判 FANUC。
        // 收紧策略：
        //   ① 参数名本身含 "FANUC" → 直接命中
        //   ② OLP_CONTROLLER_VERSION 的值含 "FANUC" / "RJ" → 命中
        //   ③ OLP_CONTROLLER_TYPE 的值含 "FANUC" → 命中
        //   ④ OLP_TOOL_NAME_x 单独存在 → 不命中（太宽，很多品牌都有）
        // 返回命中的特征描述（null = 未命中）
        private static string HasFanucInstanceParam(TxRobot robot, StringBuilder log)
        {
            try
            {
                var ie = InvokeNoArg(robot, "GetAllInstanceParameters") as IEnumerable;
                if (ie == null) return null;

                bool hasOlpToolName = false;
                foreach (object pr in ie)
                {
                    string tn = ParamTypeOrName(pr);
                    if (string.IsNullOrEmpty(tn)) continue;
                    string u = tn.ToUpperInvariant();

                    // ① 参数名含 FANUC → 直接命中
                    if (u.Contains("FANUC")) return "参数名含FANUC: " + tn;

                    // ②③ 检查参数值（OLP_CONTROLLER_VERSION / OLP_CONTROLLER_TYPE）
                    if (u.Contains("OLP_CONTROLLER_VERSION") || u.Contains("OLP_CONTROLLER_TYPE"))
                    {
                        string val = ParamValue(pr);
                        if (!string.IsNullOrEmpty(val))
                        {
                            string uv = val.ToUpperInvariant();
                            if (uv.Contains("FANUC") || uv.Contains("RJ"))
                                return tn + " 值含FANUC/RJ: " + val;
                        }
                    }

                    // 记录是否有 OLP_TOOL_NAME（但单独不判 FANUC）
                    if (u.Contains("OLP_TOOL_NAME")) hasOlpToolName = true;
                }

                // ④ OLP_TOOL_NAME_x 单独存在 → 不命中，但记录到日志便于诊断
                if (hasOlpToolName)
                    log.AppendLine("  [特征] 存在 OLP_TOOL_NAME 参数（单独不能判定 FANUC，需配合控制器版本/类型值）");
            }
            catch { }
            return null;
        }

        // 读取实例参数的值
        private static string ParamValue(object pr)
        {
            try
            {
                dynamic dp = pr;
                // 尝试多种值属性名
                foreach (var vn in new[] { "Value", "StringValue", "DefaultValue", "ParamValue", "Text" })
                {
                    try { object v = GetProp(dp, vn); if (v is string s && !string.IsNullOrEmpty(s)) return s; }
                    catch { }
                }
                // 最后尝试 ToString
                try { return dp.ToString() as string; } catch { }
            }
            catch { }
            return null;
        }

        // FANUC 型号特征库（已规范化：大写、去空格/下划线/连字符；保留判别用的 'I'）
        // 覆盖：ARC Mate / LR Mate / CR / CRX / SR(SCARA) / R-1000/2000 /
        //       M-1iA·2iA·3iA / M-10i·20i·410·710·800·810·900·950·1000·2000 /
        //       P·Paint Mate(喷涂) / F-200i / DR-3i / LR-10
        private static readonly string[] FanucTokens = new[]
        {
            "ARCMATE", "LRMATE", "CRX", "PAINTMATE",
            "R1000I", "R2000I",
            "M1IA", "M2IA", "M3IA",
            "M10I", "M20I", "M410I", "M710I", "M800I", "M810I",
            "M900I", "M950I", "M1000I", "M2000I",
            "CR4I", "CR7I", "CR14I", "CR15I", "CR35I",
            "SR3I", "SR6I", "SR12I", "SR20I",
            "P40I", "P50I", "P250I", "P350I", "F200I", "DR3I", "LR10I"
        };

        private static bool LooksLikeFanucName(string rawName)
        {
            if (string.IsNullOrEmpty(rawName)) return false;
            string n = NormalizeName(rawName);
            if (n.Contains("FANUC")) return true;
            foreach (var tok in FanucTokens)
                if (n.Contains(tok)) return true;
            return false;
        }

        // 规范化：大写 + 去掉 空格/下划线/连字符（"ARC Mate"→"ARCMATE"，"M-710iC"→"M710IC"）
        private static string NormalizeName(string s)
        {
            var sb = new StringBuilder(s.Length);
            foreach (char ch in s.ToUpperInvariant())
                if (ch != ' ' && ch != '_' && ch != '-') sb.Append(ch);
            return sb.ToString();
        }

        // ---------- 期望 BASE0（用于对比，返回 Pose6 数值） ----------
        public static Pose6 ComputeExpectedPose(TxRobot robot, TxTransformation selfTx, Pose6 self,
            string brand, StringBuilder log)
        {
            var exp = new Pose6 { RX = self.RX, RY = self.RY, RZ = self.RZ, RotValid = self.RotValid };

            if (brand == "FANUC" && TryComputeFanucWorldOrigin(robot, log, out V3 o))
            {
                exp.X = o.X; exp.Y = o.Y; exp.Z = o.Z; exp.PosValid = true;
                exp.Source = "Expected(FANUC J1∩J2)";
                return exp;
            }

            // Generic，或 FANUC 计算失败回退：取自身坐标平移
            exp.X = self.X; exp.Y = self.Y; exp.Z = self.Z; exp.PosValid = self.PosValid;
            exp.Source = brand == "FANUC" ? "Expected(FANUC计算失败→Self)" : "Expected(Generic=Self)";
            return exp;
        }

        // ---------- 期望 BASE0（用于同步写入，返回 TxTransformation） ----------
        public static TxTransformation ComputeExpectedTx(TxRobot robot, string brand, StringBuilder log)
        {
            TxTransformation selfTx = SafeTx(() => robot.AbsoluteLocation);
            if (selfTx == null) return null;

            if (brand == "FANUC" && TryComputeFanucWorldOrigin(robot, log, out V3 o))
            {
                var tx = BuildTxWithTranslation(selfTx, o);
                if (tx != null) return tx;
                log.AppendLine("  [同步] FANUC 目标变换构造失败，退回 Self");
            }
            return selfTx; // Generic 或回退
        }

        // ---------- 核心：J1∩J2 ----------
        public static bool TryComputeFanucWorldOrigin(TxRobot robot, StringBuilder log, out V3 origin)
        {
            origin = default;
            var axes = GetJointAxes(robot, log);
            if (axes.Count < 2)
            {
                log.AppendLine("  [FANUC] 可用关节轴不足 2 个，无法计算 J1∩J2");
                return false;
            }

            // 竖直参考：机器人基座 Z 轴（地装非倒挂时≈世界Z）
            V3 baseZ = GetBaseUp(robot);

            // J1 = 轴向最“竖直”的关节；J2 = 列表顺序中 J1 之后第一个“水平”关节
            // （J2/J3 都水平且互相平行，必须按链顺序取 J1 之后的第一个，否则会选成 J3）
            int i1 = -1; double best = -1;
            for (int i = 0; i < axes.Count; i++)
            {
                double v = Math.Abs(axes[i].D.Norm().Dot(baseZ));
                if (v > best) { best = v; i1 = i; }
            }
            int i2 = -1;
            for (int i = i1 + 1; i < axes.Count; i++)
                if (Math.Abs(axes[i].D.Norm().Dot(baseZ)) < 0.5) { i2 = i; break; }
            if (i2 < 0)
                for (int i = 0; i < i1; i++)
                    if (Math.Abs(axes[i].D.Norm().Dot(baseZ)) < 0.5) { i2 = i; break; }

            if (i1 < 0 || i2 < 0)
            {
                log.AppendLine("  [FANUC] 未能识别出 竖直J1 / 水平J2，回退");
                return false;
            }

            var j1 = axes[i1];
            var j2 = axes[i2];
            if (!ClosestPointOnL1ToL2(j1.P, j1.D, j2.P, j2.D, out origin))
            {
                log.AppendLine("  [FANUC] J1/J2 轴平行或退化，无法求交点");
                return false;
            }

            log.AppendLine($"  [FANUC] J1(#{i1 + 1}) P={j1.P} D={j1.D.Norm()} ({j1.How})");
            log.AppendLine($"  [FANUC] J2(#{i2 + 1}) P={j2.P} D={j2.D.Norm()} ({j2.How})");
            log.AppendLine($"  [FANUC] World 原点(J1∩J2) = {origin}");
            return true;
        }

        // 基座“上”方向（取机器人自身坐标 Z 轴；失败用世界 Z）
        private static V3 GetBaseUp(TxRobot robot)
        {
            var selfTx = SafeTx(() => robot.AbsoluteLocation);
            if (selfTx != null && TxZAxis(selfTx, out V3 z) && z.Len > 1e-9) return z.Norm();
            return new V3(0, 0, 1);
        }

        // 直线1(P1+t·D1) 上离 直线2(P2+s·D2) 最近的点
        private static bool ClosestPointOnL1ToL2(V3 P1, V3 D1, V3 P2, V3 D2, out V3 result)
        {
            result = P1;
            D1 = D1.Norm(); D2 = D2.Norm();
            V3 w0 = P1 - P2;
            double a = D1.Dot(D1), b = D1.Dot(D2), c = D2.Dot(D2);
            double d = D1.Dot(w0), e = D2.Dot(w0);
            double denom = a * c - b * b;
            if (Math.Abs(denom) < 1e-9) return false;  // 平行
            double t = (b * e - c * d) / denom;
            result = P1 + D1 * t;
            return true;
        }

        // ---------- 关节轴提取（轴上一点 + 方向） ----------
        private struct Axis { public V3 P; public V3 D; public string How; }

        private static List<Axis> GetJointAxes(TxRobot robot, StringBuilder log)
        {
            var result = new List<Axis>();
            IEnumerable joints = GetJoints(robot, log);
            if (joints == null) return result;

            int idx = 0;
            bool dumped = false;
            foreach (object j in joints)
            {
                idx++;
                if (j == null) continue;
                if (TryJointAxis(j, out V3 p, out V3 dvec, out string how, log, !dumped))
                {
                    result.Add(new Axis { P = p, D = dvec, How = how });
                    log.AppendLine($"  J{idx} 轴: P={p} D={dvec.Norm()} ({how})");
                }
                else
                {
                    log.AppendLine($"  J{idx} 轴: 提取失败");
                    dumped = true; // 只 dump 第一个失败关节，避免刷屏
                }
            }
            return result;
        }

        private static IEnumerable GetJoints(TxRobot robot, StringBuilder log)
        {
            foreach (var m in new[] { "ListJoints", "GetAllJoints", "GetJoints" })
            {
                try
                {
                    var r = InvokeNoArg(robot, m) as IEnumerable;
                    if (r != null) { log.AppendLine("  关节来源(方法): " + m); return r; }
                }
                catch { }
            }
            foreach (var p in new[] { "Joints", "DrivingJoints", "Axes" })
            {
                try
                {
                    var r = GetProp(robot, p) as IEnumerable;
                    if (r != null) { log.AppendLine("  关节来源(属性): " + p); return r; }
                }
                catch { }
            }
            log.AppendLine("  未能获取关节集合（请把下面的类型 dump 反馈）");
            DumpProps(robot, "TxRobot", log);
            return null;
        }

        private static bool TryJointAxis(object joint, out V3 P, out V3 D, out string how,
            StringBuilder log, bool dumpOnFail)
        {
            P = default; D = default; how = "";

            // ① 首选：GetAxisPoints()（世界坐标的 from/to，直接给出轴上点 + 方向）
            if (TryAxisPoints(joint, out P, out D, out how)) return true;

            // ② 退化：Axis 属性给方向 + 连杆原点给点
            foreach (var an in new[] { "Axis", "RotationAxis", "AxisDirection", "Direction" })
            {
                object o = GetProp(joint, an);
                if (ToV3(o, out V3 dv) && dv.Len > 1e-9) { D = dv; how = "Axis(" + an + ")"; break; }
            }

            bool gotP = false;
            foreach (var pn in new[] { "AxisPoint", "Origin", "Point", "Position" })
            {
                object o = GetProp(joint, pn);
                if (ToV3(o, out V3 pv)) { P = pv; gotP = true; how += "+" + pn; break; }
            }
            if (!gotP)
            {
                foreach (var ln in new[] { "ParentLink", "ChildLink" })
                {
                    object link = GetProp(joint, ln);
                    TxTransformation lt = JointWorldTx(link);
                    if (lt != null && TxOrigin(lt, out V3 lp)) { P = lp; gotP = true; how += "+" + ln + ".O"; break; }
                }
            }

            bool ok = gotP && D.Len > 1e-9;
            if (!ok && dumpOnFail)
            {
                DumpProps(joint, "Joint", log);
                object ax = GetProp(joint, "Axis");
                if (ax != null) DumpProps(ax, "Joint.Axis", log);
                try
                {
                    var mi = joint.GetType().GetMethod("GetAxisPoints");
                    log.AppendLine("  GetAxisPoints 方法: " + (mi != null ? "存在" : "不存在"));
                }
                catch { }
            }
            how = how.TrimStart('+');
            return ok;
        }

        // GetAxisPoints()：返回世界坐标的轴 from/to，兼容多种返回形态
        private static bool TryAxisPoints(object joint, out V3 P, out V3 D, out string how)
        {
            P = default; D = default; how = "";
            if (joint == null) return false;
            try
            {
                var mi = joint.GetType().GetMethod("GetAxisPoints");
                if (mi == null) return false;
                var ps = mi.GetParameters();

                V3 from, to;
                if (ps.Length == 0)
                {
                    object r = mi.Invoke(joint, null);
                    if (!TryTwoVectors(r, out from, out to)) return false;
                }
                else if (ps.Length == 2 && ps[0].IsOut && ps[1].IsOut)
                {
                    var args = new object[] { null, null };
                    mi.Invoke(joint, args);
                    if (!ToV3(args[0], out from) || !ToV3(args[1], out to)) return false;
                }
                else return false;

                P = from;
                D = to - from;
                how = "GetAxisPoints";
                return D.Len > 1e-9;
            }
            catch { return false; }
        }

        // 从“两个点”的各种返回形态里取出 a/b
        private static bool TryTwoVectors(object r, out V3 a, out V3 b)
        {
            a = default; b = default;
            if (r == null) return false;

            if (r is IEnumerable en && !(r is string))
            {
                var list = new List<object>();
                foreach (var o in en) list.Add(o);
                if (list.Count >= 2) return ToV3(list[0], out a) && ToV3(list[1], out b);
            }
            foreach (var pair in new[]
            {
                new[]{"Item1","Item2"}, new[]{"From","To"}, new[]{"First","Second"},
                new[]{"Point1","Point2"}, new[]{"StartPoint","EndPoint"}, new[]{"Start","End"}
            })
            {
                object o1 = GetProp(r, pair[0]);
                object o2 = GetProp(r, pair[1]);
                if (o1 != null && o2 != null && ToV3(o1, out a) && ToV3(o2, out b)) return true;
            }
            return false;
        }

        private static TxTransformation JointWorldTx(object joint)
        {
            foreach (var m in new[] { "AbsoluteLocation", "LocationRelativeToWorld", "Location", "Transformation" })
            {
                var t = AsTx(InvokeOrGetSafe(joint, m));
                if (t != null) return t;
            }
            // DrivingJoint 包装：内部 .Joint / .Frame
            foreach (var inner in new[] { "Joint", "Frame" })
            {
                object io = GetProp(joint, inner);
                if (io == null) continue;
                foreach (var m in new[] { "AbsoluteLocation", "Location", "Transformation" })
                {
                    var t = AsTx(InvokeOrGetSafe(io, m));
                    if (t != null) return t;
                }
            }
            return null;
        }

        // ---------- TxTransformation 工具 ----------
        private static bool TxOrigin(TxTransformation t, out V3 p)
        {
            p = default;
            try { dynamic d = t; var tr = d.Translation; p = new V3((double)tr.X, (double)tr.Y, (double)tr.Z); return true; }
            catch { return false; }
        }

        private static bool TxZAxis(TxTransformation t, out V3 z)
        {
            z = default;
            // 1) 直接属性
            foreach (var name in new[] { "ZAxis", "ZDirection", "AxisZ" })
            {
                object o = GetProp(t, name);
                if (ToV3(o, out z) && z.Len > 1e-9) return true;
            }
            // 2) 变换法：Z 方向 = T·(0,0,1) - T·(0,0,0)
            //    通过把局部点经变换得到世界点，二者之差即方向（消去平移）。
            try
            {
                if (TryTransformPoint(t, new V3(0, 0, 0), out V3 o0) &&
                    TryTransformPoint(t, new V3(0, 0, 1), out V3 o1))
                {
                    V3 d = o1 - o0;
                    if (d.Len > 1e-9) { z = d.Norm(); return true; }
                }
            }
            catch { }
            return false;
        }

        // 把局部点经变换映射到世界（尝试多种 API）
        private static bool TryTransformPoint(TxTransformation t, V3 local, out V3 world)
        {
            world = default;
            try
            {
                dynamic dt = t;
                var v = new TxVector(local.X, local.Y, local.Z);
                // 常见命名尝试
                object r = null;
                foreach (var m in new[] { "TransformPoint", "Transform", "Apply", "Multiply" })
                {
                    try { r = InvokeOneArg(t, m, v); if (r != null) break; } catch { }
                }
                if (r == null) return false;
                return ToV3(r, out world);
            }
            catch { return false; }
        }

        // 克隆 self 旋转，替换平移为 pos
        public static TxTransformation BuildTxWithTranslation(TxTransformation baseRot, V3 pos)
        {
            if (baseRot == null) return null;
            var posv = new TxVector(pos.X, pos.Y, pos.Z);

            // 1) 反射调用 Clone() 方法（若存在）+ 设置 Translation
            try
            {
                var mi = baseRot.GetType().GetMethod("Clone", Type.EmptyTypes);
                if (mi != null)
                {
                    var clone = mi.Invoke(baseRot, null) as TxTransformation;
                    if (clone != null && TrySetTranslation(clone, posv)) return clone;
                }
            }
            catch { }

            // 2) 拷贝构造 + 设置 Translation
            try
            {
                var t = (TxTransformation)Activator.CreateInstance(typeof(TxTransformation), baseRot);
                if (TrySetTranslation(t, posv)) return t;
            }
            catch { }

            return null;
        }

        private static bool TrySetTranslation(TxTransformation t, TxVector posv)
        {
            try
            {
                var pi = t.GetType().GetProperty("Translation");
                if (pi != null && pi.CanWrite) { pi.SetValue(t, posv, null); return true; }
            }
            catch { }
            return false;
        }

        // ---------- 反射/类型小工具 ----------
        private static bool ToV3(object txv, out V3 v)
        {
            v = default;
            if (txv == null) return false;
            try { dynamic d = txv; v = new V3((double)d.X, (double)d.Y, (double)d.Z); return true; }
            catch { return false; }
        }

        private static TxTransformation AsTx(object o) => o as TxTransformation;

        private static TxTransformation SafeTx(Func<TxTransformation> f)
        {
            try { return f(); } catch { return null; }
        }

        private static object GetProp(object obj, string name)
        {
            if (obj == null) return null;
            try { var pi = obj.GetType().GetProperty(name); return pi != null ? pi.GetValue(obj, null) : null; }
            catch { return null; }
        }

        private static object InvokeNoArg(object obj, string name)
        {
            if (obj == null) return null;
            var mi = obj.GetType().GetMethod(name, Type.EmptyTypes);
            return mi != null ? mi.Invoke(obj, null) : null;
        }

        private static object InvokeOneArg(object obj, string name, object arg)
        {
            if (obj == null) return null;
            var mi = obj.GetType().GetMethod(name, new[] { arg.GetType() });
            return mi != null ? mi.Invoke(obj, new[] { arg }) : null;
        }

        private static object InvokeOrGetSafe(object obj, string name)
        {
            if (obj == null) return null;
            try
            {
                var t = obj.GetType();
                var pi = t.GetProperty(name);
                if (pi != null) return pi.GetValue(obj, null);
                var mi = t.GetMethod(name, Type.EmptyTypes);
                if (mi != null) return mi.Invoke(obj, null);
            }
            catch { }
            return null;
        }

        private static string ParamTypeOrName(object pr)
        {
            try { dynamic dp = pr; return (dp.Type as string) ?? (dp.Name as string); }
            catch { return null; }
        }

        private static void DumpProps(object o, string tag, StringBuilder log)
        {
            if (o == null) return;
            try
            {
                var t = o.GetType();
                var names = new List<string>();
                foreach (var pi in t.GetProperties()) names.Add(pi.Name);
                log.AppendLine($"  [类型dump] {tag} = {t.FullName}");
                log.AppendLine("  [属性dump] " + string.Join(", ", names));
            }
            catch { }
        }
    }
}
