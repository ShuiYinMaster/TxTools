using System;
using Tecnomatix.Engineering;

namespace TxTools.WeldGunDefiner.Math
{
    /// <summary>三维向量工具</summary>
    public struct Vec3
    {
        public double X, Y, Z;
        public Vec3(double x, double y, double z) { X = x; Y = y; Z = z; }
        public Vec3(TxVector v) { X = v.X; Y = v.Y; Z = v.Z; }
        public static Vec3 operator +(Vec3 a, Vec3 b) => new Vec3(a.X+b.X, a.Y+b.Y, a.Z+b.Z);
        public static Vec3 operator -(Vec3 a, Vec3 b) => new Vec3(a.X-b.X, a.Y-b.Y, a.Z-b.Z);
        public static Vec3 operator *(Vec3 v, double s) => new Vec3(v.X*s, v.Y*s, v.Z*s);
        public static Vec3 operator *(double s, Vec3 v) => v * s;
        public double Length => System.Math.Sqrt(X*X + Y*Y + Z*Z);
        public Vec3 Normalized() { double l = Length; return l < 1e-12 ? this : new Vec3(X/l, Y/l, Z/l); }
        public static double Dot(Vec3 a, Vec3 b) => a.X*b.X + a.Y*b.Y + a.Z*b.Z;
        public static Vec3 Cross(Vec3 a, Vec3 b) => new Vec3(
            a.Y*b.Z - a.Z*b.Y, a.Z*b.X - a.X*b.Z, a.X*b.Y - a.Y*b.X);
        public TxVector ToTxVector() => new TxVector(X, Y, Z);
        public override string ToString() => $"({X:F3}, {Y:F3}, {Z:F3})";
    }

    /// <summary>二维向量（平面内计算用）</summary>
    public struct Vec2
    {
        public double X, Y;
        public Vec2(double x, double y) { X = x; Y = y; }
        public static Vec2 operator -(Vec2 a, Vec2 b) => new Vec2(a.X-b.X, a.Y-b.Y);
        public static Vec2 operator +(Vec2 a, Vec2 b) => new Vec2(a.X+b.X, a.Y+b.Y);
        public static Vec2 operator *(Vec2 v, double s) => new Vec2(v.X*s, v.Y*s);
        public double Length => System.Math.Sqrt(X*X + Y*Y);
        public override string ToString() => $"({X:F3}, {Y:F3})";
    }

    /// <summary>
    /// 参考平面：由法向量 N 和原点 Origin 定义
    /// </summary>
    public class ReferencePlane
    {
        public Vec3 Normal  { get; }
        public Vec3 Origin  { get; }
        public Vec3 AxisX   { get; }
        public Vec3 AxisY   { get; }

        public ReferencePlane(Vec3 normal, Vec3 origin)
        {
            Normal = normal.Normalized();
            Origin = origin;
            Vec3 up = System.Math.Abs(Vec3.Dot(Normal, new Vec3(0,0,1))) < 0.9
                ? new Vec3(0,0,1) : new Vec3(1,0,0);
            AxisX = Vec3.Cross(up, Normal).Normalized();
            AxisY = Vec3.Cross(Normal, AxisX).Normalized();
        }

        public Vec3 Project(Vec3 p)
        {
            double dist = Vec3.Dot(p - Origin, Normal);
            return p - Normal * dist;
        }

        public double DistanceTo(Vec3 p)
            => System.Math.Abs(Vec3.Dot(p - Origin, Normal));

        public Vec2 To2D(Vec3 p)
        {
            Vec3 proj = Project(p);
            Vec3 d = proj - Origin;
            return new Vec2(Vec3.Dot(d, AxisX), Vec3.Dot(d, AxisY));
        }
    }

    /// <summary>
    /// X型焊枪（RPRR曲柄滑块）机构参数
    ///
    /// 对应PS RPRR向导的三点：
    ///   A = Fixed-Input Joint   （活塞铰点，固定在枪体）
    ///   C = Coupler-Output Joint（连杆铰点，在动臂上）
    ///   O = Output Joint        （动臂旋转轴，固定在枪体）
    ///
    /// PS 公式语法（已从手册确认）：
    ///   D(关节名)  → 取关节当前值
    ///   (((D(j1))>=0)*((D(j1))/2))      → 条件乘法（等同 if j1>=0 then j1/2 else 0）
    ///   ((D(j1))*stroke/openGap)         → 线性比例
    /// </summary>
    public class GunMechanismParams
    {
        // ── 三铰点（平面内2D坐标，投影后）──
        public Vec2 O  { get; set; }   // Output Joint（动臂旋转轴）
        public Vec2 A  { get; set; }   // Fixed-Input Joint（活塞铰点）
        public Vec2 C0 { get; set; }   // Coupler-Output Joint（连杆铰点，Close状态）

        // ── 自动推算的几何参数 ──
        public double d_AO { get; private set; }   // A到O距离 (mm)
        public double r    { get; private set; }   // O到C距离（动臂连杆铰点到转轴）(mm)
        public double L    { get; private set; }   // 连杆长度 A→C (mm)（RPRR中活塞直接到C）
        public double Alpha { get; private set; }  // OA连线与活塞方向夹角 (rad)

        // ── 电极臂参数 ──
        public double R_arm    { get; set; }       // 动臂长度：O → 动TCP (mm)
        public double d_static { get; set; }       // O → 静TCP 距离 (mm)
        public double Theta_static { get; set; }   // 静TCP相对O的角度 (rad，平面内)
        public double Theta0_arm   { get; set; }   // 动臂Close状态角度 (rad)

        // ── 活塞方向 ──
        public Vec2 PistonDir { get; set; }        // 活塞伸出方向（平面内单位向量）

        // ── 用户输入参数 ──
        public double S0           { get; set; } = 0.0;    // Close状态活塞伸出量 (mm)
        public double PistonStroke { get; set; } = 100.0;  // 活塞杆行程 (mm)，用户手填
        public double OpenGap      { get; set; } = 185.0;  // 焊钳开口量 (mm)，用户手填
        public double WearAllowance { get; set; } = 10.0;  // 磨损补偿量 (mm)，J1 High limit

        /// <summary>
        /// 从三个投影点推算几何参数
        /// 注意：RPRR中活塞为Prismatic，A是固定铰点，C是连杆铰点
        /// </summary>
        public void ComputeFromPoints(Vec2 o, Vec2 a, Vec2 c0, Vec2 pistonDir, double s0)
        {
            O  = o;
            A  = a;
            C0 = c0;
            PistonDir = pistonDir;
            S0 = s0;

            d_AO = (A - O).Length;
            r    = (C0 - O).Length;

            // RPRR: 活塞（Prismatic）从A出发沿pistonDir运动，末端直接就是C
            // 所以L = 当前A到C0的距离（这是活塞杆在Close状态的伸出量对应的连杆长）
            // 但在RPRR中，Prismatic joint 的 offset = 活塞伸出量 s
            // C 的位置 = A + s * pistonDir（活塞杆末端直接铰接动臂）
            // 所以 L 实际是 s0（Close时伸出量），连杆长度即活塞当前长度，是变量
            // 为推算方便，记录初始状态下 A→C0 向量
            Vec2 AC0 = new Vec2(C0.X - A.X, C0.Y - A.Y);
            L = AC0.Length;  // Close状态下的活塞伸出长度

            // 结构角：OA 与 pistonDir 的夹角
            Vec2 OA = new Vec2(A.X - O.X, A.Y - O.Y);
            double dot = (OA.X * pistonDir.X + OA.Y * pistonDir.Y)
                       / (OA.Length * pistonDir.Length + 1e-12);
            dot = System.Math.Max(-1.0, System.Math.Min(1.0, dot));
            Alpha = System.Math.Acos(dot);

            // Close状态下动臂角度（O→C0方向）
            Theta0_arm = System.Math.Atan2(C0.Y - O.Y, C0.X - O.X);
        }

        /// <summary>
        /// 生成 PS Joint Dependency 公式（对齐手册格式）
        ///
        /// 手册确认的公式语义：
        ///   j1 = Prismatic主动关节（活塞伸出量，负值=开口，正值=磨量）
        ///        Low limit = -OpenGap, High limit = +WearAllowance
        ///   j2 = 磨量补偿关节：(((D(j1))>=0)*((D(j1))/2))
        ///        含义：j1>=0时 j2=j1/2（双侧补偿），j1<0时 j2=0
        ///   input_j1（动臂旋转关节）= ((D(j1))*PistonStroke/OpenGap)
        ///        含义：j1线性映射到活塞实际行程
        /// </summary>
        public GunFormulas BuildFormulas(string j1Name, string j2Name, string inputJ1Name)
        {
            // J2：磨量补偿（双侧电极对称补偿，各补偿一半）
            // (((D(j1))>=0)*((D(j1))/2))
            string j2Formula = $"(((D({j1Name}))>=0)*((D({j1Name}))/2))";

            // input_j1：活塞行程线性换算
            // ((D(j1))*PistonStroke/OpenGap)
            string inputJ1Formula = $"((D({j1Name}))*{PistonStroke:F4}/{OpenGap:F4})";

            // J1 的 Limit：
            //   High = +WearAllowance（磨损补偿，正值，枪闭合方向）
            //   Low  = -OpenGap      （开口量，负值，枪张开方向）
            double j1High = WearAllowance;
            double j1Low  = -OpenGap;

            return new GunFormulas
            {
                J1_Name       = j1Name,
                J2_Name       = j2Name,
                InputJ1_Name  = inputJ1Name,
                J1_LimitHigh  = j1High,
                J1_LimitLow   = j1Low,
                J2_Formula    = j2Formula,
                InputJ1_Formula = inputJ1Formula,
                PistonStroke  = PistonStroke,
                OpenGap       = OpenGap,
                WearAllowance = WearAllowance,
            };
        }

        /// <summary>
        /// 三铰点的世界坐标（供界面显示给用户，让用户填写RPRR向导）
        /// </summary>
        public string FormatHingePointsForWizard(Vec3 worldO, Vec3 worldA, Vec3 worldC)
        {
            return
                $"RPRR 向导输入坐标（请按顺序填写）:\r\n" +
                $"\r\n" +
                $"  点1 Fixed-Input Joint (A):\r\n" +
                $"    X={worldA.X:F2}  Y={worldA.Y:F2}  Z={worldA.Z:F2}\r\n" +
                $"\r\n" +
                $"  点2 Coupler-Output Joint (C):\r\n" +
                $"    X={worldC.X:F2}  Y={worldC.Y:F2}  Z={worldC.Z:F2}\r\n" +
                $"\r\n" +
                $"  点3 Output Joint (O):\r\n" +
                $"    X={worldO.X:F2}  Y={worldO.Y:F2}  Z={worldO.Z:F2}\r\n";
        }
    }

    /// <summary>生成的PS公式集合</summary>
    public class GunFormulas
    {
        // 关节名称
        public string J1_Name       { get; set; }  // 主动Prismatic关节（活塞）
        public string J2_Name       { get; set; }  // 磨量补偿关节
        public string InputJ1_Name  { get; set; }  // 动臂旋转关节（RPRR生成的input_j1）

        // J1 Limit
        public double J1_LimitHigh  { get; set; }  // +磨量 (mm)
        public double J1_LimitLow   { get; set; }  // -开口量 (mm)

        // 公式字符串
        public string J2_Formula       { get; set; }  // J2 的 Joint Dependency 公式
        public string InputJ1_Formula  { get; set; }  // input_j1 的 Joint Dependency 公式

        // 原始参数
        public double PistonStroke  { get; set; }
        public double OpenGap       { get; set; }
        public double WearAllowance { get; set; }

        /// <summary>格式化为人类可读的完整参数卡</summary>
        public string ToParamCard()
        {
            return
                $"== J1 ({J1_Name}) Limit ==\r\n" +
                $"  High (磨量): +{J1_LimitHigh:F2} mm\r\n" +
                $"  Low  (开口): {J1_LimitLow:F2} mm\r\n" +
                $"\r\n" +
                $"== J2 ({J2_Name}) Joint Dependency ==\r\n" +
                $"  {J2_Formula}\r\n" +
                $"  含义: J1>=0 时 J2=J1/2 (双侧电极对称补偿)\r\n" +
                $"\r\n" +
                $"== {InputJ1_Name} Joint Dependency ==\r\n" +
                $"  {InputJ1_Formula}\r\n" +
                $"  含义: 活塞行程 {PistonStroke:F0}mm / 开口 {OpenGap:F0}mm 线性换算\r\n";
        }
    }
}
