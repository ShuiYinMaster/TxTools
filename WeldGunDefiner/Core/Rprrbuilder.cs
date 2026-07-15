using System;
using System.Collections.Generic;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.DataTypes;
using TxTools.WeldGunDefiner.Math;

namespace TxTools.WeldGunDefiner.Core
{
    /// <summary>
    /// 完整复刻 PS RPRR 向导 + 旁支驱动结构。
    ///
    /// 真实结构（从工程数据 FFM130-X-0646 确认）：
    ///
    /// 主链（曲柄滑块，活塞行程→动臂转角）：
    ///   fixed_link  --fixed_Input_j1[Rev,公式]-->  input_link
    ///   input_link  --input_j1[Pri,公式]-->        coupler_link
    ///   coupler_link--coup_output_j1[Rev,公式]-->  output_link
    ///   output_link --output_j1[Rev,公式]-->       dummy_link
    ///
    /// 旁支（焊钳开口→活塞行程，用户输入轴）：
    ///   fixed_link  --j1[Pri,主动,Low=-开口,High=+磨量]--> lnk1（空）
    ///   fixed_link  --j2[Pri,公式磨量补偿]-->              lnk2（空）
    ///
    /// 驱动链：
    ///   用户拖 j1（焊钳开口量）
    ///     → input_j1 = (D(j1)*行程/开口)         开口换算活塞行程
    ///     → fixed_Input_j1 = acos(余弦定理,D(input_j1)) 活塞→连杆角
    ///     → coup_output_j1 = atan2(...T(fixed_Input_j1)...)
    ///     → output_j1 = (-1)(T(coup_output_j1)+T(fixed_Input_j1))
    ///
    /// 公式中 D()=驱动值，T()=理论总值。
    /// </summary>
    public class RprrBuilder
    {
        private readonly ITxKinematicsModellable _device;
        private readonly RprrParams _p;
        private TxTransformation _invDeviceTx;
        private bool _enumDiagnosed;

        // 创建后的关节引用
        public TxJoint J1 { get; private set; }  // 主动：焊钳开口
        public TxJoint J2 { get; private set; }  // 磨量补偿
        public TxJoint InputJl { get; private set; }  // 活塞 Prismatic
        public TxJoint FixedInputJl { get; private set; }  // 连杆角 Revolute
        public TxJoint CoupOutputJl { get; private set; }
        public TxJoint OutputJl { get; private set; }

        public string Log { get; private set; } = "";

        public RprrBuilder(ITxKinematicsModellable device, RprrParams p)
        {
            _device = device;
            _p = p;
            try
            {
                var loc = (device as ITxLocatableObject)?.AbsoluteLocation;
                if (loc != null) _invDeviceTx = loc.Inverse;
            }
            catch { _invDeviceTx = null; }
        }

        // ═════════════════════════════════════════════════════════════════
        // 主流程
        // ═════════════════════════════════════════════════════════════════
        public bool Build(out string error)
        {
            error = null;
            try
            {
                AddLog($"IsOpenForKinematicsModeling = {_device.IsOpenForKinematicsModeling}");
                AddLog($"设备逆矩阵: {(_invDeviceTx != null ? "已计算" : "null")}");

                // 先进入建模状态，再清理残留——否则非建模状态下
                // _device.Joints 可能返回不完整，残留Joint清不掉
                if (!_device.IsOpenForKinematicsModeling)
                {
                    try { (_device as ITxComponent)?.SetModelingScope(); } catch { }
                    AddLog($"SetModelingScope后 IsOpen = {_device.IsOpenForKinematicsModeling}");
                }

                CleanupPrevious();

                // ── 创建 7 个 Link ──
                var fixedLink = CreateLink("fixed_link", _p.FixedLinkBodies);
                var inputLink = CreateLink("input_link", _p.InputLinkBodies);
                var couplerLink = CreateLink("coupler_link", _p.CouplerLinkBodies);
                var outputLink = CreateLink("output_link", _p.OutputLinkBodies);
                var dummyLink = CreateLink("dummy_link", new List<ITxObject>());
                var lnk1 = CreateLink("lnk1", new List<ITxObject>());
                var lnk2 = CreateLink("lnk2", _p.Lnk2Bodies ?? new List<ITxObject>());
                if (fixedLink == null || inputLink == null || couplerLink == null ||
                    outputLink == null || dummyLink == null || lnk1 == null || lnk2 == null)
                { error = "创建 Link 失败"; return false; }

                // ── 计算公式 + 常数（在创建Joint前算好）──
                var f = BuildFormulas();

                // ── 三铰点投影到参考平面（统一轴向位置，保证共面）──
                // 选点时三点可能不严格共面，投影后机构面与焊钳侧面平行
                Vec3 projA = _p.Plane != null ? _p.Plane.Project(_p.WorldA) : _p.WorldA;
                Vec3 projC = _p.Plane != null ? _p.Plane.Project(_p.WorldC) : _p.WorldC;
                Vec3 projO = _p.Plane != null ? _p.Plane.Project(_p.WorldO) : _p.WorldO;

                // 活塞方向投影到平面内（去除轴向分量）
                Vec3 pistonDirProj = _p.PistonDir;
                if (_p.Plane != null)
                {
                    Vec3 pEnd = _p.Plane.Project(_p.WorldA + _p.PistonDir);
                    pistonDirProj = (pEnd - projA).Normalized();
                }

                AddLog($"铰点投影: A偏移={(projA - _p.WorldA).Length:F3}mm C偏移={(projC - _p.WorldC).Length:F3}mm O偏移={(projO - _p.WorldO).Length:F3}mm");
                AddLog($"平面法向 N = ({_p.N.X:F4}, {_p.N.Y:F4}, {_p.N.Z:F4})  [所有Revolute轴方向]");
                AddLog($"活塞方向(投影) = ({pistonDirProj.X:F4}, {pistonDirProj.Y:F4}, {pistonDirProj.Z:F4})  [input_j1轴]");

                // ── 主链 4 个 Joint（用投影后的共面点，世界坐标）──
                // fixed_Input_j1：Revolute，fixed_link→input_link，轴过A
                // 注意：fixed_Input_j1 和 coup_output_j1 轴方向取 -N（与output_j1相反）
                // 使这两个公式驱动轴的转动正方向与几何期望一致
                Vec3 negN = new Vec3(-_p.N.X, -_p.N.Y, -_p.N.Z);
                FixedInputJl = CreateJointInternal("fixed_Input_j1", "Revolute",
                    projA.ToTxVector(), (projA + negN).ToTxVector(),
                    0, 0, fixedLink, inputLink);
                if (FixedInputJl == null) { error = "创建 fixed_Input_j1 失败"; return false; }

                // input_j1：Prismatic，input_link→coupler_link，活塞方向（投影后）
                // OPEN适配模式：初始限位放宽到±行程，允许阶段B正向驱动回闭合
                double inJ1Low = _p.OpenStateAdapt ? -_p.PistonStroke : f.InputJl_Low;
                double inJ1High = _p.OpenStateAdapt ? _p.PistonStroke : f.InputJl_High;
                InputJl = CreateJointInternal("input_j1", "Prismatic",
                    projA.ToTxVector(), (projA + pistonDirProj).ToTxVector(),
                    inJ1Low, inJ1High, inputLink, couplerLink);
                if (InputJl == null) { error = "创建 input_j1 失败"; return false; }

                // coup_output_j1：Revolute，coupler_link→output_link，轴过C，轴向-N
                CoupOutputJl = CreateJointInternal("coup_output_j1", "Revolute",
                    projC.ToTxVector(), (projC + negN).ToTxVector(),
                    0, 0, couplerLink, outputLink);
                if (CoupOutputJl == null) { error = "创建 coup_output_j1 失败"; return false; }

                // output_j1：Revolute，output_link→dummy_link，轴过O，轴向+N
                OutputJl = CreateJointInternal("output_j1", "Revolute",
                    projO.ToTxVector(), (projO + _p.N).ToTxVector(),
                    0, 0, outputLink, dummyLink);
                if (OutputJl == null) { error = "创建 output_j1 失败"; return false; }

                // ── 旁支 2 个 Joint ──
                // j1：Prismatic 主动，fixed_link→lnk1，Low=-开口 High=+磨量
                // 轴方向 = 静电极→动电极
                Vec3 gunAxis = (_p.WorldMovingTip - _p.WorldStaticTip).Normalized();
                if (gunAxis.Length < 0.5) gunAxis = _p.PistonDir;
                J1 = CreateJointInternal("j1", "Prismatic",
                    _p.WorldStaticTip.ToTxVector(),
                    (_p.WorldStaticTip + gunAxis).ToTxVector(),
                    f.J1_Low, f.J1_High, fixedLink, lnk1);
                if (J1 == null) { error = "创建 j1 失败"; return false; }

                // j2:Prismatic 磨量补偿，fixed_link→lnk2
                J2 = CreateJointInternal("j2", "Prismatic",
                    _p.WorldStaticTip.ToTxVector(),
                    (_p.WorldStaticTip + gunAxis).ToTxVector(),
                    0, 0, fixedLink, lnk2);
                if (J2 == null) { error = "创建 j2 失败"; return false; }

                // ── 验证 PS 实际分配的关节名（CreationData.Name 应已生效）──
                string nFixedInput = SafeName(FixedInputJl);
                string nInput = SafeName(InputJl);
                string nCoupOut = SafeName(CoupOutputJl);
                string nOutput = SafeName(OutputJl);
                string nJ1 = SafeName(J1);
                string nJ2 = SafeName(J2);

                AddLog("");
                AddLog("=== 关节名验证 ===");
                AddLog($"期望 fixed_Input_j1 → 实际 {nFixedInput}");
                AddLog($"期望 input_j1       → 实际 {nInput}");
                AddLog($"期望 coup_output_j1 → 实际 {nCoupOut}");
                AddLog($"期望 output_j1      → 实际 {nOutput}");
                AddLog($"期望 j1             → 实际 {nJ1}");
                AddLog($"期望 j2             → 实际 {nJ2}");

                // ── 用实际名生成公式（即使PS改了名也能对应上）──
                var f2 = BuildFormulasWithNames(nFixedInput, nInput, nCoupOut, nOutput, nJ1, nJ2);

                if (_p.OpenStateAdapt)
                {
                    // ═══ 需求4：OPEN状态适配，分阶段写入 ═══
                    // OPEN几何读的|AC|是张开活塞长(228)。
                    //   阶段A驱动公式用 s0=|AC|(228)：让input_j1=0对应当前张开几何，
                    //     正向+行程才能正确驱动到闭合(活塞330.5)。
                    //   阶段E最终公式用 s0=|AC|+行程(330.5)：置零后几何已闭合，
                    //     此值才是正确的闭合活塞长，运动学不偏移。
                    AddLog("");
                    AddLog("=== OPEN状态适配模式：分阶段写入 ===");

                    Vec2 _O2 = _p.Plane.To2D(_p.WorldO);
                    Vec2 _A2 = _p.Plane.To2D(_p.WorldA);
                    Vec2 _C2 = _p.Plane.To2D(_p.WorldC);
                    double acOpen = (_C2 - _A2).Length;        // OPEN读取的|AC|(张开活塞长)
                    double s0Drive = acOpen;                   // 阶段A驱动用
                    double s0Final = acOpen + _p.PistonStroke; // 阶段E最终用(闭合活塞长)
                    AddLog($"[i] OPEN|AC|={acOpen:F2} → 驱动s0={s0Drive:F2}, 最终s0={s0Final:F2}");

                    // 阶段A：用驱动s0(原始|AC|)写从动公式，让几何能正确驱动回闭合
                    var fDrive = BuildFormulasWithNames(nFixedInput, nInput, nCoupOut, nOutput, nJ1, nJ2, s0Drive);
                    WriteFormula(FixedInputJl, fDrive.FixedInputJl_Formula, nFixedInput);
                    WriteFormula(CoupOutputJl, fDrive.CoupOutputJl_Formula, nCoupOut);
                    WriteFormula(OutputJl, fDrive.OutputJl_Formula, nOutput);
                    AddLog("[A] 主链从动公式已写(驱动s0,input_j1暂独立)");

                    // input_j1 临时限位放宽，允许正向驱动回闭合
                    try
                    {
                        var hl = new TxJointConstantHardLimits(-_p.PistonStroke, _p.PistonStroke);
                        InputJl.HardLimits = hl;
                        InputJl.LowerSoftLimit = -_p.PistonStroke;
                        InputJl.UpperSoftLimit = _p.PistonStroke;
                    }
                    catch (Exception ex) { AddLog($"[A] input_j1临时限位失败:{ex.Message}"); }

                    // 阶段B：正方向驱动 input_j1 = +行程 → 几何回到闭合(CLOSE)
                    try
                    {
                        InputJl.CurrentValue = _p.PistonStroke;
                        AddLog($"[B] 已驱动 input_j1 = +{_p.PistonStroke:F2}(几何回闭合)");
                    }
                    catch (Exception ex) { AddLog($"[B] 驱动input_j1失败:{ex.Message}"); }

                    // 阶段C：DefineZeroPosition——当前闭合位姿定为零位
                    try
                    {
                        bool can = true;
                        try { can = _device.CanDefineZeroPosition(); } catch { can = true; }
                        if (can) { _device.DefineZeroPosition(); AddLog("[C] DefineZeroPosition 成功(闭合=0位)"); }
                        else AddLog("[C] CanDefineZeroPosition=false，跳过");
                    }
                    catch (Exception ex) { AddLog($"[C] DefineZeroPosition失败:{ex.Message}"); }

                    // 阶段D：重写 input_j1 限位为 [-行程, 0]（闭合=0，张开=-行程）
                    try
                    {
                        var hl = new TxJointConstantHardLimits(-_p.PistonStroke, 0);
                        InputJl.HardLimits = hl;
                        InputJl.LowerSoftLimit = -_p.PistonStroke;
                        InputJl.UpperSoftLimit = 0;
                        AddLog($"[D] input_j1限位重写为 [{-_p.PistonStroke:F2}, 0]");
                    }
                    catch (Exception ex) { AddLog($"[D] 重写input_j1限位失败:{ex.Message}"); }

                    // 阶段E：置零后用【最终s0=|AC|+行程】重刷全部公式
                    // 此时几何已闭合、零点已重置，s0=闭合活塞长，运动学不偏移。
                    var fFinal = BuildFormulasWithNames(nFixedInput, nInput, nCoupOut, nOutput, nJ1, nJ2, s0Final);
                    WriteFormula(FixedInputJl, fFinal.FixedInputJl_Formula, nFixedInput);
                    WriteFormula(CoupOutputJl, fFinal.CoupOutputJl_Formula, nCoupOut);
                    WriteFormula(OutputJl, fFinal.OutputJl_Formula, nOutput);
                    WriteFormula(InputJl, fFinal.InputJl_Formula, nInput);
                    WriteFormula(J2, fFinal.J2_Formula, nJ2);
                    AddLog("[E] 置零后用最终s0重刷全部公式 + 接j1驱动链");
                }
                else
                {
                    // ═══ 默认：一次性写入全部公式 ═══
                    WriteFormula(InputJl, f2.InputJl_Formula, nInput);
                    WriteFormula(FixedInputJl, f2.FixedInputJl_Formula, nFixedInput);
                    WriteFormula(CoupOutputJl, f2.CoupOutputJl_Formula, nCoupOut);
                    WriteFormula(OutputJl, f2.OutputJl_Formula, nOutput);
                    WriteFormula(J2, f2.J2_Formula, nJ2);
                }

                AddLog("");
                AddLog("=== 公式汇总（已用真实关节名）===");
                AddLog($"{nJ1} (主动): Low={f.J1_Low:F2} High={f.J1_High:F2}");
                AddLog($"{nInput}     = {f2.InputJl_Formula}");
                AddLog($"{nFixedInput} = {f2.FixedInputJl_Formula}");
                AddLog($"{nCoupOut} = {f2.CoupOutputJl_Formula}");
                AddLog($"{nOutput}    = {f2.OutputJl_Formula}");
                AddLog($"{nJ2}           = {f2.J2_Formula}");

                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                if (ex.InnerException != null) error += " | " + ex.InnerException.Message;
                return false;
            }
        }

        // ═════════════════════════════════════════════════════════════════
        // 公式计算（核心）—— 复刻真实公式形式，常数从几何算出
        // ═════════════════════════════════════════════════════════════════
        // 无参版本：用默认名（仅供 UI 预览）
        public RprrFormulas BuildFormulas()
            => BuildFormulasWithNames("fixed_Input_j1", "input_j1", "coup_output_j1", "output_j1", "j1", "j2");

        // 带真实关节名版本：公式引用与实际 Joint 名一致
        // s0Override: 显式指定s0(活塞闭合长)；null则按OpenStateAdapt自动算
        public RprrFormulas BuildFormulasWithNames(
            string nFixedInput, string nInput, string nCoupOut, string nOutput, string nJ1, string nJ2,
            double? s0Override = null)
        {
            var plane = _p.Plane;

            // 三铰点投影到平面 2D
            Vec2 O2 = plane.To2D(_p.WorldO);
            Vec2 A2 = plane.To2D(_p.WorldA);
            Vec2 C2 = plane.To2D(_p.WorldC);

            // 活塞方向 2D 单位向量
            Vec2 pistonEnd2 = plane.To2D(_p.WorldA + _p.PistonDir);
            Vec2 pd = new Vec2(pistonEnd2.X - A2.X, pistonEnd2.Y - A2.Y);
            double pdLen = pd.Length; if (pdLen < 1e-9) pdLen = 1;
            pd = new Vec2(pd.X / pdLen, pd.Y / pdLen);

            // ── 几何边长（严格对齐系统RPRR标准公式）──
            double a = (A2 - O2).Length;                  // a = |OA| 活塞铰点到动臂转轴

            // s0 = 活塞闭合长度（写入公式的常数）
            // ★已与PS内置RPRR向导逐字符对比确认：PS官方用 s0=|AC|(闭合活塞长)，不减行程！
            // 物理：闭合时活塞末端在C点(|AC|=活塞长)，张开时活塞缩短，
            //       input_j1从0变到-行程，活塞长d=D(input_j1)+s0 从330.5降到228。
            // CLOSE模式：几何是闭合的，读的|AC|=330.5即闭合活塞长，直接用。
            // OPEN模式：几何是张开的，读的|AC|=228是张开活塞长(已减行程)，
            //           须加回行程 s0=|AC|+行程=330.5 才是闭合活塞长。
            double acRead = (C2 - A2).Length;
            double s0 = s0Override.HasValue
                ? s0Override.Value
                : (_p.OpenStateAdapt ? (acRead + _p.PistonStroke) : acRead);

            double r = (C2 - O2).Length;                  // r = |OC| 动臂半径

            // b = 连杆长 = |OC|。对所有焊钳成立（C在动臂上，|OC|恒定）
            double b = (C2 - O2).Length;

            double a2 = a * a, b2 = b * b;

            // ── 公共子式 d = D(input_j1)+s0 （活塞当前总长）──
            string d = $"(D({nInput})+({s0:F6}))";

            // ── offsetA = acos((a²+s0²-b²)/(2·a·s0)) 初始装配角 ──
            double cosInitA = (a2 + s0 * s0 - b2) / (2 * a * s0);
            cosInitA = System.Math.Max(-1, System.Math.Min(1, cosInitA));
            double offsetA = System.Math.Acos(cosInitA);

            // ── fixed_Input_j1 公式（逐字符对齐系统标准，已验证完全一致）──
            // (acos(((( a² )+( {d}{d} )+(-1)( b² ))/(2( a )( d ))))+(-1)( offsetA ))
            // acos后4个左括号；d自带括号，{d}{d}隐式相邻乘；分母(2(a)(d))
            string fixedInputFormula =
                $"(acos(((({a2:F6})+({d}{d})+(-1)({b2:F6}))/(2({a:F6})({d}))))+(-1)({offsetA:F6}))";

            // ── input_j1 公式：D(j1)*行程/开口 ──
            string inputFormula =
                $"((D({nJ1})){_p.PistonStroke:F6}/{_p.OpenGap:F6})";

            // ── coup_output_j1 公式（精确复刻系统格式）──
            // (atan2((( a )sin(T(fixed)+offsetA)),(((( d )( d ))+( innerConst ))/(2( d ))))+(-1)( offsetC ))
            // innerConst = b² - a²（已验证，系统=-49628）
            // sin参数 = T(fixed_Input_j1)+offsetA（加回offsetA还原原始角）
            double innerConst = b2 - a2;

            // offsetC = 初始状态(D(input_j1)=0)时 atan2 的值，使 coup_output_j1 初始为0
            // 初始: d=s0, T(fixed_Input_j1)=0, sin参数=offsetA
            // offsetC = atan2(a·sin(offsetA), (s0²+innerConst)/(2·s0))
            // 已验证与系统标准 1.534441 完全一致
            double atan2_y_init = a * System.Math.Sin(offsetA);
            double atan2_x_init = (s0 * s0 + innerConst) / (2 * s0);
            double offsetC = System.Math.Atan2(atan2_y_init, atan2_x_init);

            string coupFormula =
                $"(atan2((({a:F6})sin((T({nFixedInput})+({offsetA:F6})))),((({d}{d})+({innerConst:F6}))/(2{d})))+(-1)({offsetC:F6}))";

            // ── output_j1 公式：(-1)(T(coup_output_j1)+T(fixed_Input_j1)) ──
            string outputFormula =
                $"((-1)(T({nCoupOut})+T({nFixedInput})))";

            // ── j2 磨量补偿：用 D(nJ1) ──
            string j2Formula = $"(((D({nJ1}))>=0)*((D({nJ1}))/2))";

            // ── 限位 ──
            double j1Low = -_p.OpenGap;
            double j1High = _p.WearAllowance;
            double inputLow = -_p.PistonStroke;
            double inputHigh = _p.PistonStroke;

            return new RprrFormulas
            {
                FixedInputJl_Formula = fixedInputFormula,
                InputJl_Formula = inputFormula,
                CoupOutputJl_Formula = coupFormula,
                OutputJl_Formula = outputFormula,
                J2_Formula = j2Formula,
                J1_Low = j1Low,
                J1_High = j1High,
                InputJl_Low = inputLow,
                InputJl_High = inputHigh,
                Const_a = a,
                Const_b = b,
                Const_r = r,
                Const_s0 = s0,
                OffsetA = offsetA,
                OffsetC = offsetC,
            };
        }

        // ═════════════════════════════════════════════════════════════════
        // 底层 SDK 调用（已验证）
        // ═════════════════════════════════════════════════════════════════
        private void CleanupPrevious()
        {
            // 诊断：打印设备现有所有 Link 和 Joint
            try
            {
                var allLinks = _device.Links;
                AddLog($"=== 设备现有 Link ({(allLinks == null ? "null" : "有")}) ===");
                if (allLinks != null)
                    foreach (ITxObject lo in allLinks)
                    {
                        var l = lo as TxKinematicLink;
                        if (l != null) AddLog($"  Link: '{l.Name}'");
                    }
                var allJoints = _device.Joints;
                AddLog($"=== 设备现有 Joint ({(allJoints == null ? "null" : "有")}) ===");
                if (allJoints != null)
                    foreach (ITxObject jo in allJoints)
                    {
                        var j = jo as TxJoint;
                        if (j != null)
                        {
                            string pn = "?", cn = "?";
                            try { pn = j.ParentLink?.Name; cn = j.ChildLink?.Name; } catch { }
                            AddLog($"  Joint: '{j.Name}' ({pn}→{cn})");
                        }
                    }
            }
            catch (Exception ex) { AddLog($"诊断设备内容异常: {ex.Message}"); }

            var ourLinkNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "fixed_link", "input_link", "coupler_link", "output_link", "dummy_link", "lnk1", "lnk2" };
            try
            {
                var joints = _device.Joints;
                if (joints == null) return;
                var toDelete = new List<TxJoint>();
                foreach (ITxObject jo in joints)
                {
                    var j = jo as TxJoint;
                    if (j == null) continue;
                    bool del = false;
                    string[] pre = { "fixed_Input_j1", "fixed_input_jl", "input_j1", "input_jl",
                                     "coup_output_j1", "coup_output_jl", "output_j1", "output_jl", "j1", "j2" };
                    foreach (var p in pre)
                        if (j.Name != null && j.Name.Equals(p, StringComparison.OrdinalIgnoreCase))
                        { del = true; break; }
                    if (!del)
                    {
                        try
                        {
                            string pn = j.ParentLink?.Name, cn = j.ChildLink?.Name;
                            if ((pn != null && ourLinkNames.Contains(pn)) ||
                                (cn != null && ourLinkNames.Contains(cn))) del = true;
                        }
                        catch { }
                    }
                    if (del) toDelete.Add(j);
                }
                foreach (var j in toDelete)
                {
                    string jn = "?"; try { jn = j.Name; } catch { }
                    try { j.Delete(); AddLog($"清理残留 Joint: {jn}"); }
                    catch (Exception ex) { AddLog($"删除 Joint '{jn}' 失败: {ex.Message}"); }
                }
            }
            catch (Exception ex) { AddLog($"清理异常: {ex.Message}"); }
        }

        private TxKinematicLink FindExistingLink(string name)
        {
            try
            {
                var links = _device.Links;
                if (links != null)
                    foreach (ITxObject lo in links)
                    {
                        var l = lo as TxKinematicLink;
                        if (l != null && l.Name == name) return l;
                    }
            }
            catch { }
            return null;
        }

        private TxKinematicLink CreateLink(string name, List<ITxObject> bodies)
        {
            try
            {
                TxKinematicLink link = FindExistingLink(name);
                if (link != null) AddLog($"  Link '{name}' 复用");
                else
                {
                    var data = new TxKinematicLinkCreationData();
                    data.Name = name;
                    link = _device.CreateLink(data);
                    if (link == null) return null;
                }
                if (bodies != null)
                    foreach (var b in bodies)
                        if (b != null) try { link.AddObject(b); } catch { }
                AddLog($"  Link '{name}' OK ({bodies?.Count ?? 0} 几何体)");
                return link;
            }
            catch (Exception ex) { AddLog($"  [X] Link '{name}': {ex.Message}"); return null; }
        }

        private TxJoint CreateJointInternal(string name, string typeName,
            TxVector from, TxVector to, double low, double high,
            TxKinematicLink parentLink, TxKinematicLink childLink)
        {
            try
            {
                var data = new TxJointCreationData();
                // 文档确认：TxJointCreationData 有 Name 属性（get;set;）
                data.Name = name;
                data.JointType = typeName == "Prismatic"
                    ? TxJoint.TxJointType.Prismatic : TxJoint.TxJointType.Revolute;
                if (parentLink != null) data.ParentLink = parentLink;
                if (childLink != null) data.ChildLink = childLink;
                data.SetAxisPoints(from, to);

                // 诊断：打印轴坐标和Link
                double axisLen = System.Math.Sqrt(
                    (to.X - from.X) * (to.X - from.X) +
                    (to.Y - from.Y) * (to.Y - from.Y) +
                    (to.Z - from.Z) * (to.Z - from.Z));
                AddLog($"  创建{name}: from=({from.X:F1},{from.Y:F1},{from.Z:F1}) to=({to.X:F1},{to.Y:F1},{to.Z:F1})");
                AddLog($"    parent={parentLink?.Name} child={childLink?.Name} 轴长={axisLen:F4}");

                TxJoint joint = _device.CreateJoint(data);
                if (joint == null) { AddLog($"  [X] {name}: CreateJoint 返回 null"); return null; }

                if (high != 0 || low != 0)
                {
                    try
                    {
                        // HardLimits = Constant 模式：构造函数(lowerLimit, upperLimit)
                        var hardLimits = new TxJointConstantHardLimits(low, high);
                        joint.HardLimits = hardLimits;
                    }
                    catch (Exception exh) { AddLog($"  [!] {name} HardLimits失败: {exh.Message}"); }
                    try { joint.LowerSoftLimit = low; joint.UpperSoftLimit = high; }
                    catch (Exception exl) { AddLog($"  [!] {name} SoftLimits失败: {exl.Message}"); }
                    AddLog($"  限位: [{low:F2}, {high:F2}]");
                }

                AddLog($"  [OK] {name} ({typeName}) → 实际名: {SafeName(joint)}");
                return joint;
            }
            catch (Exception ex)
            {
                AddLog($"  [X] {name} 异常: {ex.Message}");
                if (ex.InnerException != null) AddLog($"      内部: {ex.InnerException.Message}");
                AddLog($"      StackTrace: {ex.StackTrace?.Split('\n')[0]}");
                return null;
            }
        }

        /// <summary>
        /// 安全读取 Joint 名（TxJoint.Name 的 set 抛 TxNotImplementedException，
        /// 无法改名，只能读取 PS 自动分配的名字）。
        /// </summary>
        private string SafeName(TxJoint joint)
        {
            if (joint == null) return "?";
            try { return joint.Name; } catch { return "?"; }
        }

        private bool WriteFormula(TxJoint joint, string formula, string label)
        {
            if (joint == null || string.IsNullOrEmpty(formula)) return false;
            try { joint.KinematicsFunction = formula; AddLog($"  {label} 公式 OK"); return true; }
            catch (Exception ex) { AddLog($"  {label} 公式失败({ex.Message}): {formula}"); return false; }
        }

        private void AddLog(string msg) => Log += msg + "\n";
    }

    // ═════════════════════════════════════════════════════════════════════
    public class RprrParams
    {
        public Vec3 WorldO { get; set; }
        public Vec3 WorldA { get; set; }
        public Vec3 WorldC { get; set; }
        public Vec3 N { get; set; }
        public Vec3 PistonDir { get; set; }
        public ReferencePlane Plane { get; set; }

        public List<ITxObject> FixedLinkBodies { get; set; } = new List<ITxObject>();
        public List<ITxObject> InputLinkBodies { get; set; } = new List<ITxObject>();
        public List<ITxObject> CouplerLinkBodies { get; set; } = new List<ITxObject>();
        public List<ITxObject> OutputLinkBodies { get; set; } = new List<ITxObject>();
        public List<ITxObject> Lnk2Bodies { get; set; } = new List<ITxObject>();

        public Vec3 WorldStaticTip { get; set; }
        public Vec3 WorldMovingTip { get; set; }
        public double R_arm { get; set; }

        public double S0 { get; set; } = 0;
        public double PistonStroke { get; set; } = 100;
        public double OpenGap { get; set; } = 170;
        public double WearAllowance { get; set; } = 10;

        // 需求4：OPEN状态适配。厂家焊钳模型若是OPEN状态(开口最大,活塞最短)，
        // 勾选后分阶段写入：先建主链独立驱动input_j1回闭合→DefineZeroPosition→
        // 重写input_j1限位为[-行程,0]→再建旁支j1/j2接驱动链。
        public bool OpenStateAdapt { get; set; } = false;
    }

    public class RprrFormulas
    {
        public string FixedInputJl_Formula { get; set; }
        public string InputJl_Formula { get; set; }
        public string CoupOutputJl_Formula { get; set; }
        public string OutputJl_Formula { get; set; }
        public string J2_Formula { get; set; }

        public double J1_Low, J1_High;
        public double InputJl_Low, InputJl_High;

        public double Const_a, Const_b, Const_r, Const_s0;
        public double OffsetA, OffsetC;
    }
}