// AutoPathPlannerForm.cs — C# 7.3
//
// v5.5 全参数 GUI 可调:
// - 规划参数区改 TabControl 四组: 常用 / 定向搜索L1 / 动态检查 / RRT&采样
// - 25 个可调量全部暴露 (原来只有 5 个), 每项带悬浮说明
// - CollisionWorld 的动态检查参数经 Planner 透传, 不再需要改代码
// - "恢复默认"按钮
//
// v5.4 后撤门槛 + 动态验证前移 (语义澄清: 后撤=沿枪-X, 非进出枪-Z):
// - GunFrameSearch 新增 MinBackoutForSide(200mm): 抬升/横移前必须先后撤够,
//   否则枪体还埋在夹具里, 一抬就扫 (日志"后撤A40+上60mm通过"就是这个 bug)
// - 动态扫掠验证前移到 L1 生成阶段 (IsEdgeSafeDynamic), 静态可行但枪体扫掠的
//   候选直接拒掉并加深后撤, 不再留给阶段2被动修
// - 侧向机动深度由深到浅遍历 (PreferDeepBackoutForSide)
// - CheckJointMotion 步数改 关节角/笛卡尔位移 双判据取大, 上限 24 → 64
//   (枪长, 腕部转5°枪尖可扫过100mm, 原采样密度会漏检)
// - 动态修复新增策略⓪: 双侧深退门形 (200~800mm), 置于最优先
//
// v5.3 深退修复 + 共线剪枝:
// - 动态修复退让阶梯 80mm → 600mm, 新增"深退+抬升"组合与双侧门形退让
//   (图1: 枪深埋夹具, 退 80mm 脱不了困, 抬升瞬间撞)
// - 修复预算 固定8 → 按链长动态 12~40 (原来 15 个违例只修了 5 个)
// - 进/出枪默认 20/10mm → 80/30mm, 上限 200 → 600
// - 新增阶段2.5 共线过渡点剪枝: 直线段上的连续 Via 只留首尾 (图2)
//   删除后复验静态边+动态扫掠, 不安全则回滚
//
// v5.2 动态干涉检查修复:
// - DynamicRefine 捕获改分段容错: 摆位失败节点标记为"洞", 只跳过相邻边,
//   其余边正常扫掠 (原: 一个节点失败就 return, 79个Via一条边都没检)
// - CheckJointMotion 关节读/写失败不再静默当"通过", 返回"未验证"让上层处理
// - ApplyJoints 失败不再全局停用动态检查; 增加 SetJointValue 备用写入路径
// - 新增 SelfTestJointIO(): 精修前做关节读→写→读往返自检, 不可用则明确告知
//
// v5.1 绑定资源完整收集:
// - FirstList 补齐 PrimaryLocator(底座 fupa) / Toolbox / AttachmentParent / Links
//   (之前只有 5 项, fupa 落在障碍方 → 机器人与自己底座永久报干涉)
// - SecondList 用 HashSet 引用级硬排除全部绑定资源 + 最终安全网校验交集为空
//
// v5.0 致命 bug 修复:
// - 干涉集查找一直用 root.CollisionPairs (不存在的属性!) → 永远返回 null
//   → "永远未找到干涉集"。改用正确的 root.Pairs (ArrayList) 强类型读取。
// - 新建 pair 时自动开启 root.CheckCollisions (实测场景中该全局开关为 false)
//
// v4.10 附着模式:
// - 规划器只复用现有干涉集 (attachOnly), 找不到跳过, 绝不新建 (修复堆积)
// - UI"自动创建"按钮仍可主动新建
//
// v4.9 鲁棒化:
// - MakeFingerprint 用 dynamic 拿位置 (兼容更多 SDK 版本)
// - CollectFirstSide 用 属性+方法+IEnumerable 三层反射 (兼容更多 pair 内部实现)
// - TryFindExistingPairName 找不到时 dump 诊断日志, 便于定位
//
// v4.5 修复 (API 优化):
// - 障碍方首层剔除 ITxComponent.IsEquipment 容器 (官方 API 语义: 逻辑容器,
//   无独立几何), 避免父级 Cell/Group 以整个边界参与检测导致大量误报
//
// v4.4 简化:
// - 障碍方一律剔除场景所有 TxRobot 与焊枪 (类型名含 Gun) 子树
// - 规划器不再自动创建干涉集 (必须先点"自动创建干涉集"按钮)
//
// v4.1 修正:
// - "自动创建干涉集"改为从操作网格反查 op.Robot (与规划器同源, 修复多机器人/同名场景不匹配)
// - 支持一次为多台机器人创建/复用干涉集, 逐台弹窗决策
// - 日志同步打印机器人身份 (Name + HashCode + 基座坐标), 便于人工核对
//
// v4 更新:
// - 移除干涉物网格 (UI 简化, 干涉物 = 场景自动收集, 不再手动追加)
// - 新增"自动创建干涉集"按钮 + 状态标签
// - 检测已有干涉集时弹窗: 强制新建 / 复用 / 取消
// - 规划完不再删除干涉集 (KeepPairOnDispose), 可复用与 PS 侧检视
//
// v3 重构说明（基于 ExportGun 模式）：
// [重构1] Theme 靜态类：集中管理所有颜色常量
// [重构2] LogLevel 枚举 + RichTextBox 彩色日志
// [重构3] #region 分组控件字段 / 状态字段
// [重构4] GroupBox 卡片布局替代 flat Panel
// [重构5] 移除 PS 原生规划选项 (方案不可行，仅保留 RRT)
// [重构6] DPI 处理与 ExportGun 一致
// [重构7] 移除 dynamic 调用不存在的 MoveChildToFirst/MoveChildToLast
//

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Tecnomatix.Engineering;
using Tecnomatix.Engineering.Ui;

// CS0104 防御别名
using Button = System.Windows.Forms.Button;
using CheckBox = System.Windows.Forms.CheckBox;
using Label = System.Windows.Forms.Label;
using TextBox = System.Windows.Forms.TextBox;
using ComboBox = System.Windows.Forms.ComboBox;
// v5.5: Tab 参数区新增控件, 同样需要别名 (解决方案含 WPF 引用会歧义)
using TabControl = System.Windows.Forms.TabControl;
using TabPage = System.Windows.Forms.TabPage;
using ToolTip = System.Windows.Forms.ToolTip;
using Panel = System.Windows.Forms.Panel;
using Control = System.Windows.Forms.Control;
using ProgressBar = System.Windows.Forms.ProgressBar;

namespace TxTools.AutoPathPlanner
{
    public partial class AutoPathPlannerForm : TxForm
    {
        // ════════════════════════════════════════════════════════════
        //  Theme：集中管理所有颜色常量
        // ════════════════════════════════════════════════════════════
        private static class Theme
        {
            // 卡片标题色
            public static readonly Color CardOps = Color.FromArgb(0, 100, 140);
            public static readonly Color CardObs = Color.FromArgb(80, 120, 140);
            public static readonly Color CardParams = Color.FromArgb(155, 120, 0);
            public static readonly Color CardLog = Color.FromArgb(50, 120, 60);

            // 功能按钮色
            public static readonly Color BtnPrimary = Color.FromArgb(0, 100, 167);
            public static readonly Color BtnSecondary = Color.FromArgb(80, 120, 140);
            public static readonly Color BtnMuted = Color.FromArgb(120, 124, 135);
            public static readonly Color BtnDanger = Color.FromArgb(130, 50, 50);

            // 日志面板
            public static readonly Color LogBg = Color.FromArgb(20, 22, 27);
            public static readonly Color LogText = Color.FromArgb(178, 200, 178);
            public static readonly Color LogOk = Color.FromArgb(90, 210, 110);
            public static readonly Color LogErr = Color.FromArgb(228, 88, 88);
            public static readonly Color LogWarn = Color.FromArgb(228, 180, 70);
            public static readonly Color LogPs = Color.FromArgb(110, 180, 228);
            public static readonly Color LogDebug = Color.FromArgb(140, 140, 160);
            public static readonly Color LogInfo = Color.FromArgb(178, 200, 178);
        }

        // ════════════════════════════════════════════════════════════
        //  日志级别
        // ════════════════════════════════════════════════════════════
        private enum LogLevel { Info, Ok, Warn, Error, Ps, Debug }

        // ════════════════════════════════════════════════════════════
        //  状态字段
        // ════════════════════════════════════════════════════════════
        #region State
        private static AutoPathPlannerForm _instance;
        private volatile bool _stopRequested;
        private bool _dpiApplied;
        private readonly Size _designSize = new Size(760, 812);   // v5.5: Tab参数区
        #endregion

        // ════════════════════════════════════════════════════════════
        //  控件字段（按区域分组）
        // ════════════════════════════════════════════════════════════
        #region UI Controls - 操作选择
        private TxObjGridCtrl _gridOps;
        private Button _btnAddOps;
        private Button _btnClearOps;
        #endregion

        #region UI Controls - 干涉集
        private Button _btnAutoCollisionSet;
        private Label _lblCollisionSetStatus;
        #endregion

        #region UI Controls - 规划参数 (v5.5: 全量暴露)
        // --- 进/出枪 (沿枪 -Z) ---
        private NumericUpDown _nudApproach;      // ApproachRetractDistance
        private NumericUpDown _nudApproachMin;   // ApproachRetractMin
        private CheckBox _chkWorldZ;             // UseWorldZForApproach
        private CheckBox _chkAppRet;             // GenerateApproachRetract

        // --- 定向搜索 L1 (沿枪 -X 后撤 + 侧向) ---
        private NumericUpDown _nudBackoutStep;   // GunBackoutStep
        private NumericUpDown _nudBackoutMax;    // GunBackoutMax
        private NumericUpDown _nudMinBackout;    // GunMinBackoutForSide  ★核心
        private NumericUpDown _nudSideStep;      // GunSideStep
        private NumericUpDown _nudSideMax;       // GunSideMax

        // --- 动态干涉检查 (关节扫掠) ---
        private CheckBox _chkDynamic;            // DynamicCheckEnabled
        private NumericUpDown _nudJointQuantum;  // DynamicJointQuantum
        private NumericUpDown _nudCartQuantum;   // DynamicCartesianQuantum
        private NumericUpDown _nudMaxSweep;      // MaxSweepSteps
        private NumericUpDown _nudConfigJump;    // ConfigJumpThreshold

        // --- RRT ---
        private NumericUpDown _nudStep;          // RrtStepSize
        private NumericUpDown _nudIter;          // RrtMaxIterations
        private NumericUpDown _nudGoalBias;      // RrtGoalBias (×100 存百分数)
        private NumericUpDown _nudEdgeRes;       // EdgeCheckResolution

        // --- 采样空间膨胀 ---
        private NumericUpDown _nudInflateXy;     // SampleBoundsInflateXy
        private NumericUpDown _nudInflateZUp;    // SampleBoundsInflateZUp
        private NumericUpDown _nudInflateZDn;    // SampleBoundsInflateZDown

        // --- 姿态变体 / 后处理 ---
        private CheckBox _chkVariants;           // OrientationVariantsEnabled
        private NumericUpDown _nudMaxVariants;   // MaxVariantTries
        private NumericUpDown _nudCollinearTol;  // CollinearTolerance
        private CheckBox _chkPrune;              // 共线剪枝开关

        // --- v6.0 节拍优化 ---
        private CheckBox _chkWeldOrder;          // WeldOrderOptEnabled
        private NumericUpDown _nudLockFirst;     // WeldOrderLockFirst
        private NumericUpDown _nudLockLast;      // WeldOrderLockLast
        private CheckBox _chkMotionJoint;        // SetTransitMotionJoint

        // --- v6.0 焊钳外部轴 ---
        private CheckBox _chkGunAxis;            // GunAxisWriteEnabled
        private NumericUpDown _nudTargetOpening; // TargetGunOpening
        private CheckBox _chkAdaptiveGun;        // AdaptiveGunOpening
        private NumericUpDown _nudTransitOpening;// TransitGunOpening
        private ComboBox _cboGunDir;             // v6.4 开口方向覆盖
        private NumericUpDown _nudGunMaxOpen;    // v6.4 最大开口幅值覆盖

        private CheckBox _chkCache;              // v6.3 查询缓存

        private Button _btnResetParams;
        #endregion

        #region UI Controls - 进度 (v6.3)
        private ProgressBar _progressBar;
        private Label _lblStage;
        private DateTime _runStart;
        #endregion

        #region UI Controls - 操作按钮
        private Button _btnRun;
        private Button _btnStop;
        #endregion

        #region UI Controls - 日志
        private RichTextBox _rtbLog;
        #endregion

        // ════════════════════════════════════════════════════════════
        //  构造 / 单例
        // ════════════════════════════════════════════════════════════
        public static void ShowSingleton()
        {
            if (_instance == null || _instance.IsDisposed)
                _instance = new AutoPathPlannerForm();
            _instance.Show();
            _instance.BringToFront();
        }

        public AutoPathPlannerForm()
        {
            SemiModal = false;
            BuildUI();
        }

        public override void OnInitTxForm()
        {
            base.OnInitTxForm();
            SemiModal = false;
            Name = GetType().FullName;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            if (!_dpiApplied)
            {
                _dpiApplied = true;
                try
                {
                    Size = _designSize;
                    float sc = CreateGraphics().DpiX / 96f;
                    if (sc < 1f) sc = 1f;
                    if (sc > 1.01f) Scale(new SizeF(sc, sc));
                }
                catch { }
            }
        }

        // ════════════════════════════════════════════════════════════
        //  界面搭建
        // ════════════════════════════════════════════════════════════
        private void BuildUI()
        {
            SuspendLayout();
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.None;
            Text = "自动路径规划器 v6.5.1 · API 校正";
            Size = _designSize;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = false;
            BackColor = SystemColors.Control;
            Font = SystemFonts.MessageBoxFont;

            // ---- 关闭 Siemens flat style 皮肤 ----
            try
            {
                var flatStyleProp = GetType().GetProperty("FlatStyleEnabled");
                if (flatStyleProp != null && flatStyleProp.CanWrite)
                    flatStyleProp.SetValue(this, false, null);
            }
            catch { }

            // ══════════ 卡片1: 操作选择 ══════════
            var cardOps = MakeCard("焊接操作 (规划范围)", Theme.CardOps, 180);
            {
                _gridOps = new TxObjGridCtrl
                {
                    Location = new Point(10, 24),
                    Size = new Size(690, 120)
                };
                cardOps.Controls.Add(_gridOps);

                _btnAddOps = MakeBtn("添加选中", Theme.BtnPrimary, 10, 150);
                _btnAddOps.Click += delegate { AppendOpsSelection(); };
                cardOps.Controls.Add(_btnAddOps);

                _btnClearOps = MakeBtn("清空", Theme.BtnMuted, 110, 150);
                _btnClearOps.Click += delegate { ClearGrid(_gridOps); };
                cardOps.Controls.Add(_btnClearOps);
            }
            Controls.Add(cardOps);

            // ══════════ 卡片2: 干涉集 (机器人 vs 场景, 自动收集) ══════════
            var cardCS = MakeCard("干涉集 (机器人 vs 场景, 自动收集)", Theme.CardObs, 90);
            {
                _btnAutoCollisionSet = MakeBtn("自动创建干涉集", Theme.BtnSecondary,
                    10, 30, 160, 30);
                _btnAutoCollisionSet.Click += OnAutoCreateCollisionSet;
                cardCS.Controls.Add(_btnAutoCollisionSet);

                _lblCollisionSetStatus = new Label
                {
                    Location = new Point(180, 36),
                    AutoSize = true,
                    Text = "状态: 未检测 (规划前必须先创建, 规划器不再自动新建)",
                    ForeColor = Color.FromArgb(90, 90, 90)
                };
                cardCS.Controls.Add(_lblCollisionSetStatus);
            }
            Controls.Add(cardCS);

            // ══════════ 卡片3: 规划参数 (v5.5 全量暴露, Tab 分组) ══════════
            var cardParams = MakeCard("规划参数", Theme.CardParams, 232);
            {
                var tabs = new TabControl
                {
                    Location = new Point(10, 22),
                    Size = new Size(700, 172),
                    Appearance = TabAppearance.Normal
                };

                // ---------- Tab 1: 常用 ----------
                var tp1 = new TabPage("常用");
                {
                    int y = 12;
                    _nudApproach = AddRow(tp1, "进/出枪最大(mm):", 10, ref y, 20, 2, 300,
                        "沿枪 -Z 把枪嘴从板件提起的距离上限");
                    _nudApproachMin = AddRow(tp1, "进/出枪最小(mm):", 10, ref y, 10, 1, 300,
                        "进/出枪距离区间下限");

                    _chkWorldZ = AddCheck(tp1, "进出枪仅沿世界Z", 10, ref y, false,
                        "勾选后进/出枪沿世界Z, 而非枪坐标系 -Z");
                    _chkAppRet = AddCheck(tp1, "生成进/出枪 Via 点", 10, ref y, true,
                        "关闭则不在焊点前后插入进/出枪过渡点");

                    int y2 = 12;
                    _nudMinBackout = AddRow(tp1, "★抬升前最小后撤(mm):", 330, ref y2, 200, 0, 900,
                        "沿枪 -X 抽出夹具的最小深度。侧向机动(抬升/横移)前必须先退够, "
                        + "否则枪体还埋在夹具里, 一抬就扫。夹具深则调大 (300~500)。0=关闭约束");
                    _nudBackoutMax = AddRow(tp1, "最大后撤深度(mm):", 330, ref y2, 800, 100, 2000,
                        "L1 沿枪 -X 后撤的搜索上限");

                    _chkDynamic = AddCheck(tp1, "动态干涉检查 (关节扫掠)", 330, ref y2, true,
                        "关闭后只做静态点位检查, 看不见枪体扫掠 —— 强烈建议保持开启");
                    _chkPrune = AddCheck(tp1, "共线过渡点剪枝", 330, ref y2, true,
                        "直线段上的连续过渡点只保留首尾 (删除后会复验安全性)");
                }
                tabs.TabPages.Add(tp1);

                // ---------- Tab 2: 定向搜索 (L1) ----------
                var tp2 = new TabPage("定向搜索 L1");
                {
                    int y = 12;
                    _nudBackoutStep = AddRow(tp2, "后撤步进(mm):", 10, ref y, 60, 10, 200,
                        "沿枪 -X 后撤的扫描步长。小=精细但慢");
                    _nudSideStep = AddRow(tp2, "侧向步进(mm):", 10, ref y, 60, 10, 200,
                        "抬升/横移的扫描步长");
                    _nudSideMax = AddRow(tp2, "侧向最大偏移(mm):", 10, ref y, 800, 100, 2000,
                        "抬升/横移的搜索上限");

                    int y2 = 12;
                    AddNote(tp2, 330, ref y2,
                        "后撤 = 沿枪 -X 抽出枪体 (300~800mm)");
                    AddNote(tp2, 330, ref y2,
                        "进/出枪 = 沿枪 -Z 提起枪嘴 (10~20mm)");
                    AddNote(tp2, 330, ref y2, "");
                    AddNote(tp2, 330, ref y2,
                        "门形直连(纯-X)不受后撤门槛限制;");
                    AddNote(tp2, 330, ref y2,
                        "抬升/横移必须先退够门槛才允许。");
                }
                tabs.TabPages.Add(tp2);

                // ---------- Tab 3: 动态检查 ----------
                var tp3 = new TabPage("动态检查");
                {
                    int y = 12;
                    _nudJointQuantum = AddRow(tp3, "关节步进量子(°):", 10, ref y, 4, 1, 30,
                        "扫掠步数 = 关节最大跨度 / 此值。小=采样密、慢");
                    _nudCartQuantum = AddRow(tp3, "笛卡尔量子(mm):", 10, ref y, 15, 3, 100,
                        "扫掠步数 = TCP位移 / 此值。与关节判据取大者。"
                        + "枪长时腕部小角度也会让枪尖扫过很远, 这一项防漏检");
                    _nudMaxSweep = AddRow(tp3, "扫掠步数上限:", 10, ref y, 64, 8, 200,
                        "单条边的最大扫掠采样数。大=准但慢");

                    int y2 = 12;
                    _nudConfigJump = AddRow(tp3, "构型突变阈值(°):", 330, ref y2, 120, 30, 360,
                        "相邻点关节跨度超此值判定为构型翻转 (奇异/翻腕)");
                    _nudCollinearTol = AddRow(tp3, "共线容差(mm):", 330, ref y2, 3, 1, 50,
                        "中点到首尾连线垂距小于此值视为共线, 可剪除");
                }
                tabs.TabPages.Add(tp3);

                // ---------- Tab 4: RRT & 采样 ----------
                var tp4 = new TabPage("RRT & 采样");
                {
                    int y = 12;
                    _nudStep = AddRow(tp4, "RRT 步长(mm):", 10, ref y, 50, 10, 400,
                        "RRT 树扩展步长");
                    _nudIter = AddRow(tp4, "RRT 最大迭代:", 10, ref y, 5000, 200, 50000,
                        "单段 RRT 的迭代上限");
                    _nudGoalBias = AddRow(tp4, "目标偏置(%):", 10, ref y, 15, 0, 90,
                        "采样时直接朝目标的概率");
                    _nudEdgeRes = AddRow(tp4, "连边检测分辨(mm):", 10, ref y, 12, 3, 60,
                        "边采样间隔。小=防薄板隧穿但慢");

                    int y2 = 12;
                    _nudInflateXy = AddRow(tp4, "采样膨胀 XY(mm):", 330, ref y2, 300, 0, 2000,
                        "RRT 采样盒在 XY 方向的膨胀量");
                    _nudInflateZUp = AddRow(tp4, "采样膨胀 Z上(mm):", 330, ref y2, 400, 0, 2000,
                        "采样盒向上膨胀");
                    _nudInflateZDn = AddRow(tp4, "采样膨胀 Z下(mm):", 330, ref y2, 100, 0, 2000,
                        "采样盒向下膨胀");

                    _chkVariants = AddCheck(tp4, "姿态变体搜索", 330, ref y2, true,
                        "点位不可达时尝试绕枪轴旋转等姿态变体");
                    _chkCache = AddCheck(tp4, "★位姿查询缓存", 330, ref y2, true,
                        "碰撞查询 ~11ms/次, 是唯一瓶颈。SDK 非线程安全无法并行,\r\n"
                        + "只能靠缓存\"少查\"。阶梯搜索/复验会大量重复查同一位姿");
                    _nudMaxVariants = AddRow(tp4, "变体尝试上限:", 330, ref y2, 13, 1, 60,
                        "每个点位最多尝试的姿态变体数");
                }
                tabs.TabPages.Add(tp4);

                // ---------- Tab 5: 节拍优化 & 焊钳 (v6.0) ----------
                var tp5 = new TabPage("节拍 & 焊钳");
                {
                    int y = 12;
                    _chkWeldOrder = AddCheck(tp5, "★焊点顺序优化 (2-opt)", 10, ref y, false,
                        "以关节节拍为边权重排焊接顺序, 节拍收益最大。\r\n"
                        + "⚠ 会改变焊接顺序 —— 若工艺有定位焊/防变形约束, 请用下方锁定选项");
                    _nudLockFirst = AddRow(tp5, "锁定开头(个):", 10, ref y, 0, 0, 50,
                        "开头 N 个焊点不参与重排 (定位焊必须先焊)");
                    _nudLockLast = AddRow(tp5, "锁定末尾(个):", 10, ref y, 0, 0, 50,
                        "末尾 N 个焊点不参与重排");
                    _chkMotionJoint = AddCheck(tp5, "过渡点用 Joint (PTP)", 10, ref y, true,
                        "过渡点运动类型设为 Joint —— 各关节以最高效方式移动。\r\n"
                        + "用 Linear 是节拍杀手 (TCP 走直线, 关节被迫绕路)");

                    int y2 = 12;
                    _chkGunAxis = AddCheck(tp5, "焊钳开口写入外部轴", 330, ref y2, true,
                        "把开口值写进 Via 点的 RobotExternalAxesData。\r\n"
                        + "需要焊钳已注册为机器人外部轴 (否则日志会提示)");
                    _nudTargetOpening = AddRow(tp5, "进/出枪开口(mm):", 330, ref y2, 30, 0, 300,
                        "进/出枪点的开口 —— 够越过翻边即可");
                    _chkAdaptiveGun = AddCheck(tp5, "过渡点按需最小开口", 330, ref y2, true,
                        "从小到大试开口, 第一个不干涉的就用。\r\n"
                        + "开口大=伺服行程长=慢, 且枪臂张开包络更大反而更容易撞");
                    _nudTransitOpening = AddRow(tp5, "过渡点开口(mm):", 330, ref y2, 60, 0, 300,
                        "关闭\"按需最小开口\"时, 过渡点统一用此开口");

                    // v6.4 开口方向手动覆盖
                    tp5.Controls.Add(new Label
                    {
                        Text = "★开口方向:",
                        Location = new Point(330, y2 + 3),
                        AutoSize = true
                    });
                    _cboGunDir = new ComboBox
                    {
                        Location = new Point(480, y2),
                        Size = new Size(100, 24),
                        DropDownStyle = ComboBoxStyle.DropDownList
                    };
                    _cboGunDir.Items.AddRange(new object[] { "自动探测", "正向 (+)", "负向 (−)" });
                    _cboGunDir.SelectedIndex = 0;
                    _tip.SetToolTip(_cboGunDir,
                        "伺服枪不保证正向开启。常见 CLOSE=0 / OPEN=-60 (负向)。\r\n"
                        + "自动探测失败时 (日志报\"符号约定未探到\"), 在这里手动指定。\r\n"
                        + "填正值给负向枪 = 往闭合方向压 = 夹住钣金/超限位");
                    tp5.Controls.Add(_cboGunDir);
                    y2 += 28;

                    _nudGunMaxOpen = AddRow(tp5, "最大开口幅值(mm):", 330, ref y2, 0, 0, 500,
                        "枪的物理最大开口。0 = 用探测值 (探不到则默认 200)");
                }
                tabs.TabPages.Add(tp5);

                cardParams.Controls.Add(tabs);

                _btnResetParams = MakeBtn("恢复默认", Theme.BtnMuted, 600, 198, 110, 26);
                _btnResetParams.Click += delegate { ResetParamsToDefault(); };
                cardParams.Controls.Add(_btnResetParams);
            }
            Controls.Add(cardParams);

            // ══════════ 按钮行 ══════════
            var pnlBtn = new Panel { Dock = DockStyle.Top, Height = 42, Padding = new Padding(6) };
            {
                _btnRun = MakeBtn("开始规划", Theme.BtnPrimary, 6, 6, 150, 28);
                _btnRun.Click += OnRun;
                pnlBtn.Controls.Add(_btnRun);

                _btnStop = MakeBtn("停止", Theme.BtnDanger, 164, 6, 100, 28);
                _btnStop.Enabled = false;
                _btnStop.Click += delegate
                {
                    _stopRequested = true;
                    _btnStop.Enabled = false;
                    Log("[请求中止] 将在当前检查点安全停止...", LogLevel.Warn);
                };
                pnlBtn.Controls.Add(_btnStop);

                // v6.3 进度条
                _progressBar = new ProgressBar
                {
                    Location = new Point(272, 8),
                    Size = new Size(300, 24),
                    Minimum = 0,
                    Maximum = 1000,
                    Value = 0,
                    Style = ProgressBarStyle.Continuous
                };
                pnlBtn.Controls.Add(_progressBar);

                _lblStage = new Label
                {
                    Location = new Point(580, 12),
                    AutoSize = true,
                    Text = "就绪",
                    ForeColor = Color.FromArgb(90, 90, 90)
                };
                pnlBtn.Controls.Add(_lblStage);

            }
            Controls.Add(pnlBtn);
            pnlBtn.BringToFront();

            // ══════════ 卡片4: 日志 ══════════
            var cardLog = MakeCard("运行日志", Theme.CardLog, 0); // 高度=0 → Dock.Fill
            cardLog.Dock = DockStyle.Fill;
            {
                _rtbLog = new RichTextBox
                {
                    Dock = DockStyle.Fill,
                    Location = new Point(10, 24),
                    Font = new Font("Consolas", 9f),
                    BackColor = Theme.LogBg,
                    ForeColor = Theme.LogText,
                    ReadOnly = true,
                    WordWrap = false,
                    BorderStyle = BorderStyle.None
                };
                cardLog.Controls.Add(_rtbLog);
            }
            Controls.Add(cardLog);
            cardLog.BringToFront();

            ResumeLayout(false);
            PerformLayout();
        }

        // ════════════════════════════════════════════════════════════
        //  UI 辅助
        // ════════════════════════════════════════════════════════════
        private static GroupBox MakeCard(string title, Color titleColor, int height)
        {
            var g = new GroupBox
            {
                Text = title,
                ForeColor = titleColor,
                Dock = height == 0 ? DockStyle.Fill : DockStyle.Top,
                Font = new Font("Microsoft YaHei UI", 9.5f, FontStyle.Bold),
                Padding = new Padding(6, 4, 6, 4)
            };
            if (height > 0) g.Height = height;
            return g;
        }

        private static Button MakeBtn(string text, Color backColor, int x, int y,
            int w = 90, int h = 26)
        {
            var b = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(w, h),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = SystemFonts.MessageBoxFont
            };
            b.FlatAppearance.BorderColor = backColor;
            return b;
        }

        /// <summary>共享 ToolTip 实例 (参数说明悬浮提示)</summary>
        private readonly ToolTip _tip = new ToolTip
        {
            AutoPopDelay = 20000,
            InitialDelay = 400,
            ReshowDelay = 200,
            ShowAlways = true
        };

        /// <summary>参数行: 标签 + NumericUpDown (+ 悬浮说明)</summary>
        private NumericUpDown AddRow(Control parent, string label, int x,
            ref int y, decimal val, decimal min, decimal max, string tip = null)
        {
            var lbl = new Label
            {
                Text = label,
                Location = new Point(x, y + 3),
                AutoSize = true
            };
            parent.Controls.Add(lbl);

            var nud = new NumericUpDown
            {
                Location = new Point(x + 150, y),
                Size = new Size(80, 24),
                Minimum = min,
                Maximum = max,
                Value = val
            };
            parent.Controls.Add(nud);

            if (!string.IsNullOrEmpty(tip))
            {
                _tip.SetToolTip(lbl, tip);
                _tip.SetToolTip(nud, tip);
            }

            y += 28;
            return nud;
        }

        /// <summary>参数行: CheckBox (+ 悬浮说明)</summary>
        private CheckBox AddCheck(Control parent, string text, int x,
            ref int y, bool chk, string tip = null)
        {
            var cb = new CheckBox
            {
                Text = text,
                Location = new Point(x, y + 2),
                AutoSize = true,
                Checked = chk,
                FlatStyle = FlatStyle.Standard
            };
            parent.Controls.Add(cb);
            if (!string.IsNullOrEmpty(tip)) _tip.SetToolTip(cb, tip);
            y += 26;
            return cb;
        }

        /// <summary>说明文字行 (非交互)</summary>
        private void AddNote(Control parent, int x, ref int y, string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                parent.Controls.Add(new Label
                {
                    Text = text,
                    Location = new Point(x, y + 2),
                    AutoSize = true,
                    ForeColor = Color.FromArgb(100, 100, 100)
                });
            }
            y += 22;
        }

        /// <summary>恢复所有参数到出厂默认</summary>
        private void ResetParamsToDefault()
        {
            // 进/出枪
            _nudApproach.Value = 20;
            _nudApproachMin.Value = 10;
            _chkWorldZ.Checked = false;
            _chkAppRet.Checked = true;

            // 定向搜索 L1
            _nudBackoutStep.Value = 60;
            _nudBackoutMax.Value = 800;
            _nudMinBackout.Value = 200;
            _nudSideStep.Value = 60;
            _nudSideMax.Value = 800;

            // 动态检查
            _chkDynamic.Checked = true;
            _nudJointQuantum.Value = 4;
            _nudCartQuantum.Value = 15;
            _nudMaxSweep.Value = 64;
            _nudConfigJump.Value = 120;
            _nudCollinearTol.Value = 3;
            _chkPrune.Checked = true;

            // RRT & 采样
            _nudStep.Value = 50;
            _nudIter.Value = 5000;
            _nudGoalBias.Value = 15;
            _nudEdgeRes.Value = 12;
            _nudInflateXy.Value = 300;
            _nudInflateZUp.Value = 400;
            _nudInflateZDn.Value = 100;
            _chkVariants.Checked = true;
            _nudMaxVariants.Value = 13;
            _chkCache.Checked = true;

            // v6.0 节拍 & 焊钳
            _chkWeldOrder.Checked = false;
            _nudLockFirst.Value = 0;
            _nudLockLast.Value = 0;
            _chkMotionJoint.Checked = true;
            _chkGunAxis.Checked = true;
            _nudTargetOpening.Value = 30;
            _chkAdaptiveGun.Checked = true;
            _nudTransitOpening.Value = 60;
            _cboGunDir.SelectedIndex = 0;
            _nudGunMaxOpen.Value = 0;

            Log("参数已恢复默认", LogLevel.Info);
        }

        // ════════════════════════════════════════════════════════════
        //  彩色日志
        // ════════════════════════════════════════════════════════════
        private void Log(string msg) { Log(msg, LogLevel.Info); }

        private void Log(string msg, LogLevel level)
        {
            if (_rtbLog == null) return;
            Color color;
            switch (level)
            {
                case LogLevel.Ok: color = Theme.LogOk; break;
                case LogLevel.Warn: color = Theme.LogWarn; break;
                case LogLevel.Error: color = Theme.LogErr; break;
                case LogLevel.Ps: color = Theme.LogPs; break;
                case LogLevel.Debug: color = Theme.LogDebug; break;
                default: color = Theme.LogInfo; break;
            }
            try
            {
                _rtbLog.SelectionStart = _rtbLog.TextLength;
                _rtbLog.SelectionLength = 0;
                _rtbLog.SelectionColor = color;
                _rtbLog.AppendText(msg + Environment.NewLine);
                _rtbLog.SelectionColor = Theme.LogText;
            }
            catch { }
            Application.DoEvents();
        }

        // ════════════════════════════════════════════════════════════
        //  网格操作 (v4: 只有操作网格保留了)
        // ════════════════════════════════════════════════════════════
        private void AppendOpsSelection()
        {
            try
            {
                TxObjectList items = TxApplication.ActiveSelection.GetItems();
                int added = 0;
                foreach (ITxObject item in items)
                {
                    if (!ContainsWeldLocations(item)) continue; // 一律只加含焊点的节点
                    if (GridContains(_gridOps, item)) continue;
                    _gridOps.AppendObject(item);
                    added++;
                }
                Log(string.Format("已添加 {0} 项 (选择中共 {1} 项)", added, items.Count), LogLevel.Ps);
            }
            catch (Exception ex)
            { Log("[警告] 读取选择失败: " + ex.Message, LogLevel.Warn); }
        }

        private static bool ContainsWeldLocations(ITxObject obj)
        {
            if (obj is TxWeldLocationOperation) return true;
            try
            {
                var container = obj as ITxObjectCollection;
                if (container == null) return false;
                var filter = new TxTypeFilter(typeof(TxWeldLocationOperation));
                TxObjectList descendants = container.GetAllDescendants(filter);
                return descendants.Count > 0;
            }
            catch { return false; }
        }

        private static bool GridContains(TxObjGridCtrl grid, ITxObject obj)
        {
            try
            {
                foreach (ITxObject o in grid.Objects)
                    if (ReferenceEquals(o, obj)) return true;
            }
            catch { }
            return false;
        }

        private void ClearGrid(TxObjGridCtrl grid)
        {
            try
            {
                while (grid.Objects.Count > 0)
                    grid.DeleteRow(0);
            }
            catch (Exception ex) { Log("[警告] 清空网格失败: " + ex.Message, LogLevel.Warn); }
        }

        private static List<ITxObject> ReadGrid(TxObjGridCtrl grid)
        {
            var list = new List<ITxObject>();
            try
            {
                foreach (ITxObject o in grid.Objects)
                    if (o != null) list.Add(o);
            }
            catch { }
            return list;
        }

        // ════════════════════════════════════════════════════════════
        //  运行
        // ════════════════════════════════════════════════════════════
        private void OnRun(object sender, EventArgs e)
        {
            _rtbLog.Clear();
            _btnRun.Enabled = false;
            _stopRequested = false;
            _btnStop.Enabled = true;
            if (_progressBar != null) _progressBar.Value = 0;
            if (_lblStage != null) _lblStage.Text = "启动中...";
            try
            {
                var robot = FindRobot();
                if (robot == null)
                    Log("[提示] 界面未选中机器人 — 将依赖操作的 .Robot 属性解析", LogLevel.Warn);

                var selectedOps = ReadGrid(_gridOps);
                if (selectedOps.Count == 0)
                {
                    Log("[错误] 请先在 PS 树中选中焊接操作，点击\"添加选中\"加入规划范围", LogLevel.Error);
                    return;
                }

                // ---- RRT 自定义路径规划 ----
                // v4: 干涉集由 CollisionSetService 自动检测/复用/新建, UI 不再传自定义障碍
                var planner = new WeldPathPlanner(Log)
                {
                    // 进/出枪 (沿枪 -Z)
                    ApproachRetractDistance = (double)_nudApproach.Value,
                    ApproachRetractMin = (double)_nudApproachMin.Value,
                    UseWorldZForApproach = _chkWorldZ.Checked,
                    GenerateApproachRetract = _chkAppRet.Checked,

                    // 定向搜索 L1 (沿枪 -X 后撤 + 侧向)
                    GunBackoutStep = (double)_nudBackoutStep.Value,
                    GunBackoutMax = (double)_nudBackoutMax.Value,
                    GunMinBackoutForSide = (double)_nudMinBackout.Value,
                    GunSideStep = (double)_nudSideStep.Value,
                    GunSideMax = (double)_nudSideMax.Value,

                    // 动态干涉检查
                    DynamicCheckEnabled = _chkDynamic.Checked,
                    DynamicJointQuantum = (double)_nudJointQuantum.Value,
                    DynamicCartesianQuantum = (double)_nudCartQuantum.Value,
                    MaxSweepSteps = (int)_nudMaxSweep.Value,
                    ConfigJumpThreshold = (double)_nudConfigJump.Value,

                    // RRT
                    RrtStepSize = (double)_nudStep.Value,
                    RrtMaxIterations = (int)_nudIter.Value,
                    RrtGoalBias = (double)_nudGoalBias.Value / 100.0,
                    EdgeCheckResolution = (double)_nudEdgeRes.Value,

                    // 采样空间膨胀
                    SampleBoundsInflateXy = (double)_nudInflateXy.Value,
                    SampleBoundsInflateZUp = (double)_nudInflateZUp.Value,
                    SampleBoundsInflateZDown = (double)_nudInflateZDn.Value,

                    // 姿态变体 / 后处理
                    OrientationVariantsEnabled = _chkVariants.Checked,
                    MaxVariantTries = (int)_nudMaxVariants.Value,
                    CollinearTolerance = (double)_nudCollinearTol.Value,
                    PruneEnabled = _chkPrune.Checked,

                    // v6.0 节拍优化
                    WeldOrderOptEnabled = _chkWeldOrder.Checked,
                    WeldOrderLockFirst = (int)_nudLockFirst.Value,
                    WeldOrderLockLast = (int)_nudLockLast.Value,
                    SetTransitMotionJoint = _chkMotionJoint.Checked,

                    // v6.0 焊钳外部轴
                    GunAxisWriteEnabled = _chkGunAxis.Checked,
                    TargetGunOpening = (double)_nudTargetOpening.Value,
                    AdaptiveGunOpening = _chkAdaptiveGun.Checked,
                    TransitGunOpening = (double)_nudTransitOpening.Value,
                    // v6.4: 0=自动 / 1=正向 / 2=负向  →  0 / +1 / -1
                    GunOpenDirectionOverride =
                        _cboGunDir.SelectedIndex == 1 ? 1 :
                        _cboGunDir.SelectedIndex == 2 ? -1 : 0,
                    GunMaxOpeningOverride = (double)_nudGunMaxOpen.Value,

                    // v6.3 性能
                    QueryCacheEnabled = _chkCache.Checked
                };
                planner.IsCancelled = delegate { return _stopRequested; };

                // v6.3 分阶段进度
                _runStart = DateTime.Now;
                var prog = new PlanningProgress();
                prog.OnProgress = delegate (PlanStage st, double frac, string text)
                {
                    // 规划全程在主线程 (SDK 非线程安全) → 直接赋值, 无需 Invoke
                    if (_progressBar != null)
                        _progressBar.Value = Math.Max(0, Math.Min(1000, (int)(frac * 1000)));

                    if (_lblStage != null)
                    {
                        var el = DateTime.Now - _runStart;
                        string eta = "";
                        if (frac > 0.02)
                        {
                            var total = TimeSpan.FromSeconds(el.TotalSeconds / frac);
                            var left = total - el;
                            if (left.TotalSeconds > 0)
                                eta = string.Format("  剩 ~{0:mm\\:ss}", left);
                        }
                        _lblStage.Text = string.Format("{0}  {1:P0}{2}",
                            text ?? PlanningProgress.StageName(st), frac, eta);
                    }
                    Application.DoEvents();   // 保活 UI + 响应停止按钮
                };
                planner.Progress = prog;

                var report = planner.ExecuteForOperations(robot, selectedOps);
                PrintReport(report);
            }
            catch (OperationCanceledException)
            { Log("[中止] 用户停止了规划 — 已插入的Via保留", LogLevel.Warn); }
            catch (Exception ex)
            { Log("[严重错误] " + ex.Message, LogLevel.Error); }
            finally
            {
                _btnRun.Enabled = true;
                _btnStop.Enabled = false;

                // v6.3 进度收尾
                if (_progressBar != null)
                    _progressBar.Value = _stopRequested ? _progressBar.Value : 1000;
                if (_lblStage != null)
                {
                    var el = DateTime.Now - _runStart;
                    _lblStage.Text = _stopRequested
                        ? string.Format("已中止  (耗时 {0:mm\\:ss})", el)
                        : string.Format("完成  (耗时 {0:mm\\:ss})", el);
                }

                try { TxApplication.RefreshDisplay(); } catch { }
            }
        }

        private void PrintReport(PlanningReport report)
        {
            foreach (var w in report.Warnings) Log("[!] " + w, LogLevel.Warn);
            Log(string.Format(
                "\n完成: {0} 个操作 / {1} 个焊点 / 插入 {2} 个Via / 耗时 {3:F1}s / 碰撞查询 {4} 次",
                report.OperationCount, report.WeldCount, report.InsertedViaCount,
                report.Elapsed.TotalSeconds, report.CollisionQueries), LogLevel.Ok);
            if (report.FailedSegments > 0)
                Log(string.Format("失败段: {0} (需人工处理)", report.FailedSegments), LogLevel.Warn);
            if (report.ClearanceSegments > 0)
                Log(string.Format("净空绕行段: {0} (端点干涉已绕行, 需确认工艺位置)", report.ClearanceSegments), LogLevel.Warn);
        }

        private TxRobot FindRobot()
        {
            try
            {
                TxObjectList items = TxApplication.ActiveSelection.GetItems();
                foreach (ITxObject item in items)
                {
                    var r = item as TxRobot;
                    if (r != null) return r;
                }
            }
            catch { }

            try
            {
                var filter = new TxTypeFilter(typeof(TxRobot));
                TxObjectList robots = TxApplication.ActiveDocument
                    .PhysicalRoot.GetAllDescendants(filter);
                foreach (ITxObject r in robots)
                {
                    var robot = r as TxRobot;
                    if (robot != null) return robot;
                }
            }
            catch { }
            return null;
        }

        // ════════════════════════════════════════════════════════════
        //  干涉集: 自动创建 (v4.1)
        //  流程: 从操作网格反查每台机器人 (op.Robot 与规划器完全一致) →
        //        引用去重 → 逐台弹窗决策 → 创建/复用/取消
        //
        //  关键修正 (v4.1): 曾用 FindRobot() 拿"场景第一台/选中的机器人",
        //  在多机器人或同名机器人场景下会与操作实际绑定的机器人不匹配,
        //  导致自动创建的干涉集与规划路径无关联。改为按操作反查即可对齐。
        // ════════════════════════════════════════════════════════════
        private void OnAutoCreateCollisionSet(object sender, EventArgs e)
        {
            try
            {
                _btnAutoCollisionSet.Enabled = false;

                // ---- 1. 优先从"焊接操作"网格反查 (op.Robot 与规划器完全同源) ----
                var selectedOps = ReadGrid(_gridOps);
                var robots = ResolveRobotsFromOps(selectedOps, out int unresolvedCount);

                if (robots.Count == 0)
                {
                    // 网格空 或 操作未绑机器人 → 请求确认后可回退到 FindRobot()
                    string prompt;
                    if (selectedOps.Count == 0)
                    {
                        prompt =
                            "尚未添加焊接操作。\n\n" +
                            "推荐做法: 先在\"焊接操作\"卡片\"添加选中\", 再点击本按钮 —\n" +
                            "这样干涉集会精确匹配每个操作绑定的机器人 (op.Robot),\n" +
                            "多机器人 / 同名机器人场景下不会错位。\n\n" +
                            "如仍要使用场景/选中的机器人作为回退, 请点\"是\"。";
                    }
                    else
                    {
                        prompt = string.Format(
                            "已添加 {0} 个操作, 但均无法解析出 op.Robot。\n\n" +
                            "可能原因: 操作未绑定机器人 (右键 → Assign Robot)。\n\n" +
                            "是否回退到场景/选中的机器人?", selectedOps.Count);
                    }

                    var wantFallback = MessageBox.Show(prompt, "自动创建干涉集",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (wantFallback != DialogResult.Yes)
                    {
                        SetCsStatus("状态: 已取消 (建议先添加操作)",
                            Color.FromArgb(150, 100, 0));
                        return;
                    }

                    var fallback = FindRobot();
                    if (fallback == null)
                    {
                        MessageBox.Show(
                            "场景与选择中均未找到机器人, 无法创建干涉集。",
                            "自动创建干涉集",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        SetCsStatus("状态: 未找到机器人", Color.FromArgb(160, 40, 40));
                        return;
                    }
                    Log("========== 自动创建干涉集 (回退模式) ==========", LogLevel.Ok);
                    Log("  [警告] 未从操作反查 — 使用回退机器人, 请务必核对与操作是否一致",
                        LogLevel.Warn);
                    robots.Add(fallback);
                }
                else
                {
                    Log("========== 自动创建干涉集 ==========", LogLevel.Ok);
                    Log(string.Format(
                        "  从 {0} 个操作反查到 {1} 台机器人 (引用去重{2})",
                        selectedOps.Count, robots.Count,
                        unresolvedCount > 0
                            ? string.Format(", {0} 个未绑机器人已跳过", unresolvedCount)
                            : ""), LogLevel.Info);
                }

                // ---- 2. 逐台机器人处理 ----
                int createdOk = 0, reused = 0, cancelled = 0, failed = 0;
                foreach (var robot in robots)
                {
                    LogRobotIdentityBrief(robot);
                    switch (CreateForRobot(robot))
                    {
                        case CreateOutcome.Created:   createdOk++; break;
                        case CreateOutcome.Reused:    reused++;    break;
                        case CreateOutcome.Cancelled: cancelled++; break;
                        default:                      failed++;    break;
                    }
                }

                // ---- 3. 汇总状态标签 ----
                var parts = new List<string>();
                if (createdOk > 0) parts.Add(createdOk + " 新建");
                if (reused > 0) parts.Add(reused + " 复用");
                if (cancelled > 0) parts.Add(cancelled + " 取消");
                if (failed > 0) parts.Add(failed + " 失败");
                Color statusColor =
                    failed > 0 ? Color.FromArgb(160, 40, 40) :
                    createdOk > 0 ? Color.FromArgb(0, 120, 50) :
                    reused > 0 ? Color.FromArgb(0, 100, 167) :
                                 Color.FromArgb(150, 100, 0);
                SetCsStatus(
                    string.Format("状态: {0} 台机器人 ({1})",
                        robots.Count,
                        parts.Count > 0 ? string.Join(", ", parts) : "无变化"),
                    statusColor);

                try { TxApplication.RefreshDisplay(); } catch { }
            }
            catch (Exception ex)
            {
                Log("[严重错误] 自动创建干涉集: " + ex.Message, LogLevel.Error);
                SetCsStatus("状态: 异常 — 见日志", Color.FromArgb(160, 40, 40));
            }
            finally
            {
                _btnAutoCollisionSet.Enabled = true;
            }
        }

        /// <summary>单机器人处理结果</summary>
        private enum CreateOutcome { Created, Reused, Cancelled, Failed }

        /// <summary>
        /// 对单台机器人: 检测已有 → 弹窗决策 → 创建/复用/取消。
        /// 全程 KeepPairOnDispose=true 保证碰撞对留在 CollisionRoot。
        /// </summary>
        private CreateOutcome CreateForRobot(TxRobot robot)
        {
            bool forceNew = false;
            string existingName;
            if (CollisionSetService.TryFindExistingPairName(robot, out existingName, s => Log(s, LogLevel.Info)))
            {
                var res = MessageBox.Show(
                    string.Format(
                        "机器人 '{0}' (HashCode={1}) 检测到已有干涉集:\n\n" +
                        "    {2}\n\n" +
                        "是否继续新建?\n\n" +
                        "  是 = 强制新建 (与已有并存)\n" +
                        "  否 = 直接复用已有\n" +
                        "  取消 = 什么也不做",
                        robot.Name, robot.GetHashCode(), existingName),
                    "已存在干涉集",
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (res == DialogResult.Cancel)
                {
                    Log("  已取消 — 该机器人干涉集未做任何更改", LogLevel.Warn);
                    return CreateOutcome.Cancelled;
                }
                if (res == DialogResult.No)
                {
                    Log(string.Format("  复用已有干涉集: {0}", existingName), LogLevel.Info);
                    return CreateOutcome.Reused;
                }
                forceNew = true;
                Log("  用户确认新建 — 强制创建 (与已有并存)", LogLevel.Info);
            }

            using (var cs = CollisionSetService.CreateRobotVsWorld(
                robot, null, null, s => Log(s, LogLevel.Info), forceNew))
            {
                cs.KeepPairOnDispose = true;
                if (cs.IsReady)
                {
                    Log(string.Format("  [✓] 干涉集就绪 (检测方 {0} 项 / 障碍方 {1} 项)",
                        cs.CheckObjectCount, cs.ObstacleObjectCount), LogLevel.Ok);
                    return CreateOutcome.Created;
                }
                Log("  [✗] 干涉集创建失败 — 详见上方日志", LogLevel.Error);
                return CreateOutcome.Failed;
            }
        }

        /// <summary>
        /// 从操作列表反查各自绑定的机器人 (op.Robot → 首焊点.Robot),
        /// 结果按引用去重 (同名不同实例视为不同机器人; 与 RobotBaseChecker 同款约定)。
        /// </summary>
        private List<TxRobot> ResolveRobotsFromOps(
            List<ITxObject> ops, out int unresolvedCount)
        {
            var result = new List<TxRobot>();
            unresolvedCount = 0;
            foreach (var op in ops)
            {
                var r = ResolveRobotOf(op);
                if (r == null) { unresolvedCount++; continue; }
                bool dup = false;
                for (int i = 0; i < result.Count; i++)
                    if (ReferenceEquals(result[i], r)) { dup = true; break; }
                if (!dup) result.Add(r);
            }
            return result;
        }

        /// <summary>
        /// 与 WeldPathPlanner.ResolveRobotOf 完全一致的解析:
        /// op.Robot 优先, 缺失则读操作树内首个焊点的 .Robot。
        /// static 避免与 Planner 内 private 实例方法耦合, 但语义严格对齐。
        /// </summary>
        private static TxRobot ResolveRobotOf(ITxObject op)
        {
            try
            {
                var r = ((dynamic)op).Robot as TxRobot;
                if (r != null) return r;
            }
            catch { }

            try
            {
                // op 本身就是焊点位置
                var self = op as TxWeldLocationOperation;
                if (self != null)
                {
                    var r = ((dynamic)self).Robot as TxRobot;
                    if (r != null) return r;
                }
                // op 是容器: 读首个焊点的 .Robot
                var container = op as ITxObjectCollection;
                if (container != null)
                {
                    var filter = new TxTypeFilter(typeof(TxWeldLocationOperation));
                    TxObjectList kids = container.GetAllDescendants(filter);
                    foreach (ITxObject k in kids)
                    {
                        try
                        {
                            var r = ((dynamic)k).Robot as TxRobot;
                            if (r != null) return r;
                        }
                        catch { }
                    }
                }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 打印机器人身份 (名称 + HashCode + 基座位置) —
        /// 与 WeldPathPlanner.LogRobotIdentity 一致, 便于用户在多同名机器人
        /// 场景下比对"按钮建的这台"是否就是"规划器要跑的那台"。
        /// </summary>
        private void LogRobotIdentityBrief(TxRobot robot)
        {
            string basePos = "?";
            try
            {
                var t = ((ITxLocatableObject)robot).AbsoluteLocation.Translation;
                basePos = string.Format("({0:F0},{1:F0},{2:F0})", t.X, t.Y, t.Z);
            }
            catch { }
            Log(string.Format("---- 机器人: '{0}' HashCode={1} 基座{2} ----",
                robot.Name, robot.GetHashCode(), basePos), LogLevel.Ok);
        }

        /// <summary>更新干涉集状态标签的文本与颜色 (深色以适配灰色卡片背景)</summary>
        private void SetCsStatus(string text, Color color)
        {
            if (_lblCollisionSetStatus == null) return;
            _lblCollisionSetStatus.Text = text;
            _lblCollisionSetStatus.ForeColor = color;
        }
    }
}
