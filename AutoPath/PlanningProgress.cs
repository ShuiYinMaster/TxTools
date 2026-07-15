using System;

namespace TxTools.AutoPathPlanner
{
    /// <summary>规划阶段</summary>
    public enum PlanStage
    {
        Init,           // 干涉集 / 碰撞世界 / 自检
        WeldOrder,      // 焊点顺序优化
        Approach,       // 进出枪点生成
        StaticPlan,     // 阶段1 静态规划 (L0/L1/L1.5/L2)
        DynamicRefine,  // 阶段2 逐帧动态精修
        Shortcut,       // 阶段2.5 路径捷径化
        ViaInsert,      // Via 点插入 + 外部轴写入
        Done
    }

    /// <summary>
    /// 规划进度上报 (v6.3)。
    ///
    /// 规划全程跑在 PS 主线程 (SDK 不是线程安全的), 所以进度回调也在主线程 —
    /// UI 更新直接赋值即可, 不需要 Invoke。靠 Application.DoEvents() 保活。
    ///
    /// 两级进度:
    ///   阶段级 (Stage)  — 大阶段, 每个阶段有预设权重
    ///   阶段内 (0..1)   — 该阶段的细粒度进度
    /// 总进度 = 已完成阶段权重和 + 当前阶段权重 × 阶段内进度
    /// </summary>
    public sealed class PlanningProgress
    {
        /// <summary>各阶段的相对权重 (按实测耗时占比调的)</summary>
        private static readonly double[] Weights =
        {
            0.05,  // Init        — 干涉集 + 自检
            0.08,  // WeldOrder   — 焊序优化 (需要每点 IK)
            0.05,  // Approach    — 进出枪点
            0.32,  // StaticPlan  — 静态规划 (L1 动态复验很贵)
            0.35,  // DynamicRefine — 精修 (最贵)
            0.10,  // Shortcut    — 捷径化 (每条捷径都要动态验证)
            0.05,  // ViaInsert   — 插点
            0.00   // Done
        };

        /// <summary>进度回调: (阶段, 总进度 0..1, 描述文本)</summary>
        public Action<PlanStage, double, string> OnProgress;

        private PlanStage _stage = PlanStage.Init;
        private double _inStage;

        /// <summary>操作级进度: 多操作时, 每个操作占总进度的一段</summary>
        private int _opIndex, _opCount = 1;

        public void SetOperationScope(int index, int count)
        {
            _opIndex = Math.Max(0, index);
            _opCount = Math.Max(1, count);
        }

        /// <summary>进入新阶段</summary>
        public void Enter(PlanStage stage, string text = null)
        {
            _stage = stage;
            _inStage = 0;
            Report(text ?? StageName(stage));
        }

        /// <summary>更新当前阶段内进度 (0..1)</summary>
        public void Update(double fraction, string text = null)
        {
            _inStage = Math.Max(0, Math.Min(1, fraction));
            Report(text);
        }

        /// <summary>按 完成数/总数 更新</summary>
        public void Update(int done, int total, string text = null)
        {
            if (total <= 0) return;
            Update((double)done / total, text);
        }

        public void Done()
        {
            _stage = PlanStage.Done;
            _inStage = 1;
            Report("完成");
        }

        private void Report(string text)
        {
            if (OnProgress == null) return;

            // 单操作内的进度
            double acc = 0;
            for (int i = 0; i < (int)_stage && i < Weights.Length; i++)
                acc += Weights[i];
            double w = (int)_stage < Weights.Length ? Weights[(int)_stage] : 0;
            double inOp = Math.Min(1.0, acc + w * _inStage);

            // 折算到多操作的总进度
            double total = (_opIndex + inOp) / _opCount;
            total = Math.Max(0, Math.Min(1, total));

            try { OnProgress(_stage, total, text); }
            catch { }
        }

        public static string StageName(PlanStage s)
        {
            switch (s)
            {
                case PlanStage.Init: return "初始化 (干涉集/自检)";
                case PlanStage.WeldOrder: return "焊点顺序优化";
                case PlanStage.Approach: return "进出枪点生成";
                case PlanStage.StaticPlan: return "静态规划";
                case PlanStage.DynamicRefine: return "动态精修";
                case PlanStage.Shortcut: return "路径捷径化";
                case PlanStage.ViaInsert: return "插入 Via 点";
                case PlanStage.Done: return "完成";
                default: return "";
            }
        }
    }
}
