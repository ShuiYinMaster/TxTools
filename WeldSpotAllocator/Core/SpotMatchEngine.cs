// SpotMatchEngine.cs  —  C# 7.3
// 匹配条件（对齐第三方规格）：两个独立开关 焊点距离匹配 / 焊点名匹配
//   仅距离：距离<阈值即匹配，不看名      仅名：同名即匹配，不看距离
//   都勾：同名且距离<阈值（更严）         都不勾：退化为仅距离
//
//   A 位置更新：操作配对（数量→名→质心）+ 操作内焊点匹配（双开关）；写 ref ← target 坐标
//   B 新焊点分配：待分配操作集展平成焊点池；遍历每条参考轨迹的焊点 → 池里找匹配
//   C 对称分配  ：同 B，但参考焊点先关于对称中心 XZ 镜像后再比距离
//   B/C 输出：每条参考轨迹一个 OpMatch（写入层据此 Paste 出 _Mapped 轨迹）

using System;
using System.Collections.Generic;
using System.Linq;

namespace MyPlugin.WeldSpotAllocator
{
    public sealed class NameOptions
    {
        public bool IgnoreCase = true;
        public bool StripDigits = false;
        public string StripPrefix = null;
        public string StripSuffix = null;

        public string Norm(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            string t = s.Trim();
            if (!string.IsNullOrEmpty(StripPrefix) && t.StartsWith(StripPrefix)) t = t.Substring(StripPrefix.Length);
            if (!string.IsNullOrEmpty(StripSuffix) && t.EndsWith(StripSuffix)) t = t.Substring(0, t.Length - StripSuffix.Length);
            if (StripDigits) t = t.TrimEnd('0','1','2','3','4','5','6','7','8','9','_','-',' ');
            if (IgnoreCase) t = t.ToUpperInvariant();
            return t;
        }
        public bool Eq(string a, string b) => Norm(a) == Norm(b) && Norm(a).Length > 0;
    }

    public sealed class MatchSettings
    {
        public NameOptions Names = new NameOptions();
        public bool ByDistance = true;
        public bool ByName = false;
        public double MaxDistMm = 5.0;
        public bool CountStrict = true;
        public int CountTolerance = 0;
        public bool DiffFrame = false;      // 目标点与参考点处于不同车件参考系（目标车身须在世界原点）

        public bool BothOn => ByDistance && ByName;
        public bool DistOnly => ByDistance && !ByName;
        public bool NameOnly => !ByDistance && ByName;
    }

    public static class SpotMatchEngine
    {
        // ── 入口 ──────────────────────────────────────────────────────────
        public static AllocPlan Build(AllocMode mode,
            IList<OpData> refOps, IList<OpData> targetOps,
            MatchSettings cfg, FlipAxis flip,
            bool copyVias, bool copyRot, bool copyParams, bool consume, Action<string> log)
        {
            log = log ?? (s => { });
            var plan = new AllocPlan
            {
                Mode = mode, Flip = flip, CopyVias = copyVias, CopyRotation = copyRot,
                CopyParams = copyParams, ConsumeTargets = consume
            };

            switch (mode)
            {
                case AllocMode.UpdatePosition: BuildUpdate(plan, refOps, targetOps, cfg, log); break;
                case AllocMode.NewSpot:        BuildAssign(plan, refOps, targetOps, cfg, log); break;
                case AllocMode.Symmetric:      BuildAssign(plan, refOps, targetOps, cfg, log); break;
            }
            return plan;
        }

        // ── 双开关命中判定 ────────────────────────────────────────────────
        // rPos：参考点用于比距离的位置（C 模式为镜像后位置）
        private static bool TryMatch(SpotData r, double[] rPos, SpotData t, MatchSettings cfg, out double dist, out MatchBy by)
        {
            dist = Math.Sqrt(SymmetryMath.Dist2(rPos, t.Position));
            bool nameEq = cfg.Names.Eq(r.Name, t.Name);
            bool distOk = dist <= cfg.MaxDistMm;
            by = MatchBy.None;

            if (cfg.BothOn) { if (nameEq && distOk) { by = MatchBy.Name; return true; } return false; }
            if (cfg.NameOnly) { if (nameEq) { by = MatchBy.Name; dist = 0; return true; } return false; }
            if (distOk) { by = MatchBy.Position; return true; }   // DistOnly / 都不勾
            return false;
        }

        // ── A 位置更新 ────────────────────────────────────────────────────
        private static void BuildUpdate(AllocPlan plan, IList<OpData> refOps, IList<OpData> targetOps, MatchSettings cfg, Action<string> log)
        {
            if (refOps == null || targetOps == null) { plan.Warnings.Add("参考/目标操作集为空"); return; }
            var usedT = new HashSet<OpData>();

            foreach (var rOp in refOps)
            {
                OpData best = PickOpPartner(rOp, targetOps, usedT, cfg);
                if (best == null) { plan.Warnings.Add($"操作[{rOp.Name}] 未找到配对（数量 {rOp.SpotCount}）"); continue; }
                usedT.Add(best);

                var om = new OpMatch { RefOp = rOp, TargetOp = best };
                MatchSpotsWithin(rOp.Spots, best.Spots, cfg, om);
                plan.OpMatches.Add(om);
            }
            foreach (var t in targetOps) if (!usedT.Contains(t))
                plan.Warnings.Add($"新版操作[{t.Name}] 未被匹配（数量 {t.SpotCount}）");
        }

        private static OpData PickOpPartner(OpData rOp, IList<OpData> targets, HashSet<OpData> used, MatchSettings cfg)
        {
            IEnumerable<OpData> pool = targets.Where(t => !used.Contains(t));
            pool = cfg.CountStrict ? pool.Where(t => t.SpotCount == rOp.SpotCount)
                                   : pool.Where(t => Math.Abs(t.SpotCount - rOp.SpotCount) <= cfg.CountTolerance);
            var cand = pool.ToList();
            if (cand.Count == 0) return null;
            if (cand.Count == 1) return cand[0];

            var byName = cand.FirstOrDefault(t => cfg.Names.Eq(t.Name, rOp.Name));
            if (byName != null) return byName;

            var rc = rOp.Centroid();
            if (rc == null) return cand[0];
            OpData best = null; double bd = double.MaxValue;
            foreach (var t in cand) { var tc = t.Centroid(); if (tc == null) continue; double d = SymmetryMath.Dist2(rc, tc); if (d < bd) { bd = d; best = t; } }
            return best ?? cand[0];
        }

        // 操作内焊点匹配（双开关），方向：ref ← target。
        private static void MatchSpotsWithin(IList<SpotData> refs, IList<SpotData> targets, MatchSettings cfg, OpMatch om)
        {
            var usedT = new HashSet<SpotData>();
            foreach (var r in refs)
            {
                SpotData best = null; double bd = double.MaxValue; MatchBy bby = MatchBy.None;
                foreach (var t in targets)
                {
                    if (usedT.Contains(t)) continue;
                    if (TryMatch(r, r.Position, t, cfg, out double d, out MatchBy by) && d < bd) { bd = d; best = t; bby = by; }
                }
                if (best != null) { usedT.Add(best); om.Matches.Add(new SpotMatch { Ref = r, Target = best, Dist = bd, By = bby }); }
            }
        }

        // ── B / C 分配：每条参考轨迹的焊点 → 待分配焊点池 ────────────────────
        private static void BuildAssign(AllocPlan plan, IList<OpData> refOps, IList<OpData> targetOps, MatchSettings cfg, Action<string> log)
        {
            var pool = new List<SpotData>();
            if (targetOps != null) foreach (var op in targetOps) pool.AddRange(op.Spots);
            if (refOps == null || pool.Count == 0) { plan.Warnings.Add("参考操作集 / 待分配焊点为空"); return; }

            bool needCenter = plan.Mode == AllocMode.Symmetric || cfg.DiffFrame; // 需要车件坐标
            var usedT = new HashSet<SpotData>();
            foreach (var refOp in refOps)
            {
                var om = new OpMatch { RefOp = refOp };
                foreach (var r in refOp.Spots)
                {
                    if (needCenter && r.SymCenter == null) { plan.Warnings.Add($"参考焊点[{r.Name}] 无车件坐标(对称中心/分身)，跳过"); continue; }
                    double[] refRes = SymmetryMath.ResolveRef(r.Matrix, r.SymCenter, plan.Mode, cfg.DiffFrame, plan.Flip);
                    double[] rPos = { refRes[3], refRes[7], refRes[11] };

                    SpotData best = null; double bd = double.MaxValue; MatchBy bby = MatchBy.None;
                    foreach (var t in pool)
                    {
                        if (usedT.Contains(t)) continue;
                        if (TryMatch(r, rPos, t, cfg, out double d, out MatchBy by) && d < bd) { bd = d; best = t; bby = by; }
                    }
                    if (best != null)
                    {
                        usedT.Add(best);
                        om.Matches.Add(new SpotMatch { Ref = r, Target = best, Dist = bd, By = bby, Mirrored = needCenter, RefMirror = needCenter ? refRes : null });
                    }
                }
                plan.OpMatches.Add(om);
            }
            foreach (var t in pool) if (!usedT.Contains(t)) plan.Warnings.Add($"待分配焊点[{t.Name}] 未匹配到参考焊点");
            log($"[Match] 命中 {plan.TotalMatches} 点，参考轨迹 {plan.OpMatches.Count} 条");
        }
    }
}
