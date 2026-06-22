// SpotWriter.cs  —  C# 7.3
// 写入层。A 走原地更新；B/C 走 Paste 路线（对齐第三方“为每条参考轨迹建对应新轨迹”）。
//
// B/C 流程（每条参考轨迹）：
//   1. parent.Paste(参考轨迹) → _Mapped 副本（via/焊点/机器人工具绑定自动带）
//   2. 改名 + 复读结构，按 index 把参考焊点位对到 _Mapped 焊点位
//   3. 有匹配的焊点位 → 写成待分配焊点的位置（+ 可选复制 ref[镜像]姿态、轨迹参数）
//                       → 可选“挪”：删除待分配轨迹里被用掉的焊点
//   4. via（C 模式）按对称中心镜像位置
//   5. 解绑零件分身（文档：新轨迹不带分身，需用户自绑——此处尽力，失败仅提示）
//   全程包在一个 Undo 事务里，失败回滚。
//
// runtime 待验证点（首次运行看日志）：Paste 行为、GetParent、Delete、Undo 事务方法名。

using System;
using System.Collections;
using System.Linq;
using Tecnomatix.Engineering;
using MyPlugin.ExportGun;

namespace MyPlugin.WeldSpotAllocator
{
    public sealed class WriteReport
    {
        public int Updated, Created, ViasMirrored, Consumed, Failed;
        public readonly System.Collections.Generic.List<string> Errors = new System.Collections.Generic.List<string>();
        public override string ToString()
            => $"更新 {Updated} · 挪入新轨迹 {Consumed} · 删参考占位 {Created} · 镜像via {ViasMirrored} · 失败 {Failed}";
    }

    public static class SpotWriter
    {
        public static WriteReport ApplyPlan(AllocPlan plan, bool aPosOnly, Action<string> log)
        {
            log = log ?? (s => { });
            return plan.Mode == AllocMode.UpdatePosition ? ApplyUpdate(plan, aPosOnly, log)
                                                         : ApplyAssignByMove(plan, log);
        }

        // ════════════════════════════════════════════════════════════
        //  A 位置更新（原地写 ref ← target）
        // ════════════════════════════════════════════════════════════
        private static WriteReport ApplyUpdate(AllocPlan plan, bool posOnly, Action<string> log)
        {
            var rep = new WriteReport();
            object um = OpenUndo("焊点位置更新", log);
            try
            {
                foreach (var om in plan.OpMatches)
                    foreach (var mt in om.Matches)
                    {
                        try
                        {
                            var wp = mt.Ref.Raw as TxWeldPoint;
                            if (wp == null) { Fail(rep, $"[{mt.Ref.Name}] 非 TxWeldPoint"); continue; }
                            double[] m = posOnly ? MergePosKeepRot(wp, mt.Target.Position) : mt.Target.Matrix;
                            wp.AbsoluteLocation = PsReader.ArrToTxPublic(m);
                            rep.Updated++;
                        }
                        catch (Exception ex) { Fail(rep, $"[{mt.Ref?.Name}] {ex.Message}"); }
                    }
                CommitUndo(um, log);
            }
            catch (Exception ex) { AbortUndo(um, log); Fail(rep, "事务异常：" + ex.Message); }
            log("[Writer] " + rep);
            return rep;
        }

        // ════════════════════════════════════════════════════════════
        //  B / C  Paste 路线
        // ════════════════════════════════════════════════════════════
        // ════════════════════════════════════════════════════════════
        //  B / C  移动路线：Paste 参考轨迹拿结构 → 把待分配焊点 location 移入 _Mapped → 删参考占位
        //    待分配焊点 MFG 全程被引用，不会变灰；参考 MFG 因原轨迹仍引用也不灰。
        // ════════════════════════════════════════════════════════════
        private static WriteReport ApplyAssignByMove(AllocPlan plan, Action<string> log)
        {
            var rep = new WriteReport();
            object um = OpenUndo("焊点分配", log);
            try
            {
                foreach (var om in plan.OpMatches)
                {
                    if (om.Matches.Count == 0) { log($"[Writer] [{om.RefOp.Name}] 无匹配，跳过"); continue; }
                    ITxObject refOp = om.RefOp.Raw;
                    ITxObject parent = GetParent(refOp, log);
                    if (parent == null) { Fail(rep, $"[{om.RefOp.Name}] 取不到父节点，无法 Paste"); continue; }

                    // 1) 复制参考轨迹结构（via / 参数 / 顺序 / 参考焊点占位）
                    ITxObject mapped = PasteInto(parent, refOp, log);
                    if (mapped == null) { Fail(rep, $"[{om.RefOp.Name}] Paste 失败"); continue; }
                    SetName(mapped, om.RefOp.Name + "_Mapped");

                    var md = SpotReader.ReadOneWeldOp(mapped, log);   // _Mapped 里参考焊点副本(占位)
                    int n = Math.Min(om.RefOp.Spots.Count, md.Spots.Count);
                    for (int i = 0; i < n; i++)
                    {
                        var refSpot = om.RefOp.Spots[i];
                        var placeholder = md.Spots[i];                // 参考占位副本(用完即删)
                        var match = om.Matches.FirstOrDefault(x => ReferenceEquals(x.Ref, refSpot));
                        if (match == null) continue;                  // 该参考点无匹配 → 保留占位

                        try
                        {
                            ITxObject tgtLoc = match.Target.LocOp ?? match.Target.Raw;  // 待分配焊点 location
                            ITxObject anchor = placeholder.LocOp;                       // 移到占位之后

                            // 2) 把待分配焊点 location 移入 _Mapped（= 树里拖动），紧跟占位之后
                            if (!MoveLocationInto(mapped, tgtLoc, anchor, log))
                            { Fail(rep, $"[{match.Target?.Name}] 移入 _Mapped 失败"); continue; }
                            rep.Consumed++;

                            // 3) 可选：姿态借参考（保持待分配焊点自身位置）
                            if (plan.CopyRotation)
                            {
                                var wp = match.Target.Raw as TxWeldPoint;
                                if (wp != null) wp.AbsoluteLocation = PsReader.ArrToTxPublic(ComposeAssignMatrix(plan, match));
                            }
                            // 4) 可选：复制轨迹参数到移入的 location
                            if (plan.CopyParams) CopyTrajectoryParams(refSpot.LocOp, tgtLoc, log);

                            // 5) 删参考占位副本（参考 MFG 原轨迹仍引用，不会灰）
                            if (DeleteObject(placeholder.LocOp ?? placeholder.Raw, log)) rep.Created++;
                        }
                        catch (Exception ex) { Fail(rep, $"[{match.Target?.Name}] {ex.Message}"); }
                    }

                    if (plan.Mode == AllocMode.Symmetric)
                        MirrorViasInPlace(md, om.RefOp, plan.Flip, rep, log);

                    UnbindAppearance(mapped, log);
                }
                CommitUndo(um, log);
            }
            catch (Exception ex) { AbortUndo(um, log); Fail(rep, "事务异常：" + ex.Message); }
            log("[Writer] " + rep);
            return rep;
        }

        // B/C 焊点最终位姿：位置恒用 target；姿态按选项（C 用 ref 镜像姿态）
        private static double[] ComposeAssignMatrix(AllocPlan plan, SpotMatch mt)
        {
            if (!plan.CopyRotation) return mt.Target.Matrix;
            double[] rotSrc = mt.RefMirror ?? mt.Ref.Matrix;   // RefMirror 已含 局部对齐/镜像
            return MergeRotKeepPos(rotSrc, mt.Target.Position);
        }

        // _Mapped 里的 via 关于对称中心镜像（C）。via 自身常无分身，借本轨迹焊点的对称中心。
        private static void MirrorViasInPlace(OpData mapped, OpData refOp, FlipAxis flip, WriteReport rep, Action<string> log)
        {
            double[] center = refOp.Spots.FirstOrDefault(s => s.SymCenter != null)?.SymCenter;
            if (center == null) { log($"[Writer] [{mapped.Name}] 无对称中心，via 不镜像"); return; }
            foreach (var via in mapped.Vias)
            {
                try
                {
                    double[] m = SymmetryMath.MirrorWorld(via.Matrix, center, flip);
                    dynamic d = via.Raw; d.AbsoluteLocation = PsReader.ArrToTxPublic(m);
                    rep.ViasMirrored++;
                }
                catch (Exception ex) { log("[Writer] via 镜像失败：" + ex.Message); }
            }
        }

        // ── 轨迹参数复制（ref → mapped 焊点位）─────────────────────────────
        private static void CopyTrajectoryParams(ITxObject refLoc, ITxObject mapLoc, Action<string> log)
        {
            if (refLoc == null || mapLoc == null) return;
            try
            {
                dynamic rd = refLoc, md = mapLoc;
                ArrayList ps = rd.Parameters as ArrayList;   // ITxRoboticOperation.Parameters
                if (ps == null) return;
                foreach (var p in ps) { try { md.SetParameter(p); } catch { } }
                // 备选（更稳，需 ControllerName）：TxOlpRoboticParametersManager.CopyValue(name, refLoc, mapLoc)
            }
            catch (Exception ex) { log("[Writer] 复制轨迹参数失败：" + ex.Message); }
        }

        // ── Paste / 树操作 ─────────────────────────────────────────────────
        private static readonly TxTypeFilter AnyFilter = new TxTypeFilter(typeof(ITxObject));

        private static ITxObject GetParent(ITxObject o, Action<string> log)
        {
            try { dynamic d = o; var p = d.Parent as ITxObject; if (p != null) return p; } catch { }
            try { dynamic d = o; var p = d.ParentOperation as ITxObject; if (p != null) return p; } catch { }
            // 扫描操作树找“谁的直接子里有 o”，只用已验证的 GetDirectDescendants
            try
            {
                var root = TxApplication.ActiveDocument.OperationRoot as ITxObject;
                var found = ScanForParent(root, o, 0);
                if (found != null) return found;
            }
            catch (Exception ex) { log("[Writer] 扫描父节点异常：" + ex.Message); }
            return null;
        }

        private static ITxObject ScanForParent(ITxObject node, ITxObject target, int depth)
        {
            if (node == null || depth > 30) return null;
            var coll = node as ITxObjectCollection;
            if (coll == null) return null;
            TxObjectList direct;
            try { direct = coll.GetDirectDescendants(AnyFilter); } catch { return null; }
            if (direct == null) return null;
            foreach (ITxObject k in direct) if (IsSame(k, target)) return node;
            foreach (ITxObject k in direct) { var p = ScanForParent(k, target, depth + 1); if (p != null) return p; }
            return null;
        }

        private static bool IsSame(ITxObject a, ITxObject b)
        {
            if (ReferenceEquals(a, b)) return true;
            try { if (a.Equals(b)) return true; } catch { }
            return false;
        }

        // 把一条 location 移入目标操作 mapped，紧跟 after 之后（= 树里拖动跨操作）。
        // 多候选，首次运行看 [Move] 日志确认哪个命中后即可固化。
        private static bool MoveLocationInto(ITxObject mapped, ITxObject loc, ITxObject after, Action<string> log)
        {
            try { dynamic d = mapped; d.AddObjectAfter(loc, after); log("[Move] AddObjectAfter ✓"); return true; } catch (Exception e) { log("[Move] AddObjectAfter×：" + e.Message); }
            try { dynamic d = mapped; d.MoveObjectAfter(loc, after); log("[Move] MoveObjectAfter ✓"); return true; } catch (Exception e) { log("[Move] MoveObjectAfter×：" + e.Message); }
            try { dynamic d = mapped; d.AddObject(loc); try { d.MoveChildAfter(loc, after); } catch { } log("[Move] AddObject(+MoveChildAfter) ✓"); return true; } catch (Exception e) { log("[Move] AddObject×：" + e.Message); }
            try { dynamic d = mapped; var l = new TxObjectList(); l.Add(loc); d.AddObjects(l); log("[Move] AddObjects ✓"); return true; } catch (Exception e) { log("[Move] AddObjects×：" + e.Message); }
            return false;
        }

        private static ITxObject PasteInto(ITxObject parent, ITxObject op, Action<string> log)
        {
            try
            {
                var list = new TxObjectList(); list.Add(op);
                dynamic p = parent;
                var res = p.Paste(list) as TxObjectList;
                if (res != null && res.Count > 0) return res[0] as ITxObject;
                log("[Writer] Paste 返回空");
            }
            catch (Exception ex) { log("[Writer] Paste 异常：" + ex.Message); }
            return null;
        }

        private static void SetName(ITxObject o, string name)
        {
            try { dynamic d = o; d.Name = name; } catch { }
        }

        private static bool DeleteObject(ITxObject o, Action<string> log)
        {
            if (o == null) return false;
            try { dynamic d = o; d.Delete(); return true; } catch { }
            try { dynamic d = o; d.delete(); return true; } catch { }
            try { dynamic d = o; d.Remove(); return true; } catch { }
            log("[Writer] 删除待分配焊点失败（无 Delete/Remove）");
            return false;
        }

        // 解绑零件分身：API 不确定，尽力尝试，失败仅提示（文档允许用户手动绑）
        private static void UnbindAppearance(ITxObject mapped, Action<string> log)
        {
            try { dynamic d = mapped; d.ClearPartAppearances(); return; } catch { }
            try { dynamic d = mapped; d.DetachAll(); return; } catch { }
            log("[Writer] 提示：_Mapped 轨迹未自动解绑零件分身，请按需手动绑定。");
        }

        // ── Undo 事务（cascading 方法名）──────────────────────────────────
        private static object OpenUndo(string name, Action<string> log)
        {
            try
            {
                dynamic um = TxApplication.ActiveUndoManager;
                if (um == null) return null;
                try { um.OpenUndoTransaction(name); return um; } catch { }
                try { um.OpenTransaction(name); return um; } catch { }
                try { um.StartTransaction(name); return um; } catch { }
                try { um.BeginUndoTransaction(name); return um; } catch { }
            }
            catch { }
            log("[Writer] 未开启 Undo 事务（PS 仍有自动撤销，可 Ctrl+Z 回退）");
            return null;
        }
        private static void CommitUndo(object um, Action<string> log)
        {
            if (um == null) return;
            try { dynamic d = um; try { d.CommitUndoTransaction(); return; } catch { } try { d.CommitTransaction(); return; } catch { } try { d.Commit(); return; } catch { } } catch { }
        }
        private static void AbortUndo(object um, Action<string> log)
        {
            if (um == null) return;
            try { dynamic d = um; try { d.AbortUndoTransaction(); return; } catch { } try { d.AbortTransaction(); return; } catch { } try { d.Rollback(); return; } catch { } } catch { }
            log("[Writer] 已尝试回滚事务");
        }

        // ── 矩阵小工具 ────────────────────────────────────────────────────
        private static double[] MergePosKeepRot(TxWeldPoint wp, double[] newPos)
        {
            double[] cur = PsReader.TxToArr(wp.AbsoluteLocation);
            cur[3] = newPos[0]; cur[7] = newPos[1]; cur[11] = newPos[2];
            return cur;
        }
        private static double[] MergeRotKeepPos(double[] rotSrc, double[] pos)
        {
            double[] m = (double[])rotSrc.Clone();
            m[3] = pos[0]; m[7] = pos[1]; m[11] = pos[2];
            return m;
        }

        private static void Fail(WriteReport rep, string msg) { rep.Failed++; rep.Errors.Add(msg); }
    }
}
