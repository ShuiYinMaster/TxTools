// ============================================================================
// LocationGeometry.cs
//
// 点位枚举 / 反查 / 变换矩阵 / 平移分量提取 / 矩阵克隆。
// 来自原文件 EnumerateLocations / FindLocationInDoc / GetLocationTransform /
//          ExtractTranslation 的整体迁移。
// ============================================================================
using System;
using System.Collections.Generic;
using Tecnomatix.Engineering;
using TxTools.RobotReachabilityChecker.Diagnostics;

namespace TxTools.RobotReachabilityChecker.Services
{
    public static class LocationEnumerator
    {
        /// <summary>枚举操作下的所有路径点。复合操作会展开。</summary>
        public static List<ITxRoboticLocationOperation> EnumerateLocations(
            ITxObject operation, ILogger log = null)
        {
            log = log ?? NullLogger.Instance;
            var list = new List<ITxRoboticLocationOperation>();
            if (operation == null) return list;

            // 路径1：ITxCompoundOperation.GetAllDescendants
            if (operation is ITxCompoundOperation comp)
            {
                try
                {
                    var objs = comp.GetAllDescendants(
                        new TxTypeFilter(typeof(ITxRoboticLocationOperation)));
                    foreach (ITxObject o in objs)
                        if (o is ITxRoboticLocationOperation l) list.Add(l);
                }
                catch (Exception ex) { log.Log($"  GetAllDescendants 异常: {ex.Message}", "WARN"); }
            }

            // 路径2：dynamic GetAllDescendants（兼容 TxRoboticOperation 非接口）
            if (list.Count == 0)
            {
                try
                {
                    dynamic dop = operation;
                    TxObjectList objs = dop.GetAllDescendants(
                        new TxTypeFilter(typeof(ITxRoboticLocationOperation)));
                    foreach (ITxObject o in objs)
                        if (o is ITxRoboticLocationOperation l) list.Add(l);
                }
                catch { }
            }

            // 路径3：操作本身就是一个 LocationOperation
            if (list.Count == 0 && operation is ITxRoboticLocationOperation self)
                list.Add(self);

            return list;
        }

        /// <summary>按 (操作名, 点位名) 反查点位对象。</summary>
        public static ITxObject FindLocationInDoc(
            TxDocument doc, string opName, string pointName, ILogger log = null)
        {
            log = log ?? NullLogger.Instance;
            if (doc == null || string.IsNullOrEmpty(pointName)) return null;

            try
            {
                // 优先：按操作名定位 → 在其下找点位名
                if (!string.IsNullOrEmpty(opName))
                {
                    var allOps = doc.OperationRoot.GetAllDescendants(
                        new TxTypeFilter(typeof(ITxObject)));
                    foreach (ITxObject obj in allOps)
                    {
                        if (obj.Name != opName) continue;
                        var locs = EnumerateLocations(obj, log);
                        foreach (var l in locs)
                            if ((l as ITxObject)?.Name == pointName) return l as ITxObject;
                    }
                }
                // 兜底：全文档找同名 ITxRoboticLocationOperation
                var allDesc = doc.OperationRoot.GetAllDescendants(
                    new TxTypeFilter(typeof(ITxObject)));
                foreach (ITxObject obj in allDesc)
                {
                    if (!(obj is ITxRoboticLocationOperation)) continue;
                    if (obj.Name == pointName) return obj;
                }
            }
            catch (Exception ex) { log.Log($"FindLocationInDoc 异常: {ex.Message}", "ERR"); }
            return null;
        }
    }

    /// <summary>变换矩阵 / 平移提取 / 矩阵克隆等几何辅助。</summary>
    public static class LocationGeometry
    {
        /// <summary>
        /// 获取点位世界坐标变换矩阵（IK 必须用世界坐标，相对坐标会全部无解）。
        /// 优先级：AbsoluteLocation > AbsoluteFrame > LocationInWorld > Location > Frame
        /// </summary>
        public static TxTransformation GetLocationTransform(
            ITxRoboticLocationOperation loc, ILogger log = null)
        {
            log = log ?? NullLogger.Instance;
            if (loc == null) return null;

            // 策略1：AbsoluteLocation（最标准）
            try { dynamic d = loc; var v = d.AbsoluteLocation; if (v is TxTransformation t && t != null) return t; } catch { }
            // 策略2：AbsoluteFrame
            try { dynamic d = loc; var v = d.AbsoluteFrame; if (v is TxTransformation t && t != null) return t; } catch { }
            // 策略3：LocationInWorld
            try { dynamic d = loc; var v = d.LocationInWorld; if (v is TxTransformation t && t != null) return t; } catch { }
            // 策略4：ITxLocatableObject.AbsoluteLocation
            try
            {
                if (loc is ITxLocatableObject lobj)
                {
                    TxTransformation tx = lobj.AbsoluteLocation;
                    if (tx != null) return tx;
                }
            }
            catch { }
            // 策略5/6：Location / Frame（相对坐标，仅备用）
            try { dynamic d = loc; var v = d.Location; if (v is TxTransformation t && t != null) { log.Log("    Location(相对坐标，IK可能无解)", "WARN"); return t; } } catch { }
            try { dynamic d = loc; var v = d.Frame; if (v is TxTransformation t && t != null) { log.Log("    Frame(相对坐标)", "WARN"); return t; } } catch { }
            return null;
        }

        /// <summary>从 TxTransformation 提取平移分量 [X, Y, Z]（mm）。</summary>
        public static double[] ExtractTranslation(TxTransformation tx)
        {
            if (tx == null) return null;
            dynamic dt = tx;
            try { return new double[] { Convert.ToDouble(dt[0, 3]), Convert.ToDouble(dt[1, 3]), Convert.ToDouble(dt[2, 3]) }; } catch { }
            try { return new double[] { Convert.ToDouble(dt.X), Convert.ToDouble(dt.Y), Convert.ToDouble(dt.Z) }; } catch { }
            try { dynamic t = dt.Translation; return new double[] { Convert.ToDouble(t.X), Convert.ToDouble(t.Y), Convert.ToDouble(t.Z) }; } catch { }
            try { dynamic t = dt.GetTranslation(); return new double[] { Convert.ToDouble(t.X), Convert.ToDouble(t.Y), Convert.ToDouble(t.Z) }; } catch { }
            return null;
        }

        /// <summary>克隆变换矩阵（三策略链：拷贝构造 → Clone() → 手工 4×4 复制）。</summary>
        public static TxTransformation CloneTransformation(TxTransformation baseTx)
        {
            try { return new TxTransformation(baseTx); } catch { }
            try { dynamic db = baseTx; return db.Clone() as TxTransformation; } catch { }
            try
            {
                var tx = new TxTransformation();
                dynamic src = baseTx;
                dynamic dst = tx;
                for (int r = 0; r < 4; r++)
                    for (int c = 0; c < 4; c++)
                        try { dst[r, c] = Convert.ToDouble(src[r, c]); } catch { }
                return tx;
            }
            catch { return null; }
        }
    }
}
