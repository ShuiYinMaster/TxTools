// PsReader.cs  —  C# 7.3
// Process Simulate 数据读取：操作解析 / 焊点提取 / 焊枪信息 / 参考坐标系 / 焊点外观绑定
//
// [补充] 新增方法（供重构后的 ExportGunForm 原生选择器使用）：
//   - ParsePickedObjectToOperations(dynamic)  → TxPickListener 拾取后解析
//   - WrapAsOperationInfo(dynamic)            → TxObjGridCtrl 插入后包装
//   - GetFrameMatrixFromObject(dynamic)       → TxFrameEditBoxCtrl 坐标提取
//
// [新增] 焊点外观绑定（用于提取参考坐标系）：
//   - GetAppearancesFromWeldPoint(wp, log)     → 单焊点 → 绑定的 Part/Appearance 列表
//   - FillAppearancesForOperation(op, log)     → 批量为 op.Points 填充外观信息
//   焊点路径的 MakePt 自动填充 PointInfo.AppearanceName / AppearanceMatrix / PartName

using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Tecnomatix.Engineering;

namespace TxTools.ExportGun
{
    public enum PointType { WeldPoint, PathPoint, ContinuousPoint, All }

    public class PointInfo
    {
        public string Name;
        public PointType Type;
        public double[] TCPMatrix;
        public double[] Position;
        public double[] Normal;

        // —— 外观绑定信息（可为 null；用于参考坐标系提取） ——
        // 焊点可能绑定到零件(TxPart)或某个具体外观(TxPartAppearance)。
        // 一个焊点可绑定多个外观；此处记录主外观（第一个），其余见 AllAppearances。
        public string AppearanceName;              // 外观名字（优先），若仅有 Part 则存 Part 名
        public double[] AppearanceMatrix;          // 外观自身的 AbsoluteLocation (4x4)
        public string PartName;                    // 外观所属的父零件名（如能解析）
        public List<AppearanceRef> AllAppearances; // 所有绑定（含 Part / Appearance）
    }

    /// <summary>焊点绑定的零件外观引用（名字 + 世界系 4x4 矩阵）。</summary>
    public class AppearanceRef
    {
        public string Name;           // 绑定对象名字
        public double[] Matrix;       // AbsoluteLocation (4x4，行主序，长度 16)
        public string TypeName;       // 原始类型名（诊断用）
        public string ParentPartName; // 若能反查到父零件名
        public ITxObject RawObject;   // 底层 ITxObject 引用（用于 Display/Blank 控制）
    }

    public class GunInfo
    {
        public string Name;
        public string ModelPath;
        public double[] ToolMatrix;     // 工具安装坐标（世界系），= CGR 自然原点
        public double[] TcpWorldMatrix; // TCP 在世界系的绝对位置
        public double[] TcpRelTool;     // TCP 相对工具自身的局部固定偏移 = Inv(ToolMatrix)*TcpWorldMatrix
    }

    public class OperationInfo
    {
        public string Name;
        public string TypeLabel;
        public ITxObject PsObject;
        public List<PointInfo> Points = new List<PointInfo>();
        public GunInfo Gun;
    }
    /// <summary>
    /// 焊点标注截图的点位数据。
    /// PS 坐标由 PsReader 填写；屏幕坐标由 PsReader.ProjectPointsToScreen 填写。
    /// </summary>
    public class WeldAnnotationPoint
    {
        // ── PS 数据（由 GetPointsFromSelection 填充）────────────────────────
        public int Index;
        public string Name;
        public string OpName;
        public double X, Y, Z;                  // 世界坐标 mm
        public TxTransformation WorldTx;        // 完整变换矩阵（备用）

        // ── 板厚检测：绑定的不重复零件数（用于自动分类） ────────────────
        // 由 GetPointsFromSelection 在收集焊点时统计 AllAppearances 的
        // 唯一 ParentPartName（无 ParentPartName 则用 Appearance 名作 fallback）。
        // 0 = 未检测到绑定；1 = 单层；2 = 两层；3+ = 多层
        public int BoundPartsCount;

        // ── 截图投影数据（由 ProjectPointsToScreen 填充）───────────────────
        public float ScreenX;
        public float ScreenY;
        public bool InViewport;
    }

    public static class PsReader
    {
        // ════════════════════════════════════════════════════════════
        //  1. 解析选中对象 → 操作列表
        // ════════════════════════════════════════════════════════════
        public static List<OperationInfo> GetOperationsFromSelection(Action<string> log)
        {
            if (log == null) log = Nop;
            var result = new List<OperationInfo>();
            try
            {
                TxObjectList sel = TxApplication.ActiveSelection.GetItems();
                if (sel != null && sel.Count > 0)
                    foreach (ITxObject obj in sel)
                        ParseItem(obj, result, log);
            }
            catch (Exception ex) { log($"[PS] 读取选中异常：{ex.Message}"); }

            if (result.Count == 0)
            {
                log("[PS] 无有效选中，读取 OperationRoot...");
                ReadRoot(result, log);
            }
            log($"[PS] 解析完成：{result.Count} 个操作");
            return result;
        }

        private static void ParseItem(ITxObject obj, List<OperationInfo> list, Action<string> log)
        {
            if (obj == null) return;
            string tn = obj.GetType().Name;

            if (tn == "TxWeldOperation" || (tn.Contains("TxWeld") && obj is ITxRoboticOperation))
            { list.Add(MakeOp(obj, "焊接操作")); return; }

            if (obj is TxCompoundOperation)
            { ExpandCompound((TxCompoundOperation)obj, list, log); return; }

            if (obj is ITxRoboticOperation)
            { list.Add(MakeOp(obj, LabelOp(tn))); return; }

            if (obj is TxWeldPoint wp)
            {
                // 选中单个焊点：向上找父 WeldOperation，保留 Robot/Tool 信息
                var wpTx = SafeGetTx(() => wp.AbsoluteLocation);
                ITxObject parentOp = FindParentWeldOperation(wp);
                OperationInfo op = parentOp != null
                    ? MakeOp(parentOp, "焊接操作(单点)")
                    : MakeOp(obj, "焊点");
                op.Name = wp.Name;
                if (wpTx != null) op.Points.Add(MakePt(wpTx, wp.Name, PointType.WeldPoint, wp, Nop));
                list.Add(op);
                return;
            }

            if (obj is ITxMfgFeature)
            { list.Add(MakeOp(obj, "MFG特征")); return; }

            TxObjectList kids = GetKids(obj);
            if (kids != null && kids.Count > 0)
                foreach (ITxObject child in kids)
                    ParseItem(child, list, log);
        }

        private static void ExpandCompound(TxCompoundOperation comp, List<OperationInfo> list, Action<string> log)
        {
            TxObjectList kids = GetKids(comp);
            if (kids == null || kids.Count == 0) { list.Add(MakeOp(comp, "复合操作")); return; }
            bool added = false;
            foreach (ITxObject child in kids)
            {
                string tn = child.GetType().Name;
                if (tn == "TxWeldOperation" || tn.Contains("TxWeld"))
                { added = true; list.Add(MakeOp(child, "焊接操作")); }
                else if (child is TxCompoundOperation)
                { added = true; ExpandCompound((TxCompoundOperation)child, list, log); }
                else if (child is ITxRoboticOperation)
                { added = true; list.Add(MakeOp(child, LabelOp(tn))); }
            }
            if (!added) list.Add(MakeOp(comp, "复合操作"));
        }

        private static void ReadRoot(List<OperationInfo> list, Action<string> log)
        {
            try
            {
                TxDocument doc = TxApplication.ActiveDocument;
                if (doc == null) return;
                TxObjectList kids = GetKidsFromRoot(doc.OperationRoot);
                if (kids == null) return;
                foreach (ITxObject obj in kids) ParseItem(obj, list, log);
            }
            catch (Exception ex) { log($"[PS] ReadRoot 异常：{ex.Message}"); }
        }

        // 通过 WeldLocationOperations 反向找父 TxWeldOperation
        private static ITxObject FindParentWeldOperation(TxWeldPoint wp)
        {
            try
            {
                TxObjectList weldLocOps = wp.WeldLocationOperations;
                if (weldLocOps == null || weldLocOps.Count == 0) return null;
                foreach (ITxObject locOp in weldLocOps)
                {
                    try
                    {
                        dynamic d = locOp;
                        ITxObject parent = null;
                        try { parent = d.Operation as ITxObject; } catch { }
                        if (parent == null) try { parent = d.ParentOperation as ITxObject; } catch { }
                        if (parent == null) try { parent = d.Owner as ITxObject; } catch { }
                        if (parent != null && (parent.GetType().Name.Contains("Weld") || parent is ITxRoboticOperation))
                            return parent;
                        if (locOp is ITxRoboticOperation) return locOp;
                    }
                    catch { }
                }
            }
            catch { }
            return null;
        }

        // ════════════════════════════════════════════════════════════
        //  2. 填充焊点数据
        // ════════════════════════════════════════════════════════════
        public static void FillPoints(OperationInfo op, PointType ptFilter, bool useMfg, Action<string> log)
        {
            if (log == null) log = Nop;
            if (op?.PsObject == null) return;
            if (op.Points.Count > 0) return;

            bool wWeld = ptFilter == PointType.WeldPoint || ptFilter == PointType.All;
            bool wPath = ptFilter == PointType.PathPoint || ptFilter == PointType.All;
            bool wCont = ptFilter == PointType.ContinuousPoint || ptFilter == PointType.All;

            string tn = op.PsObject.GetType().Name;
            if (tn == "TxWeldOperation" || tn.Contains("TxWeld"))
                L1_WeldOpEnum(op, useMfg, wWeld, wPath, wCont, log);

            if (wPath || wCont) L2_Locations(op, wPath, wCont, log);
            if (wWeld) L3_MfgFeatures(op, useMfg, log);  // 始终尝试，不依赖前层是否找到点
            if (op.Points.Count == 0) L4_DeepWalk(op.PsObject, op.Points, wWeld, wPath, wCont, 0, log);

            log($"[PS] [{op.Name}] {op.Points.Count} 个焊点");
        }

        private static void L1_WeldOpEnum(OperationInfo op, bool useMfg,
            bool wWeld, bool wPath, bool wCont, Action<string> log)
        {
            try
            {
                IEnumerable src = op.PsObject as IEnumerable;
                if (src == null) return;
                foreach (object item in src)
                {
                    if (item == null) continue;
                    try
                    {
                        PointType ptKind = PmKindOf(item.GetType().Name);
                        bool want = (ptKind == PointType.WeldPoint && wWeld)
                                 || (ptKind == PointType.PathPoint && wPath)
                                 || (ptKind == PointType.ContinuousPoint && wCont);
                        if (!want) continue;
                        TxTransformation tx = GetTxFromPm(item);
                        if (tx == null) continue;
                        string nm = GetNameFromPm(item, useMfg) ?? $"{op.Name}_{op.Points.Count + 1}";
                        if (Dup(op.Points, nm, ptKind)) continue;
                        // 焊点优先用 TxWeldPoint（更规范的绑定查询入口）；
                        // 路径点或焊点取不到 TxWeldPoint 时，直接传 PM item 自己
                        object source = item;
                        if (ptKind == PointType.WeldPoint)
                        {
                            TxWeldPoint wpFromPm = GetWpFromPm(item);
                            if (wpFromPm != null) source = wpFromPm;
                        }
                        op.Points.Add(MakePt(tx, nm, ptKind, source, log));
                    }
                    catch { }
                }
            }
            catch (Exception ex) { log($"[PS] L1 异常：{ex.Message}"); }
        }

        private static void L2_Locations(OperationInfo op, bool wPath, bool wCont, Action<string> log)
        {
            try
            {
                dynamic d = op.PsObject;
                IEnumerable locsEnum = null;
                TryGetEnum(d, "Locations", ref locsEnum);
                if (locsEnum == null) TryGetEnum(d, "LocationList", ref locsEnum);
                if (locsEnum == null) TryGetEnum(d, "RoboticLocations", ref locsEnum);
                if (locsEnum == null) return;
                foreach (object loc in locsEnum)
                {
                    try
                    {
                        PointType k = KindOf(loc.GetType().Name);
                        if (k == PointType.WeldPoint) continue;
                        if (k == PointType.PathPoint && !wPath) continue;
                        if (k == PointType.ContinuousPoint && !wCont) continue;
                        TxTransformation tx = GetTx(loc);
                        if (tx == null) continue;
                        string nm = SafeNameObj(loc);
                        if (!Dup(op.Points, nm, k)) op.Points.Add(MakePt(tx, nm, k));
                    }
                    catch { }
                }
            }
            catch (Exception ex) { log($"[PS] L2 异常：{ex.Message}"); }
        }

        private static void L3_MfgFeatures(OperationInfo op, bool useMfg, Action<string> log)
            => L3_Recurse(op.PsObject, op.Points, useMfg, 0, log);

        private static void L3_Recurse(ITxObject node, List<PointInfo> pts, bool useMfg, int depth, Action<string> log)
        {
            if (node == null || depth > 20) return;
            bool found = false;
            try
            {
                dynamic d = node;
                object raw = null;
                try { raw = d.AssignedMfgFeatures; } catch { }
                if (raw is IEnumerable ie)
                {
                    found = true;
                    foreach (object f in ie)
                    {
                        if (!(f is TxWeldPoint wp)) continue;
                        TxTransformation tx = SafeGetTx(() => wp.AbsoluteLocation);
                        if (tx == null) continue;
                        string nm = useMfg ? (MfgOpName(wp) ?? wp.Name) : wp.Name;
                        if (!Dup(pts, nm, PointType.WeldPoint)) pts.Add(MakePt(tx, nm, PointType.WeldPoint, wp, log));
                    }
                }
            }
            catch { }
            if (found) return;
            TxObjectList kids = GetKids(node);
            if (kids == null) return;
            foreach (ITxObject child in kids) L3_Recurse(child, pts, useMfg, depth + 1, log);
        }

        private static void L4_DeepWalk(ITxObject node, List<PointInfo> pts,
            bool wWeld, bool wPath, bool wCont, int depth, Action<string> log)
        {
            if (node == null || depth > 30) return;
            string tn = node.GetType().Name;

            if (node is TxWeldPoint wp)
            {
                if (wWeld)
                {
                    TxTransformation tx = SafeGetTx(() => wp.AbsoluteLocation);
                    if (tx != null && !Dup(pts, wp.Name, PointType.WeldPoint))
                        pts.Add(MakePt(tx, wp.Name, PointType.WeldPoint, wp, log));
                }
                return;
            }

            if (IsLocNode(tn))
            {
                PointType k = KindOf(tn);
                bool want = (k == PointType.WeldPoint && wWeld)
                         || (k == PointType.PathPoint && wPath)
                         || (k == PointType.ContinuousPoint && wCont);
                if (want)
                {
                    TxTransformation tx = GetTx(node);
                    string nm = SafeName(node);
                    if (tx != null && !Dup(pts, nm, k)) pts.Add(MakePt(tx, nm, k));
                }
            }
            else
            {
                IEnumerable asEnum = node as IEnumerable;
                if (asEnum != null)
                {
                    foreach (object item in asEnum)
                    {
                        if (item is ITxObject child) { L4_DeepWalk(child, pts, wWeld, wPath, wCont, depth + 1, log); continue; }
                        if (item == null) continue;
                        PointType pk = PmKindOf(item.GetType().Name);
                        bool want = (pk == PointType.WeldPoint && wWeld)
                                 || (pk == PointType.PathPoint && wPath)
                                 || (pk == PointType.ContinuousPoint && wCont);
                        if (!want) continue;
                        TxTransformation tx = GetTxFromPm(item) ?? GetTx(item);
                        if (tx == null) continue;
                        string nm = GetNameFromPm(item, false) ?? SafeNameObj(item);
                        if (Dup(pts, nm, pk)) continue;
                        object source = item;
                        if (pk == PointType.WeldPoint)
                        {
                            TxWeldPoint wpFromPm = GetWpFromPm(item);
                            if (wpFromPm != null) source = wpFromPm;
                        }
                        pts.Add(MakePt(tx, nm, pk, source, log));
                    }
                }
            }
            TxObjectList kids = GetKids(node);
            if (kids == null) return;
            foreach (ITxObject child in kids) L4_DeepWalk(child, pts, wWeld, wPath, wCont, depth + 1, log);
        }

        // ════════════════════════════════════════════════════════════
        //  3. 焊枪信息
        // ════════════════════════════════════════════════════════════
        public static GunInfo GetGunFromOperation(OperationInfo op, string modelPath, Action<string> log)
        {
            if (log == null) log = Nop;
            try
            {
                object toolObj = null;

                if (op.PsObject is TxWeldOperation weldOp)
                {
                    try { var t = weldOp.Tool; if (t != null) toolObj = t; } catch { }
                    if (toolObj == null) try { var g = weldOp.Gun; if (g != null) toolObj = g; } catch { }
                }
                if (toolObj == null)
                {
                    dynamic dOp = op.PsObject;
                    try { toolObj = dOp.Tool; } catch { }
                    if (toolObj == null) try { toolObj = dOp.Gun; } catch { }
                    if (toolObj == null) try { toolObj = dOp.ActiveTool; } catch { }
                }

                TxRobot robot = FindRobot(op.PsObject);
                if (toolObj == null && robot != null)
                {
                    try { dynamic dr = robot; toolObj = dr.ActiveTool; } catch { }
                    if (toolObj == null) try { dynamic dr = robot; toolObj = dr.CurrentTool; } catch { }
                    if (toolObj == null) toolObj = FindGunTool(op.PsObject);
                }

                if (toolObj == null && robot == null)
                { log($"[PS] [{op.Name}] 未找到工具对象"); return null; }

                string toolName = robot?.Name ?? "UnknownTool";
                try { dynamic dt = toolObj; toolName = (dt.Name as string) ?? toolName; } catch { }

                string resolvedPath = ResolveModelPath(toolObj, toolName, modelPath, log);

                TxTransformation toolTx = null;
                if (toolObj != null)
                {
                    try { dynamic dt = toolObj; toolTx = dt.AbsoluteLocation as TxTransformation; } catch { }
                    if (toolTx == null) try { dynamic dt = toolObj; toolTx = dt.LocationRelativeToWorld as TxTransformation; } catch { }
                }
                if (toolTx == null && robot != null)
                    try { toolTx = robot.Baseframe.AbsoluteLocation; } catch { }
                if (toolTx == null && robot != null)
                    try { toolTx = robot.AbsoluteLocation; } catch { }

                TxTransformation tcpWorldTx = null;
                if (toolObj != null)
                {
                    dynamic dt = toolObj;
                    try { tcpWorldTx = dt.TCPF.AbsoluteLocation as TxTransformation; } catch { }
                    if (tcpWorldTx == null) try { tcpWorldTx = dt.AbsoluteTCPLocation as TxTransformation; } catch { }
                    if (tcpWorldTx == null) try { tcpWorldTx = dt.TCPFrame.AbsoluteLocation as TxTransformation; } catch { }
                }
                if (tcpWorldTx == null && robot != null)
                    try { tcpWorldTx = robot.TCPF.AbsoluteLocation; } catch { }
                if (tcpWorldTx == null) tcpWorldTx = toolTx;

                TxTransformation tcpRelTool = null;
                if (toolTx != null && tcpWorldTx != null)
                    try { tcpRelTool = TxTransformation.Multiply(toolTx.Inverse, tcpWorldTx); } catch { }

                return new GunInfo
                {
                    Name = toolName,
                    ModelPath = resolvedPath,
                    ToolMatrix = TxToArr(toolTx),
                    TcpWorldMatrix = TxToArr(tcpWorldTx),
                    TcpRelTool = TxToArr(tcpRelTool)
                };
            }
            catch (Exception ex)
            { log($"[PS] GetGun 异常：{ex.Message}"); return null; }
        }

        // ════════════════════════════════════════════════════════════
        //  3b. 多 TCP（工具坐标）枚举与解析
        //
        //  来源（与可达性验证一致）：
        //    · 默认：robot.TCPF.AbsoluteLocation（机器人当前 TCP）
        //    · 各路径点 location.Parameters 里 Type=="RRS_TOOL_FRAME" 的
        //      TxRoboticTxObjectParam.Value（TxFrame），按名字去重
        //    · 工具对象（Gun/Tool）下挂的 TxFrame 子对象（兜底）
        // ════════════════════════════════════════════════════════════
        public const string DefaultTcpLabel = "默认 (机器人当前TCP)";
        private const string TYPE_RRS_TOOL_FRAME = "RRS_TOOL_FRAME";

        public class TcpOption
        {
            public string Name;          // 显示名（默认项为 DefaultTcpLabel）
            public double[] WorldMatrix; // TCP 世界系 4x4
            public bool IsDefault;       // 是否为机器人当前 TCP
        }

        /// <summary>枚举该操作可用的 TCP 选项（默认 + 各点工具坐标 + 工具子坐标系）。</summary>
        public static List<TcpOption> EnumerateTcpOptions(OperationInfo op, Action<string> log)
        {
            if (log == null) log = Nop;
            var result = new List<TcpOption>();
            var seen = new HashSet<string>();
            if (op?.PsObject == null) return result;

            try
            {
                // (1) 默认 TCP —— robot.TCPF
                TxRobot robot = FindRobot(op.PsObject);
                if (robot != null)
                {
                    TxTransformation tcpf = null;
                    try { tcpf = robot.TCPF.AbsoluteLocation; } catch { }
                    if (tcpf != null)
                    {
                        result.Add(new TcpOption { Name = DefaultTcpLabel, WorldMatrix = TxToArr(tcpf), IsDefault = true });
                        seen.Add(DefaultTcpLabel);
                    }
                }

                // (2) 控制器定义的系统坐标系 —— robot.GetAllSystemFrames()
                //     返回机器人树下 <robot>.<frame> 的全部坐标系（工具/用户/基),
                //     是跨厂商(ABB/KUKA/FANUC/...)提取 TCP 的主力来源。
                CollectSystemFrames(robot, result, seen, log);

                // (3) 各路径点的 RRS_TOOL_FRAME（按名字去重）
                foreach (object loc in EnumLocations(op.PsObject))
                {
                    TxFrame f = ReadLocationToolFrame(loc);
                    if (f == null) continue;
                    string nm = SafeNameObj(f);
                    if (string.IsNullOrEmpty(nm) || !seen.Add(nm)) continue;
                    TxTransformation abs = SafeGetTx(() => f.AbsoluteLocation);
                    if (abs == null) continue;
                    result.Add(new TcpOption { Name = nm, WorldMatrix = TxToArr(abs) });
                }

                // (4) 工具对象下挂的 TxFrame 子坐标系（兜底，按名字去重）
                object toolObj = ResolveToolObject(op);
                if (toolObj is ITxObject toolTx)
                {
                    CollectChildFrames(toolTx, result, seen, 0);
                }

                // 仅有默认项时，输出诊断帮助定位（实例参数 / OLP 定义）
                if (result.Count <= 1) LogRobotParamDiagnostics(robot, log);
            }
            catch (Exception ex) { log($"[TCP] 枚举异常：{ex.Message}"); }

            log($"[TCP] [{op.Name}] 可用 TCP {result.Count} 个");
            return result;
        }

        /// <summary>读取 robot.GetAllSystemFrames() 的系统坐标系，加入候选（按名去重）。</summary>
        private static void CollectSystemFrames(TxRobot robot, List<TcpOption> result,
            HashSet<string> seen, Action<string> log)
        {
            if (robot == null) return;
            IEnumerable frames = null;
            try { dynamic dr = robot; frames = dr.GetAllSystemFrames() as IEnumerable; }
            catch (Exception ex) { log($"[TCP] GetAllSystemFrames 异常：{ex.Message}"); return; }
            if (frames == null) { log("[TCP] GetAllSystemFrames 返回空"); return; }

            int n = 0;
            foreach (object fobj in frames)
            {
                if (fobj == null) continue;
                string nm = null;
                try { dynamic df = fobj; nm = df.Name as string; } catch { }
                if (string.IsNullOrEmpty(nm)) nm = SafeNameObj(fobj);
                if (string.IsNullOrEmpty(nm) || seen.Contains(nm)) continue;
                double[] m = GetFrameWorldMatrix(fobj);
                if (m == null) continue;
                seen.Add(nm);
                result.Add(new TcpOption { Name = nm, WorldMatrix = m });
                n++;
            }
            log($"[TCP] 系统坐标系 {n} 个（GetAllSystemFrames）");
        }

        /// <summary>从坐标系对象提取世界系 4x4（多属性兜底）。</summary>
        private static double[] GetFrameWorldMatrix(object frame)
        {
            if (frame == null) return null;
            if (frame is TxFrame fr)
            {
                var t = SafeGetTx(() => fr.AbsoluteLocation);
                return t != null ? TxToArr(t) : null;
            }
            TxTransformation tx = null;
            try { dynamic d = frame; tx = d.AbsoluteLocation as TxTransformation; } catch { }
            if (tx == null) try { dynamic d = frame; tx = d.LocationRelativeToWorld as TxTransformation; } catch { }
            if (tx == null) try { dynamic d = frame; tx = d.Location as TxTransformation; } catch { }
            if (tx == null) try { dynamic d = frame; tx = d.Transformation as TxTransformation; } catch { }
            if (tx == null) try { dynamic d = frame; var inner = d.Frame as TxFrame; if (inner != null) tx = SafeGetTx(() => inner.AbsoluteLocation); } catch { }
            return tx != null ? TxToArr(tx) : null;
        }

        /// <summary>诊断：输出机器人实例参数 / OLP 数据定义可用性，便于定位 TCP 提取失败。</summary>
        private static void LogRobotParamDiagnostics(TxRobot robot, Action<string> log)
        {
            if (robot == null) { log("[TCP][诊断] 未关联机器人"); return; }
            try
            {
                // 全部实例参数名
                try
                {
                    dynamic dr = robot;
                    object allp = null;
                    try { allp = dr.GetAllInstanceParameters(); } catch { }
                    if (allp is IEnumerable ie)
                    {
                        var names = new List<string>();
                        foreach (object pr in ie)
                        {
                            if (pr == null) continue;
                            string pn = null;
                            try { dynamic dp = pr; pn = (dp.Type as string) ?? (dp.Name as string); } catch { }
                            if (string.IsNullOrEmpty(pn)) pn = SafeNameObj(pr);
                            if (!string.IsNullOrEmpty(pn)) names.Add(pn);
                        }
                        log("[TCP][诊断] 实例参数：" + (names.Count > 0 ? string.Join(", ", names) : "(空)"));
                    }
                    else log("[TCP][诊断] GetAllInstanceParameters 不可用或为空");
                }
                catch (Exception ex) { log("[TCP][诊断] 枚举实例参数异常：" + ex.Message); }

                // ABB：OLP_LOCAL_DATA_DEFINITIONS（含 tooldata/wobjdata 原生语法）
                try
                {
                    dynamic dr = robot;
                    object v = null;
                    try { v = dr.GetInstanceParameter("OLP_LOCAL_DATA_DEFINITIONS"); } catch { }
                    string sv = v as string;
                    if (!string.IsNullOrEmpty(sv))
                        log($"[TCP][诊断] OLP_LOCAL_DATA_DEFINITIONS 可用（{sv.Length} 字符，ABB 可解析 tooldata）");
                }
                catch { }
            }
            catch (Exception ex) { log("[TCP][诊断] 异常：" + ex.Message); }
        }

        /// <summary>按名称解析 TCP 世界矩阵；name 为默认标签时返回 robot.TCPF；找不到返回 null。</summary>
        public static double[] ResolveTcpWorldByName(OperationInfo op, string name, Action<string> log)
        {
            if (log == null) log = Nop;
            if (op?.PsObject == null || string.IsNullOrEmpty(name)) return null;
            var opts = EnumerateTcpOptions(op, log);
            foreach (var o in opts)
                if (string.Equals(o.Name, name, StringComparison.Ordinal)) return o.WorldMatrix;
            // 退化：按名字在工具子坐标系里再找一次
            foreach (var o in opts)
                if (o.Name != null && o.Name.Equals(name, StringComparison.OrdinalIgnoreCase)) return o.WorldMatrix;
            return null;
        }

        /// <summary>解析操作绑定的工具对象（Tool/Gun/ActiveTool）。</summary>
        private static object ResolveToolObject(OperationInfo op)
        {
            object toolObj = null;
            try
            {
                if (op.PsObject is TxWeldOperation weldOp)
                {
                    try { var t = weldOp.Tool; if (t != null) toolObj = t; } catch { }
                    if (toolObj == null) try { var g = weldOp.Gun; if (g != null) toolObj = g; } catch { }
                }
                if (toolObj == null)
                {
                    dynamic dOp = op.PsObject;
                    try { toolObj = dOp.Tool; } catch { }
                    if (toolObj == null) try { toolObj = dOp.Gun; } catch { }
                    if (toolObj == null) try { toolObj = dOp.ActiveTool; } catch { }
                }
                if (toolObj == null)
                {
                    TxRobot robot = FindRobot(op.PsObject);
                    if (robot != null)
                    {
                        try { dynamic dr = robot; toolObj = dr.ActiveTool; } catch { }
                        if (toolObj == null) try { dynamic dr = robot; toolObj = dr.CurrentTool; } catch { }
                    }
                    if (toolObj == null) toolObj = FindGunTool(op.PsObject);
                }
            }
            catch { }
            return toolObj;
        }

        /// <summary>枚举操作下的路径点（Locations / LocationList / RoboticLocations）。</summary>
        private static IEnumerable EnumLocations(ITxObject opObj)
        {
            IEnumerable locs = null;
            try
            {
                dynamic d = opObj;
                TryGetEnum(d, "Locations", ref locs);
                if (locs == null) TryGetEnum(d, "LocationList", ref locs);
                if (locs == null) TryGetEnum(d, "RoboticLocations", ref locs);
            }
            catch { }
            return locs ?? new object[0];
        }

        /// <summary>读取一个路径点绑定的工具坐标系（RRS_TOOL_FRAME）；无则 null。</summary>
        private static TxFrame ReadLocationToolFrame(object loc)
        {
            if (loc == null) return null;
            try
            {
                ArrayList paramList = null;
                try { dynamic d = loc; paramList = d.Parameters as ArrayList; } catch { }
                if (paramList == null) return null;
                for (int i = 0; i < paramList.Count; i++)
                {
                    var prm = paramList[i] as TxRoboticParam;
                    if (prm == null) continue;
                    if (prm.Type != TYPE_RRS_TOOL_FRAME) continue;
                    var objParam = prm as TxRoboticTxObjectParam;
                    if (objParam == null) continue;
                    return objParam.Value as TxFrame;
                }
            }
            catch { }
            return null;
        }

        /// <summary>递归收集工具对象下的 TxFrame 子坐标系（按名去重，深度限 6）。</summary>
        private static void CollectChildFrames(ITxObject node, List<TcpOption> result,
            HashSet<string> seen, int depth)
        {
            if (node == null || depth > 6) return;
            TxObjectList kids = GetKids(node);
            if (kids == null) return;
            foreach (ITxObject child in kids)
            {
                if (child is TxFrame fr)
                {
                    string nm = SafeName(fr);
                    if (!string.IsNullOrEmpty(nm) && seen.Add(nm))
                    {
                        TxTransformation abs = SafeGetTx(() => fr.AbsoluteLocation);
                        if (abs != null) result.Add(new TcpOption { Name = nm, WorldMatrix = TxToArr(abs) });
                    }
                }
                CollectChildFrames(child, result, seen, depth + 1);
            }
        }

        // ── CGR 路径解析 ──────────────────────────────────────────────
        private static string ResolveModelPath(object toolObj, string toolName, string fallback, Action<string> log)
        {
            // 策略1：官方 API
            if (toolObj is ITxStorable storable)
            {
                try
                {
                    TxStorage storage = storable.StorageObject;
                    TxLibraryStorage libStorage = storage as TxLibraryStorage;
                    if (libStorage != null)
                    {
                        string fullPath = libStorage.FullPath;
                        if (!string.IsNullOrEmpty(fullPath))
                        {
                            if (File.Exists(fullPath) && IsSupportedModel(fullPath)) return fullPath;
                            if (Directory.Exists(fullPath))
                            {
                                string found = SearchCgrInDir(fullPath, toolName, log);
                                if (found != null) return found;
                            }
                            string dir = Path.GetDirectoryName(fullPath);
                            if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
                            {
                                string same = Path.Combine(dir, Path.GetFileNameWithoutExtension(fullPath) + ".cgr");
                                if (File.Exists(same)) return same;
                                string found = SearchCgrInDir(dir, toolName, log);
                                if (found != null) return found;
                            }
                        }
                    }
                    else if (storage != null)
                    {
                        try
                        {
                            dynamic dyn = storage;
                            string p = dyn.FullPath as string;
                            if (!string.IsNullOrEmpty(p))
                            {
                                if (File.Exists(p) && IsSupportedModel(p)) return p;
                                string dir = Path.GetDirectoryName(p);
                                if (Directory.Exists(dir)) { string f = SearchCgrInDir(dir, toolName, log); if (f != null) return f; }
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            }

            // 策略2：反射属性
            string[] pathProps = {
                "ExternalFilePath", "FilePath", "ModelFilePath", "SourceFilePath",
                "CgrFilePath", "GeometryFilePath", "ResourceFilePath", "ResourcePath",
                "ExternalFile", "FileLocation", "DataFilePath", "JtFilePath"
            };
            foreach (string prop in pathProps)
            {
                try
                {
                    var pi = toolObj.GetType().GetProperty(prop, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                    if (pi == null) continue;
                    string s = pi.GetValue(toolObj) as string;
                    if (string.IsNullOrEmpty(s)) continue;
                    if (File.Exists(s) && IsSupportedModel(s)) return s;
                    string ext = Path.GetExtension(s).ToLowerInvariant();
                    string rawDir = Path.GetDirectoryName(s);
                    if ((ext == ".jt" || ext == ".xml") && Directory.Exists(rawDir))
                    {
                        string same = Path.Combine(rawDir, Path.GetFileNameWithoutExtension(s) + ".cgr");
                        if (File.Exists(same)) return same;
                        string found = SearchCgrInDir(rawDir, toolName, log);
                        if (found != null) return found;
                    }
                    if (Directory.Exists(s)) { string found = SearchCgrInDir(s, toolName, log); if (found != null) return found; }
                }
                catch { }
            }

            // 策略3：StudyPath 搜索
            foreach (string root in GetPsSearchRoots())
            {
                string toolDir = Path.Combine(root, toolName);
                if (Directory.Exists(toolDir)) { string f = SearchCgrInDir(toolDir, toolName, log); if (f != null) return f; }
                { string f = SearchCgrInDir(root, toolName, log); if (f != null) return f; }
            }

            return fallback;
        }

        private static string SearchCgrInDir(string dir, string toolName, Action<string> log)
        {
            if (!Directory.Exists(dir)) return null;
            foreach (string ext in new[] { "*.cgr", "*.CATPart", "*.CATProduct" })
            {
                try
                {
                    var files = Directory.GetFiles(dir, ext, SearchOption.TopDirectoryOnly);
                    if (files.Length == 0) continue;
                    foreach (string f in files)
                        if (string.Equals(Path.GetFileNameWithoutExtension(f), toolName, StringComparison.OrdinalIgnoreCase))
                            return f;
                    foreach (string f in files)
                    {
                        string stem = Path.GetFileNameWithoutExtension(f);
                        if (stem.IndexOf(toolName, StringComparison.OrdinalIgnoreCase) >= 0 ||
                            toolName.IndexOf(stem, StringComparison.OrdinalIgnoreCase) >= 0)
                            return f;
                    }
                    if (files.Length == 1) return files[0];
                }
                catch { }
            }
            return null;
        }

        private static bool IsSupportedModel(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            return ext == ".cgr" || ext == ".catpart" || ext == ".catproduct";
        }

        private static string[] GetPsSearchRoots()
        {
            var roots = new List<string>();
            void TryAdd(string p) { if (!string.IsNullOrEmpty(p) && Directory.Exists(p) && !roots.Contains(p)) roots.Add(p); }
            try
            {
                dynamic doc = TxApplication.ActiveDocument;
                if (doc != null)
                {
                    foreach (string prop in new[] { "StudyPath", "LibraryPath", "RootPath", "FilePath", "FolderPath" })
                        try { TryAdd(doc.GetType().GetProperty(prop)?.GetValue(doc) as string); } catch { }
                    try { string fp = doc.FilePath as string; TryAdd(Path.GetDirectoryName(fp)); } catch { }
                }
            }
            catch { }
            try { TryAdd(Directory.GetCurrentDirectory()); } catch { }
            return roots.ToArray();
        }

        private static object FindGunTool(ITxObject root)
        {
            if (root == null) return null;
            try { string tn = root.GetType().Name; if (tn.Contains("Gun") || tn.Contains("Tool") || tn.Contains("Gripper")) return root; } catch { }
            TxObjectList kids = GetKids(root);
            if (kids == null) return null;
            foreach (ITxObject child in kids) { var f = FindGunTool(child); if (f != null) return f; }
            return null;
        }

        private static TxRobot FindRobot(ITxObject obj)
        {
            if (obj == null) return null;
            try { dynamic d = obj; TxRobot r = d.Robot as TxRobot; if (r != null) return r; } catch { }
            TxObjectList kids = GetKids(obj);
            if (kids == null) return null;
            foreach (ITxObject child in kids) { TxRobot r = FindRobot(child); if (r != null) return r; }
            return null;
        }

        /// <summary>公开包装，供 Form 构造"仅显示"白名单时查找绑定机器人。</summary>
        public static TxRobot FindRobotForOperation(ITxObject op) => FindRobot(op);

        // ════════════════════════════════════════════════════════════
        //  4. 参考坐标系
        // ════════════════════════════════════════════════════════════
        public static Tuple<string, double[]> GetReferenceFrame()
        {
            try
            {
                TxObjectList sel = TxApplication.ActiveSelection.GetItems();
                if (sel != null && sel.Count > 0)
                {
                    foreach (ITxObject obj in sel)
                    {
                        if (obj == null) continue;
                        string tn = obj.GetType().Name;
                        if (obj is TxFrame fr)
                        {
                            TxTransformation tx = SafeGetTx(() => fr.AbsoluteLocation);
                            if (tx != null && !IsIdentity(TxToArr(tx)))
                                return Tuple.Create(fr.Name, TxToArr(tx));
                        }
                        if (tn.Contains("Component") || tn.Contains("Product") || tn.Contains("Physical") || tn.Contains("Resource"))
                        {
                            TxTransformation tx = null;
                            try { dynamic d = obj; tx = d.AbsoluteLocation as TxTransformation; } catch { }
                            if (tx == null) try { dynamic d = obj; tx = d.LocationRelativeToWorkingFrame as TxTransformation; } catch { }
                            if (tx != null && !IsIdentity(TxToArr(tx)))
                                return Tuple.Create(SafeName(obj) + "(组件)", TxToArr(tx));
                        }
                        try
                        {
                            dynamic d = obj;
                            TxTransformation tx = d.AbsoluteLocation as TxTransformation;
                            if (tx != null && !IsIdentity(TxToArr(tx)))
                                return Tuple.Create(SafeName(obj), TxToArr(tx));
                        }
                        catch { }
                    }
                }
            }
            catch { }
            try
            {
                TxDocument doc = TxApplication.ActiveDocument;
                if (doc != null)
                {
                    TxTransformation wf = doc.WorkingFrame;
                    if (wf != null) { double[] arr = TxToArr(wf); if (!IsIdentity(arr)) return Tuple.Create("产品坐标系", arr); }
                }
            }
            catch { }
            return Tuple.Create("世界坐标系", new double[] { 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1 });
        }

        // ════════════════════════════════════════════════════════════
        //  5. 公共矩阵工具
        // ════════════════════════════════════════════════════════════
        public static double[] ToRelative(double[] absTx, double[] refMatrix)
        {
            if (refMatrix == null || IsIdentity(refMatrix)) return absTx;
            try
            {
                TxTransformation rel = TxTransformation.Multiply(ArrToTx(refMatrix).Inverse, ArrToTx(absTx));
                return TxToArr(rel);
            }
            catch { return absTx; }
        }

        public static bool IsIdentity(double[] m)
        {
            if (m == null || m.Length < 16) return true;
            return Math.Abs(m[0] - 1) < 1e-9 && Math.Abs(m[5] - 1) < 1e-9 && Math.Abs(m[10] - 1) < 1e-9
                && Math.Abs(m[1]) < 1e-9 && Math.Abs(m[2]) < 1e-9 && Math.Abs(m[4]) < 1e-9
                && Math.Abs(m[6]) < 1e-9 && Math.Abs(m[8]) < 1e-9 && Math.Abs(m[9]) < 1e-9
                && Math.Abs(m[3]) < 1e-9 && Math.Abs(m[7]) < 1e-9 && Math.Abs(m[11]) < 1e-9;
        }

        public static TxTransformation ArrToTxPublic(double[] m) => ArrToTx(m);

        public static double[] TxToArr(TxTransformation tx)
        {
            if (tx == null) return new double[] { 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1 };
            return new double[] {
                tx[0,0], tx[0,1], tx[0,2], tx[0,3],
                tx[1,0], tx[1,1], tx[1,2], tx[1,3],
                tx[2,0], tx[2,1], tx[2,2], tx[2,3],
                tx[3,0], tx[3,1], tx[3,2], tx[3,3]
            };
        }

        internal static void MatrixToEulerDeg(double[] m, out double rx, out double ry, out double rz)
        {
            rx = ry = rz = 0.0;
            if (m == null || m.Length < 12) return;
            double sinY = -m[8];
            if (sinY > 1.0) sinY = 1.0;
            if (sinY < -1.0) sinY = -1.0;
            ry = Math.Asin(sinY);
            const double GimbalLock = 1.0 - 1e-6;
            if (Math.Abs(sinY) < GimbalLock)
            {
                rx = Math.Atan2(m[9], m[10]);
                rz = Math.Atan2(m[4], m[0]);
            }
            else
            {
                rx = 0.0;
                rz = sinY > 0 ? Math.Atan2(-m[1], m[5]) : Math.Atan2(m[1], -m[5]);
            }
            const double Rad2Deg = 180.0 / Math.PI;
            rx *= Rad2Deg; ry *= Rad2Deg; rz *= Rad2Deg;
        }

        // ════════════════════════════════════════════════════════════
        //  6. 附加公共工具方法（供 GUI 调用）
        // ════════════════════════════════════════════════════════════

        // 获取操作绑定的工具名（供日志显示）
        public static string GetToolNameFromOperation(OperationInfo op)
        {
            if (op?.PsObject == null) return null;
            try
            {
                object toolObj = null;
                if (op.PsObject is TxWeldOperation weldOp)
                {
                    try { toolObj = weldOp.Tool; } catch { }
                    if (toolObj == null) try { toolObj = weldOp.Gun; } catch { }
                }
                if (toolObj == null)
                {
                    dynamic d = op.PsObject;
                    try { toolObj = d.Tool; } catch { }
                    if (toolObj == null) try { toolObj = d.Gun; } catch { }
                    if (toolObj == null) try { toolObj = d.ActiveTool; } catch { }
                }
                TxRobot robot = FindRobot(op.PsObject);
                if (toolObj == null && robot != null)
                {
                    try { dynamic dr = robot; toolObj = dr.ActiveTool; } catch { }
                    if (toolObj == null) try { dynamic dr = robot; toolObj = dr.CurrentTool; } catch { }
                }
                if (toolObj != null)
                    try { dynamic dt = toolObj; string nm = dt.Name as string; if (!string.IsNullOrEmpty(nm)) return nm; } catch { }
                if (robot != null) return robot.Name + "（机器人）";
            }
            catch { }
            return null;
        }

        // 从操作绑定的外观 Component 获取参考坐标
        /// <summary>
        /// 从操作获取参考坐标系。严格策略：
        ///   - 仅取操作内第一个焊点绑定的零件/外观矩阵（PointInfo.AppearanceMatrix）
        ///   - 没有外观则返回 null（由调用方决定是否回退为世界系）
        /// 显示名自带来源标签："<名字>(外观)" 或 "<名字>(零件)"。
        /// </summary>
        public static Tuple<string, double[]> GetRefFrameFromOperation(OperationInfo op, Action<string> log = null)
        {
            if (op == null || op.Points == null || op.Points.Count == 0) return null;
            if (log == null) log = Nop;

            // 需要已填过外观数据。若未填，主动补一次（幂等）。
            bool anyHasApp = false;
            foreach (var p in op.Points) { if (p.AppearanceMatrix != null) { anyHasApp = true; break; } }
            if (!anyHasApp) FillAppearancesForOperation(op, log);

            // 严格取第一个点的外观
            PointInfo first = op.Points[0];
            if (first == null || first.AppearanceMatrix == null) return null;

            // 按点类型决定语义：焊点=零件，路径点=夹具
            string baseTag = (first.Type == PointType.WeldPoint) ? "零件" : "夹具";
            string suffix = IsIdentity(first.AppearanceMatrix) ? "(" + baseTag + "/单位阵)" : "(" + baseTag + ")";

            string name = string.IsNullOrEmpty(first.AppearanceName) ? "参考坐标系" : first.AppearanceName;
            return Tuple.Create(name + suffix, first.AppearanceMatrix);
        }

        /// <summary>参考坐标系来源枚举（供 GUI 日志区分）。</summary>
        public enum RefFrameSource { None, Appearance, Part, Fixture, World }

        /// <summary>
        /// 参考坐标系解析结果。
        /// Source=None 表示操作内无焊点外观，调用方应决定是否回退；
        /// 其余情况 Matrix/Name 均非空。
        /// NeedsUserChoice=true 表示零件与夹具坐标不一致，GUI 应弹窗让用户选。
        /// </summary>
        public class RefFrameResult
        {
            public RefFrameSource Source;
            public string Name;        // 已含来源后缀，可直接作为 GUI 显示
            public double[] Matrix;    // 4x4 (长度 16)；World 时为单位阵
            public string PointName;   // 关联的焊点名（仅 Appearance / Part / Fixture 情况有值）

            // 候选集合（供冲突消解/弹窗使用）
            public List<AppearanceRef> PartCandidates;     // 所有零件候选
            public List<AppearanceRef> FixtureCandidates;  // 所有夹具候选
            public bool NeedsUserChoice;                   // 零件与夹具坐标不一致：GUI 应弹窗
            public string ConflictReason;                  // 冲突说明（仅 NeedsUserChoice=true 时有值）
        }

        /// <summary>坐标一致性容差。可按需调节。</summary>
        public const double RefFramePosTol = 0.01;   // 位置容差 mm
        public const double RefFrameRotTol = 1e-4;   // 旋转元素容差

        /// <summary>
        /// 判断两个 4x4 变换矩阵是否在容差内一致（位置+旋转）。
        /// 矩阵按行主序长度 16，0..2/4..6/8..10 为旋转部分，3/7/11 为平移。
        /// </summary>
        public static bool MatricesNearlyEqual(double[] a, double[] b, double posTol = -1, double rotTol = -1)
        {
            if (a == null || b == null || a.Length < 12 || b.Length < 12) return false;
            if (posTol < 0) posTol = RefFramePosTol;
            if (rotTol < 0) rotTol = RefFrameRotTol;

            // 位置
            if (Math.Abs(a[3] - b[3]) > posTol) return false;
            if (Math.Abs(a[7] - b[7]) > posTol) return false;
            if (Math.Abs(a[11] - b[11]) > posTol) return false;

            // 旋转 3x3
            int[] rotIdx = { 0, 1, 2, 4, 5, 6, 8, 9, 10 };
            foreach (int i in rotIdx)
                if (Math.Abs(a[i] - b[i]) > rotTol) return false;

            return true;
        }

        /// <summary>
        /// GUI 友好的参考坐标解析：
        /// 1) 严格取第一个焊点的外观/零件绑定；
        /// 2) 拿不到则按 fallbackToWorld 决定是否返回世界系；
        /// 3) Name 字段自带来源后缀，PointName 用于日志"基于 xx 号焊点的外观"。
        /// </summary>
        public static RefFrameResult ResolveOperationRefFrame(OperationInfo op, bool fallbackToWorld = true, Action<string> log = null)
        {
            if (log == null) log = Nop;
            var r = new RefFrameResult { Source = RefFrameSource.None };
            if (op == null || op.Points == null || op.Points.Count == 0)
            {
                if (fallbackToWorld)
                {
                    r.Source = RefFrameSource.World;
                    r.Name = "世界坐标系";
                    r.Matrix = new double[] { 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1 };
                }
                return r;
            }

            // 懒填一次外观
            bool anyHasApp = false;
            foreach (var p in op.Points) { if (p.AppearanceMatrix != null) { anyHasApp = true; break; } }
            if (!anyHasApp) FillAppearancesForOperation(op, log);

            // 1) 找第一个"焊点"的第一个绑定对象 → 零件（默认参考）
            PointInfo firstWeldPt = null; int firstWeldIdx = -1;
            for (int i = 0; i < op.Points.Count; i++)
            {
                var p = op.Points[i];
                if (p != null && p.Type == PointType.WeldPoint
                    && p.AllAppearances != null && p.AllAppearances.Count > 0)
                { firstWeldPt = p; firstWeldIdx = i; break; }
            }

            // 2) 找第一个"路径点/过渡点"的第一个绑定对象 → 夹具（用于对比）
            PointInfo firstPathPt = null; int firstPathIdx = -1;
            for (int i = 0; i < op.Points.Count; i++)
            {
                var p = op.Points[i];
                if (p != null && (p.Type == PointType.PathPoint || p.Type == PointType.ContinuousPoint)
                    && p.AllAppearances != null && p.AllAppearances.Count > 0)
                { firstPathPt = p; firstPathIdx = i; break; }
            }

            // 收集候选（仅用于 GUI 弹窗时展示所有选项）
            var parts = new List<AppearanceRef>();      // 焊点绑定（叫零件）
            var fixtures = new List<AppearanceRef>();   // 路径点绑定（叫夹具）
            if (firstWeldPt != null)
                foreach (var a in firstWeldPt.AllAppearances) if (a != null) parts.Add(a);
            if (firstPathPt != null)
                foreach (var a in firstPathPt.AllAppearances) if (a != null) fixtures.Add(a);
            r.PartCandidates = parts;
            r.FixtureCandidates = fixtures;

            // 精简日志：只保留摘要行
            if (firstWeldPt != null)
                log("[坐标] 首焊点 '" + firstWeldPt.Name + "'：" + parts.Count + " 个零件候选");
            if (firstPathPt != null)
                log("[坐标] 首路径点 '" + firstPathPt.Name + "'：" + fixtures.Count + " 个夹具候选");

            // 3) 选主：焊点零件优先；无焊点则用路径点夹具
            AppearanceRef chosen = null;
            bool chosenIsFixture = false;
            string chosenPointName = null;
            if (parts.Count > 0) { chosen = parts[0]; chosenIsFixture = false; chosenPointName = firstWeldPt.Name; }
            else if (fixtures.Count > 0) { chosen = fixtures[0]; chosenIsFixture = true; chosenPointName = firstPathPt.Name; }

            if (chosen == null)
            {
                log("[坐标] 未找到任何点的零件/夹具绑定");
                if (fallbackToWorld)
                {
                    r.Source = RefFrameSource.World;
                    r.Name = "世界坐标系";
                    r.Matrix = new double[] { 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1 };
                }
                return r;
            }

            r.Matrix = chosen.Matrix;
            r.PointName = chosenPointName;
            r.Source = chosenIsFixture ? RefFrameSource.Fixture : RefFrameSource.Part;

            string srcTag = chosenIsFixture ? "夹具" : "零件";
            string suffix = IsIdentity(chosen.Matrix) ? "(" + srcTag + "/单位阵)" : "(" + srcTag + ")";
            r.Name = (string.IsNullOrEmpty(chosen.Name) ? "参考坐标系" : chosen.Name) + suffix;

            // 4) 一致性检查：只在同时有焊点零件和路径点夹具时，且都只取第一个
            if (parts.Count > 0 && fixtures.Count > 0)
            {
                var partFirst = parts[0];
                var fixFirst = fixtures[0];
                if (!MatricesNearlyEqual(partFirst.Matrix, fixFirst.Matrix))
                {
                    r.NeedsUserChoice = true;
                    r.ConflictReason = "焊点零件 '" + partFirst.Name + "' 与路径点夹具 '"
                                     + fixFirst.Name + "' 坐标不一致";
                    log("[坐标] ⚠ " + r.ConflictReason);
                    // 冲突时列出所有候选，方便用户在弹窗里决策
                    foreach (var a in parts) log("[坐标]   [零件] " + a.Name);
                    foreach (var a in fixtures) log("[坐标]   [夹具] " + a.Name);
                }
            }

            return r;
        }

        // 功能4：从 PS 当前选中获取操作列表，若选中的不是操作则返回 null
        public static List<OperationInfo> GetOperationsFromSelectionIfAny(out string selKey)
        {
            selKey = null;
            var result = new List<OperationInfo>();
            try
            {
                TxObjectList sel = TxApplication.ActiveSelection.GetItems();
                if (sel == null || sel.Count == 0) return null;

                foreach (ITxObject obj in sel)
                {
                    if (obj == null) continue;
                    string tn = obj.GetType().Name;

                    bool isOp = obj is ITxRoboticOperation
                             || obj is TxCompoundOperation
                             || obj is ITxOperation
                             || tn.Contains("Operation");
                    if (!isOp) return null;

                    ParseItem(obj, result, Nop);
                }

                if (result.Count == 0) return null;

                var sb = new System.Text.StringBuilder();
                foreach (OperationInfo op in result) sb.Append(op.Name).Append("|");
                selKey = sb.ToString();
                return result;
            }
            catch { return null; }
        }

        // ════════════════════════════════════════════════════════════
        //  7. [补充] 供重构后 ExportGunForm 原生选择器使用的新方法
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 将 TxPickListener 拾取到的对象解析为操作列表。
        /// 
        /// 策略（防御式，兼容不同 PS 版本）：
        ///   1. pickedObj 本身是 ITxObject → 走已有 ParseItem 逻辑
        ///   2. pickedObj 是 PickedEventArgs 包装 → 尝试 .PickedObject 解包
        ///   3. pickedObj 是集合 → 逐个解析
        ///   4. pickedObj 不是操作类型 → 返回空列表（不抛异常）
        /// 
        /// 必须在 PS 主线程（STA）上调用。
        /// </summary>
        public static List<OperationInfo> ParsePickedObjectToOperations(dynamic pickedObj)
        {
            var result = new List<OperationInfo>();
            if (pickedObj == null) return result;

            try
            {
                // ── 尝试1：直接作为 ITxObject ──────────────────────────
                ITxObject txObj = pickedObj as ITxObject;
                if (txObj != null)
                {
                    ParseItemForPick(txObj, result);
                    return result;
                }

                // ── 尝试2：从事件参数中解包 ──────────────────────────────
                // TxPickListener_PickedEventArgs.PickedObject 可能返回
                // ITxObject 或其他包装类型
                try
                {
                    object innerObj = pickedObj.PickedObject;
                    if (innerObj is ITxObject innerTx)
                    {
                        ParseItemForPick(innerTx, result);
                        return result;
                    }
                    // 如果 innerObj 不是 ITxObject，尝试 dynamic 解析
                    if (innerObj != null)
                    {
                        ITxObject castTx = innerObj as ITxObject;
                        if (castTx != null)
                        {
                            ParseItemForPick(castTx, result);
                            return result;
                        }
                    }
                }
                catch { }

                // ── 尝试3：pickedObj 是集合（多选拾取） ──────────────────
                try
                {
                    if (pickedObj is IEnumerable enumerable)
                    {
                        foreach (object item in enumerable)
                        {
                            ITxObject itemTx = item as ITxObject;
                            if (itemTx != null)
                                ParseItemForPick(itemTx, result);
                        }
                        return result;
                    }
                }
                catch { }

                // ── 尝试4：TxObjectList（PS 原生集合） ───────────────────
                try
                {
                    TxObjectList txList = pickedObj as TxObjectList;
                    if (txList != null)
                    {
                        foreach (ITxObject item in txList)
                            ParseItemForPick(item, result);
                        return result;
                    }
                }
                catch { }

                // ── 尝试5：dynamic 属性探测 ─────────────────────────────
                // 某些 PS 版本的拾取对象可能需要通过特殊属性获取
                try
                {
                    object obj = pickedObj.Object;
                    if (obj is ITxObject objTx)
                    {
                        ParseItemForPick(objTx, result);
                        return result;
                    }
                }
                catch { }

                try
                {
                    object obj = pickedObj.Item;
                    if (obj is ITxObject objTx)
                    {
                        ParseItemForPick(objTx, result);
                        return result;
                    }
                }
                catch { }
            }
            catch { }

            return result;
        }

        /// <summary>
        /// 将单个 PS 对象包装为 OperationInfo。
        /// 供 TxObjGridCtrl 的 ObjectInserted 事件使用。
        /// 
        /// 与 ParsePickedObjectToOperations 的区别：
        ///   - 本方法只处理单个对象，返回单个 OperationInfo（或 null）
        ///   - 不递归展开 CompoundOperation
        ///   - 非操作类型直接返回 null
        /// 
        /// 必须在 PS 主线程（STA）上调用。
        /// </summary>
        public static OperationInfo WrapAsOperationInfo(dynamic obj)
        {
            if (obj == null) return null;

            try
            {
                // 优先尝试转为 ITxObject
                ITxObject txObj = obj as ITxObject;
                if (txObj != null)
                    return WrapTxObject(txObj);

                // dynamic 探测
                try
                {
                    ITxObject inner = obj.Object as ITxObject;
                    if (inner != null) return WrapTxObject(inner);
                }
                catch { }

                // 最后兜底：尝试用 dynamic 直接读取 Name
                try
                {
                    string name = obj.Name as string;
                    if (!string.IsNullOrEmpty(name))
                    {
                        // 创建一个"最小化"的 OperationInfo
                        // PsObject 为 null 意味着后续 FillPoints 等方法将跳过
                        return new OperationInfo
                        {
                            Name = name,
                            TypeLabel = "未知类型(从网格插入)",
                            PsObject = obj as ITxObject // 可能为 null
                        };
                    }
                }
                catch { }
            }
            catch { }

            return null;
        }

        /// <summary>
        /// 从任意 PS 对象中提取坐标系矩阵。
        /// 供 TxFrameEditBoxCtrl 的 ValidFrameSet 事件使用。
        /// 
        /// 返回 Tuple(坐标系名称, 4x4矩阵)，失败返回 null。
        /// 
        /// 策略：
        ///   1. TxFrame → AbsoluteLocation
        ///   2. Component/Product → AbsoluteLocation
        ///   3. 任意对象 → dynamic AbsoluteLocation
        ///   4. 兜底 → null
        /// 
        /// 必须在 PS 主线程（STA）上调用。
        /// </summary>
        public static Tuple<string, double[]> GetFrameMatrixFromObject(dynamic obj)
        {
            if (obj == null) return null;

            try
            {
                // ── TxFrame（坐标系对象） ─────────────────────────────
                TxFrame frame = obj as TxFrame;
                if (frame != null)
                {
                    TxTransformation tx = SafeGetTx(() => frame.AbsoluteLocation);
                    if (tx != null)
                    {
                        double[] arr = TxToArr(tx);
                        if (!IsIdentity(arr))
                            return Tuple.Create(frame.Name ?? "Frame", arr);
                    }
                }

                // ── ITxObject 通用路径 ───────────────────────────────
                ITxObject txObj = obj as ITxObject;
                if (txObj != null)
                {
                    string tn = txObj.GetType().Name;
                    string name = SafeName(txObj);

                    // Component / Product / Resource
                    if (tn.Contains("Component") || tn.Contains("Product")
                        || tn.Contains("Physical") || tn.Contains("Resource"))
                    {
                        TxTransformation tx = null;
                        try { dynamic d = txObj; tx = d.AbsoluteLocation as TxTransformation; } catch { }
                        if (tx == null)
                            try { dynamic d = txObj; tx = d.LocationRelativeToWorkingFrame as TxTransformation; } catch { }
                        if (tx != null)
                        {
                            double[] arr = TxToArr(tx);
                            if (!IsIdentity(arr))
                                return Tuple.Create(name + "(组件)", arr);
                        }
                    }

                    // 通用 AbsoluteLocation
                    try
                    {
                        dynamic d = txObj;
                        TxTransformation tx = d.AbsoluteLocation as TxTransformation;
                        if (tx != null)
                        {
                            double[] arr = TxToArr(tx);
                            if (!IsIdentity(arr))
                                return Tuple.Create(name, arr);
                        }
                    }
                    catch { }
                }

                // ── 非 ITxObject 的 dynamic 兜底 ────────────────────
                try
                {
                    TxTransformation tx = obj.AbsoluteLocation as TxTransformation;
                    if (tx != null)
                    {
                        double[] arr = TxToArr(tx);
                        string name = "未知坐标系";
                        try { name = obj.Name as string ?? name; } catch { }
                        if (!IsIdentity(arr))
                            return Tuple.Create(name, arr);
                    }
                }
                catch { }
            }
            catch { }

            return null;
        }

        // ════════════════════════════════════════════════════════════
        //  7.1 内部辅助（供新增方法使用）
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 拾取专用的 ParseItem —— 与原版 ParseItem 类似，但：
        ///   - 不在非操作类型上递归（拾取只应处理用户明确点击的对象）
        ///   - 对非操作对象静默跳过而非递归子项
        /// </summary>
        private static void ParseItemForPick(ITxObject obj, List<OperationInfo> result)
        {
            if (obj == null) return;
            string tn = obj.GetType().Name;

            // 焊接操作
            if (tn == "TxWeldOperation" || (tn.Contains("TxWeld") && obj is ITxRoboticOperation))
            {
                result.Add(MakeOp(obj, "焊接操作"));
                return;
            }

            // 复合操作 → 展开子操作
            if (obj is TxCompoundOperation comp)
            {
                ExpandCompound(comp, result, Nop);
                return;
            }

            // 其他机器人操作
            if (obj is ITxRoboticOperation)
            {
                result.Add(MakeOp(obj, LabelOp(tn)));
                return;
            }

            // ITxOperation（非 Robotic 的通用操作接口）
            if (obj is ITxOperation)
            {
                result.Add(MakeOp(obj, LabelOp(tn)));
                return;
            }

            // 类型名包含 "Operation" 的对象（兼容未知操作类型）
            if (tn.Contains("Operation"))
            {
                result.Add(MakeOp(obj, LabelOp(tn)));
                return;
            }

            // 单个焊点 → 向上查找父操作
            if (obj is TxWeldPoint wp)
            {
                var wpTx = SafeGetTx(() => wp.AbsoluteLocation);
                ITxObject parentOp = FindParentWeldOperation(wp);
                OperationInfo op = parentOp != null
                    ? MakeOp(parentOp, "焊接操作(单点)")
                    : MakeOp(obj, "焊点");
                op.Name = wp.Name;
                if (wpTx != null)
                    op.Points.Add(MakePt(wpTx, wp.Name, PointType.WeldPoint, wp, Nop));
                result.Add(op);
                return;
            }

            // MFG 特征
            if (obj is ITxMfgFeature)
            {
                result.Add(MakeOp(obj, "MFG特征"));
                return;
            }

            // 其他类型：不递归，静默忽略
            // 拾取模式下用户应明确选择操作对象
        }

        /// <summary>
        /// 将单个 ITxObject 包装为 OperationInfo。
        /// 非操作类型返回 null。
        /// </summary>
        private static OperationInfo WrapTxObject(ITxObject obj)
        {
            if (obj == null) return null;
            string tn = obj.GetType().Name;

            // 判断是否为操作类型
            bool isOperation = obj is ITxRoboticOperation
                            || obj is TxCompoundOperation
                            || obj is ITxOperation
                            || tn.Contains("Operation")
                            || tn.Contains("Weld");

            if (!isOperation)
            {
                // 非操作类型：尝试作为焊点处理
                if (obj is TxWeldPoint wp)
                {
                    ITxObject parentOp = FindParentWeldOperation(wp);
                    if (parentOp != null)
                        return MakeOp(parentOp, "焊接操作(单点)");
                    var opInfo = MakeOp(obj, "焊点");
                    opInfo.Name = wp.Name;
                    var wpTx = SafeGetTx(() => wp.AbsoluteLocation);
                    if (wpTx != null)
                        opInfo.Points.Add(MakePt(wpTx, wp.Name, PointType.WeldPoint, wp, Nop));
                    return opInfo;
                }

                if (obj is ITxMfgFeature)
                    return MakeOp(obj, "MFG特征");

                // 其他非操作类型返回 null
                return null;
            }

            // 操作类型：直接包装
            return MakeOp(obj, LabelOp(tn));
        }

        // ════════════════════════════════════════════════════════════
        //  私有工具方法
        // ════════════════════════════════════════════════════════════
        private static TxTransformation ArrToTx(double[] m)
        {
            var tx = new TxTransformation();
            for (int r = 0; r < 4; r++) for (int c = 0; c < 4; c++) tx[r, c] = m[r * 4 + c];
            return tx;
        }

        private static TxTransformation GetTxFromPm(object item)
        {
            dynamic d = item;
            try { TxTransformation tx = d.AbsoluteLocation as TxTransformation; if (tx != null) return tx; } catch { }
            try { object loc = d.Location; if (loc is TxTransformation t) return t; } catch { }
            try { dynamic ld = d.LocationData; TxTransformation tx = ld?.Frame as TxTransformation; if (tx != null) return tx; } catch { }
            try { TxWeldPoint wp = d.MfgFeature as TxWeldPoint; if (wp != null) return wp.AbsoluteLocation; } catch { }
            try { TxWeldPoint wp = d.Feature as TxWeldPoint; if (wp != null) return wp.AbsoluteLocation; } catch { }
            try { TxWeldPoint wp = d.WeldPoint as TxWeldPoint; if (wp != null) return wp.AbsoluteLocation; } catch { }
            return null;
        }

        /// <summary>从 Planning 层的 PM item 提取其对应的 TxWeldPoint（若存在）。</summary>
        private static TxWeldPoint GetWpFromPm(object item)
        {
            if (item == null) return null;
            if (item is TxWeldPoint wpDirect) return wpDirect;
            try { dynamic d = item; TxWeldPoint wp = d.MfgFeature as TxWeldPoint; if (wp != null) return wp; } catch { }
            try { dynamic d = item; TxWeldPoint wp = d.Feature as TxWeldPoint; if (wp != null) return wp; } catch { }
            try { dynamic d = item; TxWeldPoint wp = d.WeldPoint as TxWeldPoint; if (wp != null) return wp; } catch { }
            try { dynamic d = item; TxWeldPoint wp = d.AssignedMfgFeature as TxWeldPoint; if (wp != null) return wp; } catch { }
            return null;
        }

        private static TxTransformation GetTx(object node)
        {
            TxTransformation tx = null;
            try { dynamic d = node; tx = d.AbsoluteLocation as TxTransformation; } catch { }
            if (tx != null) return tx;
            try { dynamic d = node; tx = d.LocationData.Frame as TxTransformation; } catch { }
            if (tx != null) return tx;
            try { dynamic d = node; tx = d.Location as TxTransformation; } catch { }
            return tx;
        }

        private static string GetNameFromPm(object item, bool useMfg)
        {
            dynamic d = item;
            if (useMfg)
            {
                try { TxWeldPoint wp = d.MfgFeature as TxWeldPoint; if (wp != null) return MfgOpName(wp) ?? wp.Name; } catch { }
                try { TxWeldPoint wp = d.Feature as TxWeldPoint; if (wp != null) return MfgOpName(wp) ?? wp.Name; } catch { }
            }
            try { string nm = d.Name as string; if (!string.IsNullOrEmpty(nm)) return nm; } catch { }
            try { string nm = d.DisplayName as string; if (!string.IsNullOrEmpty(nm)) return nm; } catch { }
            return null;
        }

        private static void TryGetEnum(dynamic d, string prop, ref IEnumerable result)
        {
            if (result != null) return;
            try { object raw = d.GetType().GetProperty(prop)?.GetValue(d); if (raw is IEnumerable ie) result = ie; } catch { }
        }

        private static TxObjectList GetKids(ITxObject node)
        {
            if (node == null) return null;
            TxTypeFilter f = new TxTypeFilter(typeof(ITxObject));
            if (node is TxCompoundOperation co) try { return co.GetDirectDescendants(f); } catch { }
            if (node is TxOperationRoot ro) try { return ro.GetDirectDescendants(f); } catch { }
            try { dynamic d = node; return d.GetDirectDescendants(f) as TxObjectList; } catch { }
            return null;
        }

        private static TxObjectList GetKidsFromRoot(TxOperationRoot root)
        {
            if (root == null) return null;
            TxTypeFilter f = new TxTypeFilter(typeof(ITxObject));
            try { return root.GetDirectDescendants(f); } catch { }
            try { dynamic d = root; return d.GetDirectDescendants(f) as TxObjectList; } catch { }
            return null;
        }

        private static string MfgOpName(TxWeldPoint wp)
        {
            try { TxObjectList ops = wp.WeldLocationOperations; if (ops?.Count > 0) { dynamic f = ops[0]; string s = f.Name as string; if (!string.IsNullOrEmpty(s)) return s; } } catch { }
            return null;
        }

        private static PointInfo MakePt(TxTransformation tx, string name, PointType pt)
        {
            double[] m = TxToArr(tx);
            return new PointInfo { Name = name, Type = pt, TCPMatrix = m, Position = new[] { m[3], m[7], m[11] }, Normal = new[] { m[2], m[6], m[10] } };
        }

        /// <summary>带外观绑定信息的焊点构造：从 TxWeldPoint 读取外观并一并填充。</summary>
        private static PointInfo MakePt(TxTransformation tx, string name, PointType pt, TxWeldPoint wp, Action<string> log)
            => MakePt(tx, name, pt, (object)wp, log);

        /// <summary>
        /// 带绑定信息的点构造：从任意对象（TxWeldPoint、PM item、ITxObject）读取 AssignedParts 等。
        /// 路径点 / 过渡点也可以传入它们对应的 PM item，以便获取绑定的夹具/设备。
        /// </summary>
        private static PointInfo MakePt(TxTransformation tx, string name, PointType pt, object source, Action<string> log)
        {
            PointInfo info = MakePt(tx, name, pt);
            if (source == null) return info;
            try
            {
                List<AppearanceRef> apps = GetAppearancesFromObject(source);
                if (apps != null && apps.Count > 0)
                {
                    info.AllAppearances = apps;
                    AppearanceRef main = apps[0];
                    info.AppearanceName = main.Name;
                    info.AppearanceMatrix = main.Matrix;
                    info.PartName = main.ParentPartName;
                }
            }
            catch { }
            return info;
        }

        private static OperationInfo MakeOp(ITxObject obj, string label)
            => new OperationInfo { Name = SafeName(obj), TypeLabel = label, PsObject = obj };

        private static TxTransformation SafeGetTx(Func<TxTransformation> fn) { try { return fn(); } catch { return null; } }
        private static string SafeName(ITxObject obj) { try { dynamic d = obj; return (d.Name as string) ?? obj.GetType().Name; } catch { return obj.GetType().Name; } }
        private static string SafeNameObj(object obj) { try { dynamic d = obj; return (d.Name as string) ?? obj.GetType().Name; } catch { return obj.GetType().Name; } }
        private static bool IsLocNode(string tn) => tn.Contains("Location") || tn.Contains("Via") || tn.Contains("Arc") || tn.Contains("PathPoint") || tn.Contains("WeldLoc");
        private static PointType KindOf(string tn) { if (tn.Contains("Weld") || tn.Contains("Seam")) return PointType.WeldPoint; if (tn.Contains("Continuous")) return PointType.ContinuousPoint; return PointType.PathPoint; }
        private static PointType PmKindOf(string tn) { if (tn.Contains("Weld")) return PointType.WeldPoint; if (tn.Contains("Via") || tn.Contains("Arc")) return PointType.PathPoint; if (tn.Contains("Continuous")) return PointType.ContinuousPoint; return PointType.WeldPoint; }
        private static bool Dup(List<PointInfo> pts, string nm, PointType t) { foreach (var p in pts) if (p.Name == nm && p.Type == t) return true; return false; }
        private static string LabelOp(string tn) { if (tn.Contains("Weld")) return "焊接操作"; if (tn.Contains("Compound")) return "复合操作"; if (tn.Contains("Robotic")) return "机器人操作"; return "操作"; }
        private static void Nop(string s) { }

        // ════════════════════════════════════════════════════════════
        //  7. 焊点外观绑定提取（新增）
        //  一个 TxWeldPoint 可绑定到零件(TxPart)或具体外观(TxPartAppearance)，
        //  通常用于从零件局部坐标系反推焊点参考位姿。
        //  策略：先试官方 API 成员 → 失败降级到 dynamic 反射常见属性名。
        //  失败全部吞异常，只在 log 中追加一行定位信息，不抛。
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 从任意对象（TxWeldPoint、PM item、ITxObject）获取其绑定的零件/外观/夹具列表。
        /// 失败返回空列表而不是 null。静默，不输出日志。
        /// </summary>
        public static List<AppearanceRef> GetAppearancesFromObject(object source)
        {
            var list = new List<AppearanceRef>();
            if (source == null) return list;

            // 策略 A：官方 API 候选
            string[] officialCandidates =
            {
                "WeldedAppearances", "Appearances", "AssignedAppearances",
                "WeldedParts", "AssignedParts", "Parts", "WeldedComponents",
                "AssignedComponents", "AssignedDevices", "AssignedFixtures",
                "AttachedObjects", "Attachments"
            };

            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string propName in officialCandidates)
            {
                try
                {
                    object raw = TryGetMemberValue(source, propName);
                    if (raw == null) continue;
                    CollectAppearanceCandidates(raw, list, seen);
                }
                catch { }
            }

            // 策略 B：反射兜底（仅当策略 A 一无所获）
            if (list.Count == 0)
            {
                try
                {
                    var props = source.GetType().GetProperties(
                        System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                    foreach (var pi in props)
                    {
                        string pn = pi.Name;
                        if (pn.IndexOf("Part", StringComparison.OrdinalIgnoreCase) < 0
                            && pn.IndexOf("Appearance", StringComparison.OrdinalIgnoreCase) < 0
                            && pn.IndexOf("Component", StringComparison.OrdinalIgnoreCase) < 0
                            && pn.IndexOf("Fixture", StringComparison.OrdinalIgnoreCase) < 0
                            && pn.IndexOf("Device", StringComparison.OrdinalIgnoreCase) < 0
                            && pn.IndexOf("Attach", StringComparison.OrdinalIgnoreCase) < 0) continue;

                        object val = null;
                        try { val = pi.GetValue(source); } catch { continue; }
                        if (val == null) continue;

                        CollectAppearanceCandidates(val, list, seen);
                    }
                }
                catch { }
            }

            return list;
        }

        /// <summary>向后兼容：从焊点取外观。内部转发给 GetAppearancesFromObject。</summary>
        public static List<AppearanceRef> GetAppearancesFromWeldPoint(TxWeldPoint wp)
            => GetAppearancesFromObject(wp);

        /// <summary>通用收集器：把候选对象（单体或集合）展开为 AppearanceRef。</summary>
        private static void CollectAppearanceCandidates(
            object raw, List<AppearanceRef> list, HashSet<string> seen)
        {
            if (raw == null) return;

            // 集合：枚举每个元素递归收集
            IEnumerable asEnum = raw as IEnumerable;
            if (asEnum != null && !(raw is string))
            {
                foreach (object item in asEnum)
                    CollectOneCandidate(item, list, seen);
                return;
            }

            // 单体
            CollectOneCandidate(raw, list, seen);
        }

        /// <summary>从单个候选对象提取外观信息。不做类型判别，统一收集。无矩阵对象静默跳过。</summary>
        private static void CollectOneCandidate(
            object item, List<AppearanceRef> list, HashSet<string> seen)
        {
            if (item == null) return;

            string tn;
            try { tn = item.GetType().Name; } catch { return; }

            string name = SafeNameObj(item);
            TxTransformation tx = null;
            try { dynamic d = item; tx = d.AbsoluteLocation as TxTransformation; } catch { }
            if (tx == null)
                try { dynamic d = item; tx = d.LocationRelativeToWorkingFrame as TxTransformation; } catch { }
            if (tx == null)
                try { dynamic d = item; tx = d.LocationRelativeToWorld as TxTransformation; } catch { }

            // 无矩阵对象（如 TxAttachmentType 等元信息对象）静默跳过
            if (tx == null) return;

            // 去重键：类型名+名字+矩阵前几位
            double[] m = TxToArr(tx);
            string key = tn + "|" + name + "|" + m[3].ToString("F3") + "," + m[7].ToString("F3") + "," + m[11].ToString("F3");
            if (!seen.Add(key)) return;

            string parent = TryGetParentPartName(item);

            list.Add(new AppearanceRef
            {
                Name = name,
                Matrix = m,
                TypeName = tn,
                ParentPartName = parent,
                RawObject = item as ITxObject
            });
        }

        /// <summary>若候选是外观，尝试反查其父零件名；否则返回 null。</summary>
        private static string TryGetParentPartName(object item)
        {
            if (item == null) return null;
            string[] parentProps = { "Part", "ParentPart", "Owner", "OwnerPart", "Parent" };
            foreach (string pn in parentProps)
            {
                try
                {
                    object p = TryGetMemberValue(item, pn);
                    if (p == null) continue;
                    string ptn = p.GetType().Name;
                    // 只接受含 "Part"、不含 "Appearance" 的父对象
                    if (ptn.IndexOf("Part", StringComparison.OrdinalIgnoreCase) >= 0
                        && ptn.IndexOf("Appearance", StringComparison.OrdinalIgnoreCase) < 0)
                        return SafeNameObj(p);
                }
                catch { }
            }
            return null;
        }

        /// <summary>反射取属性/字段值，找不到返回 null（不抛）。</summary>
        private static object TryGetMemberValue(object obj, string memberName)
        {
            if (obj == null || string.IsNullOrEmpty(memberName)) return null;
            var t = obj.GetType();
            var flags = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance;
            try
            {
                var pi = t.GetProperty(memberName, flags);
                if (pi != null && pi.CanRead) return pi.GetValue(obj);
            }
            catch { }
            try
            {
                var fi = t.GetField(memberName, flags);
                if (fi != null) return fi.GetValue(obj);
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 批量为一个操作的所有焊点填充外观信息。
        /// 在 FillPoints 之后调用；对已有外观数据的点跳过。
        /// </summary>
        public static void FillAppearancesForOperation(OperationInfo op, Action<string> log = null)
        {
            if (log == null) log = Nop;
            if (op == null || op.Points == null || op.Points.Count == 0) return;

            int filled = 0, missing = 0;
            foreach (PointInfo p in op.Points)
            {
                if (p.AppearanceMatrix != null) { filled++; continue; }

                // 通过名字在 MFG 特征中反查焊点对象
                TxWeldPoint wp = FindWeldPointByName(op.PsObject, p.Name, 0);
                if (wp == null) { missing++; continue; }

                var apps = GetAppearancesFromWeldPoint(wp);
                if (apps.Count == 0) { missing++; continue; }

                AppearanceRef main = apps[0];
                p.AppearanceName = main.Name;
                p.AppearanceMatrix = main.Matrix;
                p.PartName = main.ParentPartName;
                p.AllAppearances = apps;
                filled++;
            }
            // 仅当有缺失时才输出，成功情况静默
            if (missing > 0)
                log("[APP] 操作 '" + op.Name + "'：" + filled + " / " + op.Points.Count + " 个点已绑定（" + missing + " 个缺失）");
        }

        /// <summary>在操作子树中按名字查找焊点（宽度有限的深搜）。</summary>
        private static TxWeldPoint FindWeldPointByName(ITxObject root, string name, int depth)
        {
            if (root == null || depth > 25 || string.IsNullOrEmpty(name)) return null;
            try
            {
                if (root is TxWeldPoint wp && string.Equals(wp.Name, name, StringComparison.Ordinal))
                    return wp;
            }
            catch { }

            // 从 AssignedMfgFeatures 里找（最常见来源）
            try
            {
                dynamic d = root;
                object raw = null;
                try { raw = d.AssignedMfgFeatures; } catch { }
                if (raw is IEnumerable ie)
                {
                    foreach (object f in ie)
                    {
                        if (f is TxWeldPoint wpf && string.Equals(wpf.Name, name, StringComparison.Ordinal))
                            return wpf;
                    }
                }
            }
            catch { }

            TxObjectList kids = GetKids(root);
            if (kids == null) return null;
            foreach (ITxObject child in kids)
            {
                var r = FindWeldPointByName(child, name, depth + 1);
                if (r != null) return r;
            }
            return null;
        }

        // ════════════════════════════════════════════════════════════
        //  CGR 模糊匹配（新增）
        //  由 ExportGunForm 在导出前调用，判断工具同级目录下的 CGR
        //  是否与工具名匹配。高相似度自动使用，低相似度交由用户确认。
        // ════════════════════════════════════════════════════════════
        public class CgrMatch
        {
            public string Path;            // 绝对路径
            public string CandidateName;   // 不含扩展名的文件名
            public double Similarity;      // 0..1，越大越相似
        }

        public class CgrLookupResult
        {
            public string ToolName;                  // 解析到的工具名（可能为空）
            public string SameDir;                   // 同级目录（可能为 null）
            public List<CgrMatch> Candidates;        // 相似度降序排列
            public string Reason;                    // 诊断信息（查不到时的说明）
        }

        /// <summary>从一个操作解析其工具 CGR 的同级目录 + 候选文件（按相似度降序）。</summary>
        /// <remarks>必须在 PS 主线程调用（InvokeOnPs）。</remarks>
        public static CgrLookupResult LookupCgrForOperation(OperationInfo op, Action<string> log = null)
        {
            if (log == null) log = Nop;
            var result = new CgrLookupResult
            {
                Candidates = new List<CgrMatch>(),
                ToolName = null,
                SameDir = null,
                Reason = null
            };

            if (op == null || op.PsObject == null) { result.Reason = "操作为空或 PsObject 为空"; return result; }

            // 1) 解析工具对象（与 GetGunFromOperation 同源）
            object toolObj = null;
            try
            {
                if (op.PsObject is TxWeldOperation weldOp)
                {
                    try { var t = weldOp.Tool; if (t != null) toolObj = t; } catch { }
                    if (toolObj == null) try { var g = weldOp.Gun; if (g != null) toolObj = g; } catch { }
                }
                if (toolObj == null)
                {
                    dynamic dOp = op.PsObject;
                    try { toolObj = dOp.Tool; } catch { }
                    if (toolObj == null) try { toolObj = dOp.Gun; } catch { }
                    if (toolObj == null) try { toolObj = dOp.ActiveTool; } catch { }
                }
                if (toolObj == null)
                {
                    TxRobot robot = FindRobot(op.PsObject);
                    if (robot != null)
                    {
                        try { dynamic dr = robot; toolObj = dr.ActiveTool; } catch { }
                        if (toolObj == null) try { dynamic dr = robot; toolObj = dr.CurrentTool; } catch { }
                        if (toolObj == null) toolObj = FindGunTool(op.PsObject);
                    }
                }
            }
            catch (Exception ex) { log("[CGR] 解析工具对象异常：" + ex.Message); }

            if (toolObj == null) { result.Reason = "未找到绑定工具"; return result; }

            // 2) 工具名
            try { dynamic dt = toolObj; result.ToolName = (dt.Name as string) ?? null; } catch { }
            if (string.IsNullOrEmpty(result.ToolName)) result.ToolName = "UnknownTool";

            // 3) 解析同级目录
            string sameDir = TryGetToolStorageDir(toolObj);
            result.SameDir = sameDir;

            if (string.IsNullOrEmpty(sameDir) || !Directory.Exists(sameDir))
            { result.Reason = string.IsNullOrEmpty(sameDir) ? "无法解析工具同级目录" : ("同级目录不存在：" + sameDir); return result; }

            // 4) 收集 .cgr 候选（仅扫描当前目录）
            string[] files = null;
            try { files = Directory.GetFiles(sameDir, "*.cgr", SearchOption.TopDirectoryOnly); } catch { }
            if (files == null || files.Length == 0)
            { result.Reason = "同级目录下未发现 CGR 文件：" + sameDir; return result; }

            foreach (var f in files)
            {
                string stem = Path.GetFileNameWithoutExtension(f);
                double sim = SimilarityRatio(stem, result.ToolName);
                result.Candidates.Add(new CgrMatch { Path = f, CandidateName = stem, Similarity = sim });
            }
            result.Candidates.Sort(delegate (CgrMatch a, CgrMatch b) { return b.Similarity.CompareTo(a.Similarity); });
            return result;
        }

        /// <summary>从工具对象解析其 CGR 所在目录（仅取同级，不递归子目录/StudyPath）。</summary>
        private static string TryGetToolStorageDir(object toolObj)
        {
            if (toolObj == null) return null;

            // 策略1：ITxStorable.StorageObject
            try
            {
                if (toolObj is ITxStorable storable)
                {
                    TxStorage storage = storable.StorageObject;
                    TxLibraryStorage libStorage = storage as TxLibraryStorage;
                    if (libStorage != null)
                    {
                        string fullPath = libStorage.FullPath;
                        if (!string.IsNullOrEmpty(fullPath))
                        {
                            if (Directory.Exists(fullPath)) return fullPath;
                            string dir = Path.GetDirectoryName(fullPath);
                            if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir)) return dir;
                        }
                    }
                    else if (storage != null)
                    {
                        try
                        {
                            dynamic dyn = storage;
                            string p = dyn.FullPath as string;
                            if (!string.IsNullOrEmpty(p))
                            {
                                if (Directory.Exists(p)) return p;
                                string dir = Path.GetDirectoryName(p);
                                if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir)) return dir;
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }

            // 策略2：反射常见路径属性
            string[] pathProps = {
                "ExternalFilePath", "FilePath", "ModelFilePath", "SourceFilePath",
                "CgrFilePath", "GeometryFilePath", "ResourceFilePath", "ResourcePath",
                "ExternalFile", "FileLocation", "DataFilePath", "JtFilePath"
            };
            foreach (string prop in pathProps)
            {
                try
                {
                    var pi = toolObj.GetType().GetProperty(prop, System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                    if (pi == null) continue;
                    string s = pi.GetValue(toolObj) as string;
                    if (string.IsNullOrEmpty(s)) continue;
                    if (Directory.Exists(s)) return s;
                    string dir = Path.GetDirectoryName(s);
                    if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir)) return dir;
                }
                catch { }
            }
            return null;
        }

        /// <summary>基于归一化 Levenshtein 距离的相似度：0 (完全不同) ~ 1 (完全相同)。</summary>
        public static double SimilarityRatio(string a, string b)
        {
            if (string.IsNullOrEmpty(a) && string.IsNullOrEmpty(b)) return 1.0;
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return 0.0;
            string na = NormalizeForMatch(a);
            string nb = NormalizeForMatch(b);
            if (na.Length == 0 && nb.Length == 0) return 1.0;
            if (na.Length == 0 || nb.Length == 0) return 0.0;
            if (string.Equals(na, nb, StringComparison.Ordinal)) return 1.0;

            // 子串包含给一定加分（用于工具名带版本号/后缀的情况）
            double contain = 0.0;
            if (na.IndexOf(nb, StringComparison.Ordinal) >= 0 || nb.IndexOf(na, StringComparison.Ordinal) >= 0)
                contain = 0.15;

            int dist = Lev(na, nb);
            int maxLen = Math.Max(na.Length, nb.Length);
            double base0 = 1.0 - (double)dist / maxLen;
            double total = base0 + contain;
            if (total > 1.0) total = 1.0;
            if (total < 0.0) total = 0.0;
            return total;
        }

        private static string NormalizeForMatch(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            var sb = new System.Text.StringBuilder(s.Length);
            foreach (char c in s)
            {
                if (char.IsLetterOrDigit(c)) sb.Append(char.ToLowerInvariant(c));
            }
            return sb.ToString();
        }

        private static int Lev(string a, string b)
        {
            if (a == null) a = string.Empty;
            if (b == null) b = string.Empty;
            int n = a.Length, m = b.Length;
            if (n == 0) return m;
            if (m == 0) return n;
            var prev = new int[m + 1];
            var curr = new int[m + 1];
            for (int j = 0; j <= m; j++) prev[j] = j;
            for (int i = 1; i <= n; i++)
            {
                curr[0] = i;
                for (int j = 1; j <= m; j++)
                {
                    int cost = (a[i - 1] == b[j - 1]) ? 0 : 1;
                    int del = prev[j] + 1;
                    int ins = curr[j - 1] + 1;
                    int sub = prev[j - 1] + cost;
                    int min = del < ins ? del : ins;
                    if (sub < min) min = sub;
                    curr[j] = min;
                }
                var tmp = prev; prev = curr; curr = tmp;
            }
            return prev[m];
        }

        // =============================================================================
        // PsReader.cs  ──  Section 8 新增内容（追加到 PsReader 类末尾，关闭 } 前）
        //
        // 供 WeldAnnotatorForm 使用的专用方法：
        //   GetCompoundOpsFromRoot  — 获取根操作节点列表（填充 TxObjGridCtrl）
        //   GetPointsFromSelection  — 从 ActiveSelection 提取焊点坐标
        //   CaptureActiveViewer     — 截取 PS 视口位图
        //   ProjectPointsToScreen   — 将世界坐标批量投影到截图像素坐标
        // =============================================================================

        // ─── 以下代码追加到 PsReader 类内部（紧接 Section 7 之后，类的 } 之前）───────

        // ════════════════════════════════════════════════════════════════════
        //  8. 焊点标注截图  —  TxTools.WeldAnnotator 专用
        // ════════════════════════════════════════════════════════════════════

        // ── 8.0  数据模型 ─────────────────────────────────────────────────────
        // （类声明独立放到 PsReader 类外部，与 OperationInfo / PointInfo 并列）

        // ── 8.1  获取文档根下的所有复合操作节点 ──────────────────────────────
        /// <summary>
        /// 返回 OperationRoot 直接子级中所有 ITxCompoundOperation 对象列表，
        /// 供 TxObjGridCtrl.AppendObject 逐个填充。
        /// </summary>
        public static TxObjectList GetCompoundOpsFromRoot(Action<string> log)
        {
            if (log == null) log = Nop;
            TxObjectList result = new TxObjectList();
            try
            {
                TxDocument doc = TxApplication.ActiveDocument;
                if (doc == null) { log("[Annotator] 无活动文档"); return result; }
                TxObjectList kids = GetKidsFromRoot(doc.OperationRoot);
                if (kids == null) return result;
                foreach (ITxObject o in kids)
                {
                    if (o is ITxCompoundOperation) result.Add(o);
                }
                log(string.Format("[Annotator] 根节点下找到 {0} 个复合操作", result.Count));
            }
            catch (Exception ex) { log("[Annotator] GetCompoundOpsFromRoot 异常：" + ex.Message); }
            return result;
        }

        // ── 8.2  从 ActiveSelection 提取焊点 ─────────────────────────────────
        /// <summary>
        /// 读取 TxApplication.ActiveSelection，提取其中所有焊点/路径点的坐标。
        /// 支持直接选中 TxWeldPoint、ITxRoboticLocationOperation、
        /// TxWeldOperation、TxCompoundOperation（自动展开子点）。
        /// </summary>
        /// <param name="opNameHint">操作名提示（来自TxObjGridCtrl确认的OP节点名）</param>
        public static List<WeldAnnotationPoint> GetPointsFromSelection(
            string opNameHint, Action<string> log)
        {
            if (log == null) log = Nop;
            var result = new List<WeldAnnotationPoint>();
            try
            {
                TxObjectList sel = TxApplication.ActiveSelection.GetItems();
                if (sel == null || sel.Count == 0)
                {
                    log("[Annotator] ActiveSelection 为空");
                    return result;
                }
                int idx = 1;
                foreach (ITxObject obj in sel)
                    CollectAnnotationPoints(obj, opNameHint, ref idx, result, log);

                log(string.Format("[Annotator] 从 {0} 个选中对象中提取到 {1} 个点位",
                    sel.Count, result.Count));
            }
            catch (Exception ex) { log("[Annotator] GetPointsFromSelection 异常：" + ex.Message); }
            return result;
        }

        private static void CollectAnnotationPoints(
            ITxObject obj, string opName, ref int idx,
            List<WeldAnnotationPoint> list, Action<string> log)
        {
            if (obj == null) return;
            string tn = obj.GetType().Name;

            // ── 单个 TxWeldPoint ──────────────────────────────────────────
            if (obj is TxWeldPoint wp)
            {
                TxTransformation tx = SafeGetTx(() => wp.AbsoluteLocation);
                if (tx != null)
                {
                    int boundParts = CountBoundParts(wp);
                    list.Add(MakeAnnotPt(idx++, wp.Name, opName, tx, boundParts));
                }
                return;
            }

            // ── ITxRoboticLocationOperation ───────────────────────────────
            if (obj is ITxRoboticLocationOperation loc)
            {
                TxTransformation tx = null;
                try { dynamic d = loc; tx = d.AbsoluteLocation as TxTransformation; } catch { }
                if (tx != null)
                    list.Add(MakeAnnotPt(idx++, loc.Name, opName, tx));
                return;
            }

            // ── TxWeldOperation / TxCompoundOperation → 展开子点 ─────────
            if (tn == "TxWeldOperation" || tn.Contains("TxWeld") || obj is ITxCompoundOperation)
            {
                TxObjectList kids = GetKids(obj);
                if (kids != null)
                    foreach (ITxObject child in kids)
                        CollectAnnotationPoints(child, opName, ref idx, list, log);
                return;
            }

            // ── 其他：尝试递归子节点 ──────────────────────────────────────
            TxObjectList misc = GetKids(obj);
            if (misc != null && misc.Count > 0)
                foreach (ITxObject child in misc)
                    CollectAnnotationPoints(child, opName, ref idx, list, log);
        }

        private static WeldAnnotationPoint MakeAnnotPt(
            int idx, string name, string opName, TxTransformation tx, int boundParts = 0)
        {
            return new WeldAnnotationPoint
            {
                Index = idx,
                Name = name,
                OpName = opName,
                X = tx[0, 3],
                Y = tx[1, 3],
                Z = tx[2, 3],
                WorldTx = tx,
                BoundPartsCount = boundParts
            };
        }

        /// <summary>
        /// 数焊点绑定的不重复零件数。来源：GetAppearancesFromObject 返回的
        /// AppearanceRef 列表。去重键优先用 ParentPartName（多个外观可能挂在同一
        /// 零件下，应算 1 层），fallback 用 Appearance Name。
        /// </summary>
        private static int CountBoundParts(ITxObject wp)
        {
            if (wp == null) return 0;
            try
            {
                List<AppearanceRef> apps = GetAppearancesFromObject(wp);
                if (apps == null || apps.Count == 0) return 0;
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var ar in apps)
                {
                    if (ar == null) continue;
                    string key = !string.IsNullOrEmpty(ar.ParentPartName)
                        ? ar.ParentPartName
                        : (ar.Name ?? "");
                    if (!string.IsNullOrEmpty(key)) seen.Add(key);
                }
                return seen.Count;
            }
            catch { return 0; }
        }

        // ── 8.3  截取 PS 视口位图 ─────────────────────────────────────────────
        /// <summary>
        /// 调用 TxGraphicViewer.GetImage(Size, bool) 截图。
        /// size 若为 Empty 则自动从 ContainerWindow.Bounds 读取视口实际尺寸。
        /// 截图前临时关闭 PS 视图右上角的导航辅助控件（导航方块、三轴指示器、罗盘），
        /// 截图后通过 try/finally 恢复原状。
        /// </summary>
        public static System.Drawing.Bitmap CaptureActiveViewer(
            System.Drawing.Size size, bool transparent, Action<string> log)
        {
            if (log == null) log = Nop;
            try
            {
                TxViewersManager mgr = TxApplication.ViewersManager;
                if (mgr == null) throw new InvalidOperationException("ViewersManager 不可用");
                TxGraphicViewer viewer = mgr.GraphicViewer;
                if (viewer == null) throw new InvalidOperationException("未找到活动 GraphicViewer");

                if (size.IsEmpty || size.Width <= 0)
                {
                    try
                    {
                        TxViewerWindow tvw = viewer.ContainerWindow;
                        if (tvw != null && tvw.Bounds.Width > 0)
                            size = new System.Drawing.Size(tvw.Bounds.Width, tvw.Bounds.Height);
                    }
                    catch { }
                    if (size.IsEmpty || size.Width <= 0)
                        size = new System.Drawing.Size(1920, 1080);
                }

                log(string.Format("[Annotator] 截图尺寸：{0}×{1}", size.Width, size.Height));

                // ── 关闭导航辅助控件 → 截图 → 恢复 ────────────────────────
                var savedFlags = HideViewerAids(viewer, log);
                System.Drawing.Bitmap bmp;
                try
                {
                    // 强制刷新一次，让关闭操作生效
                    try { TxApplication.RefreshDisplay(); } catch { }
                    bmp = viewer.GetImage(size, transparent);
                    if (bmp == null) throw new Exception("GetImage 返回 null");
                }
                finally
                {
                    RestoreViewerAids(viewer, savedFlags, log);
                    try { TxApplication.RefreshDisplay(); } catch { }
                }
                return bmp;
            }
            catch (Exception ex)
            {
                log("[Annotator] CaptureActiveViewer 异常：" + ex.Message);
                return null;
            }
        }

        // ── 视图辅助控件的关闭/恢复（导航方块 / 导航框架） ─────────────────
        //
        //  根据 PS SDK 文档，TxGraphicViewer 暴露两个 STATIC bool 属性：
        //    ShowNavigationCube  — 导航方块（右上角立方体）
        //    ShowNavigationFrame — 导航框架
        //  这两个都是静态属性（影响所有 viewer 实例），直接读写即可。
        //  截图前关闭、截图后按原值恢复。

        private struct ViewerAidState
        {
            public bool SavedShowNavigationCube;
            public bool SavedShowNavigationFrame;
            public bool HadCube;
            public bool HadFrame;
        }

        private static ViewerAidState HideViewerAids(
            TxGraphicViewer viewer, Action<string> log)
        {
            var saved = new ViewerAidState();

            // 静态属性 → 直接访问，无需反射
            try
            {
                saved.SavedShowNavigationCube = TxGraphicViewer.ShowNavigationCube;
                saved.HadCube = true;
                if (saved.SavedShowNavigationCube)
                {
                    TxGraphicViewer.ShowNavigationCube = false;
                    log("[Annotator] 已临时关闭：ShowNavigationCube");
                }
            }
            catch (Exception ex) { log("[Annotator] ShowNavigationCube 不可用：" + ex.Message); }

            try
            {
                saved.SavedShowNavigationFrame = TxGraphicViewer.ShowNavigationFrame;
                saved.HadFrame = true;
                if (saved.SavedShowNavigationFrame)
                {
                    TxGraphicViewer.ShowNavigationFrame = false;
                    log("[Annotator] 已临时关闭：ShowNavigationFrame");
                }
            }
            catch (Exception ex) { log("[Annotator] ShowNavigationFrame 不可用：" + ex.Message); }

            // 注：DisplayOrientationFrameOfReference 已被 SDK 弃用，
            // 其功能已合并到 ShowNavigationCube / ShowNavigationFrame 两个属性中，
            // 上面两段代码已覆盖。

            return saved;
        }

        private static void RestoreViewerAids(
            TxGraphicViewer viewer, ViewerAidState saved, Action<string> log)
        {
            if (saved.HadCube && saved.SavedShowNavigationCube)
            {
                try { TxGraphicViewer.ShowNavigationCube = true; }
                catch (Exception ex) { log("[Annotator] 恢复 ShowNavigationCube 失败：" + ex.Message); }
            }
            if (saved.HadFrame && saved.SavedShowNavigationFrame)
            {
                try { TxGraphicViewer.ShowNavigationFrame = true; }
                catch (Exception ex) { log("[Annotator] 恢复 ShowNavigationFrame 失败：" + ex.Message); }
            }
        }

        // ── 8.4  批量投影世界坐标 → 截图像素坐标 ──────────────────────────────
        /// <summary>
        /// 对 points 列表中每个点调用 PointScreenCoordinates，
        /// 将结果写入 ScreenX / ScreenY / InViewport。
        /// imgW / imgH 为截图像素尺寸，用于缩放比计算。
        /// </summary>
        public static void ProjectPointsToScreen(
            List<WeldAnnotationPoint> points, int imgW, int imgH, Action<string> log)
        {
            if (log == null) log = Nop;
            if (points == null || points.Count == 0) return;
            try
            {
                TxViewersManager mgr = TxApplication.ViewersManager;
                if (mgr == null) { log("[Annotator] ViewersManager 不可用"); return; }
                TxGraphicViewer viewer = mgr.GraphicViewer;
                if (viewer == null) { log("[Annotator] 未找到 GraphicViewer"); return; }

                // 计算截图与视口的缩放比
                float scaleX = 1f, scaleY = 1f;
                try
                {
                    TxViewerWindow tvw = viewer.ContainerWindow;
                    if (tvw != null && tvw.Bounds.Width > 0)
                    {
                        scaleX = (float)imgW / tvw.Bounds.Width;
                        scaleY = (float)imgH / tvw.Bounds.Height;
                    }
                }
                catch { }

                int ok = 0, fail = 0;
                foreach (WeldAnnotationPoint pt in points)
                {
                    pt.InViewport = false;
                    try
                    {
                        TxVector world = new TxVector(pt.X, pt.Y, pt.Z);
                        TxVector screen = viewer.PointScreenCoordinates(world);
                        pt.ScreenX = (float)(screen.X * scaleX);
                        pt.ScreenY = (float)(screen.Y * scaleY);
                        pt.InViewport = pt.ScreenX >= 0 && pt.ScreenX <= imgW
                                       && pt.ScreenY >= 0 && pt.ScreenY <= imgH;
                        ok++;
                    }
                    catch { fail++; }
                }
                log(string.Format("[Annotator] 投影：{0} 成功，{1} 失败", ok, fail));
            }
            catch (Exception ex) { log("[Annotator] ProjectPointsToScreen 异常：" + ex.Message); }
        }

        // ─── Section 8 结束 ────────────────────────────────────────────────────────────
        // =============================================================================
        // PsReader.cs  ──  Section 9 追加内容（仅快照相关，无重复方法）
        //
        // 追加位置：PsReader 类内部，Section 8 之后，类的 } 之前。
        //
        // Section 8 已包含：GetCompoundOpsFromRoot / GetPointsFromSelection /
        //                   MakeAnnotPt / CollectAnnotPts / CaptureActiveViewer /
        //                   ProjectPointsToScreen
        // Section 9 仅新增：显示状态快照/恢复/仅显示操作外观/恢复全部显示
        // =============================================================================

        // ════════════════════════════════════════════════════════════════════
        //  9. 显示状态快照  —  TxTools.WeldAnnotator 专用
        // ════════════════════════════════════════════════════════════════════

        // ── 9.1  拍摄快照：记录所有 ITxDisplayableObject 的当前 Visibility 状态 ──
        //
        //  实现基于官方 ITxDisplayableObject 接口：
        //    Visibility 属性可读（ALL / PARTIAL / NONE）
        //    Display() / Blank() 可写
        //  本函数遍历场景所有可显示对象，用 Visibility 属性记录初始状态；
        //  Restore 时按记录逐个 Display()/Blank()。
        public static List<Tuple<ITxObject, bool>> SnapshotDisplayStates(Action<string> log)
        {
            if (log == null) log = Nop;
            _visErrorLogged = false;
            var result = new List<Tuple<ITxObject, bool>>();
            try
            {
                List<ITxObject> all = EnumDisplayableObjects(log);
                int visibleCount = 0, hiddenCount = 0, unsupported = 0;
                foreach (ITxObject obj in all)
                {
                    bool? vis = GetObjVisible(obj, log);
                    if (!vis.HasValue) { unsupported++; continue; }
                    if (vis.Value) visibleCount++; else hiddenCount++;
                    result.Add(Tuple.Create(obj, vis.Value));
                }
                log(string.Format(
                    "[Snapshot] 已记录 {0} 个对象（显示 {1}，隐藏 {2}，不支持 {3}）",
                    result.Count, visibleCount, hiddenCount, unsupported));
            }
            catch (Exception ex) { log("[Snapshot] SnapshotDisplayStates: " + ex.Message); }
            return result;
        }

        // ── 9.2  恢复快照：按记录逐个 Display/Blank ───────────────────────────
        public static void RestoreDisplayStates(
            List<Tuple<ITxObject, bool>> snapshot, Action<string> log)
        {
            if (log == null) log = Nop;
            if (snapshot == null || snapshot.Count == 0)
            {
                log("[Snapshot] 无快照可恢复");
                return;
            }
            int changed = 0, alreadyOk = 0, failed = 0;
            foreach (Tuple<ITxObject, bool> t in snapshot)
            {
                // 先读当前状态；若已与记录一致则跳过，避免无谓调用
                bool? curr = GetObjVisible(t.Item1, log);
                if (curr.HasValue && curr.Value == t.Item2) { alreadyOk++; continue; }
                if (SetObjVisible(t.Item1, t.Item2, log)) changed++;
                else failed++;
            }
            try { TxApplication.RefreshDisplay(); } catch { }
            log(string.Format(
                "[Snapshot] 已恢复 {0} 个对象（变更 {1}，无需变更 {2}，失败 {3}）",
                snapshot.Count, changed, alreadyOk, failed));
        }

        // ── 9.3  仅显示操作绑定的外观 ────────────────────────────────────────
        //  本次会话中被 Hide 的对象列表（用于 RestoreFromLastHide 恢复）。
        //  保存的是 Blank() 时我们主动调用过的对象 —— 这样 Display() 它们即可。
        private static readonly List<ITxObject> _lastHiddenObjects = new List<ITxObject>();

        /// <summary>是否有未恢复的"仅显示操作外观"隐藏状态。</summary>
        public static bool HasUnrestoredHide
        {
            get { return _lastHiddenObjects.Count > 0; }
        }

        /// <summary>
        /// 隐藏场景所有对象，仅显示传入的 whitelist（通常 = 机器人 + 工具 + 焊点绑定外观）。
        /// 本版本不再从焊点反查 —— 白名单由调用方（Form）提供，Form 在读焊点时
        /// 已经通过 FillAppearancesForOperation 把绑定对象填到了 AppearanceRef.RawObject 里。
        /// 若 whitelist 为 null 或空，抛 InvalidOperationException。
        /// </summary>
        public static void HideAllExcept(
            IEnumerable<ITxObject> whitelist, Action<string> log)
        {
            if (log == null) log = Nop;
            if (whitelist == null)
                throw new InvalidOperationException("未提供白名单");

            // 白名单去重
            var keep = new System.Collections.Generic.HashSet<int>();
            var keepList = new List<ITxObject>();
            foreach (ITxObject o in whitelist)
            {
                if (o == null) continue;
                int id = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(o);
                if (keep.Add(id)) keepList.Add(o);
            }
            if (keepList.Count == 0)
                throw new InvalidOperationException("白名单为空，已取消操作（否则场景会全部被隐藏）。");

            // 枚举所有可显示对象
            List<ITxObject> all = EnumDisplayableObjects(log);

            // 重置"本次隐藏集合"。
            // 关键：先读对象当前 Visibility —— 如果原本就已隐藏（NONE），
            // 跳过不 Blank，也不加入 _lastHiddenObjects（否则 Restore 时
            // 会把用户手动隐藏过的对象一起 Display 出来）。
            _lastHiddenObjects.Clear();
            int hidden = 0, shown = 0, skippedAlreadyHidden = 0, failed = 0;
            foreach (ITxObject obj in all)
            {
                int id = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
                if (keep.Contains(id))
                {
                    // 白名单：确保显示
                    SetObjVisible(obj, true, log);
                    shown++;
                }
                else
                {
                    bool? curr = GetObjVisible(obj, log);
                    if (curr.HasValue && !curr.Value)
                    {
                        // 已经隐藏 — 保持现状，不纳入本次记录
                        skippedAlreadyHidden++;
                        continue;
                    }
                    // 当前显示中，或读不到状态 → Blank 它
                    if (SetObjVisible(obj, false, log))
                    {
                        _lastHiddenObjects.Add(obj);
                        hidden++;
                    }
                    else
                    {
                        failed++;
                    }
                }
            }
            try { TxApplication.RefreshDisplay(); } catch { }
            log(string.Format(
                "[Snapshot] 仅显示白名单：实际隐藏 {0} 个，显示 {1} 个，" +
                "原本就隐藏 {2} 个（已跳过），调用失败 {3} 个",
                hidden, shown, skippedAlreadyHidden, failed));
        }

        /// <summary>
        /// 恢复上一次 HideAllExcept 所隐藏的对象（Display 它们）。
        /// </summary>
        public static void RestoreFromLastHide(Action<string> log)
        {
            if (log == null) log = Nop;
            if (_lastHiddenObjects.Count == 0)
            {
                log("[Snapshot] 无上一次隐藏记录，改为显示全部。");
                ShowAllDevices(log);
                return;
            }
            int n = 0;
            foreach (ITxObject obj in _lastHiddenObjects)
            {
                SetObjVisible(obj, true, log);
                n++;
            }
            try { TxApplication.RefreshDisplay(); } catch { }
            log(string.Format("[Snapshot] 已恢复 {0} 个被隐藏的对象", n));
            _lastHiddenObjects.Clear();
        }

        /// <summary>
        /// 向后兼容旧签名：调用方（Form）应改为传入 AppearanceRef.RawObject 组成的列表。
        /// 若调用本重载且 knownWeldPointNames 非空，会尝试按名 find，但通常不可靠。
        /// </summary>
        public static void ShowOnlyOperationAppearance(
            ITxObject opObj, List<string> knownWeldPointNames, Action<string> log)
        {
            if (log == null) log = Nop;
            log("[Snapshot] ⚠ 使用了旧签名 ShowOnlyOperationAppearance — 请改用 HideAllExcept(whitelist)。");
            throw new InvalidOperationException(
                "旧签名已废弃：请从 Form 端构造白名单列表（机器人+工具+焊点绑定 RawObject）后调用 HideAllExcept。");
        }

        /// <summary>最旧签名（只传 opObj） — 同样废弃。</summary>
        public static void ShowOnlyOperationAppearance(ITxObject opObj, Action<string> log)
        {
            ShowOnlyOperationAppearance(opObj, null, log);
        }

        // ── 9.4  恢复全部显示 ────────────────────────────────────────────────
        public static void ShowAllDevices(Action<string> log)
        {
            if (log == null) log = Nop;
            try
            {
                List<ITxObject> all = EnumDisplayableObjects(log);
                foreach (ITxObject obj in all) SetObjVisible(obj, true, log);
                try { TxApplication.RefreshDisplay(); } catch { }
                log(string.Format("[Snapshot] 已恢复全部 {0} 个对象显示", all.Count));
            }
            catch (Exception ex) { log("[Snapshot] ShowAllDevices: " + ex.Message); }
        }

        // ── 9.x  私有辅助 ────────────────────────────────────────────────────

        /// <summary>
        /// 枚举场景组件树中所有实现 ITxDisplayableObject 接口的对象。
        /// 策略：
        ///   1. 通过 dynamic 直接读取 doc 上的 PhysicalRoot / ComponentRoot /
        ///      ResourceRoot 三个根（OperationRoot 下是逻辑操作，不纳入）；
        ///   2. 对每个根先尝试 GetAllDescendants（分别用 TxTypeFilter(ITxObject)
        ///      和无参重载）拿完整后代列表；
        ///   3. 若 GetAllDescendants 失败或返回空，回退到递归 Children 遍历；
        ///   4. 只保留 obj is ITxDisplayableObject 的对象（官方接口契约保证可用
        ///      Display/Blank/Visibility）。
        /// 每一步都打日志，便于诊断"0 个对象"的根因。
        /// </summary>
        private static List<ITxObject> EnumDisplayableObjects(Action<string> log)
        {
            var result = new List<ITxObject>();
            var seen = new System.Collections.Generic.HashSet<int>();
            try
            {
                TxDocument doc = TxApplication.ActiveDocument;
                if (doc == null) { log("[Snapshot] ActiveDocument 为 null"); return result; }

                // 只枚举 3D 物理/组件/资源根 —— OperationRoot 下是逻辑操作节点
                // （TxDeviceOperation、TxWeldOperation 等），不支持 Display/Blank，
                // 试图 Blank 会抛异常且无意义；因此不纳入枚举。
                string[] rootNames = { "PhysicalRoot", "ComponentRoot", "ResourceRoot" };
                dynamic dDoc = doc;
                foreach (string rn in rootNames)
                {
                    object root = null;
                    try
                    {
                        switch (rn)
                        {
                            case "PhysicalRoot": root = dDoc.PhysicalRoot; break;
                            case "ComponentRoot": root = dDoc.ComponentRoot; break;
                            case "ResourceRoot": root = dDoc.ResourceRoot; break;
                        }
                    }
                    catch { /* 该版本没有此根 */ }
                    if (root == null) { log(string.Format("[Snapshot] {0} 不存在", rn)); continue; }

                    int before = result.Count;

                    // ── Step A: GetAllDescendants(TxTypeFilter) ───────────
                    TxObjectList kids = null;
                    try
                    {
                        dynamic dRoot = root;
                        kids = dRoot.GetAllDescendants(
                            new TxTypeFilter(typeof(ITxObject))) as TxObjectList;
                    }
                    catch (Exception ex)
                    { log(string.Format("[Snapshot] {0}.GetAllDescendants(ITxObject) 异常: {1}", rn, ex.Message)); }

                    // ── Step B: GetAllDescendants() 无参 ───────────────────
                    if (kids == null || kids.Count == 0)
                    {
                        try
                        {
                            dynamic dRoot = root;
                            kids = dRoot.GetAllDescendants() as TxObjectList;
                        }
                        catch (Exception ex)
                        { log(string.Format("[Snapshot] {0}.GetAllDescendants() 异常: {1}", rn, ex.Message)); }
                    }

                    // ── Step C: 把拿到的对象加入结果（仅 ITxDisplayableObject）─
                    //   只保留实现了 ITxDisplayableObject 接口的对象 —— 它们才有
                    //   Visibility 属性和 Display/Blank 方法。非可显示对象（如操作、
                    //   坐标系 owner 等）加进来也会导致 Blank 失败或读 Visibility 异常。
                    if (kids != null && kids.Count > 0)
                    {
                        foreach (ITxObject obj in kids)
                        {
                            if (obj == null) continue;
                            if (!(obj is ITxDisplayableObject)) continue;
                            int id = System.Runtime.CompilerServices
                                         .RuntimeHelpers.GetHashCode(obj);
                            if (!seen.Add(id)) continue;
                            result.Add(obj);
                        }
                    }
                    else
                    {
                        // ── Step D: 递归 Children 遍历（兜底） ──────────────
                        log(string.Format(
                            "[Snapshot] {0}.GetAllDescendants 返回空，改用递归遍历", rn));
                        try
                        {
                            WalkChildren(root, result, seen, 0, log);
                        }
                        catch (Exception ex)
                        { log(string.Format("[Snapshot] {0} 递归遍历异常: {1}", rn, ex.Message)); }
                    }

                    log(string.Format(
                        "[Snapshot] {0}: 累计 +{1}（该根添加到结果）",
                        rn, result.Count - before));
                }
                log(string.Format("[Snapshot] EnumDisplayableObjects 合计: {0} 个", result.Count));
            }
            catch (Exception ex) { log("[Snapshot] EnumDisplayableObjects 顶层异常: " + ex.Message); }
            return result;
        }

        /// <summary>
        /// 递归遍历 root 的 Children（或 GetAllChildren）填充 result。
        /// 仅加入有 Visible 属性且未见过的对象。深度限制 25 层防环路。
        /// </summary>
        private static void WalkChildren(object root, List<ITxObject> result,
            System.Collections.Generic.HashSet<int> seen, int depth, Action<string> log)
        {
            if (root == null || depth > 25) return;

            TxObjectList kids = null;
            try
            {
                dynamic d = root;
                object raw = null;
                try { raw = d.Children; } catch { }
                if (raw == null) try { raw = d.GetAllChildren(); } catch { }
                kids = raw as TxObjectList;
            }
            catch { }
            if (kids == null || kids.Count == 0) return;

            foreach (ITxObject child in kids)
            {
                if (child == null) continue;
                int id = System.Runtime.CompilerServices
                             .RuntimeHelpers.GetHashCode(child);
                if (seen.Add(id) && child is ITxDisplayableObject)
                    result.Add(child);
                WalkChildren(child, result, seen, depth + 1, log);
            }
        }

        // ═══════════════════════════════════════════════════════════════════
        //  显示控制 API（基于 ITxDisplayableObject 官方接口）
        //
        //  Tecnomatix.Engineering.ITxDisplayableObject 接口提供：
        //    - Display()   : 显示
        //    - Blank()     : 隐藏
        //    - Visibility  : 读取当前可见状态（TxDisplayableObjectVisibility 枚举）
        //                    ALL     = 对象及所有子对象可见
        //                    NONE    = 对象及所有子对象都隐藏
        //                    PARTIAL = 子对象部分可见（仅集合对象可能返回）
        //    - DisplayMode : 显示模式（线框/着色等），不用于 visibility 判断
        //
        //  任何实现 ITxDisplayableObject 的对象（如 TxComponent、TxPart、TxFrame、
        //  TxRobot、TxDevice 等）均可使用；操作类（TxOperation）不实现此接口。
        // ═══════════════════════════════════════════════════════════════════

        private static bool _visErrorLogged = false;

        /// <summary>
        /// 如果 obj 实现 ITxDisplayableObject，则返回 cast 后的接口；否则 null。
        /// </summary>
        private static ITxDisplayableObject AsDisplayable(ITxObject obj)
        {
            return obj as ITxDisplayableObject;
        }

        /// <summary>
        /// 读取对象当前可见状态：true = 至少部分可见（ALL 或 PARTIAL），false = 完全隐藏（NONE）。
        /// 对不支持 Visibility 的对象返回 null（调用方决定如何处理）。
        /// </summary>
        private static bool? GetObjVisible(ITxObject obj, Action<string> log)
        {
            var d = AsDisplayable(obj);
            if (d == null) return null;
            try
            {
                TxDisplayableObjectVisibility v = d.Visibility;
                // ALL(0)/PARTIAL(1)/NONE(2) 枚举值的实际数值取决于 SDK 版本，
                // 直接按枚举名比较最稳妥
                return v != TxDisplayableObjectVisibility.None;
            }
            catch (Exception ex)
            {
                if (!_visErrorLogged)
                {
                    _visErrorLogged = true;
                    (log ?? Nop)(string.Format(
                        "[Snapshot] ⚠ 读取 Visibility 失败 [{0}]: {1}（后续同类失败静默）",
                        obj.GetType().Name, ex.Message));
                }
                return null;
            }
        }

        /// <summary>
        /// 对 obj 调用 Display() 或 Blank()。返回 true 表示调用成功。
        /// </summary>
        private static bool SetObjVisible(ITxObject obj, bool visible, Action<string> log)
        {
            var d = AsDisplayable(obj);
            if (d == null) return false;
            try
            {
                if (visible) d.Display();
                else d.Blank();
                return true;
            }
            catch (Exception ex)
            {
                if (!_visErrorLogged)
                {
                    _visErrorLogged = true;
                    (log ?? Nop)(string.Format(
                        "[Snapshot] ⚠ Display/Blank 调用失败 [{0}]: {1}（后续同类失败静默）",
                        obj.GetType().Name, ex.Message));
                }
                return false;
            }
        }

        /// <summary>向后兼容重载（不带 log）。</summary>
        private static bool SetObjVisible(ITxObject obj, bool visible)
        {
            return SetObjVisible(obj, visible, Nop);
        }


        // ── Section 9 end ─────────────────────────────────────────────────────
    }
}