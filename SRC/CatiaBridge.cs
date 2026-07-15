// CatiaBridge.cs  —  C# 7.3
// CATIA V5 COM 互操作：导出点球 / 导出插枪
//
// 依赖 COM 引用（位于 CATIA 安装目录 intel_a\code\bin\）：
//   INFITF.dll / MECMOD.dll / HybridShapeTypeLib.dll / ProductStructureTypeLib.dll
//
// 修改说明：
// 1. 保持所有公开方法签名不变（Connect, ExportBalls, ExportGuns）。
// 2. 点球导出：在几何集中创建由原点 + 三条直线段组成的“基准坐标系”（通过 AddNewLinePtPt）。
// 3. 插枪导出：勾选“导出TCP坐标”时，为每个焊枪操作创建公共零件，
//    在其中为每个焊点创建由原点和三轴线表示的 TCP 坐标系。
// 4. 解决 AddNewLinePtDir 重载错误，统一使用 AddNewLinePtPt 生成轴线。
// 5. Connect() 不再自动启动 CATIA：仅尝试附加到已运行的 CATIA 实例，
//    未检测到时返回 false 并提示用户先打开 CATIA。
// 6. 其他资源管理、性能优化与异常处理。
// 7. 修复 ExportToCurrentDoc=true + 活动文档为 Product 时新建子 Part 失败：
//    AddNewComponent("Part",...) 不会切换 _catia.ActiveDocument，
//    必须经 newProd.ReferenceProduct.Parent 取宿主 PartDocument。
//    新增 ResolvePartDocFromComponent + TryFindNewlyAddedPartDoc 辅助方法。
//    同时把 ExportGuns 里 TCP 公共零件的同类逻辑一起修了。
// 8. ExportToCurrentDoc=true 且活动文档已是 Part 时，若用户填了自定义零件名，
//    现在会应用到当前 Part 上（仅 set_Name，不动 PartNumber），与其它分支一致。

using HybridShapeTypeLib;
using INFITF;
using MECMOD;
using ProductStructureTypeLib;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms.DataVisualization.Charting;
using Tecnomatix.Engineering;
using SysDir = System.IO.Directory;
using SysFile = System.IO.File;
using SysPath = System.IO.Path;

namespace TxTools.ExportGun
{
    public enum ExportFormat { Xml3d, CATProduct }
    public enum BallExportOption { TrajectoryAndBall, TrajectoryOnly, BallOnly }

    /// <summary>
    /// 焊枪导出模式（体积 vs 命名的权衡）。
    /// SharedGeometry：Copy+Paste 共享 CGR 几何，3DXML 只写一份几何；
    ///                 CATIA COM 层不允许改 CGR 实例名，所有实例名保持 CGR 原名（如 FFM130-....1/.2/.3）
    /// IndependentNaming：每个焊点复制一份独立的 CGR 副本，通过文件名让 CATIA 生成不同的 Reference；
    ///                 实例名 = 焊点名，代价是 3DXML 体积 ≈ 焊点数 × CGR
    /// </summary>
    public enum GunExportMode
    {
        SharedGeometry,
        IndependentNaming
    }

    public class GunExportParams
    {
        public List<OperationInfo> Operations;
        public bool ExportTCP;
        public bool GunOriginAtTCP;
        public string CustomModelPath;
        public Dictionary<string, string> PerOpModelPaths;
        public string CustomProductName;
        public ExportFormat Format;
        public string OutputPath;
        public double[] RefMatrix;
        public string RefName;
        public PointType PointFilter;
        public bool UseMfgName;

        // —— 选中点白名单（key=PointKey(opName,ptName)）；null=不过滤(全部) ——
        public HashSet<string> SelectedKeys;

        // —— TCP 覆盖：TcpCustomMatrix 优先；否则按 TcpName 在该操作工具上解析 ——
        //    两者都为空 → 沿用 GunInfo 中已解析的默认 TCP
        public string TcpName;
        public double[] TcpCustomMatrix;

        // —— 焊枪导出模式（默认共享几何，体积最小） ——
        public GunExportMode ExportMode;
    }

    public class BallExportParams
    {
        public List<OperationInfo> Operations;
        public bool ExportToCurrentDoc;
        public BallExportOption Option;
        public double BallDiameter;
        public string OutputPath;
        public PointType PointFilter;
        public bool UseMfgName;
        public string GeomSetName;
        public string NamePrefix;
        public string CustomPartName;
        public double[] RefMatrix;
        public string RefName;

        // —— 选中点白名单（key=PointKey(opName,ptName)）；null=不过滤(全部) ——
        public HashSet<string> SelectedKeys;
    }

    public class ExportProgress
    {
        public int Total;
        public int Current;
        public string CurrentItem;
    }

    public class CatiaBridge : IDisposable
    {
        private INFITF.Application _catia;
        private bool _disposed;

        // ════════════════════════════════════════════════════════════
        //  选中点过滤辅助
        // ════════════════════════════════════════════════════════════
        /// <summary>选中点唯一键：操作名 + 点名（与 UI 侧保持一致）。</summary>
        public static string PointKey(string opName, string ptName)
        {
            return (opName ?? "") + "\u0001" + (ptName ?? "");
        }

        /// <summary>白名单为 null 时视为全选；否则按 key 判断。</summary>
        private static bool IsPointSelected(HashSet<string> keys, string opName, string ptName)
        {
            return keys == null || keys.Contains(PointKey(opName, ptName));
        }

        /// <summary>计算 TCP 相对工具的局部偏移 = Inv(toolWorld) * tcpWorld。</summary>
        private static double[] ComputeTcpRelTool(double[] toolWorld, double[] tcpWorld)
        {
            try
            {
                if (toolWorld == null || tcpWorld == null) return null;
                var rel = TxTransformation.Multiply(
                    PsReader.ArrToTxPublic(toolWorld).Inverse,
                    PsReader.ArrToTxPublic(tcpWorld));
                return PsReader.TxToArr(rel);
            }
            catch { return null; }
        }

        /// <summary>
        /// 尝试连接到 <b>已经运行</b> 的 CATIA V5 实例。
        /// 若 CATIA 未启动，本方法不会自动启动 CATIA，而是返回 false 并通过 error 给出提示。
        /// 调用方在失败时必须停止后续 Export* 调用。
        /// </summary>
        public bool Connect(out string error)
        {
            error = null;
            try
            {
                _catia = (INFITF.Application)Marshal.GetActiveObject("CATIA.Application");
                // 仅在成功附加到已运行实例时尝试设置可见性，避免任何额外副作用。
                try { _catia.Visible = true; } catch { }
                return true;
            }
            catch (COMException)
            {
                // ROT 中未找到 CATIA.Application（典型 HRESULT MK_E_UNAVAILABLE 0x800401E3），
                // 说明 CATIA 没有运行；按需求不再自动启动 CATIA。
                _catia = null;
                error = "未检测到正在运行的 CATIA V5，请先手动打开 CATIA 后再执行此操作。";
                return false;
            }
            catch (Exception ex)
            {
                _catia = null;
                error = "连接 CATIA V5 失败：" + ex.Message + "。请确认 CATIA V5 已完全启动后再试。";
                return false;
            }
        }

        // ════════════════════════════════════════════════════════════
        //  导出点球（含基准坐标系）
        // ════════════════════════════════════════════════════════════
        public void ExportBalls(BallExportParams p, Action<ExportProgress> onProgress, Action<string> onLog)
        {
            if (_catia == null) throw new InvalidOperationException("CATIA 未连接");

            INFITF.Window initialWindow = null;
            try { initialWindow = _catia.ActiveWindow; } catch { }

            string refInfo = p.RefMatrix != null ? p.RefName ?? "自定义参考系" : "世界坐标系";
            onLog($"[Catia] 导出点球，参考坐标系：{refInfo}");

            string customPartName = string.IsNullOrWhiteSpace(p.CustomPartName) ? null : p.CustomPartName.Trim();

            var allPts = new List<PointInfo>();
            foreach (var op in p.Operations)
            {
                if (op?.Points == null) continue;
                foreach (var pt in op.Points)
                    if (IsPointSelected(p.SelectedKeys, op.Name, pt.Name))
                        allPts.Add(pt);
            }
            if (allPts.Count == 0) { onLog("  ! 无可导出的点（请检查勾选）"); return; }

            int total = allPts.Count;
            string geomSetName = string.IsNullOrEmpty(p.GeomSetName) ? "Geometry_Spheres" : p.GeomSetName;
            string namePrefix = string.IsNullOrEmpty(p.NamePrefix) ? "SPHERE" : p.NamePrefix;
            double radius = p.BallDiameter / 2.0;

            PartDocument partDoc = null;
            try
            {
                // —— 获取/创建 PartDocument ——
                if (p.ExportToCurrentDoc)
                {
                    Document active = _catia.ActiveDocument;
                    if (active is PartDocument pd)
                    {
                        partDoc = pd;
                        // 用户在 PartDocument 下指定了自定义零件名，应用到当前 Part 上
                        // 仅改 Part.Name，不动 PartNumber，避免破坏用户已有的零件号体系
                        if (customPartName != null)
                            try { partDoc.Part.set_Name(customPartName); } catch { }
                        onLog("[Catia] 使用当前 Part：" + active.get_Name());
                    }
                    else if (active is ProductDocument prodDoc)
                    {
                        // 在 Product 下新建 Part 子组件
                        string newPartNum = customPartName ?? "";
                        Product newProd = prodDoc.Product.Products.AddNewComponent("Part", newPartNum);
                        if (customPartName != null) try { newProd.set_Name(customPartName); } catch { }

                        // 关键修复：AddNewComponent 不会切换 _catia.ActiveDocument，
                        // 直接 _catia.ActiveDocument as PartDocument 会拿到 null（仍是 Product 文档）。
                        // 正确做法：从新建的 Product 上经 ReferenceProduct.Parent 取宿主 PartDocument。
                        partDoc = ResolvePartDocFromComponent(newProd, onLog);

                        // 退化路径：极个别版本若 ReferenceProduct.Parent 不可用，尝试从 Documents 集合中按文件名匹配新建文档
                        if (partDoc == null) partDoc = TryFindNewlyAddedPartDoc(newProd, onLog);

                        if (partDoc != null && customPartName != null)
                            try { partDoc.Part.set_Name(customPartName); } catch { }

                        onLog("[Catia] 在 Product 下新建 Part：" + (customPartName ?? newProd.get_PartNumber()));
                    }
                }
                else
                {
                    partDoc = (PartDocument)_catia.Documents.Add("Part");
                    if (customPartName != null)
                    {
                        try { partDoc.Part.set_Name(customPartName); } catch { }
                        try { partDoc.Product.set_PartNumber(customPartName); } catch { }
                        try { partDoc.Product.set_Name(customPartName); } catch { }
                    }
                    onLog("[Catia] 新建 Part：" + partDoc.get_Name());
                }

                if (partDoc == null) { onLog("[Catia] x 无法获取/创建 PartDocument"); return; }

                Part part = partDoc.Part;
                HybridShapeFactory sf = (HybridShapeFactory)part.HybridShapeFactory;

                // —— 获取/创建几何集 ——
                HybridBody geomSet = null;
                HybridBodies bodies = part.HybridBodies;
                for (int i = 1; i <= bodies.Count; i++)
                {
                    HybridBody b = bodies.Item(i);
                    if (b.get_Name() == geomSetName) { geomSet = b; break; }
                }
                if (geomSet == null) { geomSet = bodies.Add(); geomSet.set_Name(geomSetName); }

                // [新增] 创建基准坐标系（由原点 + 三条线段表示）
                double[] identity = new double[] { 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1 };
                CreateAxisVisual(part, geomSet, sf, p.RefMatrix ?? identity, "参考坐标系", onLog);

                // —— 阶段1：创建点 ——
                onLog($"[Catia] 创建 {total} 个点...");
                var createdPoints = new List<HybridShapePointCoord>();
                for (int i = 0; i < allPts.Count; i++)
                {
                    try
                    {
                        double[] m = allPts[i].TCPMatrix;
                        if (p.RefMatrix != null && !PsReader.IsIdentity(p.RefMatrix))
                            m = PsReader.ToRelative(m, p.RefMatrix);
                        HybridShapePointCoord pt = sf.AddNewPointCoord(m[3], m[7], m[11]);
                        geomSet.AppendHybridShape(pt);
                        createdPoints.Add(pt);
                    }
                    catch (Exception ex) { onLog($"  ! 点{i + 1} 异常：{ex.Message}"); continue; }
                    onProgress?.Invoke(new ExportProgress { Total = total * 2, Current = i + 1, CurrentItem = allPts[i].Name });
                }
                onLog($"[Catia] 已创建 {createdPoints.Count} 个点");
                part.Update();

                // —— 阶段2：创建球体 ——
                onLog($"[Catia] 创建球体，半径={radius}mm...");
                int ok = 0, fail = 0;
                for (int i = 0; i < createdPoints.Count; i++)
                {
                    string ballName = Sanitize(namePrefix + "_" + (i < allPts.Count ? allPts[i].Name : (i + 1).ToString()));
                    try
                    {
                        Reference ptRef = part.CreateReferenceFromObject(createdPoints[i]);
                        HybridShapeSphere sphere = sf.AddNewSphere(ptRef, null, radius, -90.0, 90.0, 0.0, 360.0);
                        geomSet.AppendHybridShape(sphere);
                        sphere.set_Name(ballName);
                        ok++;
                        Marshal.ReleaseComObject(ptRef);
                    }
                    catch (Exception ex) { onLog($"  ! [{ballName}] {ex.Message}"); fail++; continue; }
                    onProgress?.Invoke(new ExportProgress { Total = createdPoints.Count * 2, Current = createdPoints.Count + i + 1, CurrentItem = ballName });
                }
                part.Update();
                onLog($"[Catia] 点球完成：成功 {ok}，失败 {fail}");

                if (!p.ExportToCurrentDoc)
                {
                    string outFile = BuildPath(p.OutputPath, Sanitize(customPartName ?? "WeldPoints_Spheres"), "CATPart");
                    partDoc.SaveAs(outFile);
                    onLog("[Catia] 已保存：" + outFile);
                }
                else partDoc.Save();
            }
            catch (Exception ex) { onLog($"[Catia] x 导出点球异常：{ex.Message}"); }
            finally { if (initialWindow != null) try { initialWindow.Activate(); } catch { } }
        }

        // [新增] 创建由原点 + 三轴线段表示的坐标系（使用 AddNewLinePtPt）
        private void CreateAxisVisual(Part part, HybridBody geomSet, HybridShapeFactory sf,
            double[] matrix, string baseName, Action<string> onLog)
        {
            try
            {
                double ox = matrix[3], oy = matrix[7], oz = matrix[11];
                double xx = matrix[0], xy = matrix[4], xz = matrix[8];
                double yx = matrix[1], yy = matrix[5], yz = matrix[9];
                double zx = matrix[2], zy = matrix[6], zz = matrix[10];
                const double LENGTH = 5.0;

                // 原点
                HybridShapePointCoord origin = sf.AddNewPointCoord(ox, oy, oz);
                geomSet.AppendHybridShape(origin);
                origin.set_Name(baseName + "_原点");

                // 创建 X 轴线
                CreateAxisLine(part, geomSet, sf, origin, ox, oy, oz, xx, xy, xz, LENGTH, baseName + "_X轴");
                // 创建 Y 轴线
                CreateAxisLine(part, geomSet, sf, origin, ox, oy, oz, yx, yy, yz, LENGTH, baseName + "_Y轴");
                // 创建 Z 轴线
                CreateAxisLine(part, geomSet, sf, origin, ox, oy, oz, zx, zy, zz, LENGTH, baseName + "_Z轴");

                part.Update();
                onLog($"  基准坐标系 '{baseName}' 已创建（原点 ({ox:F3},{oy:F3},{oz:F3})）");
            }
            catch (Exception ex) { onLog($"  ! 创建基准坐标系失败：{ex.Message}"); }
        }

        // 辅助：创建一条轴线（从原点沿方向创建端点，再连直线）
        private void CreateAxisLine(Part part, HybridBody geomSet, HybridShapeFactory sf,
            HybridShapePointCoord origin,
            double ox, double oy, double oz,
            double dx, double dy, double dz, double length, string name)
        {
            double ex = ox + dx * length, ey = oy + dy * length, ez = oz + dz * length;
            HybridShapePointCoord endPt = sf.AddNewPointCoord(ex, ey, ez);
            geomSet.AppendHybridShape(endPt);

            Reference refOrigin = part.CreateReferenceFromObject(origin);
            Reference refEnd = part.CreateReferenceFromObject(endPt);
            HybridShapeLinePtPt line = sf.AddNewLinePtPt(refOrigin, refEnd);
            geomSet.AppendHybridShape(line);
            line.set_Name(name);

            Marshal.ReleaseComObject(refEnd);
            Marshal.ReleaseComObject(refOrigin);
        }

        // ════════════════════════════════════════════════════════════
        //  从一个新建的 Part 子组件解析其宿主 PartDocument
        //
        //  背景：ProductDocument.Product.Products.AddNewComponent("Part", ...)
        //        会创建一个新的 Part 子组件，但 _catia.ActiveDocument 仍指向
        //        外层的 Product，直接 cast 为 PartDocument 会得到 null。
        //
        //  正确链路（CATIA V5 标准）：
        //        Product (子组件) → ReferenceProduct → Parent → 即对应的 PartDocument
        // ════════════════════════════════════════════════════════════
        private PartDocument ResolvePartDocFromComponent(Product component, Action<string> onLog)
        {
            if (component == null) return null;
            try
            {
                Product refProd = component.ReferenceProduct;   // 实际承载几何的 Product
                if (refProd == null) return null;
                Document parent = refProd.Parent as Document;   // 该 Product 所属文档
                PartDocument pd = parent as PartDocument;
                if (pd != null) return pd;
                if (parent != null && onLog != null)
                    onLog("    ! ReferenceProduct.Parent 不是 PartDocument，实际类型：" + parent.GetType().Name);
            }
            catch (Exception ex)
            {
                if (onLog != null) onLog("    ! ReferenceProduct.Parent 解析失败：" + ex.Message);
            }
            return null;
        }

        // 退化路径：遍历 Documents 集合按候选名匹配（仅在 ReferenceProduct.Parent 失败时使用）
        private PartDocument TryFindNewlyAddedPartDoc(Product component, Action<string> onLog)
        {
            if (component == null || _catia == null) return null;
            string partNumber = null;
            string name = null;
            try { partNumber = component.get_PartNumber(); } catch { }
            try { name = component.get_Name(); } catch { }

            try
            {
                Documents docs = _catia.Documents;
                int count = docs.Count;
                // 从后向前找，新建的文档通常排在末尾
                for (int i = count; i >= 1; i--)
                {
                    Document d;
                    try { d = docs.Item(i); } catch { continue; }
                    PartDocument pd = d as PartDocument;
                    if (pd == null) continue;

                    string dname = null;
                    try { dname = d.get_Name(); } catch { }
                    // CATIA 文档名通常形如 "PartXX.CATPart" 或 "<PartNumber>.CATPart"
                    if (!string.IsNullOrEmpty(dname))
                    {
                        string stem = SysPath.GetFileNameWithoutExtension(dname);
                        if (!string.IsNullOrEmpty(partNumber) &&
                            string.Equals(stem, partNumber, StringComparison.OrdinalIgnoreCase)) return pd;
                        if (!string.IsNullOrEmpty(name) &&
                            string.Equals(stem, name, StringComparison.OrdinalIgnoreCase)) return pd;
                    }
                }
                if (onLog != null) onLog("    ! 遍历 Documents 未匹配到新建 Part 文档");
            }
            catch (Exception ex)
            {
                if (onLog != null) onLog("    ! Documents 集合遍历失败：" + ex.Message);
            }
            return null;
        }

        // ════════════════════════════════════════════════════════════
        //  导出插枪（分派器）
        // ════════════════════════════════════════════════════════════
        public void ExportGuns(GunExportParams p, Action<ExportProgress> onProgress, Action<string> onLog)
        {
            if (_catia == null) throw new InvalidOperationException("CATIA 未连接");

            if (p.ExportMode == GunExportMode.IndependentNaming)
            {
                onLog?.Invoke("[Catia] 导出模式：独立命名（每焊点独立几何，实例名=焊点名，体积随焊点数线性增长）");
                ExportGunsIndependent(p, onProgress, onLog);
            }
            else
            {
                onLog?.Invoke("[Catia] 导出模式：共享几何（Copy+Paste 实例复用，体积最小；实例名保持 CGR 原名）");
                ExportGunsShared(p, onProgress, onLog);
            }
        }

        // ════════════════════════════════════════════════════════════
        //  模式 A：共享几何（体积最小，实例名保持 CGR 原名）
        // ════════════════════════════════════════════════════════════
        private void ExportGunsShared(GunExportParams p, Action<ExportProgress> onProgress, Action<string> onLog)
        {
            // 保存并关闭弹窗 / 刷新
            bool savedDisplayAlerts = true;
            bool savedRefreshDisplay = true;
            try { savedDisplayAlerts = (bool)((dynamic)_catia).DisplayFileAlerts; ((dynamic)_catia).DisplayFileAlerts = false; } catch { }
            try { savedRefreshDisplay = (bool)((dynamic)_catia).RefreshDisplay; ((dynamic)_catia).RefreshDisplay = false; } catch { }

            INFITF.Window initialWindow = null;
            try { initialWindow = _catia.ActiveWindow; } catch { }

            try
            {
                // 1. Product 文档：优先复用活动 Product，否则新建
                ProductDocument productDoc;
                if (_catia.ActiveDocument is ProductDocument existingPd)
                {
                    productDoc = existingPd;
                    onLog?.Invoke("[Catia] 使用当前 Product 文档");
                }
                else
                {
                    productDoc = (ProductDocument)_catia.Documents.Add("Product");
                    if (!string.IsNullOrWhiteSpace(p.CustomProductName))
                        try { productDoc.Product.set_PartNumber(p.CustomProductName.Trim()); } catch { }
                    onLog?.Invoke("[Catia] 新建 Product 文档");
                }

                Product rootProduct = productDoc.Product;
                Products rootProducts = rootProduct.Products;
                Selection sel = productDoc.Selection;

                // 2. 统计总数
                int total = 0;
                foreach (var op in p.Operations)
                {
                    if (op.Gun == null || op.Points == null) continue;
                    foreach (var pt in op.Points)
                        if (IsPointSelected(p.SelectedKeys, op.Name, pt.Name)) total++;
                }
                int current = 0;

                // 3. Reference 复用缓存（跨 Operation）
                var sourceCache = new Dictionary<string, Product>(StringComparer.OrdinalIgnoreCase);

                // 4. 主循环
                foreach (var op in p.Operations)
                {
                    var gun = op.Gun;
                    if (gun == null || op.Points == null) continue;

                    string modelPath = ResolveModelPath(p, op, gun);
                    if (modelPath == null)
                    {
                        onLog?.Invoke($"  ! [{op.Name}] 未找到模型文件");
                        continue;
                    }

                    var opPts = new List<PointInfo>();
                    foreach (var pt in op.Points)
                        if (IsPointSelected(p.SelectedKeys, op.Name, pt.Name)) opPts.Add(pt);
                    if (opPts.Count == 0) continue;

                    // TCP 覆盖
                    ApplyTcpOverride(p, op, gun, onLog);

                    // 容器：同名冲突自动加后缀
                    string containerName = GetUniqueChildName(rootProducts, op.Name);
                    Product container = rootProducts.AddNewComponent("Product", containerName);
                    try { container.set_Name(containerName); } catch { }
                    try { container.set_PartNumber(containerName); } catch { }
                    Products targetProducts = container.Products;

                    // 准备源实例
                    Product sourceInst;
                    int startIdx = 0;

                    if (!sourceCache.TryGetValue(modelPath, out sourceInst))
                    {
                        try
                        {
                            targetProducts.AddComponentsFromFiles(new object[] { modelPath }, "All");
                        }
                        catch (Exception ex)
                        {
                            onLog?.Invoke($"    x [{op.Name}] 加载 CGR 失败: {ex.Message}");
                            continue;
                        }

                        sourceInst = targetProducts.Item(targetProducts.Count);
                        sourceCache[modelPath] = sourceInst;

                        // 首焊点：源实例定位（不改名）
                        var pt0 = opPts[0];
                        current++;
                        onProgress?.Invoke(new ExportProgress { Total = total, Current = current, CurrentItem = pt0.Name });

                        try
                        {
                            double[] placed0 = ComputePlaced(p, gun, pt0);
                            sourceInst.Position.SetComponents(ToSetComp(placed0));
                        }
                        catch (Exception ex)
                        {
                            onLog?.Invoke($"    x 首实例设置失败 [{pt0.Name}]: {ex.Message}");
                        }

                        startIdx = 1;
                        onLog?.Invoke($"    Reference 加载: {System.IO.Path.GetFileName(modelPath)}");
                    }
                    else
                    {
                        onLog?.Invoke($"    Reference 复用: {System.IO.Path.GetFileName(modelPath)}");
                    }

                    // Copy 源
                    try
                    {
                        sel.Clear();
                        sel.Add(sourceInst);
                        sel.Copy();
                    }
                    catch (Exception ex)
                    {
                        onLog?.Invoke($"    x [{op.Name}] Copy 源实例失败: {ex.Message}");
                        continue;
                    }

                    // Paste 剩余焊点：每次 Paste 后立即定位
                    for (int i = startIdx; i < opPts.Count; i++)
                    {
                        var pt = opPts[i];
                        current++;
                        onProgress?.Invoke(new ExportProgress { Total = total, Current = current, CurrentItem = pt.Name });

                        try
                        {
                            int beforeCount = targetProducts.Count;

                            sel.Clear();
                            sel.Add(container);
                            sel.Paste();

                            int afterCount = targetProducts.Count;
                            if (afterCount <= beforeCount)
                                throw new Exception("Paste 后实例数未增加");

                            Product inst = targetProducts.Item(afterCount);
                            double[] placed = ComputePlaced(p, gun, pt);
                            inst.Position.SetComponents(ToSetComp(placed));
                        }
                        catch (Exception ex)
                        {
                            onLog?.Invoke($"    x 插入失败 [{pt.Name}]: {ex.Message}");
                        }
                    }

                    onLog?.Invoke($"  [{op.Name}] 已插入 {opPts.Count} 个焊点实例（共享 Reference）");
                }

                try { rootProduct.Update(); } catch { }
                onLog?.Invoke($"[Catia] 插枪完成（共享几何：{sourceCache.Count} 份几何，{current}/{total} 个实例）");
            }
            finally
            {
                try { ((dynamic)_catia).RefreshDisplay = savedRefreshDisplay; } catch { }
                try { ((dynamic)_catia).DisplayFileAlerts = savedDisplayAlerts; } catch { }
                if (initialWindow != null) try { initialWindow.Activate(); } catch { }
            }
        }

        // ════════════════════════════════════════════════════════════
        //  模式 B：独立命名（每焊点独立 CGR 副本，实例名=焊点名）
        // ════════════════════════════════════════════════════════════
        private void ExportGunsIndependent(GunExportParams p, Action<ExportProgress> onProgress, Action<string> onLog)
        {
            INFITF.Window initialWindow = null;
            try { initialWindow = _catia.ActiveWindow; } catch { }

            // 1. Product 文档：优先复用活动 Product，否则新建
            ProductDocument productDoc;
            if (_catia.ActiveDocument is ProductDocument existingPd)
            {
                productDoc = existingPd;
                onLog?.Invoke("[Catia] 使用当前 Product 文档");
            }
            else
            {
                productDoc = (ProductDocument)_catia.Documents.Add("Product");
                if (!string.IsNullOrWhiteSpace(p.CustomProductName))
                    try { productDoc.Product.set_PartNumber(p.CustomProductName.Trim()); } catch { }
                onLog?.Invoke("[Catia] 新建 Product 文档");
            }

            Product rootProduct = productDoc.Product;
            Products rootProducts = rootProduct.Products;

            // 2. 统计总数
            int total = 0;
            foreach (var op in p.Operations)
            {
                if (op.Gun == null || op.Points == null) continue;
                foreach (var pt in op.Points)
                    if (IsPointSelected(p.SelectedKeys, op.Name, pt.Name)) total++;
            }
            int current = 0;

            // 3. 临时目录：为每个焊点生成独立命名的 CGR 副本
            string tempDir = SysPath.Combine(SysPath.GetTempPath(), "CatiaGunExport_" + Guid.NewGuid().ToString("N"));
            SysDir.CreateDirectory(tempDir);

            try
            {
                foreach (var op in p.Operations)
                {
                    var gun = op.Gun;
                    if (gun == null || op.Points == null) continue;

                    string modelPath = ResolveModelPath(p, op, gun);
                    if (modelPath == null)
                    {
                        onLog?.Invoke($"  ! [{op.Name}] 未找到模型文件");
                        continue;
                    }

                    var opPts = new List<PointInfo>();
                    foreach (var pt in op.Points)
                        if (IsPointSelected(p.SelectedKeys, op.Name, pt.Name)) opPts.Add(pt);
                    if (opPts.Count == 0) continue;

                    // TCP 覆盖
                    ApplyTcpOverride(p, op, gun, onLog);

                    // 容器
                    string containerName = GetUniqueChildName(rootProducts, op.Name);
                    Product container = rootProducts.AddNewComponent("Product", containerName);
                    try { container.set_Name(containerName); } catch { }
                    try { container.set_PartNumber(containerName); } catch { }
                    Products targetProducts = container.Products;

                    string modelExt = SysPath.GetExtension(modelPath);
                    onLog?.Invoke($"  [{op.Name}] {opPts.Count} 个焊点，模型：{SysPath.GetFileName(modelPath)}");

                    foreach (var pt in opPts)
                    {
                        current++;
                        onProgress?.Invoke(new ExportProgress { Total = total, Current = current, CurrentItem = pt.Name });

                        try
                        {
                            // 关键：为每个焊点复制一份 CGR，用焊点名命名
                            // CATIA 加载后 Reference 名 = 文件名 = 焊点名，实例名同步为焊点名
                            string safeName = Sanitize(pt.Name);
                            string ptFile = SysPath.Combine(tempDir, safeName + modelExt);
                            SysFile.Copy(modelPath, ptFile, true);

                            targetProducts.AddComponentsFromFiles(new object[] { ptFile }, "All");
                            Product inst = targetProducts.Item(targetProducts.Count);
                            try { inst.set_Name(safeName); } catch { }

                            double[] placed = ComputePlaced(p, gun, pt);
                            inst.Position.SetComponents(ToSetComp(placed));
                        }
                        catch (Exception ex)
                        {
                            onLog?.Invoke($"    x 插入失败 [{pt.Name}]: {ex.Message}");
                        }
                    }
                }

                try { rootProduct.Update(); } catch { }
                onLog?.Invoke($"[Catia] 插枪完成（独立命名：{current}/{total} 个实例，几何数=实例数）");
            }
            finally
            {
                try { if (SysDir.Exists(tempDir)) SysDir.Delete(tempDir, true); }
                catch { onLog?.Invoke($"    ! 临时目录清理失败：{tempDir}"); }
                if (initialWindow != null) try { initialWindow.Activate(); } catch { }
            }
        }

        // ---------- 辅助：解析模型路径（自定义 > 每 Op 覆盖 > GunInfo 默认） ----------
        private static string ResolveModelPath(GunExportParams p, OperationInfo op, GunInfo gun)
        {
            if (p.PerOpModelPaths != null &&
                p.PerOpModelPaths.TryGetValue(op.Name, out string perOpPath) &&
                !string.IsNullOrEmpty(perOpPath) && SysFile.Exists(perOpPath))
                return perOpPath;

            if (!string.IsNullOrEmpty(p.CustomModelPath) && SysFile.Exists(p.CustomModelPath))
                return p.CustomModelPath;

            if (!string.IsNullOrEmpty(gun.ModelPath) && SysFile.Exists(gun.ModelPath))
                return gun.ModelPath;

            return null;
        }

        // ---------- 辅助：TCP 覆盖（GunOriginAtTCP=true 且指定了 TCP 时应用） ----------
        private static void ApplyTcpOverride(GunExportParams p, OperationInfo op, GunInfo gun, Action<string> onLog)
        {
            if (!p.GunOriginAtTCP) return;
            if (p.TcpCustomMatrix == null && string.IsNullOrEmpty(p.TcpName)) return;

            double[] tcpWorld = p.TcpCustomMatrix ?? PsReader.ResolveTcpWorldByName(op, p.TcpName, onLog);
            if (tcpWorld == null)
            {
                onLog?.Invoke($"    ! 未能解析所选 TCP（{p.TcpName}），改用默认 TCP");
                return;
            }

            gun.TcpWorldMatrix = tcpWorld;
            double[] rel = ComputeTcpRelTool(gun.ToolMatrix, tcpWorld);
            if (rel != null) gun.TcpRelTool = rel;
            onLog?.Invoke($"    TCP 覆盖：{(p.TcpCustomMatrix != null ? "自定义坐标" : p.TcpName)}");
        }

        // ---------- 辅助：容器同名冲突处理（自动加 _2 / _3 后缀） ----------
        private static string GetUniqueChildName(Products parent, string baseName)
        {
            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int cnt = 0;
            try { cnt = parent.Count; } catch { }
            for (int i = 1; i <= cnt; i++)
            {
                try
                {
                    Product ch = parent.Item(i);
                    try { existing.Add(ch.get_Name()); } catch { }
                    try { existing.Add(ch.get_PartNumber()); } catch { }
                }
                catch { }
            }

            if (!existing.Contains(baseName)) return baseName;

            for (int i = 2; i <= 999; i++)
            {
                string tryName = baseName + "_" + i;
                if (!existing.Contains(tryName)) return tryName;
            }
            return baseName + "_" + Guid.NewGuid().ToString("N").Substring(0, 6);
        }


        // ---------- 辅助：位置矩阵计算 ----------
        private double[] ComputePlaced(GunExportParams p, dynamic gun, PointInfo pt)
        {
            double[] placed = p.GunOriginAtTCP
                ? CalcGunPlacedMatrix(pt.TCPMatrix, gun.ToolMatrix, gun.TcpWorldMatrix, gun.TcpRelTool)
                : pt.TCPMatrix;

            if (p.RefMatrix != null && !PsReader.IsIdentity(p.RefMatrix))
                placed = PsReader.ToRelative(placed, p.RefMatrix);

            return placed;
        }

        // ════════════════════════════════════════════════════════════
        //  矩阵计算
        // ════════════════════════════════════════════════════════════
        private static double[] CalcGunPlacedMatrix(double[] weldPt, double[] toolWorld, double[] tcpWorld, double[] tcpRelTool)
        {
            try
            {
                var weldT = PsReader.ArrToTxPublic(weldPt);
                TxTransformation tRel;
                if (tcpRelTool != null && !PsReader.IsIdentity(tcpRelTool))
                    tRel = PsReader.ArrToTxPublic(tcpRelTool);
                else if (toolWorld != null && tcpWorld != null)
                    tRel = TxTransformation.Multiply(PsReader.ArrToTxPublic(toolWorld).Inverse, PsReader.ArrToTxPublic(tcpWorld));
                else return weldPt;
                return PsReader.TxToArr(TxTransformation.Multiply(weldT, tRel.Inverse));
            }
            catch { return weldPt; }
        }

        private static object[] ToSetComp(double[] m) => new object[]
        {
            (object)m[0], (object)m[4], (object)m[8],
            (object)m[1], (object)m[5], (object)m[9],
            (object)m[2], (object)m[6], (object)m[10],
            (object)m[3], (object)m[7], (object)m[11]
        };

        private void InsertMarker(Products products, string name, double[] m)
        {
            try
            {
                Product p = products.AddNewComponent("Part", "");
                p.set_Name(Sanitize(name));
                p.Position.SetComponents(ToSetComp(m));
            }
            catch { }
        }

        // ════════════════════════════════════════════════════════════
        //  工具方法
        // ════════════════════════════════════════════════════════════
        private static string Sanitize(string n)
        {
            if (string.IsNullOrEmpty(n)) return "Item";
            n = Regex.Replace(n, @"[/\\:*?""<>|]", "_");
            return n.Length > 80 ? n.Substring(0, 80) : n;
        }

        private static string BuildPath(string basePath, string name, string ext)
        {
            if (string.IsNullOrEmpty(basePath))
                basePath = SysPath.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "CatiaExport");
            if (!ext.StartsWith(".")) ext = "." + ext;
            if (!SysDir.Exists(basePath)) SysDir.CreateDirectory(basePath);
            return SysPath.Combine(basePath, Sanitize(name) + ext);
        }

        // ════════════════════════════════════════════════════════════
        //  IDisposable
        // ════════════════════════════════════════════════════════════
        public void Dispose()
        {
            if (!_disposed)
            {
                if (_catia != null) { try { Marshal.ReleaseComObject(_catia); } catch { } _catia = null; }
                _disposed = true;
            }
            GC.SuppressFinalize(this);
        }
        ~CatiaBridge() { Dispose(); }
    }
}