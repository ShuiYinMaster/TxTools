using System;
using System.Collections.Generic;
using System.Reflection;
using Tecnomatix.Engineering;

namespace TxTools.FenceBuilder
{
    public class BuildResult
    {
        public List<TxSolid> CreatedSolids = new List<TxSolid>();
        public int PostCount;
        public int FrameRailCount;
        public int MeshPanelCount;
        public int BaseplateCount;
        public int TextureSuccessCount;
        public int TextureFallbackCount;
        public TxComponent Container;
        public bool UsedActiveModeling;   // true=容器是用户的建模资源(撤销时不删容器), false=新建 Resource
        public List<ITxObject> PanelGroups = new List<ITxObject>();
        public List<ITxObject> PostGroups = new List<ITxObject>();
    }

    /// <summary>
    /// 围栏几何创建器。
    /// 
    /// 资源层级:
    ///   Fence_yyyyMMdd_HHmmss (Resource, TxComponent)
    ///     ├── 所有 TxSolid 直接挂在 Resource 下(立柱/底板/方管/薄板)
    ///     └── 视图组 (TxGroup):
    ///         ├── Post_S0_0, Post_S0_1, ...    每根立柱+底板一个 group
    ///         ├── Post_Corner_5, ...           拐角立柱
    ///         └── Panel_0_0, Panel_0_1, ...    每片网片(5 个 Solid)一个 group
    ///
    /// 每个网片包含独立完整的 5 个几何:下沿+上沿+左边+右边+薄板,4 框首尾相连。
    /// 每根立柱包含 1-2 个几何:立柱本体 (+ 可选底板)。
    /// </summary>
    public static class FenceGeometryBuilder
    {
        public static BuildResult Build(
            FenceLayout layout,
            FenceParameters p,
            string textureFilePath,
            Action<string> log)
        {
            BuildResult result = new BuildResult();
            if (layout == null || layout.Posts.Count + layout.Panels.Count == 0)
            {
                log?.Invoke("[Builder] 空布局,跳过");
                return result;
            }
            AppearanceHelper.ResetTextureLogState();

            // 1) 顶层 Resource 容器
            // 优先级:
            //   - 如果 CreateUnderActiveModeling 启用,先尝试找"当前正在建模"的 component 作为父级
            //     (该 component 的 IsOpenForModeling=true,围栏几何直接挂在它内部)
            //   - 否则在 PhysicalRoot 下 CreateResource 新建独立容器
            TxComponent topComp = null;
            bool useActiveModeling = false;

            if (p.CreateUnderActiveModeling)
            {
                topComp = TryFindActiveModelingComponent(log);
                if (topComp != null)
                {
                    useActiveModeling = true;
                    log?.Invoke("[Builder] 使用当前建模资源: " + TrySafeName(topComp));
                    result.Container = topComp;
                    result.UsedActiveModeling = true;
                }
            }

            if (topComp == null)
            {
                try
                {
                    var root = TxApplication.ActiveDocument.PhysicalRoot;
                    string resName = (string.IsNullOrEmpty(p.GroupName) ? "Fence" : p.GroupName)
                                    + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

                    object creationData = TryCreateCreationData("TxResourceCreationData", resName, log);
                    if (creationData == null)
                    {
                        log?.Invoke("[Builder] ERR 无法构造 TxResourceCreationData");
                        return result;
                    }

                    MethodInfo miCreate = root.GetType().GetMethod("CreateResource", new[] { creationData.GetType() });
                    if (miCreate == null)
                    {
                        foreach (var m in root.GetType().GetMethods())
                        {
                            if (m.Name == "CreateResource" && m.GetParameters().Length == 1)
                            {
                                miCreate = m; break;
                            }
                        }
                    }
                    if (miCreate == null)
                    {
                        log?.Invoke("[Builder] ERR PhysicalRoot.CreateResource 方法未找到");
                        return result;
                    }

                    topComp = miCreate.Invoke(root, new[] { creationData }) as TxComponent;
                    if (topComp == null)
                    {
                        log?.Invoke("[Builder] ERR CreateResource 返回 null");
                        return result;
                    }
                    log?.Invoke("[Builder] 顶层 Resource: " + resName);
                    result.Container = topComp;
                }
                catch (Exception ex)
                {
                    log?.Invoke("[Builder] ERR CreateResource: " + ex.Message);
                    return result;
                }
            }

            // 2) 所有几何统一建在 topComp 内
            //    如果是新建的 Resource,需要先 SetModelingScope;
            //    如果使用了当前正在建模的资源,已经在建模状态,可以省略 SetModelingScope,但调用一次也无害
            try { topComp.SetModelingScope(); }
            catch (Exception ex) { log?.Invoke("[Builder] WARN SetModelingScope: " + ex.Message); }

            try
            {
                // 立柱(含可选底板): 每根立柱创建完后,把"立柱+底板"打包成一个 TxGroup
                int postIdx = 0;
                foreach (PostPlacement post in layout.Posts)
                {
                    int beforeCount = result.CreatedSolids.Count;
                    BuildOnePost(topComp, post, p, layout, result, log);

                    var postSolids = new List<TxSolid>();
                    for (int k = beforeCount; k < result.CreatedSolids.Count; k++)
                        postSolids.Add(result.CreatedSolids[k]);

                    if (postSolids.Count > 0)
                    {
                        // 命名: Post_si_pi (段 -1 = 拐角)
                        string segTag = (post.SegmentIndex >= 0) ? ("S" + post.SegmentIndex) : "Corner";
                        string groupName = "Post_" + segTag + "_" + postIdx;
                        ITxObject group = TryCreateGroup(topComp, groupName, postSolids, log);
                        if (group != null) result.PostGroups.Add(group);
                        postIdx++;
                    }
                }

                // 网片: 每片创建完后,把这片的 Solid 打包成 TxGroup
                for (int i = 0; i < layout.Panels.Count; i++)
                {
                    MeshPanelPlacement panel = layout.Panels[i];
                    int beforeCount = result.CreatedSolids.Count;

                    BuildOnePanel(topComp, panel, p, layout, textureFilePath, result, log);

                    var panelSolids = new List<TxSolid>();
                    for (int k = beforeCount; k < result.CreatedSolids.Count; k++)
                        panelSolids.Add(result.CreatedSolids[k]);

                    if (panelSolids.Count > 0)
                    {
                        string groupName = "Panel_" + panel.SegmentIndex + "_" + panel.IndexInSegment;
                        ITxObject group = TryCreateGroup(topComp, groupName, panelSolids, log);
                        if (group != null) result.PanelGroups.Add(group);
                    }
                }
            }
            finally
            {
                // LineToSolid 既有经验: EndModeling 会抛 "External component has thrown an exception"
                // 并阻塞数秒,但 PS 实际上会自动收尾 scope。完全省略 EndModeling 调用以提速。
                // 使用当前建模资源时也不调,避免打断用户的建模会话。
                // (如果将来发现省略 EndModeling 导致几何不持久化,再恢复)
            }

            log?.Invoke("[Builder] 完成: 立柱=" + result.PostCount +
                        " 方管=" + result.FrameRailCount +
                        " 薄板=" + result.MeshPanelCount +
                        " 底板=" + result.BaseplateCount +
                        " 立柱group=" + result.PostGroups.Count +
                        " 网片group=" + result.PanelGroups.Count +
                        " (纹理OK=" + result.TextureSuccessCount +
                        " 降级=" + result.TextureFallbackCount + ")");
            return result;
        }

        /// <summary>
        /// 调用 PS 官方 API: comp.CreateGroup(TxGroupCreationData) → TxGroup
        /// 用反射构造 TxGroupCreationData(兼容不同 PS 版本的命名空间布局)。
        /// 创建后把 members 加入 group(成员需在同一 component 内)。
        /// </summary>
        private static ITxObject TryCreateGroup(
            TxComponent parent, string name, List<TxSolid> members, Action<string> log)
        {
            try
            {
                object creationData = TryCreateCreationData("TxGroupCreationData", name, log);
                if (creationData == null)
                {
                    log?.Invoke("[Builder] WARN TxGroupCreationData 未找到,group 创建失败");
                    return null;
                }

                // 调 parent.CreateGroup(creationData)
                MethodInfo miCreate = null;
                foreach (var m in parent.GetType().GetMethods())
                {
                    if (m.Name == "CreateGroup" && m.GetParameters().Length == 1)
                    {
                        miCreate = m; break;
                    }
                }
                if (miCreate == null)
                {
                    log?.Invoke("[Builder] WARN parent.CreateGroup(TxGroupCreationData) 方法未找到");
                    return null;
                }

                object group = miCreate.Invoke(parent, new[] { creationData });
                if (group == null)
                {
                    log?.Invoke("[Builder] WARN CreateGroup 返回 null");
                    return null;
                }

                // 把 members 加入 group(尝试多种 API 名)
                if (!TryAddMembersToGroup(group, members, log))
                    log?.Invoke("[Builder] WARN group '" + name + "' 已创建但成员添加失败");

                return group as ITxObject;
            }
            catch (Exception ex)
            {
                log?.Invoke("[Builder] CreateGroup 异常: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 把 solid 列表加入 TxGroup。多路径尝试常见 API 名。
        /// </summary>
        private static bool TryAddMembersToGroup(object group, List<TxSolid> members, Action<string> log)
        {
            Type t = group.GetType();
            // 路径 1: AddMembers(IList<ITxObject>) / AddMembers(TxObjectList)
            string[] methodNames = { "AddMembers", "Add", "AddObjects", "AddItems" };
            foreach (string mname in methodNames)
            {
                // 单参,接受 IEnumerable
                foreach (var mi in t.GetMethods())
                {
                    if (mi.Name != mname) continue;
                    var pars = mi.GetParameters();
                    if (pars.Length != 1) continue;
                    // 尝试传 List<TxSolid> 或 TxObjectList
                    try
                    {
                        // 优先尝试构造 TxObjectList(PS 原生集合)
                        object collection = TryBuildTxObjectList(members);
                        if (collection != null && pars[0].ParameterType.IsAssignableFrom(collection.GetType()))
                        {
                            mi.Invoke(group, new[] { collection });
                            return true;
                        }
                        // 退化:直接传 IList<ITxObject>
                        var asList = new List<ITxObject>();
                        foreach (var s in members) asList.Add(s);
                        if (pars[0].ParameterType.IsAssignableFrom(asList.GetType()))
                        {
                            mi.Invoke(group, new[] { (object)asList });
                            return true;
                        }
                    }
                    catch (Exception ex) { log?.Invoke("[Builder] AddMembers(" + mname + ", collection) 失败: " + ex.Message); }
                }
                // 单参,接受单个 ITxObject(逐个调用)
                MethodInfo singleMi = null;
                foreach (var mi in t.GetMethods())
                {
                    if (mi.Name != mname) continue;
                    var pars = mi.GetParameters();
                    if (pars.Length == 1 &&
                        (pars[0].ParameterType == typeof(ITxObject) ||
                         typeof(ITxObject).IsAssignableFrom(pars[0].ParameterType)))
                    {
                        singleMi = mi; break;
                    }
                }
                if (singleMi != null)
                {
                    try
                    {
                        foreach (var s in members) singleMi.Invoke(group, new object[] { s });
                        return true;
                    }
                    catch (Exception ex) { log?.Invoke("[Builder] AddMembers(" + mname + ", single) 失败: " + ex.Message); }
                }
            }
            return false;
        }

        /// <summary>
        /// 反射构造 TxObjectList(PS 原生集合,实现 IList<ITxObject>)。
        /// </summary>
        private static object TryBuildTxObjectList(List<TxSolid> members)
        {
            try
            {
                Type t = typeof(TxObjectList);
                ConstructorInfo ctor = t.GetConstructor(Type.EmptyTypes);
                if (ctor == null) return null;
                object inst = ctor.Invoke(null);
                MethodInfo miAdd = t.GetMethod("Add", new[] { typeof(ITxObject) })
                                ?? t.GetMethod("Add", new[] { typeof(object) });
                if (miAdd == null) return null;
                foreach (var s in members) miAdd.Invoke(inst, new object[] { s });
                return inst;
            }
            catch { return null; }
        }

        /// <summary>
        /// 通过反射构造 PS SDK 的 *CreationData 对象(例如 TxResourceCreationData / TxPartCreationData /
        /// TxCompoundCreationData)。在不同 PS 版本中,这些类型可能在以下命名空间:
        ///   - Tecnomatix.Engineering
        ///   - Tecnomatix.Engineering.Modeling
        /// 因此走反射 + 多路径 GetType 探测,避免编译时强引用。
        /// </summary>
        private static object TryCreateCreationData(string typeShortName, string nameArg, Action<string> log)
        {
            string[] nsCandidates =
            {
                "Tecnomatix.Engineering",
                "Tecnomatix.Engineering.Modeling",
                "Tecnomatix.Engineering.Olp"
            };
            string[] asmCandidates =
            {
                "Tecnomatix.Engineering",
                "Tecnomatix.Engineering.Modeling",
                "Tecnomatix.Engineering.Olp"
            };

            // 先尝试 "type, assembly" 组合
            foreach (string ns in nsCandidates)
            {
                foreach (string asm in asmCandidates)
                {
                    try
                    {
                        Type t = Type.GetType(ns + "." + typeShortName + ", " + asm, false);
                        if (t == null) continue;
                        var ctor = t.GetConstructor(new[] { typeof(string) });
                        if (ctor != null) return ctor.Invoke(new object[] { nameArg });
                    }
                    catch { }
                }
            }

            // 再扫描已加载程序集
            try
            {
                foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
                {
                    try
                    {
                        foreach (var t in asm.GetTypes())
                        {
                            if (t.Name != typeShortName) continue;
                            var ctor = t.GetConstructor(new[] { typeof(string) });
                            if (ctor != null) return ctor.Invoke(new object[] { nameArg });
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex) { log?.Invoke("[Builder] 反射扫描失败: " + ex.Message); }

            log?.Invoke("[Builder] WARN 未找到类型 " + typeShortName);
            return null;
        }

        // =============== 单立柱(含可选底板) ===============
        // PS 语义: TxBoxCreationData.absLoc 是沿 zDir 一端的端面中心。
        //   长方体从该端面沿 zDir 延伸 edgeAlong 到另一端。
        //   在 X/Y 方向(截面方向),长方体关于 absLoc 对称(各占 edgeX/2 和 edgeY/2)。
        //
        // 立柱: zDir = (0,0,1), 沿 +Z 延伸 postHeight
        //   → absLoc 是底端面中心,Z = postBaseZ
        //   XY 中心 = postCenter.XY(立柱截面对称)
        private static void BuildOnePost(TxComponent comp, PostPlacement post, FenceParameters p,
            FenceLayout layout, BuildResult result, Action<string> log)
        {
            double postHeight;
            double postBaseZ = layout.GroundZ;
            if (p.BaseplateMode == BaseplateMode.WithBaseplate)
            {
                postHeight = p.MeshHeight + p.GroundClearance + p.PostTopMargin - p.BaseplateThickness;
                postBaseZ = layout.GroundZ + p.BaseplateThickness;
            }
            else
            {
                postHeight = p.MeshHeight + p.GroundClearance + p.PostTopMargin;
            }
            if (postHeight < 1.0) return;

            TxVector absLoc = new TxVector(post.CenterXY.X, post.CenterXY.Y, postBaseZ);
            TxSolid s = CreateBoxCore(comp, "FencePost", absLoc, new TxVector(0, 0, 1),
                p.PostWidth, p.PostThickness, postHeight, log);
            if (s != null)
            {
                AppearanceHelper.TrySetColor(s, p.PostColor, log);
                result.CreatedSolids.Add(s);
                result.PostCount++;
            }

            if (p.BaseplateMode == BaseplateMode.WithBaseplate)
            {
                TxVector bpLoc = new TxVector(post.CenterXY.X, post.CenterXY.Y, layout.GroundZ);
                TxSolid bp = CreateBoxCore(comp, "FenceBaseplate", bpLoc, new TxVector(0, 0, 1),
                    p.BaseplateLength, p.BaseplateWidth, p.BaseplateThickness, log);
                if (bp != null)
                {
                    AppearanceHelper.TrySetColor(bp, p.BaseplateColor, log);
                    result.CreatedSolids.Add(bp);
                    result.BaseplateCount++;
                }
            }
        }

        // =============== 单网片: 5 个几何 ===============
        // PS 语义重申: absLoc 是沿 zDir 一端的端面中心,XY 方向(垂直 zDir 截面)关于 absLoc 对称。
        //
        // 局部坐标系 (X' 沿 dirAlong, Y' 沿 nrm 法线, Z' 沿世界 Z):
        //   网片占据 X' ∈ [0, panelW], Z' ∈ [meshBotZ, meshTopZ]
        //   网片左下角(X'=0, Y'=0, Z'=meshBotZ) = leftEdge
        //
        // 几何参数(zDir 决定 X/Y 哪两个轴是截面、哪个轴是延伸方向):
        //
        //   下沿: zDir = dirAlong, 沿 X' 延伸 panelW
        //     absLoc = X'=0 端的端面中心
        //     端面在 Y'Z' 平面上为 frame×frame, 关于 (Y'=0, Z'=meshBotZ+frame/2) 对称
        //     → absLoc = (leftEdge.X, leftEdge.Y, meshBotZ + frame/2)
        //
        //   上沿: 同下沿, Z' = meshTopZ - frame/2
        //
        //   左边: zDir = (0,0,1), 沿 Z' 延伸 innerH
        //     absLoc = Z' = meshBotZ + frame 端的端面中心
        //     端面在 X'Y' 平面上为 frame×frame, 关于 (X'=frame/2, Y'=0) 对称
        //     → absLoc = (leftEdge + dirAlong * frame/2, meshBotZ + frame)
        //
        //   右边: 同左边, X' = panelW - frame/2
        //
        //   薄板: zDir = dirAlong, 沿 X' 延伸 innerW
        //     absLoc = X'=frame 端的端面中心(从内沿起算)
        //     端面在 Y'Z' 平面上为 Thickness×innerH, 关于 (Y'=0, Z'=meshBotZ+frame+innerH/2) 对称
        //     → absLoc = (leftEdge + dirAlong * frame, leftEdge.Y, meshBotZ + frame + innerH/2)
        private static void BuildOnePanel(TxComponent comp, MeshPanelPlacement pan, FenceParameters p,
            FenceLayout layout, string textureFilePath, BuildResult result, Action<string> log)
        {
            double frameDim = Math.Max(p.MeshFrameWidth, p.MeshFrameThickness);
            double meshBotZ = layout.GroundZ + p.GroundClearance;
            double meshTopZ = meshBotZ + p.MeshHeight;
            double panelW = pan.ActualWidth;
            double innerH = p.MeshHeight - 2.0 * frameDim;
            double innerW = panelW - 2.0 * frameDim;

            TxVector dirAlong = pan.DirAlong;
            TxVector leftEdge = AddXY(pan.LeftPostCenter, dirAlong, p.PostWidth / 2.0 + p.PostGap);

            // 1) 下沿: zDir=dirAlong, 截面 X=nrm 方向, Y=世界Z 方向
            TxVector botLoc = new TxVector(leftEdge.X, leftEdge.Y, meshBotZ + frameDim / 2.0);
            TxSolid bot = CreateBoxCore(comp, "FenceFrameBottom", botLoc, dirAlong,
                frameDim, frameDim, panelW, log);
            if (bot != null)
            {
                AppearanceHelper.TrySetColor(bot, p.FrameColor, log);
                result.CreatedSolids.Add(bot); result.FrameRailCount++;
            }

            // 2) 上沿: 同下沿,Z 移到 meshTopZ - frame/2
            TxVector topLoc = new TxVector(leftEdge.X, leftEdge.Y, meshTopZ - frameDim / 2.0);
            TxSolid top = CreateBoxCore(comp, "FenceFrameTop", topLoc, dirAlong,
                frameDim, frameDim, panelW, log);
            if (top != null)
            {
                AppearanceHelper.TrySetColor(top, p.FrameColor, log);
                result.CreatedSolids.Add(top); result.FrameRailCount++;
            }

            // 3) 左边: zDir=+Z, 截面 X=dirAlong 方向, Y=nrm 方向
            if (innerH > 1.0)
            {
                TxVector leftLoc = AddXY(
                    new TxVector(leftEdge.X, leftEdge.Y, meshBotZ + frameDim),
                    dirAlong, frameDim / 2.0);
                TxSolid leftRail = CreateBoxCore(comp, "FenceFrameLeft", leftLoc, new TxVector(0, 0, 1),
                    frameDim, frameDim, innerH, log);
                if (leftRail != null)
                {
                    AppearanceHelper.TrySetColor(leftRail, p.FrameColor, log);
                    result.CreatedSolids.Add(leftRail); result.FrameRailCount++;
                }

                // 4) 右边
                TxVector rightLoc = AddXY(
                    new TxVector(leftEdge.X, leftEdge.Y, meshBotZ + frameDim),
                    dirAlong, panelW - frameDim / 2.0);
                TxSolid rightRail = CreateBoxCore(comp, "FenceFrameRight", rightLoc, new TxVector(0, 0, 1),
                    frameDim, frameDim, innerH, log);
                if (rightRail != null)
                {
                    AppearanceHelper.TrySetColor(rightRail, p.FrameColor, log);
                    result.CreatedSolids.Add(rightRail); result.FrameRailCount++;
                }
            }

            // 5) 薄板: zDir=nrm(法线), 沿 nrm 延伸 MeshThickness (薄板厚度)
            //    端面在 dirAlong-Z 平面上为 innerW × innerH,关于 absLoc 对称
            //    薄板正面中线(nrm 方向中心)对齐到网片中央
            //    absLoc 沿 -nrm 偏移 MeshThickness/2 (端面位于 nrm 一端)
            //    几何中心(沿 dirAlong): leftEdge + frame + innerW/2 = panelW/2
            //    几何中心(沿世界 Z): meshBotZ + frame + innerH/2
            //    PS 自选 X/Y 时,二者都在 dirAlong-Z 平面里,都是大尺寸,X/Y 谁朝哪边视觉上一致
            if (p.EnableMeshTexture && innerW > 1.0 && innerH > 1.0)
            {
                TxVector nrm = pan.DirNormal;
                // 中央点 (在网片几何中央)
                TxVector meshCenter = AddXY(
                    new TxVector(leftEdge.X, leftEdge.Y, meshBotZ + frameDim + innerH / 2.0),
                    dirAlong, frameDim + innerW / 2.0);
                // 沿 -nrm 偏移 MeshThickness/2 得到端面中心 (absLoc)
                TxVector meshLoc = new TxVector(
                    meshCenter.X - nrm.X * p.MeshThickness / 2.0,
                    meshCenter.Y - nrm.Y * p.MeshThickness / 2.0,
                    meshCenter.Z);
                TxSolid mesh = CreateBoxCore(comp, "FenceMesh", meshLoc, nrm,
                    innerW, innerH, p.MeshThickness, log);
                if (mesh != null)
                {
                    bool textured = AppearanceHelper.TrySetTexture(mesh, textureFilePath, log);
                    if (textured) result.TextureSuccessCount++;
                    else
                    {
                        AppearanceHelper.TrySetColor(mesh, p.MeshColor, log);
                        AppearanceHelper.TrySetTransparency(mesh, 0.5, log);
                        result.TextureFallbackCount++;
                    }
                    result.CreatedSolids.Add(mesh);
                    result.MeshPanelCount++;
                }
            }
        }

        // =============== 底层 ===============
        // PS 2402 实测验证的 TxBoxCreationData.absLoc 语义:
        //   - absLoc.Position = 长方体沿 zDir 一端(-zDir 端)的端面中心
        //   - 沿 zDir 方向: 从 absLoc 起延伸 edgeAlong 到另一端
        //   - 在 X/Y 截面(垂直 zDir 的两个轴): 关于 absLoc 对称(各占 edgeX/2 / edgeY/2)
        //   - PS 自选 X/Y 朝向, 不可控
        //
        // 沿用 LineToSolid 已验证的策略: 仅用单参 TxTransformation(point, zDir)。
        // 为规避 X/Y 朝向不可控:
        //   - 所有方管和立柱采用正方形截面(W=H), X/Y 朝向无视觉差
        //   - 薄板 zDir=nrm, X/Y 都是大尺寸(innerW×innerH), X/Y 朝向也无视觉差
        private static TxSolid CreateBoxCore(TxComponent comp, string name,
            TxVector basePoint, TxVector zDir,
            double edgeX, double edgeY, double edgeAlong, Action<string> log)
        {
            if (edgeAlong < 1e-6 || edgeX < 1e-6 || edgeY < 1e-6) return null;
            try
            {
                TxTransformation xform = new TxTransformation(basePoint, zDir);
                TxVector edgeSizes = new TxVector(edgeX, edgeY, edgeAlong);
                TxVector offset = new TxVector(0, 0, 0);
                TxBoxCreationData data = new TxBoxCreationData(name, xform, edgeSizes, offset);
                return comp.CreateSolidBox(data);
            }
            catch (Exception ex)
            {
                log?.Invoke("[Builder] CreateSolidBox FAIL (" + name + "): " + ex.Message);
                return null;
            }
        }

        private static TxVector AddXY(TxVector p, TxVector dir, double dist)
        {
            return new TxVector(p.X + dir.X * dist, p.Y + dir.Y * dist, p.Z);
        }

        /// <summary>
        /// 查找当前处于建模状态(ITxComponent.IsOpenForModeling=true)的 TxComponent。
        /// 用 4 路径递归遍历 PhysicalRoot 后代(沿用 FenceBaselineReader 同款策略):
        ///   1) GetAllDescendants(null/无参) 一次性深度遍历
        ///   2) GetDirectDescendants(null/无参) + 手动递归
        ///   3) Components/Children/SubObjects/Items/Members 属性 + 手动递归
        ///   4) IEnumerable 兜底
        /// </summary>
        private static TxComponent TryFindActiveModelingComponent(Action<string> log)
        {
            try
            {
                var root = TxApplication.ActiveDocument.PhysicalRoot;
                if (root == null) { log?.Invoke("[Builder] TryFindActiveModelingComponent: PhysicalRoot 为 null"); return null; }

                int totalScanned = 0;
                TxComponent found = SearchRecursively(root, ref totalScanned, 0, log);

                log?.Invoke("[Builder] TryFindActiveModelingComponent: 扫描 " + totalScanned + " 项, " +
                            (found != null ? "命中 " + TrySafeName(found) : "未命中(没有任何 ITxComponent.IsOpenForModeling=true)"));
                return found;
            }
            catch (Exception ex) { log?.Invoke("[Builder] TryFindActiveModelingComponent 异常: " + ex.Message); }
            return null;
        }

        private static TxComponent SearchRecursively(object container, ref int totalScanned, int depth, Action<string> log)
        {
            if (container == null || depth > 50) return null;

            // 当前节点本身就检查一次(也可能就是它)
            ITxComponent itc = container as ITxComponent;
            if (itc != null)
            {
                totalScanned++;
                try
                {
                    if (itc.IsOpenForModeling)
                    {
                        TxComponent tc = container as TxComponent;
                        if (tc != null) return tc;
                        log?.Invoke("[Builder] " + TrySafeName(itc) + " IsOpenForModeling=true 但不是 TxComponent (实际类型 " +
                                    container.GetType().Name + ")");
                    }
                }
                catch { }
            }

            // 路径 1: GetAllDescendants(null/无参)
            var found = TryEnumerateMethod(container, "GetAllDescendants");
            if (found != null)
            {
                foreach (var item in found)
                {
                    if (item == null) continue;
                    totalScanned++;
                    ITxComponent ic = item as ITxComponent;
                    if (ic == null) continue;
                    try
                    {
                        if (ic.IsOpenForModeling)
                        {
                            TxComponent tc = item as TxComponent;
                            if (tc != null) return tc;
                        }
                    }
                    catch { }
                }
                return null; // 一次性遍历,无需再递归
            }

            // 路径 2/3/4: 取直接子节点列表,然后递归
            System.Collections.IEnumerable children = TryEnumerateMethod(container, "GetDirectDescendants");
            if (children == null) children = TryEnumerateChildProperty(container);
            if (children == null) children = container as System.Collections.IEnumerable;

            if (children != null)
            {
                foreach (var child in children)
                {
                    if (child == null) continue;
                    TxComponent sub = SearchRecursively(child, ref totalScanned, depth + 1, log);
                    if (sub != null) return sub;
                }
            }
            return null;
        }

        private static System.Collections.IEnumerable TryEnumerateMethod(object container, string methodName)
        {
            try
            {
                foreach (var mi in container.GetType().GetMethods())
                {
                    if (mi.Name != methodName) continue;
                    var pars = mi.GetParameters();
                    try
                    {
                        if (pars.Length == 1)
                        {
                            var res = mi.Invoke(container, new object[] { null }) as System.Collections.IEnumerable;
                            if (res != null) return res;
                        }
                        if (pars.Length == 0)
                        {
                            var res = mi.Invoke(container, null) as System.Collections.IEnumerable;
                            if (res != null) return res;
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return null;
        }

        private static System.Collections.IEnumerable TryEnumerateChildProperty(object container)
        {
            string[] propNames = { "Components", "Children", "SubObjects", "Items", "Members", "Descendants" };
            try
            {
                foreach (string name in propNames)
                {
                    PropertyInfo pi = container.GetType().GetProperty(name);
                    if (pi == null) continue;
                    var val = pi.GetValue(container, null) as System.Collections.IEnumerable;
                    if (val != null) return val;
                }
            }
            catch { }
            return null;
        }

        private static string TrySafeName(ITxObject obj)
        {
            try { return obj.Name ?? "(unnamed)"; }
            catch { return "(unknown)"; }
        }
    }
}