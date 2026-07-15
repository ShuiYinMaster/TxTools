using System;
using System.Collections.Generic;
using System.Reflection;
using Tecnomatix.Engineering;

namespace TxTools.WeldGunDefiner.Core
{
    using Math;

    /// <summary>
    /// PS SDK 接口封装层（基于 tecnomatix_api_docs_combined.md 验证）
    ///
    /// 已确认的关键API：
    ///   TxApplication.ActiveSelection → TxSelection
    ///   TxSelection.GetItems() → TxObjectList（有GetEnumerator，支持foreach）
    ///   TxSelection.Count（属性）
    ///   TxObjectList : Collection<ITxObject>（有索引器、foreach、Count）
    ///
    ///   ITxKinematicsModellable.CreateJoint(creationData) → TxJoint
    ///   ITxKinematicsModellable.CreateLink(creationData) → 链接
    ///   TxJoint.KinematicsFunction（get/set，字符串公式）
    ///   TxJoint.Type（TxJointType枚举）
    ///   TxJoint.LowerSoftLimit / UpperSoftLimit
    ///
    ///   ITxDevice.CreatePose(TxPoseCreationData) → pose
    ///   TxPoseCreationData.Name / PoseData
    ///   TxPoseData.JointValues
    /// </summary>
    public static class PsSdkHelper
    {
        // ─────────────────────────────────────────────
        // 1. 从选中对象获取世界坐标点
        // ─────────────────────────────────────────────

        public static bool TryGetWorldPosition(object txObj, out Vec3 position)
        {
            position = default;
            if (txObj == null) return false;

            try
            {
                if (txObj is ITxLocatableObject locatable)
                {
                    var loc = locatable.AbsoluteLocation;
                    position = ExtractTranslation(loc);
                    return true;
                }

                // 反射回退
                var prop = txObj.GetType().GetProperty("AbsoluteLocation");
                if (prop?.GetValue(txObj, null) is TxTransformation tx)
                {
                    position = ExtractTranslation(tx);
                    return true;
                }
            }
            catch { }
            return false;
        }

        private static Vec3 ExtractTranslation(TxTransformation tx)
            => new Vec3(tx[0, 3], tx[1, 3], tx[2, 3]);

        // ─────────────────────────────────────────────
        // 2. 从选中的 Face 对象获取法向量（反射多路探测）
        // ─────────────────────────────────────────────

        public static bool TryGetFaceNormal(object faceObj, out Vec3 normal)
        {
            normal = default;
            if (faceObj == null) return false;

            try
            {
                foreach (var name in new[] { "Normal", "NormalVector" })
                {
                    var p = faceObj.GetType().GetProperty(name);
                    if (p?.GetValue(faceObj, null) is TxVector v)
                    {
                        normal = new Vec3(v).Normalized();
                        return true;
                    }
                }
                var mi = faceObj.GetType().GetMethod("GetNormal", Type.EmptyTypes);
                if (mi?.Invoke(faceObj, null) is TxVector v2)
                {
                    normal = new Vec3(v2).Normalized();
                    return true;
                }
            }
            catch { }
            return false;
        }

        public static string DiagnoseFaceObject(object faceObj)
        {
            if (faceObj == null) return "null";
            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"类型: {faceObj.GetType().FullName}");
            sb.AppendLine("属性:");
            foreach (var p in faceObj.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
                sb.AppendLine($"  {p.PropertyType.Name} {p.Name}");
            sb.AppendLine("方法:");
            foreach (var m in faceObj.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance))
                if (!m.Name.StartsWith("get_") && !m.Name.StartsWith("set_"))
                    sb.AppendLine($"  {m.Name}");
            return sb.ToString();
        }

        // ─────────────────────────────────────────────
        // 3. 运动学写入（基于文档确认的API）
        // ─────────────────────────────────────────────

        /// <summary>
        /// 在 ITxKinematicsModellable 上创建 Joint。
        /// CreateJoint 接受一个 CreationData 对象（反射构造，类名待运行时确认）。
        /// 返回 TxJoint，失败返回 null。
        /// </summary>
        public static TxJoint CreateJoint(ITxKinematicsModellable device, string name,
            string jointTypeName, double lowerLimit, double upperLimit,
            TxVector axisFrom, TxVector axisTo)
        {
            try
            {
                // TxJointCreationData 位于 Tecnomatix.Engineering.DataTypes 命名空间
                // CreateJoint 参数类型为 Tecnomatix.Engineering.DataTypes.TxJointCreationData
                var asmNames = new[]
                {
                    "Tecnomatix.Engineering.DataTypes, Tecnomatix.Engineering",
                    "Tecnomatix.Engineering.DataTypes, Tecnomatix.Engineering.Base",
                    "Tecnomatix.Engineering"
                };

                Type dataType = null;
                foreach (var asm in asmNames)
                {
                    dataType = Type.GetType($"Tecnomatix.Engineering.DataTypes.TxJointCreationData, {asm}");
                    if (dataType != null) break;
                }
                // 回退：直接搜所有已加载程序集
                if (dataType == null)
                {
                    foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
                    {
                        dataType = asm.GetType("Tecnomatix.Engineering.DataTypes.TxJointCreationData");
                        if (dataType != null) break;
                    }
                }
                if (dataType == null) return null;

                var data = Activator.CreateInstance(dataType);

                TrySet(data, "Name", name);
                TrySetEnum(data, "JointType", jointTypeName);
                TrySet(data, "LowerLimit", lowerLimit);
                TrySet(data, "UpperLimit", upperLimit);
                TrySet(data, "LowerSoftLimit", lowerLimit);
                TrySet(data, "UpperSoftLimit", upperLimit);

                // TxJointAxis 构造函数需要两个 TxVector 参数（无无参构造）
                if (axisFrom != null && axisTo != null)
                {
                    var axis = new TxJointAxis(axisFrom, axisTo);
                    TrySet(data, "Axis", axis);
                }

                // CreateJoint 参数类型为 DataTypes.TxJointCreationData，用反射调用避免类型不匹配
                var mi = device.GetType().GetMethod("CreateJoint");
                if (mi == null) return null;
                var result = mi.Invoke(device, new[] { data });
                return result as TxJoint;
            }
            catch { return null; }
        }

        /// <summary>
        /// 通过 TxJoint.KinematicsFunction 属性写入公式字符串
        /// （文档确认：TxJoint.KinematicsFunction get/set）
        /// </summary>
        public static bool SetJointFormula(TxJoint joint, string formula)
        {
            if (joint == null || string.IsNullOrEmpty(formula)) return false;
            try
            {
                joint.KinematicsFunction = formula;
                return true;
            }
            catch
            {
                // 回退反射
                try
                {
                    var p = joint.GetType().GetProperty("KinematicsFunction");
                    if (p != null && p.CanWrite) { p.SetValue(joint, formula, null); return true; }
                }
                catch { }
                return false;
            }
        }

        // ─────────────────────────────────────────────
        // 4. Pose 创建（基于文档确认的API）
        // 文档：ITxDevice.CreatePose(TxPoseCreationData)
        //       TxPoseCreationData.Name / PoseData
        //       TxPoseData.JointValues（类型待确认，反射探测）
        // ─────────────────────────────────────────────

        // 正确流程（文档确认）：
        //   TxDevicePoseData(device, device.CurrentPose) 包装当前位姿
        //   → SetJointValue(joint, value) 逐关节设值
        //   → devicePoseData.PoseData 取回 TxPoseData
        //   → TxPoseCreationData{Name, PoseData} → device.CreatePose
        // 之前用反射猜 JointValues 属性，TxPoseData 根本没有该属性，故静默失败。
        public static bool CreatePose(ITxDevice device, string poseName,
            Dictionary<TxJoint, double> jointValues, out string err)
        {
            err = null;
            try
            {
                if (device == null) { err = "device为null"; return false; }

                TxPoseData template = device.CurrentPose;
                if (template == null) { err = "device.CurrentPose为null"; return false; }

                TxDevicePoseData dpd = null;
                try { dpd = new TxDevicePoseData(device, template); }
                catch (Exception ex) { err = "构造TxDevicePoseData失败:" + ex.Message; }

                if (dpd != null)
                {
                    foreach (var kv in jointValues)
                    {
                        if (kv.Key == null) continue;
                        try
                        {
                            bool has = true;
                            try { has = dpd.DoesHaveJoint(kv.Key); } catch { has = true; }
                            if (has) dpd.SetJointValue(kv.Key, kv.Value);
                        }
                        catch (Exception ex) { err = $"SetJointValue({kv.Key?.Name})失败:" + ex.Message; }
                    }

                    TxPoseData finalData = dpd.PoseData;
                    var cd = new TxPoseCreationData { Name = poseName, PoseData = finalData };
                    device.CreatePose(cd);
                    return true;
                }

                var cd2 = new TxPoseCreationData { Name = poseName, PoseData = template };
                device.CreatePose(cd2);
                return true;
            }
            catch (Exception ex) { err = ex.Message; return false; }
        }

        public static bool CreatePose(ITxDevice device, string poseName,
            Dictionary<TxJoint, double> jointValues)
        {
            string _; return CreatePose(device, poseName, jointValues, out _);
        }

        // ── DefineZeroPosition：把当前位姿定为零位（需求4，对应"Define as Zero Position"）──
        public static bool DefineZeroPosition(ITxKinematicsModellable device, out string err)
        {
            err = null;
            try
            {
                if (device == null) { err = "device为null"; return false; }
                bool can = true;
                try { can = device.CanDefineZeroPosition(); } catch { can = true; }
                if (!can) { err = "CanDefineZeroPosition=false"; return false; }
                device.DefineZeroPosition();
                return true;
            }
            catch (Exception ex) { err = ex.Message; return false; }
        }

        // ── 驱动单个关节到指定值（需求4：把焊钳驱动回CLOSE）──
        public static bool SetJointCurrentValue(TxJoint joint, double value, out string err)
        {
            err = null;
            try
            {
                if (joint == null) { err = "joint为null"; return false; }
                joint.CurrentValue = value;
                return true;
            }
            catch (Exception ex) { err = ex.Message; return false; }
        }

        // ── 定义为 Servo Gun 工具（需求3焊钳定义）──
        // API确认：DefineAsTool(ITxToolDefinitionData) 需参数。
        // 用 TxServoGunDefinitionData 设 TCPF(TxFrame) / BaseframeAbsoluteLocation /
        // NonCollidingEntities(TxObjectList)，再传入 DefineAsTool。
        public static bool DefineAsServoGun(ITxKinematicsModellable device,
            TxFrame tcpFrame, TxTransformation baseLoc,
            List<ITxObject> nonCollidingEntities, out string err)
        {
            err = null;
            try
            {
                if (device == null) { err = "device为null"; return false; }

                // 不检测干涉列表（构造参数，可空列表）
                var ncList = new TxObjectList();
                if (nonCollidingEntities != null)
                    foreach (var o in nonCollidingEntities)
                        if (o != null) ncList.Add(o);

                // 构造确认：TxServoGunDefinitionData(TxFrame TCPF,
                //   TxTransformation baseframeAbsoluteLocation, TxObjectList nonCollidingEntities)
                var def = new Tecnomatix.Engineering.DataTypes.TxServoGunDefinitionData(
                    tcpFrame, baseLoc, ncList);

                // 检查可定义性（带参）
                bool can = true;
                try { can = device.CanDefineAsTool(def); } catch { can = true; }
                if (!can) err = (err ?? "") + " | CanDefineAsTool=false（运动学可能不完整或已是工具）";

                device.DefineAsTool(def);
                return true;
            }
            catch (Exception ex) { err = (err ?? "") + " | " + ex.Message; return false; }
        }

        // ─────────────────────────────────────────────
        // 5. Logic Block（反射，接口未在文档中明确）
        // ─────────────────────────────────────────────

        public static bool CreateSimpleLogicBlock(ITxDevice device,
            string signalName, string poseOn, string poseOff)
        {
            try
            {
                // 尝试通过 ITxComponent 接口找 CreateLogicBlock
                var comp = device as ITxComponent;
                if (comp == null) return false;

                var mi = comp.GetType().GetMethod("CreateLogicBlock");
                if (mi == null) return false;

                var lbType = Type.GetType("Tecnomatix.Engineering.TxLogicBlockCreationData, Tecnomatix.Engineering");
                if (lbType == null) return false;

                var lbData = Activator.CreateInstance(lbType);
                TrySet(lbData, "Name", "GunControl");
                var lb = mi.Invoke(comp, new[] { lbData });
                if (lb == null) return false;

                TryCall(lb, "AddInputSignal", signalName, "Boolean");
                TryCall(lb, "AddPoseTransition", signalName, true, poseOn);
                TryCall(lb, "AddPoseTransition", signalName, false, poseOff);
                return true;
            }
            catch { return false; }
        }

        // ─────────────────────────────────────────────
        // 内部工具
        // ─────────────────────────────────────────────

        private static void TrySet(object obj, string prop, object value)
        {
            try { obj.GetType().GetProperty(prop)?.SetValue(obj, value, null); } catch { }
        }

        private static void TrySetEnum(object obj, string prop, string enumName)
        {
            try
            {
                var p = obj.GetType().GetProperty(prop);
                if (p == null) return;
                if (p.PropertyType.IsEnum)
                {
                    var val = Enum.Parse(p.PropertyType, enumName, true);
                    p.SetValue(obj, val, null);
                }
                else
                {
                    p.SetValue(obj, enumName, null);
                }
            }
            catch { }
        }

        private static void TryCall(object obj, string method, params object[] args)
        {
            try { obj.GetType().GetMethod(method)?.Invoke(obj, args); } catch { }
        }

        // ─────────────────────────────────────────────
        // 高亮显示：选中对象变色，取消时恢复
        // ─────────────────────────────────────────────
        // ITxDisplayableObject.Color 是属性（TxColor类型），不是SetColor方法
        // 用预定义颜色常量（静态属性，避免构造函数参数类型不确定）

        /// <summary>不同 Link 的高亮配色（与图例对应）</summary>
        public static TxColor ColorFixed => TxColor.TxColorYellow;   // fixed_link
        public static TxColor ColorInput => TxColor.TxColorOrange;   // input_link
        public static TxColor ColorCoupler => TxColor.TxColorCyan;     // coupler_link
        public static TxColor ColorOutput => TxColor.TxColorMagenta;  // output_link
        public static TxColor ColorLnk2 => TxColor.TxColorPink;     // lnk2
        public static TxColor ColorTip => TxColor.TxColorGreen;    // 电极帽

        /// <summary>高亮一个对象为指定颜色（ITxDisplayableObject.Color 属性）</summary>
        public static void Highlight(ITxObject obj, TxColor color)
        {
            if (obj == null) return;
            try
            {
                var disp = obj as ITxDisplayableObject;
                if (disp != null) disp.Color = color;
            }
            catch { }
        }

        /// <summary>高亮一个对象（默认黄色）</summary>
        public static void Highlight(ITxObject obj)
        {
            Highlight(obj, TxColor.TxColorYellow);
        }

        /// <summary>取消高亮，恢复原色（ITxDisplayableObject.RestoreColor）</summary>
        public static void Unhighlight(ITxObject obj)
        {
            if (obj == null) return;
            try
            {
                var disp = obj as ITxDisplayableObject;
                if (disp != null) disp.RestoreColor();
            }
            catch { }
        }

        /// <summary>批量高亮为指定颜色</summary>
        public static void HighlightAll(System.Collections.Generic.IEnumerable<ITxObject> objs, TxColor color)
        {
            if (objs == null) return;
            foreach (var o in objs) Highlight(o, color);
        }

        /// <summary>批量高亮</summary>
        public static void HighlightAll(System.Collections.Generic.IEnumerable<ITxObject> objs)
        {
            if (objs == null) return;
            foreach (var o in objs) Highlight(o);
        }

        /// <summary>批量取消高亮</summary>
        public static void UnhighlightAll(System.Collections.Generic.IEnumerable<ITxObject> objs)
        {
            if (objs == null) return;
            foreach (var o in objs) Unhighlight(o);
        }

        // ── 在 Component 内创建 Frame（需求3：TCP点选位置时建TCPF）──
        public static TxFrame CreateFrame(ITxComponent comp, string name,
            TxTransformation absLoc, out string err)
        {
            err = null;
            try
            {
                var creator = comp as ITxFrameCreation;
                if (creator == null) { err = "组件不支持ITxFrameCreation"; return null; }
                var cd = new TxFrameCreationData();
                cd.Name = name;
                if (absLoc != null) cd.AbsoluteLocation = absLoc;
                var frame = creator.CreateFrame(cd);
                return frame;
            }
            catch (Exception ex) { err = ex.Message; return null; }
        }

        // ── 判断 obj 是否属于 comp（祖先链向上遍历）──
        // 需求2修复：之前用 GetAllDescendants 集合精确匹配，但 PickLevel.Entity 选到的
        // 几何实体(TxSolid/TxPolyhedron)往往不在该集合里，导致本体几何也被误判为外部、
        // 无法选择(且与link无关)。改为从obj向上遍历祖先，到达comp即属于本焊钳。
        public static bool IsDescendantOf(ITxObject obj, ITxComponent comp)
        {
            if (obj == null || comp == null) return false;
            try
            {
                if (ReferenceEquals(obj, comp)) return true;
                var cur = obj;
                for (int depth = 0; depth < 64 && cur != null; depth++)
                {
                    if (ReferenceEquals(cur, comp)) return true;
                    // 向上取宿主：优先反射 Parent / OwningObject，再退回 Collection 的宿主
                    object parent = GetParentObject(cur);
                    cur = parent as ITxObject;
                }
            }
            catch { }
            return false;
        }

        private static object GetParentObject(ITxObject obj)
        {
            if (obj == null) return null;
            try
            {
                var t = obj.GetType();
                // 常见祖先访问属性
                foreach (var name in new[] { "Parent", "OwningObject", "Owner", "Container" })
                {
                    var p = t.GetProperty(name);
                    if (p != null)
                    {
                        var v = p.GetValue(obj, null);
                        if (v is ITxObject && !ReferenceEquals(v, obj)) return v;
                    }
                }
                // Collection 的宿主对象
                var collProp = t.GetProperty("Collection");
                if (collProp != null)
                {
                    var coll = collProp.GetValue(obj, null);
                    if (coll != null)
                    {
                        var ownerProp = coll.GetType().GetProperty("OwningObject")
                                     ?? coll.GetType().GetProperty("Owner")
                                     ?? coll.GetType().GetProperty("Parent");
                        if (ownerProp != null)
                        {
                            var owner = ownerProp.GetValue(coll, null);
                            if (owner is ITxObject && !ReferenceEquals(owner, obj)) return owner;
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        // ── 收集 Component 下所有后代（保留兼容，但归属判断改用 IsDescendantOf）──
        public static System.Collections.Generic.HashSet<ITxObject> CollectDescendants(ITxComponent comp)
        {
            var set = new System.Collections.Generic.HashSet<ITxObject>();
            if (comp == null) return set;
            try
            {
                var coll = comp as ITxObjectCollection;
                if (coll != null)
                {
                    var filter = new TxTypeFilter(typeof(ITxObject));
                    var all = coll.GetAllDescendants(filter);
                    if (all != null)
                        foreach (ITxObject o in all) if (o != null) set.Add(o);
                }
            }
            catch { }
            set.Add(comp);   // 自身也算
            return set;
        }
    }
}