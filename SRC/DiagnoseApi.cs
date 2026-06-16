// DiagnoseApi.cs  —  诊断 PmWeldLocationOperation 等 Planning 类型的属性
//
// 用法：将此文件临时替换 ExportGunCmd.cs 中的 Execute()，
// 或者单独在 PS 插件上下文中调用 DiagnoseApi.Run(log)。
//
// 目标：找出 PmWeldLocationOperation / PmViaLocationOperation /
//        PmContinuousFeatureOperation 的属性名和坐标访问路径。

using System;
using System.Collections;
using System.Reflection;
using Tecnomatix.Engineering;

namespace MyPlugin.ExportGun
{
    public static class DiagnoseApi
    {
        public static void Run(Action<string> log)
        {
            if (log == null) log = delegate(string s) { };
            try
            {
                TxDocument doc = TxApplication.ActiveDocument;
                if (doc == null) { log("[Diag] ActiveDocument=null"); return; }

                TxOperationRoot opRoot = doc.OperationRoot;
                if (opRoot == null) { log("[Diag] OperationRoot=null"); return; }

                // 找第一个 TxWeldOperation
                TxTypeFilter f = new TxTypeFilter(typeof(ITxObject));
                TxObjectList kids = opRoot.GetDirectDescendants(f);
                FindWeldOp(kids, log, 0);
            }
            catch (Exception ex) { log("[Diag] 异常：" + ex.Message); }
        }

        private static void FindWeldOp(TxObjectList list, Action<string> log, int depth)
        {
            if (list == null || depth > 10) return;
            foreach (ITxObject obj in list)
            {
                string tn = obj.GetType().Name;
                if (tn == "TxWeldOperation" || tn.Contains("WeldOp"))
                {
                    log("[Diag] 找到 " + tn + " : " + SafeName(obj));
                    InspectEnum(obj, log);
                    return;
                }
                TxTypeFilter f = new TxTypeFilter(typeof(ITxObject));
                TxObjectList children = null;
                try
                {
                    if (obj is TxCompoundOperation)
                        children = ((TxCompoundOperation)obj).GetDirectDescendants(f);
                    else
                    {
                        dynamic d = obj;
                        children = d.GetDirectDescendants(f) as TxObjectList;
                    }
                }
                catch { }
                FindWeldOp(children, log, depth + 1);
            }
        }

        private static void InspectEnum(ITxObject weldOp, Action<string> log)
        {
            IEnumerable src = weldOp as IEnumerable;
            if (src == null) { log("[Diag] 不可枚举"); return; }

            int idx = 0;
            foreach (object item in src)
            {
                if (item == null) continue;
                Type t = item.GetType();
                log("[Diag] Child[" + idx + "] type=" + t.FullName);

                // 反射所有属性和方法名
                MemberInfo[] members = t.GetMembers(
                    BindingFlags.Public | BindingFlags.Instance);
                foreach (MemberInfo m in members)
                {
                    string mName = m.Name;
                    // 只关注可能含坐标/名称/类型的成员
                    if (mName.Contains("Location") || mName.Contains("Name")
                        || mName.Contains("Frame")  || mName.Contains("Pose")
                        || mName.Contains("Matrix") || mName.Contains("Point")
                        || mName.Contains("Feature")|| mName.Contains("Type")
                        || mName.Contains("Weld")   || mName.Contains("Via"))
                    {
                        string valStr = "";
                        if (m.MemberType == MemberTypes.Property)
                        {
                            try
                            {
                                object val = ((PropertyInfo)m).GetValue(item, null);
                                valStr = " = " + (val == null ? "null" : val.GetType().Name + ":" + val);
                            }
                            catch (Exception ex) { valStr = " [get threw: " + ex.Message + "]"; }
                        }
                        log("[Diag]   ." + mName + " (" + m.MemberType + ")" + valStr);
                    }
                }

                // 尝试直接动态访问常见属性
                dynamic d = item;
                TryLog(log, "[Diag]   dyn.Name",             delegate() { return (string)d.Name; });
                TryLog(log, "[Diag]   dyn.AbsoluteLocation", delegate() { return d.AbsoluteLocation; });
                TryLog(log, "[Diag]   dyn.Location",         delegate() { return d.Location; });
                TryLog(log, "[Diag]   dyn.LocationData",     delegate() { return d.LocationData; });
                TryLog(log, "[Diag]   dyn.MfgFeature",       delegate() { return d.MfgFeature; });
                TryLog(log, "[Diag]   dyn.Feature",          delegate() { return d.Feature; });
                TryLog(log, "[Diag]   dyn.WeldPoint",        delegate() { return d.WeldPoint; });
                TryLog(log, "[Diag]   dyn.OperationType",    delegate() { return d.OperationType; });
                TryLog(log, "[Diag]   dyn.Process",          delegate() { return d.Process; });

                if (++idx >= 3) break; // 只检查前3个
            }
        }

        private static void TryLog(Action<string> log, string label, Func<object> fn)
        {
            try
            {
                object val = fn();
                string valStr = val == null ? "null"
                    : val.GetType().Name + " : " + val.ToString();
                log(label + " = " + valStr);
            }
            catch (Exception ex)
            {
                log(label + " => 异常: " + ex.Message.Split('\n')[0]);
            }
        }

        private static string SafeName(ITxObject obj)
        {
            try { dynamic d = obj; return (string)d.Name; } catch { return "?"; }
        }
    }
}
