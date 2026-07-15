using System;
using System.Text;
using Tecnomatix.Engineering;

namespace TxTools.WeldGunDefiner.Core
{
    /// <summary>
    /// 运行时诊断：枚举已加载程序集，找出Joint/Link相关的CreationData类型
    /// 在插件启动时或创建失败时调用，将结果写入日志供开发分析
    /// </summary>
    public static class SdkDiagnostic
    {
        public static string DiagnoseJointCreation()
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== SDK Joint/Link CreationData 诊断 ===");

            // 1. 在所有已加载程序集中搜索包含 "Joint" 的公共类型
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                string asmName = asm.GetName().Name;
                if (!asmName.StartsWith("Tecnomatix", StringComparison.OrdinalIgnoreCase))
                    continue;

                try
                {
                    foreach (var t in asm.GetExportedTypes())
                    {
                        string tn = t.Name;
                        if ((tn.IndexOf("Joint", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             tn.IndexOf("Link",  StringComparison.OrdinalIgnoreCase) >= 0 ||
                             tn.IndexOf("Crank", StringComparison.OrdinalIgnoreCase) >= 0)
                            && (tn.IndexOf("Creation", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                tn.IndexOf("Data",     StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            sb.AppendLine($"  [{asmName}] {t.FullName}");

                            // 列出该类的公开属性
                            foreach (var p in t.GetProperties(
                                System.Reflection.BindingFlags.Public |
                                System.Reflection.BindingFlags.Instance))
                            {
                                sb.AppendLine($"      .{p.Name} : {p.PropertyType.Name}");
                            }
                        }
                    }
                }
                catch { }
            }

            // 2. 确认 ITxKinematicsModellable.CreateJoint 的参数类型
            sb.AppendLine("\n=== ITxKinematicsModellable.CreateJoint 签名 ===");
            var ifaceType = typeof(ITxKinematicsModellable);
            foreach (var mi in ifaceType.GetMethods())
            {
                if (mi.Name.IndexOf("Joint", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    mi.Name.IndexOf("Link",  StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    sb.Append($"  {mi.Name}(");
                    var pars = mi.GetParameters();
                    for (int i = 0; i < pars.Length; i++)
                        sb.Append($"{pars[i].ParameterType.Name} {pars[i].Name}{(i<pars.Length-1?", ":"")}");
                    sb.AppendLine($") -> {mi.ReturnType.Name}");
                }
            }

            // 3. 确认 TxJoint 的可写属性（特别是 KinematicsFunction）
            sb.AppendLine("\n=== TxJoint 属性 ===");
            try
            {
                foreach (var p in typeof(TxJoint).GetProperties(
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.Instance))
                {
                    sb.AppendLine($"  {(p.CanWrite?"rw":"ro")} {p.PropertyType.Name} {p.Name}");
                }
            }
            catch { }

            // 4. 确认 TxJointAxis 构造函数
            sb.AppendLine("\n=== TxJointAxis 构造函数 ===");
            try
            {
                foreach (var ctor in typeof(TxJointAxis).GetConstructors())
                {
                    sb.Append("  TxJointAxis(");
                    var pars = ctor.GetParameters();
                    for (int i = 0; i < pars.Length; i++)
                        sb.Append($"{pars[i].ParameterType.Name} {pars[i].Name}{(i<pars.Length-1?", ":"")}");
                    sb.AppendLine(")");
                }
            }
            catch { }

            return sb.ToString();
        }
    }
}
