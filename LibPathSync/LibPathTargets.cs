using System;
using System.Text;
using Microsoft.Win32;

namespace TxTools.LibPathSync
{
    /// <summary>
    /// 库路径同步的共享配置与注册表读写。
    /// 需要再加同步目标时，只在 Targets 表里加一行即可，两个按钮自动生效。
    /// 所有键均位于 HKCU，值类型 REG_SZ。
    /// </summary>
    internal static class LibPathTargets
    {
        // 同步源：System Root（EMS 选项）
        public const string SystemRootKey =
            @"Software\TECNOMATIX\TUNE\NewAssembler\Options\EMS";
        public const string SystemRootValue = "System Root Path";

        // 同步目标表：(子键, 键名, 说明)
        public static readonly (string Key, string Value, string Label)[] Targets =
        {
            (@"Software\TECNOMATIX\TUNE\NewAssembler\Commands Info\DnDataAdministrationCommands\Values\CCoDACDefineComponentTypeCmd",
             "LastFolder", "Define Component Type"),

            (@"Software\TECNOMATIX\TUNE\NewAssembler\Commands Info\ComponentOperations\Dialogs\2005\Default",
             "GET_COMPONENT_DIALOG_LAST_FOLDER", "Insert Component"),

            (@"Software\TECNOMATIX\TUNE\NewAssembler\Commands Info\DnProcessSimulateCommands\Dialogs\CUiImportMFGsFromFileDialog",
             "LastBrowsedDirectory", "Import Mfgs"),
        };

        /// <summary>把同一个文件夹写入所有目标键（键不存在自动创建）。</summary>
        public static void WriteAll(string folder)
        {
            foreach (var t in Targets)
                WriteValue(t.Key, t.Value, folder);
        }

        /// <summary>回读所有目标，生成多行校验汇总。</summary>
        public static string ReadBackSummary()
        {
            var sb = new StringBuilder();
            foreach (var t in Targets)
            {
                string v = ReadValue(t.Key, t.Value);
                sb.AppendLine("• " + t.Label + " (" + t.Value + ")：");
                sb.AppendLine("  " + (v ?? "<空>"));
            }
            return sb.ToString();
        }

        public static string ReadValue(string subKey, string valueName)
        {
            // HKCU\Software 不受 WOW6432Node 重定向影响
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(subKey, false))
            {
                if (key == null) return null;
                return key.GetValue(valueName) as string;
            }
        }

        public static void WriteValue(string subKey, string valueName, string value)
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(subKey))
            {
                if (key == null)
                    throw new InvalidOperationException("无法创建/打开注册表键：HKCU\\" + subKey);
                key.SetValue(valueName, value, RegistryValueKind.String);
            }
        }
    }
}
