using System;
using System.IO;
using System.Windows.Forms;
using Tecnomatix.Engineering;

namespace TxTools.LibPathSync
{
    /// <summary>
    /// 将所有目标对话框的“上次目录”同步为当前选中组件的库存储目录。
    /// 取路径链路：ITxStorable -> StorageObject -> (TxLibraryStorage).FullPath，再归一化为文件夹。
    /// 目标键集中维护在 LibPathTargets.Targets。
    /// </summary>
    public class SyncToSelectedComponentCommand : TxButtonCommand
    {
        public override string Name => ".同步到所选组件";
        public override string Category => "TxTools";
        public override string Description => "将目录路径同步为当前选中组件的库存储目录";

        public override void Execute(object cmdParams)
        {
            try
            {
                // 1. 取当前选择（ribbon 点击即在 PS 主线程，可直接调 SDK）
                TxObjectList selected = TxApplication.ActiveSelection.GetItems();
                if (selected == null || selected.Count == 0)
                {
                    MessageBox.Show("请先在树或视图中选择一个组件，再执行同步。",
                        "同步到所选组件", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ITxObject comp = selected[0] as ITxObject;

                // 2. 解析组件的原始存储路径
                string raw = TryGetStoragePath(comp);
                if (string.IsNullOrEmpty(raw))
                {
                    MessageBox.Show(
                        "无法从所选对象解析出库存储路径。\n对象类型：" +
                        (comp?.GetType().FullName ?? "<null>"),
                        "同步到所选组件", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 3. 归一化为文件夹
                string folder = NormalizeToFolder(raw);

                // 4. 写入所有目标键并回读
                LibPathTargets.WriteAll(folder);

                MessageBox.Show(
                    "组件：" + (comp.Name ?? "") + "\n" +
                    "原始存储路径：\n" + raw + "\n" +
                    "已写入目录：\n" + folder + "\n\n" +
                    LibPathTargets.ReadBackSummary(),
                    "同步完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("同步失败：\n" + ex.Message,
                    "同步到所选组件", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>沿 ITxStorable -> StorageObject -> FullPath 取路径，多重兜底。</summary>
        private static string TryGetStoragePath(ITxObject obj)
        {
            if (obj == null) return null;

            // 1) 标准链路
            try
            {
                ITxStorable storable = obj as ITxStorable;
                if (storable != null)
                {
                    dynamic storage = storable.StorageObject;   // 多为 TxLibraryStorage
                    string fp = TryReadPath(storage);
                    if (!string.IsNullOrEmpty(fp)) return fp;
                }
            }
            catch { }

            // 2) 直接在对象上探测（部分类型/版本）
            try
            {
                string fp = TryReadPath((dynamic)obj);
                if (!string.IsNullOrEmpty(fp)) return fp;
            }
            catch { }

            return null;
        }

        private static string TryReadPath(dynamic o)
        {
            if (o == null) return null;
            try { string s = o.FullPath; if (!string.IsNullOrEmpty(s)) return s; } catch { }
            try { string s = o.Path; if (!string.IsNullOrEmpty(s)) return s; } catch { }
            try { string s = o.FileName; if (!string.IsNullOrEmpty(s)) return s; } catch { }
            return null;
        }

        /// <summary>把文件/容器路径归一化为对话框可用的文件夹。</summary>
        private static string NormalizeToFolder(string path)
        {
            string folder = path;
            try
            {
                if (File.Exists(path))
                    folder = Path.GetDirectoryName(path);
                else if (!Directory.Exists(path))
                    folder = Path.GetDirectoryName(path);
            }
            catch { }

            // 若落在 .cojt / .co 组件容器内，上溯一级到组件文件夹（与示例 ...\VAN\5400710 一致）
            // 如你的库结构不需要这层处理，删掉下面这段即可。
            try
            {
                string trimmed = folder?.TrimEnd('\\', '/');
                string leaf = Path.GetFileName(trimmed ?? "");
                if (leaf.EndsWith(".cojt", StringComparison.OrdinalIgnoreCase) ||
                    leaf.EndsWith(".co", StringComparison.OrdinalIgnoreCase) ||
                    leaf.EndsWith(".jt", StringComparison.OrdinalIgnoreCase))
                {
                    string parent = Path.GetDirectoryName(trimmed);
                    if (!string.IsNullOrEmpty(parent)) folder = parent;
                }
            }
            catch { }

            return folder;
        }
    }
}