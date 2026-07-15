// OpenResourceFolderCmd.cs  —  C# 7.3   (v2: 修复 PS 2402 SDK 编译错误)
// TxTools 插件：打开选中资源所在目录
//
// 行为：
//   - 无对话框，点击按钮直接执行
//   - 读取 PS 当前选中对象（TxApplication.ActiveSelection）
//   - 解析每个对象的磁盘存储路径（优先 ITxStorable.StorageObject → TxLibraryStorage.FullPath，
//     回退反射常见路径属性）
//   - 用 explorer.exe /select,"path" 打开所在目录并高亮该文件
//   - 同一目录只打开一次，避免多次弹窗
//
// v2 修复（vs v1）:
//   ① TxButtonCommand.Name 是 abstract，必须实现 → 加 override Name
//   ② TxMessageEndpoint 在本 SDK 不存在 → 改为反射式日志器，
//      探测 TxApplication 上的静态日志方法；找不到 fallback 到 Debug.WriteLine

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Tecnomatix.Engineering;

namespace TxTools.OpenResourceFolder
{
    public class OpenResourceFolderCmd : TxButtonCommand
    {
        private const string LOG_PREFIX = "[OpenResFolder] ";

        public override string Name => ".打开资源所在目录";
        public override string Description { get { return "打开选中资源所在目录"; } }
        public override string Category => "TxTools";
        public override string LargeBitmap => "image.file.png";

        public override void Execute(object cmdParams)
        {
            try
            {
                ExecuteCore();
            }
            catch (Exception ex)
            {
                Log("顶层异常: " + ex.Message);
            }
        }

        private void ExecuteCore()
        {
            TxObjectList sel = null;
            try { sel = TxApplication.ActiveSelection.GetItems(); }
            catch (Exception ex) { Log("读取选中失败: " + ex.Message); return; }

            if (sel == null || sel.Count == 0)
            {
                Log("未选中任何对象。请在 PS 中选中一个资源（Component/Part/Tool 等）后再点击。");
                return;
            }

            // 去重：同一文件路径只打开一次；同一目录也只 explorer 一次
            var openedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var openedDirs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int okCount = 0;
            int failCount = 0;

            foreach (ITxObject obj in sel)
            {
                if (obj == null) { failCount++; continue; }

                string name = SafeName(obj);
                string typeName = obj.GetType().Name;
                string path = TryResolveStoragePath(obj);

                if (string.IsNullOrEmpty(path))
                {
                    Log(string.Format("× [{0}] {1}：未找到磁盘路径（无 StorageObject/路径属性）",
                        typeName, name));
                    failCount++;
                    continue;
                }

                // 规范化
                try { path = Path.GetFullPath(path); } catch { }

                if (openedPaths.Contains(path))
                {
                    // 同一文件已处理过，跳过即可
                    continue;
                }
                openedPaths.Add(path);

                if (File.Exists(path))
                {
                    string dir = Path.GetDirectoryName(path);
                    if (openedDirs.Add(dir ?? path))
                    {
                        if (OpenAndSelect(path))
                        {
                            Log(string.Format("√ [{0}] {1}  →  {2}", typeName, name, path));
                            okCount++;
                        }
                        else
                        {
                            failCount++;
                        }
                    }
                    else
                    {
                        // 同目录里的另一个文件——目录已打开，不再重复
                        okCount++;
                    }
                }
                else if (Directory.Exists(path))
                {
                    if (openedDirs.Add(path))
                    {
                        if (OpenFolder(path))
                        {
                            Log(string.Format("√ [{0}] {1}  →  {2}（目录）", typeName, name, path));
                            okCount++;
                        }
                        else
                        {
                            failCount++;
                        }
                    }
                }
                else
                {
                    // 路径不存在：尝试它的父目录
                    string dir = null;
                    try { dir = Path.GetDirectoryName(path); } catch { }
                    if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
                    {
                        if (openedDirs.Add(dir))
                        {
                            if (OpenFolder(dir))
                            {
                                Log(string.Format("△ [{0}] {1}：文件不存在，打开父目录 {2}",
                                    typeName, name, dir));
                                okCount++;
                            }
                            else
                            {
                                failCount++;
                            }
                        }
                    }
                    else
                    {
                        Log(string.Format("× [{0}] {1}：路径无效 {2}", typeName, name, path));
                        failCount++;
                    }
                }
            }

            Log(string.Format("完成：成功 {0}，失败 {1}（共选中 {2}）", okCount, failCount, sel.Count));
        }

        // ════════════════════════════════════════════════════════════
        //  路径解析（参考 PsReader.TryGetToolStorageDir 的策略组合）
        // ════════════════════════════════════════════════════════════

        /// <summary>
        /// 解析对象在磁盘上的存储路径（文件或目录）。
        /// 策略 1：ITxStorable.StorageObject → TxLibraryStorage.FullPath
        /// 策略 2：反射常见路径属性
        /// 失败返回 null。
        /// </summary>
        private static string TryResolveStoragePath(ITxObject obj)
        {
            if (obj == null) return null;

            // —— 策略 1：ITxStorable ——
            try
            {
                if (obj is ITxStorable storable)
                {
                    TxStorage storage = null;
                    try { storage = storable.StorageObject; } catch { }

                    if (storage != null)
                    {
                        // 1a. 强类型 TxLibraryStorage
                        TxLibraryStorage libStorage = storage as TxLibraryStorage;
                        if (libStorage != null)
                        {
                            try
                            {
                                string fp = libStorage.FullPath;
                                if (!string.IsNullOrEmpty(fp)) return fp;
                            }
                            catch { }
                        }

                        // 1b. 兜底：dynamic 取 FullPath（不同子类都有同名属性）
                        try
                        {
                            dynamic dyn = storage;
                            string p = dyn.FullPath as string;
                            if (!string.IsNullOrEmpty(p)) return p;
                        }
                        catch { }
                    }
                }
            }
            catch { }

            // —— 策略 2：反射常见路径属性 ——
            string[] pathProps = {
                "ExternalFilePath", "FilePath", "ModelFilePath", "SourceFilePath",
                "CgrFilePath", "GeometryFilePath", "ResourceFilePath", "ResourcePath",
                "ExternalFile", "FileLocation", "DataFilePath", "JtFilePath", "FullPath"
            };
            foreach (string prop in pathProps)
            {
                try
                {
                    var pi = obj.GetType().GetProperty(prop,
                        BindingFlags.Public | BindingFlags.Instance);
                    if (pi == null) continue;
                    string s = pi.GetValue(obj) as string;
                    if (!string.IsNullOrEmpty(s)) return s;
                }
                catch { }
            }

            return null;
        }

        // ════════════════════════════════════════════════════════════
        //  Explorer 调用
        // ════════════════════════════════════════════════════════════

        /// <summary>用 explorer.exe /select,"path" 打开目录并高亮指定文件。</summary>
        private bool OpenAndSelect(string fullFilePath)
        {
            try
            {
                // 注意 /select 后面用逗号分隔，路径加引号以兼容空格
                var psi = new ProcessStartInfo("explorer.exe",
                    "/select,\"" + fullFilePath + "\"")
                {
                    UseShellExecute = true
                };
                Process.Start(psi);
                return true;
            }
            catch (Exception ex)
            {
                Log("explorer /select 失败 [" + fullFilePath + "]: " + ex.Message);
                return false;
            }
        }

        /// <summary>直接打开目录（无高亮文件）。</summary>
        private bool OpenFolder(string folderPath)
        {
            try
            {
                var psi = new ProcessStartInfo("explorer.exe", "\"" + folderPath + "\"")
                {
                    UseShellExecute = true
                };
                Process.Start(psi);
                return true;
            }
            catch (Exception ex)
            {
                Log("explorer 打开目录失败 [" + folderPath + "]: " + ex.Message);
                return false;
            }
        }

        // ════════════════════════════════════════════════════════════
        //  工具
        // ════════════════════════════════════════════════════════════

        private static string SafeName(ITxObject obj)
        {
            try { return obj.Name ?? "(无名)"; }
            catch { return "(无名)"; }
        }

        // ════════════════════════════════════════════════════════════
        //  日志（反射式：尽量输出到 PS，无可用出口时静默到 Debug）
        // ════════════════════════════════════════════════════════════

        private static MethodInfo _logMethod;
        private static bool _logProbed;

        private static void Log(string msg)
        {
            string full = LOG_PREFIX + msg;

            if (!_logProbed)
            {
                _logProbed = true;
                ProbeLogMethod();
            }

            if (_logMethod != null)
            {
                try { _logMethod.Invoke(null, new object[] { full }); return; }
                catch { /* 出错就走 fallback */ }
            }

            try { Debug.WriteLine(full); } catch { }
        }

        private static void ProbeLogMethod()
        {
            // 候选：TxApplication 上的静态日志方法（不同版本命名不同）
            string[] candidates = { "WriteToOutput", "WriteMessage", "PrintMessage", "ReportInfo", "Log" };
            try
            {
                Type appT = typeof(TxApplication);
                foreach (var name in candidates)
                {
                    var m = appT.GetMethod(name,
                        BindingFlags.Public | BindingFlags.Static,
                        null, new[] { typeof(string) }, null);
                    if (m != null) { _logMethod = m; return; }
                }
            }
            catch { }
        }
    }
}