// ExcelExporter.cs  —  C# 7.3
//
// 【完全独立的工具类，不依赖任何项目类型；输入仅用 string[] / IList<string[]>】
// 通过 late-bound COM 启动 Excel、新建空白工作簿，把表头与数据行注入到工作表中。
// 不生成任何文件 —— Excel 以可见状态打开，由用户自行决定是否保存。
//
// 用法：
//   ExcelExporter.Inject(
//       new[] { "新轨迹", "参考焊点", "待分配焊点", "依据", "距离(mm)" },
//       rows,                 // IList<string[]>，每行长度可与表头一致
//       "匹配结果",            // 工作表名，可空
//       msg => Log(msg));     // 日志回调，可空

using TxTools.ExportGun;
using System;
using System.Collections.Generic;

namespace TxTools.ExcelInjectTool
{
    public static class ExcelExporter
    {
        /// <summary>启动 Excel、新建空白工作簿，把表头+数据注入到工作表（不落地文件）。</summary>
        public static void Inject(string[] headers, IList<string[]> rows, string sheetName, Action<string> log)
        {
            if (log == null) log = delegate (string s) { };

            object excelObj = null;
            try
            {
                Type t = Type.GetTypeFromProgID("Excel.Application");
                if (t == null) { log("[Excel] 未检测到 Excel（ProgID Excel.Application 缺失）"); return; }

                excelObj = Activator.CreateInstance(t);
                dynamic app = excelObj;
                app.Visible = true;
                app.DisplayAlerts = false;

                dynamic wb = app.Workbooks.Add();      // 新建空白工作簿
                dynamic ws = wb.ActiveSheet;           // 默认空白工作表
                if (!string.IsNullOrEmpty(sheetName)) { try { ws.Name = sheetName; } catch { } }

                int headerCount = headers == null ? 0 : headers.Length;

                // 表头
                for (int c = 0; c < headerCount; c++)
                    ws.Cells[1, c + 1] = headers[c] ?? "";

                // 数据行
                int r = headerCount > 0 ? 2 : 1;
                int written = 0;
                if (rows != null)
                {
                    foreach (var row in rows)
                    {
                        if (row == null) continue;
                        for (int c = 0; c < row.Length; c++)
                            ws.Cells[r, c + 1] = row[c] ?? "";
                        r++; written++;
                    }
                }

                // 美化：表头加粗 + 自动列宽
                try
                {
                    if (headerCount > 0)
                        ws.Range[ws.Cells[1, 1], ws.Cells[1, headerCount]].Font.Bold = true;
                    ws.Columns.AutoFit();
                }
                catch { }

                log("[Excel] 已注入 " + written + " 行数据到新工作簿");
            }
            catch (Exception ex)
            {
                log("[Excel] 注入异常：" + ex.Message);
                // 出错也尽量把已打开的 Excel 留给用户
                try { if (excelObj != null) ((dynamic)excelObj).Visible = true; } catch { }
            }
        }

        internal static string Export(List<OperationInfo> ops, double[] refMatrix, string folder, Action<string> nolog)
        {
            throw new NotImplementedException();
        }
    }
}
