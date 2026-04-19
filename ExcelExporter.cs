// ExcelExporter.cs  —  C# 7.3
//
// 生成 Office 365 兼容的 .xlsx 文件（Open XML / SpreadsheetML）
//
// Office365 兼容要点（WPS 宽松，Office365 严格）：
//   1. UTF-8 无 BOM — 用 new UTF8Encoding(false) 而非 Encoding.UTF8
//   2. xl/worksheets/_rels/sheet1.xml.rels — worksheet 必须有自己的 rels 文件
//   3. [Content_Types].xml 必须包含 worksheet 的 Override 条目
//   4. 列宽元素 <cols> 必须在 <sheetData> 之前（按 xlsx spec 顺序）
//   5. 数字格式 numFmtId 使用内置 id（0=General, 4=0.00 小数）
//
// 坐标转换：绝对坐标 → Inverse(refMatrix) × absTx
// 欧拉角：ZYX 顺序，单位度

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace MyPlugin.ExportGun
{
    public class ExcelRow
    {
        public string OperationName;
        public string PointName;
        public string PointType;
        public double X, Y, Z;
        public double RX, RY, RZ;
    }

    public static class ExcelExporter
    {
        // UTF-8 无 BOM（Office365 严格要求）
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

        // ════════════════════════════════════════════════════════════
        //  主入口
        // ════════════════════════════════════════════════════════════
        public static string Export(List<OperationInfo> ops,
                                     double[]            refMatrix,
                                     string              outputFolder,
                                     Action<string>      log)
        {
            if (log == null) log = delegate(string s) { };

            List<ExcelRow> rows = BuildRows(ops, refMatrix, log);
            if (rows.Count == 0) { log("  ! 无数据可导出"); return null; }

            if (string.IsNullOrEmpty(outputFolder))
                outputFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "CatiaExport");
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            string ts   = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string path = Path.Combine(outputFolder, "WeldPoints_" + ts + ".xlsx");

            WriteXlsx(rows, path);
            log("[Excel] 已导出 " + rows.Count + " 行 -> " + path);
            return path;
        }

        // ════════════════════════════════════════════════════════════
        //  构建行数据
        // ════════════════════════════════════════════════════════════
        private static List<ExcelRow> BuildRows(List<OperationInfo> ops,
                                                  double[] refMatrix,
                                                  Action<string> log)
        {
            List<ExcelRow> rows = new List<ExcelRow>();
            foreach (OperationInfo op in ops)
            {
                foreach (PointInfo pt in op.Points)
                {
                    double[] rel = PsReader.ToRelative(pt.TCPMatrix, refMatrix);

                    double rx, ry, rz;
                    PsReader.MatrixToEulerDeg(rel, out rx, out ry, out rz);

                    rows.Add(new ExcelRow
                    {
                        OperationName = op.Name      ?? "",
                        PointName     = pt.Name      ?? "",
                        PointType     = PtLabel(pt.Type),
                        X  = Math.Round(rel[3],  4),
                        Y  = Math.Round(rel[7],  4),
                        Z  = Math.Round(rel[11], 4),
                        RX = Math.Round(rx, 4),
                        RY = Math.Round(ry, 4),
                        RZ = Math.Round(rz, 4)
                    });
                }
            }
            return rows;
        }

        private static string PtLabel(PointType t)
        {
            switch (t)
            {
                case PointType.WeldPoint:       return "焊点";
                case PointType.PathPoint:       return "路径点";
                case PointType.ContinuousPoint: return "连续点";
                default:                        return "未知";
            }
        }

        // ════════════════════════════════════════════════════════════
        //  写 xlsx（严格 Office365 兼容结构）
        //
        //  ZIP 内容：
        //    [Content_Types].xml
        //    _rels/.rels
        //    xl/workbook.xml
        //    xl/_rels/workbook.xml.rels
        //    xl/worksheets/sheet1.xml
        //    xl/worksheets/_rels/sheet1.xml.rels   ← Office365 必须
        //    xl/styles.xml
        //    xl/sharedStrings.xml
        // ════════════════════════════════════════════════════════════
        private static void WriteXlsx(List<ExcelRow> rows, string path)
        {
            // ── 共享字符串表 ─────────────────────────────────────────
            List<string> ss = new List<string>();
            Func<string, int> S = delegate(string v)
            {
                int i = ss.IndexOf(v);
                if (i < 0) { i = ss.Count; ss.Add(v); }
                return i;
            };

            // 列头 shared strings
            int H_OP   = S("操作名");
            int H_PT   = S("点名");
            int H_TYPE = S("点类型");
            int H_X    = S("X (mm)");
            int H_Y    = S("Y (mm)");
            int H_Z    = S("Z (mm)");
            int H_RX   = S("RX (°)");
            int H_RY   = S("RY (°)");
            int H_RZ   = S("RZ (°)");

            // 数据行中的字符串也要入 shared strings
            foreach (ExcelRow r in rows)
            {
                S(r.OperationName); S(r.PointName); S(r.PointType);
            }

            // ── sheet1.xml ───────────────────────────────────────────
            StringBuilder sb = new StringBuilder(rows.Count * 200 + 2000);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<worksheet");
            sb.Append(" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
            sb.Append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
            sb.Append(" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"");
            sb.Append(" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"");
            sb.Append(" mc:Ignorable=\"x14ac\">");

            // cols 必须在 sheetData 之前（xlsx 规范顺序）
            sb.Append("<cols>");
            sb.Append("<col min=\"1\" max=\"2\" width=\"24\" bestFit=\"1\" customWidth=\"1\"/>");
            sb.Append("<col min=\"3\" max=\"3\" width=\"12\" bestFit=\"1\" customWidth=\"1\"/>");
            sb.Append("<col min=\"4\" max=\"9\" width=\"14\" bestFit=\"1\" customWidth=\"1\"/>");
            sb.Append("</cols>");

            sb.Append("<sheetData>");

            // 标题行（style 1 = 粗体）
            sb.Append("<row r=\"1\" spans=\"1:9\">");
            AppendSsCell(sb, "A1", H_OP,   1);
            AppendSsCell(sb, "B1", H_PT,   1);
            AppendSsCell(sb, "C1", H_TYPE, 1);
            AppendSsCell(sb, "D1", H_X,    1);
            AppendSsCell(sb, "E1", H_Y,    1);
            AppendSsCell(sb, "F1", H_Z,    1);
            AppendSsCell(sb, "G1", H_RX,   1);
            AppendSsCell(sb, "H1", H_RY,   1);
            AppendSsCell(sb, "I1", H_RZ,   1);
            sb.Append("</row>");

            for (int i = 0; i < rows.Count; i++)
            {
                int row = i + 2;
                ExcelRow dr = rows[i];
                sb.Append("<row r=\"" + row + "\" spans=\"1:9\">");
                AppendSsCell(sb, CellAddr(1, row), S(dr.OperationName), 0);
                AppendSsCell(sb, CellAddr(2, row), S(dr.PointName),     0);
                AppendSsCell(sb, CellAddr(3, row), S(dr.PointType),     0);
                AppendNumCell(sb, CellAddr(4, row), dr.X,  2);
                AppendNumCell(sb, CellAddr(5, row), dr.Y,  2);
                AppendNumCell(sb, CellAddr(6, row), dr.Z,  2);
                AppendNumCell(sb, CellAddr(7, row), dr.RX, 2);
                AppendNumCell(sb, CellAddr(8, row), dr.RY, 2);
                AppendNumCell(sb, CellAddr(9, row), dr.RZ, 2);
                sb.Append("</row>");
            }
            sb.Append("</sheetData>");
            sb.Append("</worksheet>");
            string sheetXml = sb.ToString();

            // ── sharedStrings.xml ────────────────────────────────────
            StringBuilder ssSb = new StringBuilder(ss.Count * 50 + 200);
            ssSb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            ssSb.Append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
            ssSb.Append(" count=\"" + ss.Count + "\"");
            ssSb.Append(" uniqueCount=\"" + ss.Count + "\">");
            foreach (string v in ss)
            {
                ssSb.Append("<si><t xml:space=\"preserve\">");
                ssSb.Append(XmlEsc(v));
                ssSb.Append("</t></si>");
            }
            ssSb.Append("</sst>");

            // ── 写 ZIP ──────────────────────────────────────────────
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Create, false, Encoding.UTF8))
            {
                W(zip, "[Content_Types].xml",                  ContentTypes());
                W(zip, "_rels/.rels",                          RootRels());
                W(zip, "xl/workbook.xml",                      Workbook());
                W(zip, "xl/_rels/workbook.xml.rels",           WorkbookRels());
                W(zip, "xl/worksheets/sheet1.xml",             sheetXml);
                W(zip, "xl/worksheets/_rels/sheet1.xml.rels",  SheetRels());   // Office365必须
                W(zip, "xl/styles.xml",                        Styles());
                W(zip, "xl/sharedStrings.xml",                 ssSb.ToString());
            }
        }

        // ── 单元格辅助 ────────────────────────────────────────────────
        private static void AppendSsCell(StringBuilder sb,
            string addr, int ssIdx, int style)
        {
            sb.Append("<c r=\"").Append(addr)
              .Append("\" t=\"s\" s=\"").Append(style).Append("\">")
              .Append("<v>").Append(ssIdx).Append("</v></c>");
        }

        private static void AppendNumCell(StringBuilder sb,
            string addr, double val, int style)
        {
            sb.Append("<c r=\"").Append(addr)
              .Append("\" s=\"").Append(style).Append("\">")
              .Append("<v>").Append(val.ToString("G10")).Append("</v></c>");
        }

        // 列地址：1-based col number → "A","B"...
        private static string CellAddr(int col, int row)
        {
            string colStr = "";
            int c = col;
            while (c > 0)
            {
                c--;
                colStr = (char)('A' + c % 26) + colStr;
                c /= 26;
            }
            return colStr + row;
        }

        // ── Open XML 固定内容 ─────────────────────────────────────────

        private static string ContentTypes()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                 + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                 + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                 + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                 + "<Override PartName=\"/xl/workbook.xml\""
                 +   " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                 + "<Override PartName=\"/xl/worksheets/sheet1.xml\""
                 +   " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                 + "<Override PartName=\"/xl/styles.xml\""
                 +   " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
                 + "<Override PartName=\"/xl/sharedStrings.xml\""
                 +   " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
                 + "</Types>";
        }

        private static string RootRels()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                 + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                 + "<Relationship Id=\"rId1\""
                 +   " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\""
                 +   " Target=\"xl/workbook.xml\"/>"
                 + "</Relationships>";
        }

        private static string Workbook()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                 + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                 +   " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
                 +   " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\""
                 +   " mc:Ignorable=\"x15 xr xr6 xr10 xr2\">"
                 + "<fileVersion appName=\"xl\" lastEdited=\"7\" lowestEdited=\"7\"/>"
                 + "<workbookPr defaultThemeVersion=\"166925\"/>"
                 + "<sheets>"
                 +   "<sheet name=\"焊点数据\" sheetId=\"1\" r:id=\"rId1\"/>"
                 + "</sheets>"
                 + "<calcPr calcId=\"181029\"/>"
                 + "</workbook>";
        }

        private static string WorkbookRels()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                 + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                 + "<Relationship Id=\"rId1\""
                 +   " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\""
                 +   " Target=\"worksheets/sheet1.xml\"/>"
                 + "<Relationship Id=\"rId2\""
                 +   " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\""
                 +   " Target=\"styles.xml\"/>"
                 + "<Relationship Id=\"rId3\""
                 +   " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\""
                 +   " Target=\"sharedStrings.xml\"/>"
                 + "</Relationships>";
        }

        // worksheet 的 rels 文件（Office365 必须，即使内容为空关系列表）
        private static string SheetRels()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                 + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                 + "</Relationships>";
        }

        private static string Styles()
        {
            // style 0 = 普通文本
            // style 1 = 粗体（列头）
            // style 2 = 数值，保留4位小数（numFmtId=4 = built-in "0.00"，
            //           改用自定义格式 164 = "0.0000"）
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                 + "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""
                 +   " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">"
                 + "<numFmts count=\"1\">"
                 +   "<numFmt numFmtId=\"164\" formatCode=\"0.0000\"/>"
                 + "</numFmts>"
                 + "<fonts count=\"2\">"
                 +   "<font><sz val=\"10\"/><name val=\"等线\"/><family val=\"2\"/><charset val=\"134\"/></font>"
                 +   "<font><b/><sz val=\"10\"/><name val=\"等线\"/><family val=\"2\"/><charset val=\"134\"/></font>"
                 + "</fonts>"
                 + "<fills count=\"2\">"
                 +   "<fill><patternFill patternType=\"none\"/></fill>"
                 +   "<fill><patternFill patternType=\"gray125\"/></fill>"
                 + "</fills>"
                 + "<borders count=\"1\">"
                 +   "<border><left/><right/><top/><bottom/><diagonal/></border>"
                 + "</borders>"
                 + "<cellStyleXfs count=\"1\">"
                 +   "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>"
                 + "</cellStyleXfs>"
                 + "<cellXfs count=\"3\">"
                 +   "<xf numFmtId=\"0\"   fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
                 +   "<xf numFmtId=\"0\"   fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
                 +   "<xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>"
                 + "</cellXfs>"
                 + "<cellStyles count=\"1\">"
                 +   "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>"
                 + "</cellStyles>"
                 + "</styleSheet>";
        }

        private static void W(ZipArchive zip, string name, string content)
        {
            ZipArchiveEntry entry = zip.CreateEntry(name, CompressionLevel.Fastest);
            // 注意：必须用无BOM的UTF-8，Office365对BOM非常敏感
            using (StreamWriter w = new StreamWriter(entry.Open(), Utf8NoBom))
                w.Write(content);
        }

        private static string XmlEsc(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("&", "&amp;")
                    .Replace("<", "&lt;")
                    .Replace(">", "&gt;")
                    .Replace("\"", "&quot;");
        }
    }
}
