using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.Streaming;
using NPOI.OpenXml4Net.OPC;
using NPOI.XSSF.EventUserModel;
using NPOI.XSSF.Model;

namespace SheetAppendApp
{
    public sealed class MergeReport
    {
        public int InputFiles { get; set; }
        public int FilesSucceeded { get; set; }
        public int FilesFailed { get; set; }
        public int SheetsMerged { get; set; }
        public long RowsWritten { get; set; }
        public long DataRows { get; set; }
        public long BlankRowsPreserved { get; set; }
        public long HeadersSkipped { get; set; }
        public int StylesCreated { get; set; }
    }

    public static class NpoiMerger
    {
        private static readonly string[] ValidExts = { ".xls", ".xlsx", ".xlsm", ".csv" };

        public static MergeReport MergeFast_ToOneXlsx(
            string[] inputFiles, string outputPath,
            bool trySkipHeaderIfMatchesBase, Action<string>? log = null)
        {
            log ??= _ => { };
            var files = NormalizeFiles(inputFiles);
            if (files.Length == 0)
                throw new InvalidOperationException("Không có file Excel/CSV hợp lệ.");
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("outputPath is required.", nameof(outputPath));
            if (!outputPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                outputPath += ".xlsx";

            var report = new MergeReport { InputFiles = files.Length };

            using var outWb = new SXSSFWorkbook(100);
            outWb.CompressTempFiles = true;
            int sheetIndex = 1;
            var outSheet = outWb.CreateSheet($"Merged_{sheetIndex}");

            var styleMap = new Dictionary<string, ICellStyle>();
            var fontMap = new Dictionary<string, IFont>();
            var dataFormatMap = new Dictionary<string, short>();

            bool isFirstSheet = true;
            int destRow = 0;
            string[]? baseHeader = null;

            for (int fi = 0; fi < files.Length; fi++)
            {
                var file = files[fi];
                log($"[FILE {fi + 1}/{files.Length}] {Path.GetFileName(file)}");

                long fileSize = 0;
                try { fileSize = new FileInfo(file).Length; } catch { }

                try
                {
                    var ext = Path.GetExtension(file).ToLowerInvariant();
                    if (ext == ".csv")
                    {
                        MergeCsvFile(file, outWb, ref outSheet, ref sheetIndex, ref destRow, report, log, ref isFirstSheet, ref baseHeader, trySkipHeaderIfMatchesBase);
                    }
                    else if (ext == ".xls")
                    {
                        MergeXlsFile(file, outWb, ref outSheet, ref sheetIndex, ref destRow, styleMap, fontMap, dataFormatMap, report, log, ref isFirstSheet, ref baseHeader, trySkipHeaderIfMatchesBase);
                    }
                    else if (ext == ".xlsx" || ext == ".xlsm")
                    {
                        MergeXlsxSax(file, outWb, ref outSheet, ref sheetIndex, ref destRow, styleMap, fontMap, dataFormatMap, report, log, ref isFirstSheet, ref baseHeader, trySkipHeaderIfMatchesBase);
                    }
                }
                catch (Exception ex)
                {
                    report.FilesFailed++;
                    log($"  [ERROR] Lỗi gộp: {ex.Message}");
                }

                if (fileSize > 3 * 1024 * 1024)
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }
            }

            log($"[STYLE INFO] Đã tạo cực thấp: {report.StylesCreated} định dạng độc lập (Bảo vệ an toàn chống lỗi Corrupt file của Excel).");

            if (report.RowsWritten == 0)
                throw new InvalidOperationException("Không có dữ liệu hợp lệ nào được lấy ra.");

            var outDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrWhiteSpace(outDir)) Directory.CreateDirectory(outDir);

            using (var outFs = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None))
                outWb.Write(outFs);

            outWb.Dispose();

            log($"[DONE] Đã lưu thành công: {outputPath}");
            return report;
        }

        // ======================== TÁCH STYLE THÀNH DỮ LIỆU CHUỖI FINGERPRINT ========================
        private class StyleInfo
        {
            public string FontName = "";
            public double FontHeight;
            public bool IsBold, IsItalic, IsStrikeout;
            public FontUnderlineType Underline;
            public string FontColor = "";

            public NPOI.SS.UserModel.HorizontalAlignment Align;
            public VerticalAlignment VAlign;
            public bool WrapText;

            public FillPattern FillPat;
            public string FillFg = "";
            public string FillBg = "";

            public NPOI.SS.UserModel.BorderStyle BTop, BBot, BLef, BRig;
            public string CTop = "", CBot = "", CLef = "", CRig = "";

            public string DataFmt = "General";

            public string Fingerprint() => $"{FontName}|{FontHeight}|{IsBold}|{IsItalic}|{Underline}|{IsStrikeout}|{FontColor}|" +
                $"{Align}|{VAlign}|{WrapText}|" +
                $"{FillPat}|{FillFg}|{FillBg}|" +
                $"{BTop}|{BBot}|{BLef}|{BRig}|{CTop}|{CBot}|{CLef}|{CRig}|{DataFmt}";
            
            public string FontFingerprint() => $"{FontName}|{FontHeight}|{IsBold}|{IsItalic}|{Underline}|{IsStrikeout}|{FontColor}";
        }

        private static StyleInfo ExtractStyleInfo(ICellStyle s, IWorkbook? wb)
        {
            var info = new StyleInfo {
                Align = s.Alignment,
                VAlign = s.VerticalAlignment,
                WrapText = s.WrapText,
                FillPat = s.FillPattern,
                BTop = s.BorderTop, BBot = s.BorderBottom, BLef = s.BorderLeft, BRig = s.BorderRight
            };

            info.DataFmt = s.GetDataFormatString() ?? "General";
            info.FillFg = GetColorStr(s, wb, "FillFg");
            info.FillBg = GetColorStr(s, wb, "FillBg");
            info.CTop = GetColorStr(s, wb, "BTop");
            info.CBot = GetColorStr(s, wb, "BBot");
            info.CLef = GetColorStr(s, wb, "BLef");
            info.CRig = GetColorStr(s, wb, "BRig");

            try {
                IFont? f = null;
                if (s is XSSFCellStyle xs) f = xs.GetFont();
                else if (wb != null) f = wb.GetFontAt(s.FontIndex);

                if (f != null) {
                    info.FontName = f.FontName;
                    info.FontHeight = f.FontHeightInPoints;
                    info.IsBold = f.IsBold;
                    info.IsItalic = f.IsItalic;
                    info.Underline = f.Underline;
                    info.IsStrikeout = f.IsStrikeout;
                    if (f is XSSFFont xf) {
                        var xc = xf.GetXSSFColor();
                        info.FontColor = XColorToString(xc);
                    } else if (wb is HSSFWorkbook hw) {
                        var rgb = GetHssfRgb(hw, f.Color);
                        if (rgb != null) info.FontColor = BitConverter.ToString(rgb);
                    }
                }
            } catch { }

            return info;
        }

        private static void RenderStyleToDest(
            StyleInfo info, ICellStyle ds, IWorkbook destWb,
            Dictionary<string, IFont> fontMap, Dictionary<string, short> fmtMap)
        {
            ds.Alignment = info.Align;
            ds.VerticalAlignment = info.VAlign;
            ds.WrapText = info.WrapText;
            ds.BorderTop = info.BTop;
            ds.BorderBottom = info.BBot;
            ds.BorderLeft = info.BLef;
            ds.BorderRight = info.BRig;
            ds.FillPattern = info.FillPat;

            if (ds is XSSFCellStyle xds)
            {
                if (info.FillPat != FillPattern.NoFill) {
                    var fg = StringToXColor(info.FillFg); if (fg != null) xds.FillForegroundXSSFColor = fg;
                    var bg = StringToXColor(info.FillBg); if (bg != null) xds.FillBackgroundXSSFColor = bg;
                }
                
                if (info.BTop != NPOI.SS.UserModel.BorderStyle.None) { var c = StringToXColor(info.CTop); if (c != null) xds.SetTopBorderColor(c); }
                if (info.BBot != NPOI.SS.UserModel.BorderStyle.None) { var c = StringToXColor(info.CBot); if (c != null) xds.SetBottomBorderColor(c); }
                if (info.BLef != NPOI.SS.UserModel.BorderStyle.None) { var c = StringToXColor(info.CLef); if (c != null) xds.SetLeftBorderColor(c); }
                if (info.BRig != NPOI.SS.UserModel.BorderStyle.None) { var c = StringToXColor(info.CRig); if (c != null) xds.SetRightBorderColor(c); }
            }

            if (!string.IsNullOrEmpty(info.FontName))
            {
                var ffp = info.FontFingerprint();
                if (!fontMap.TryGetValue(ffp, out var dFont)) {
                    dFont = destWb.CreateFont();
                    dFont.FontName = info.FontName;
                    dFont.FontHeightInPoints = info.FontHeight;
                    dFont.IsBold = info.IsBold;
                    dFont.IsItalic = info.IsItalic;
                    dFont.Underline = info.Underline;
                    dFont.IsStrikeout = info.IsStrikeout;
                    
                    var fc = StringToXColor(info.FontColor);
                    if (fc != null && dFont is XSSFFont xf) xf.SetColor(fc);
                    
                    fontMap[ffp] = dFont;
                }
                ds.SetFont(dFont);
            }

            if (!string.IsNullOrEmpty(info.DataFmt) && info.DataFmt != "General" && info.DataFmt != "@")
            {
                if (!fmtMap.TryGetValue(info.DataFmt, out short fidx)) {
                    fidx = destWb.CreateDataFormat().GetFormat(info.DataFmt);
                    fmtMap[info.DataFmt] = fidx;
                }
                ds.DataFormat = fidx;
            }
        }

        private static string GetColorStr(ICellStyle s, IWorkbook? wb, string type)
        {
            if (s is XSSFCellStyle xs) {
                var xc = type switch { "FillFg"=>xs.FillForegroundXSSFColor, "FillBg"=>xs.FillBackgroundXSSFColor, "BTop"=>xs.TopBorderXSSFColor, "BBot"=>xs.BottomBorderXSSFColor, "BLef"=>xs.LeftBorderXSSFColor, "BRig"=>xs.RightBorderXSSFColor, _=>null };
                return XColorToString(xc);
            }
            if (wb is HSSFWorkbook hw) {
                short idx = type switch { "FillFg"=>s.FillForegroundColor, "FillBg"=>s.FillBackgroundColor, "BTop"=>s.TopBorderColor, "BBot"=>s.BottomBorderColor, "BLef"=>s.LeftBorderColor, "BRig"=>s.RightBorderColor, _=>(short)0 };
                var rgb = GetHssfRgb(hw, idx);
                if (rgb != null) return BitConverter.ToString(rgb);
            }
            return "";
        }

        private static string XColorToString(XSSFColor? c)
        {
            if (c == null) return "";
            if (c.RGB != null && c.RGB.Length >= 3) return BitConverter.ToString(c.RGB);
            if (c.Theme >= 0) return $"T{c.Theme}_{c.Tint:G}";
            return "";
        }

        private static XSSFColor? StringToXColor(string val)
        {
            if (string.IsNullOrEmpty(val)) return null;
            if (val.StartsWith("T")) {
                var p = val.Substring(1).Split('_');
                if (p.Length == 2 && int.TryParse(p[0], out int t) && double.TryParse(p[1], out double tint)) {
                    var x = new XSSFColor(); x.Theme = t; x.Tint = tint; return x;
                }
            }
            var parts = val.Split('-');
            if (parts.Length >= 3) {
                try {
                    int len = parts.Length > 4 ? 4 : parts.Length;
                    var rgb = new byte[len];
                    for (int i = 0; i < len; i++) rgb[i] = Convert.ToByte(parts[i], 16);
                    var x = new XSSFColor(); x.SetRgb(rgb); return x;
                } catch { } 
            }
            return null;
        }

        private static byte[]? GetHssfRgb(HSSFWorkbook hw, short idx)
        {
            if (idx == 0 || idx == 64 || idx == 32767) return null;
            try { return hw.GetCustomPalette()?.GetColor(idx)?.GetTriplet(); } catch { return null; }
        }

        // ======================== DOM BASED XLS (.xls) ========================
        private static void MergeXlsFile(string file, SXSSFWorkbook outWb, ref ISheet outSheet, ref int sheetIndex, ref int destRow,
            Dictionary<string, ICellStyle> styleMap, Dictionary<string, IFont> fontMap, Dictionary<string, short> dataFormatMap,
            MergeReport rep, Action<string> log, ref bool isFirstSheet, ref string[]? baseHeader, bool trySkipHeader)
        {
            using var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var wb = new HSSFWorkbook(fs);
            bool fileMerged = false;

            for (int si = 0; si < wb.NumberOfSheets; si++)
            {
                var sh = wb.GetSheetAt(si);
                if (sh == null) continue;

                int firstRow = -1, lastRow = -1;
                for (int r = sh.FirstRowNum; r <= sh.LastRowNum; r++) {
                    var row = sh.GetRow(r);
                    if (row != null && !RowIsEmpty(row)) { if (firstRow < 0) firstRow = r; lastRow = r; }
                }

                if (firstRow < 0 || lastRow < firstRow) { log($"  [SKIP] Sheet [{sh.SheetName}] không có dữ liệu."); continue; }

                int startRow;
                if (isFirstSheet) {
                    isFirstSheet = false; startRow = firstRow; baseHeader = GetRowValues(sh.GetRow(firstRow));
                    log($"  [BASE] Sheet=[{sh.SheetName}] (Lấy Header gốc)");
                } else {
                    if (trySkipHeader) {
                        var currentHeader = GetRowValues(sh.GetRow(firstRow));
                        if (HeadersMatch(baseHeader, currentHeader)) { startRow = firstRow + 1; rep.HeadersSkipped++; }
                        else startRow = firstRow;
                    } else startRow = firstRow;
                }

                if (startRow > lastRow) continue;

                long wrote = 0;
                for (int r = startRow; r <= lastRow; r++)
                {
                    if (destRow >= 1048575) Paginate(outWb, ref outSheet, ref sheetIndex, ref destRow, baseHeader, rep);
                    
                    var sRow = sh.GetRow(r);
                    var dRow = outSheet.CreateRow(destRow);

                    if (sRow == null || RowIsEmpty(sRow)) { rep.BlankRowsPreserved++; rep.RowsWritten++; destRow++; wrote++; continue; }
                    if (sRow.Height >= 0) dRow.Height = sRow.Height;

                    foreach (var sCell in sRow.Cells) {
                        if (sCell == null) continue;
                        var dCell = dRow.CreateCell(sCell.ColumnIndex);
                        CopyCellValue(sCell, dCell);

                        var sStyle = sCell.CellStyle;
                        if (sStyle == null) continue;

                        var info = ExtractStyleInfo(sStyle, wb);
                        var fp = info.Fingerprint();
                        if (!styleMap.TryGetValue(fp, out var dStyle)) {
                            dStyle = outWb.CreateCellStyle();
                            RenderStyleToDest(info, dStyle, outWb, fontMap, dataFormatMap);
                            styleMap[fp] = dStyle; rep.StylesCreated++;
                        }
                        dCell.CellStyle = dStyle;
                    }

                    rep.DataRows++; rep.RowsWritten++; destRow++; wrote++;
                }

                if (wrote > 0) { rep.SheetsMerged++; fileMerged = true; log($"  [OK] Gộp: {wrote} dòng từ [{sh.SheetName}]"); }
            }

            if (fileMerged) rep.FilesSucceeded++;
            else rep.FilesFailed++;
        }

        private static bool RowIsEmpty(IRow row)
        {
            foreach (var cell in row.Cells) {
                if (cell == null) continue;
                switch (cell.CellType) {
                    case CellType.String: if (!string.IsNullOrWhiteSpace(cell.StringCellValue)) return false; break;
                    case CellType.Numeric: case CellType.Boolean: case CellType.Error: case CellType.Formula: return false;
                }
            }
            return true;
        }

        private static void CopyCellValue(ICell s, ICell d)
        {
            switch (s.CellType)
            {
                case CellType.String: d.SetCellValue(s.StringCellValue ?? ""); break;
                case CellType.Numeric: d.SetCellValue(s.NumericCellValue); break;
                case CellType.Boolean: d.SetCellValue(s.BooleanCellValue); break;
                case CellType.Error: d.SetCellErrorValue(s.ErrorCellValue); break;
                case CellType.Formula:
                    if (s.CachedFormulaResultType == CellType.Numeric) d.SetCellValue(s.NumericCellValue);
                    else if (s.CachedFormulaResultType == CellType.String) d.SetCellValue(s.StringCellValue ?? "");
                    else d.SetBlank();
                    break;
                default: d.SetBlank(); break;
            }
        }

        // ======================== SAX STREAMING XLSX (.xlsx) ========================
        private static void MergeXlsxSax(string file, SXSSFWorkbook outWb, ref ISheet outSheet, ref int sheetIndex, ref int destRow,
            Dictionary<string, ICellStyle> styleMap, Dictionary<string, IFont> fontMap, Dictionary<string, short> dataFormatMap,
            MergeReport rep, Action<string> log, ref bool isFirstSheet, ref string[]? baseHeader, bool trySkipHeader)
        {
            OPCPackage pkg;
            try {
                pkg = OPCPackage.Open(file, PackageAccess.READ);
            } catch (Exception) {
                // Fallback for Zip64 / Size mismatch between central header and local header
                var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                pkg = OPCPackage.Open(fs);
            }

            try {
            var reader = new XSSFReader(pkg);
            var sst = reader.SharedStringsTable;
            var styles = reader.StylesTable;
            var iter = (XSSFReader.SheetIterator)reader.GetSheetsData();

            bool fileMerged = false;

            while (iter.MoveNext())
            {
                using var stream = iter.Current;
                var sheetName = iter.SheetName;
                
                using var xmlReader = XmlReader.Create(stream);
                
                bool inSheetData = false;
                int rowNum = -1;
                long wrote = 0;
                
                var rowCells = new List<(int Col, string Val, string Type, int Style)>();

                while (xmlReader.Read())
                {
                    if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "sheetData") {
                        inSheetData = true;
                    } else if (xmlReader.NodeType == XmlNodeType.EndElement && xmlReader.Name == "sheetData") {
                        inSheetData = false;
                    }
                    
                    if (inSheetData && xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "row")
                    {
                        string rAttr = xmlReader.GetAttribute("r");
                        rowNum = string.IsNullOrEmpty(rAttr) ? rowNum + 1 : int.Parse(rAttr) - 1;
                        
                        rowCells.Clear();
                        
                        bool emptyRow = xmlReader.IsEmptyElement;
                        while (!emptyRow && xmlReader.Read())
                        {
                            if (xmlReader.NodeType == XmlNodeType.EndElement && xmlReader.Name == "row") break;
                            if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "c")
                            {
                                string cRef = xmlReader.GetAttribute("r");
                                int colIndex = GetColIndex(cRef);
                                string sType = xmlReader.GetAttribute("t") ?? "";
                                string sStyle = xmlReader.GetAttribute("s") ?? "0";
                                int sIdx = int.TryParse(sStyle, out int parsedS) ? parsedS : 0;
                                
                                string val = "";
                                bool cEmpty = xmlReader.IsEmptyElement;
                                if (!cEmpty)
                                {
                                    while (xmlReader.Read())
                                    {
                                        if (xmlReader.NodeType == XmlNodeType.EndElement && xmlReader.Name == "c") break;
                                        if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "v") {
                                            val = xmlReader.ReadElementContentAsString();
                                            if (xmlReader.NodeType == XmlNodeType.EndElement && xmlReader.Name == "c") break;
                                        } else if (xmlReader.NodeType == XmlNodeType.Element && xmlReader.Name == "t" && sType == "inlineStr") {
                                            val = xmlReader.ReadElementContentAsString();
                                            if (xmlReader.NodeType == XmlNodeType.EndElement && xmlReader.Name == "c") break;
                                        }
                                    }
                                }
                                
                                rowCells.Add((colIndex, val, sType, sIdx));
                            }
                        }

                        var currentHeader = new List<string>();
                        bool rowHasData = false;
                        foreach(var c in rowCells) {
                            string textVal = GetSaxValue(c.Val, c.Type, sst);
                            while (currentHeader.Count <= c.Col) currentHeader.Add("");
                            currentHeader[c.Col] = textVal;
                            if (!string.IsNullOrWhiteSpace(textVal)) rowHasData = true;
                        }

                        if (!rowHasData) {
                            rep.BlankRowsPreserved++; rep.RowsWritten++; destRow++; wrote++;
                            continue;
                        }

                        if (isFirstSheet) {
                            isFirstSheet = false; 
                            baseHeader = currentHeader.ToArray();
                            log($"  [BASE] Sheet=[{sheetName}] (Lấy Header gốc SAX)");
                        } else if (rowNum == 0 && trySkipHeader) {
                            if (HeadersMatch(baseHeader, currentHeader.ToArray())) {
                                rep.HeadersSkipped++;
                                continue;
                            }
                        }

                        if (destRow >= 1048575) Paginate(outWb, ref outSheet, ref sheetIndex, ref destRow, baseHeader, rep);

                        var dRow = outSheet.CreateRow(destRow);
                        foreach (var c in rowCells)
                        {
                            var dCell = dRow.CreateCell(c.Col);
                            string textVal = GetSaxValue(c.Val, c.Type, sst);
                            SetSaxCell(dCell, textVal, c.Type);

                            if (c.Style > 0 && styles != null)
                            {
                                var sStyle = styles.GetStyleAt(c.Style);
                                if (sStyle != null)
                                {
                                    var info = ExtractStyleInfo(sStyle, null);
                                    var fp = info.Fingerprint();
                                    if (!styleMap.TryGetValue(fp, out var dStyle)) {
                                        dStyle = outWb.CreateCellStyle();
                                        RenderStyleToDest(info, dStyle, outWb, fontMap, dataFormatMap);
                                        styleMap[fp] = dStyle; rep.StylesCreated++;
                                    }
                                    dCell.CellStyle = dStyle;
                                }
                            }
                        }

                        rep.DataRows++; rep.RowsWritten++; destRow++; wrote++;
                    }
                }

                if (wrote > 0) { rep.SheetsMerged++; fileMerged = true; log($"  [OK] Gộp: {wrote} dòng từ [{sheetName}] (SAX Streaming)"); }
                else log($"  [SKIP] Sheet [{sheetName}] không có dữ liệu.");
            }
            if (fileMerged) rep.FilesSucceeded++; else rep.FilesFailed++;
            } finally {
                pkg.Close();
            }
        }

        private static string GetSaxValue(string raw, string type, SharedStringsTable? sst)
        {
            if (string.IsNullOrEmpty(raw)) return "";
            if (type == "s" && sst != null && int.TryParse(raw, out int idx)) {
                return sst.GetEntryAt(idx)?.ToString() ?? "";
            }
            return raw;
        }

        private static void SetSaxCell(ICell cell, string val, string type)
        {
            if (string.IsNullOrEmpty(val)) { cell.SetBlank(); return; }
            if (type == "s" || type == "inlineStr" || type == "str") cell.SetCellValue(val);
            else if (type == "b") cell.SetCellValue(val == "1");
            else if (double.TryParse(val, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double d)) 
                cell.SetCellValue(d);
            else cell.SetCellValue(val);
        }

        private static int GetColIndex(string? cellRef)
        {
            if (string.IsNullOrEmpty(cellRef)) return 0;
            int col = 0;
            foreach (char c in cellRef) {
                if (char.IsLetter(c)) col = col * 26 + (char.ToUpper(c) - 'A' + 1);
                else break;
            }
            return col > 0 ? col - 1 : 0;
        }

        // ======================== CSV STREAMING (.csv) ========================
        private static void MergeCsvFile(string file, SXSSFWorkbook outWb, ref ISheet outSheet, ref int sheetIndex, ref int destRow, MergeReport rep, Action<string> log, ref bool isFirstSheet, ref string[]? baseHeader, bool trySkipHeader)
        {
            try {
                using var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var sr = new StreamReader(fs, Encoding.UTF8, true);
                
                long wrote = 0;
                bool isFirstRow = true;

                foreach (var cells in ReadCsvRows(sr))
                {
                    if (cells.Length == 0 || (cells.Length == 1 && string.IsNullOrWhiteSpace(cells[0]))) {
                        rep.BlankRowsPreserved++; rep.RowsWritten++; destRow++; wrote++;
                        continue;
                    }

                    if (isFirstRow) {
                        isFirstRow = false;
                        if (isFirstSheet) {
                            isFirstSheet = false; baseHeader = cells;
                            log("  [BASE] CSV (Lấy Header gốc)");
                        } else if (trySkipHeader) {
                            if (HeadersMatch(baseHeader, cells)) { rep.HeadersSkipped++; continue; }
                        }
                    }

                    if (destRow >= 1048575) Paginate(outWb, ref outSheet, ref sheetIndex, ref destRow, baseHeader, rep);

                    var dRow = outSheet.CreateRow(destRow);
                    for (int c = 0; c < cells.Length; c++) {
                        var dCell = dRow.CreateCell(c);
                        if (double.TryParse(cells[c], out double num)) dCell.SetCellValue(num);
                        else dCell.SetCellValue(cells[c]);
                    }
                    rep.DataRows++; rep.RowsWritten++; destRow++; wrote++;
                }

                if (wrote > 0) { rep.SheetsMerged++; rep.FilesSucceeded++; log($"  [OK] CSV → {wrote} dòng"); }
                else { log("  [SKIP] CSV không có dòng dữ liệu hợp lệ."); rep.FilesFailed++; }
            } catch (Exception ex) { rep.FilesFailed++; log($"  [ERROR] CSV: {ex.Message}"); }
        }

        private static IEnumerable<string[]> ReadCsvRows(StreamReader sr)
        {
            var curRow = new List<string>();
            var curCell = new StringBuilder();
            bool inQuotes = false;
            int ch;
            while ((ch = sr.Read()) != -1)
            {
                char c = (char)ch;
                if (c == '"') {
                    if (inQuotes && sr.Peek() == '"') { curCell.Append('"'); sr.Read(); }
                    else inQuotes = !inQuotes;
                }
                else if (c == ',' && !inQuotes) { curRow.Add(curCell.ToString()); curCell.Clear(); }
                else if (c == '\r' && !inQuotes) { }
                else if (c == '\n' && !inQuotes) { curRow.Add(curCell.ToString()); curCell.Clear(); yield return curRow.ToArray(); curRow.Clear(); }
                else { curCell.Append(c); }
            }
            if (curRow.Count > 0 || curCell.Length > 0) { curRow.Add(curCell.ToString()); yield return curRow.ToArray(); }
        }

        // ======================== COMMON HELPERS ========================
        private static void Paginate(SXSSFWorkbook outWb, ref ISheet dest, ref int sheetIndex, ref int destRow, string[]? baseHeader, MergeReport rep)
        {
            sheetIndex++;
            dest = outWb.CreateSheet($"Merged_{sheetIndex}");
            destRow = 0;
            if (baseHeader != null && baseHeader.Length > 0) {
                var hdrRow = dest.CreateRow(destRow);
                for (int c = 0; c < baseHeader.Length; c++) hdrRow.CreateCell(c).SetCellValue(baseHeader[c]);
                destRow++; rep.RowsWritten++;
            }
        }

        private static string[] NormalizeFiles(string[] inputFiles)
        {
            if (inputFiles == null || inputFiles.Length == 0) return Array.Empty<string>();
            return inputFiles.Where(f => !string.IsNullOrWhiteSpace(f)).Select(f => f.Trim())
                .Where(File.Exists).Where(f => ValidExts.Contains(Path.GetExtension(f).ToLowerInvariant()))
                .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToArray();
        }

        private static string[] GetRowValues(IRow? row)
        {
            if (row == null) return Array.Empty<string>();
            var list = new List<string>();
            int maxCol = row.LastCellNum;
            for (int c = 0; c < maxCol; c++) {
                var cell = row.GetCell(c);
                if (cell == null) list.Add("");
                else list.Add(GetCellStringValue(cell));
            }
            while (list.Count > 0 && string.IsNullOrWhiteSpace(list.Last())) list.RemoveAt(list.Count - 1);
            return list.ToArray();
        }

        private static string GetCellStringValue(ICell cell)
        {
            try {
                if (cell.CellType == CellType.String) return cell.StringCellValue ?? "";
                if (cell.CellType == CellType.Numeric) return cell.NumericCellValue.ToString();
                if (cell.CellType == CellType.Boolean) return cell.BooleanCellValue.ToString();
                if (cell.CellType == CellType.Formula) {
                    if (cell.CachedFormulaResultType == CellType.String) return cell.StringCellValue ?? "";
                    if (cell.CachedFormulaResultType == CellType.Numeric) return cell.NumericCellValue.ToString();
                }
            } catch { }
            return "";
        }

        private static bool HeadersMatch(string[]? h1, string[]? h2)
        {
            if (h1 == null || h2 == null) return false;
            if (h1.Length == 0 && h2.Length == 0) return true;
            if (h1.Length != h2.Length) return false;
            for (int i = 0; i < h1.Length; i++) {
                if (!string.Equals(h1[i].Trim(), h2[i].Trim(), StringComparison.OrdinalIgnoreCase)) return false;
            }
            return true;
        }
    }
}