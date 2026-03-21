using System;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.Streaming;

namespace SheetAppendApp
{
    public sealed class MergeReport
    {
        public int InputFiles { get; set; }
        public int FilesSucceeded { get; set; }
        public int FilesFailed { get; set; }
        public int SheetsMerged { get; set; }
        public long RowsWritten { get; set; }        // tổng số dòng ghi ra output (bao gồm cả dòng trống trong vùng data)
        public long DataRows { get; set; }           // số dòng có dữ liệu thực
        public long BlankRowsPreserved { get; set; } // số dòng trống được giữ lại
        public long HeadersSkipped { get; set; }     // số sheet bị bỏ 1 dòng header
    }

    public static class NpoiMerger
    {
        /// <summary>
        /// FAST merge .xls/.xlsx -> .xlsx (streaming, low RAM, values-only).
        /// - Base header: first non-empty row of first sheet of first valid file
        /// - Other sheets: skip first row ONLY if it matches base header (>=80%) AND looks like header
        /// - Preserves blank rows between firstDataRow..lastDataRow to avoid "missing 1 line".
        /// </summary>
        public static MergeReport MergeFast_ToOneXlsx(
            string[] inputFiles,
            string outputPath,
            bool trySkipHeaderIfMatchesBase,
            Action<string>? log = null)
        {
            log ??= _ => { };

            var files = NormalizeFiles(inputFiles);
            if (files.Length == 0)
                throw new InvalidOperationException("Không có file .xls/.xlsx hợp lệ.");

            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("outputPath is required.", nameof(outputPath));

            if (!outputPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                outputPath += ".xlsx";

            var rep = new MergeReport
            {
                InputFiles = files.Length
            };

            var baseXssf = new XSSFWorkbook();
            using var outWb = new SXSSFWorkbook(baseXssf, 200);
            outWb.CompressTempFiles = true;

            var outSheet = outWb.CreateSheet("Merged");

            var dateStyle = baseXssf.CreateCellStyle();
            dateStyle.DataFormat = baseXssf.CreateDataFormat().GetFormat("yyyy-mm-dd hh:mm:ss");

            string[]? baseHeader = null;
            int baseMinCol = 0;
            int baseMaxCol = 0;
            bool baseHeaderInitialized = false;

            int destRow = 0;
            var fmt = new DataFormatter(true);

            for (int fi = 0; fi < files.Length; fi++)
            {
                string file = files[fi];
                log($"[FILE {fi + 1}/{files.Length}] {Path.GetFileName(file)}");

                IWorkbook? wb = null;
                FileStream? fs = null;

                try
                {
                    fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    wb = OpenWorkbook(fs, file);

                    if (wb.NumberOfSheets <= 0)
                    {
                        log("  [SKIP] File không có sheet nào.");
                        rep.FilesFailed++;
                        continue;
                    }

                    bool fileMergedSomething = false;

                    for (int si = 0; si < wb.NumberOfSheets; si++)
                    {
                        ISheet? sh = null;

                        try
                        {
                            sh = wb.GetSheetAt(si);
                            if (sh == null)
                            {
                                log($"  [SKIP] Sheet #{si + 1} null.");
                                continue;
                            }

                            int firstDataRow = FindFirstDataRow(sh);
                            int lastDataRow = FindLastDataRow(sh);

                            if (firstDataRow < 0 || lastDataRow < firstDataRow)
                            {
                                log($"  [SKIP] Sheet [{sh.SheetName}] trống.");
                                continue;
                            }

                            bool isBaseSheet = !baseHeaderInitialized;

                            if (isBaseSheet)
                            {
                                var hdr = sh.GetRow(firstDataRow);
                                if (hdr == null)
                                {
                                    log($"  [SKIP] Sheet [{sh.SheetName}] không đọc được header.");
                                    continue;
                                }

                                (baseMinCol, baseMaxCol) = GetRowColSpan(hdr);
                                baseHeader = GetRowSignatureFast(hdr, baseMinCol, baseMaxCol, fmt);
                                baseHeaderInitialized = true;

                                log($"  [BASE] File={Path.GetFileName(file)}, Sheet=[{sh.SheetName}], Row={firstDataRow}, Cols={baseMinCol}..{baseMaxCol}");
                            }

                            bool skipHeader = false;
                            if (!isBaseSheet && trySkipHeaderIfMatchesBase && baseHeader != null)
                            {
                                var hdr = sh.GetRow(firstDataRow);
                                if (hdr != null)
                                {
                                    bool looksHeader = LooksLikeHeaderRow(hdr, baseMinCol, baseMaxCol, fmt);
                                    bool matchesBase = RowMatchesBaseHeader(hdr, baseMinCol, baseMaxCol, baseHeader, fmt);

                                    if (looksHeader && matchesBase)
                                    {
                                        skipHeader = true;
                                        rep.HeadersSkipped++;
                                    }
                                }
                            }

                            int startRow = firstDataRow + (skipHeader ? 1 : 0);
                            if (startRow > lastDataRow)
                            {
                                log($"  [SKIP] Sheet [{sh.SheetName}] chỉ có header.");
                                continue;
                            }

                            long wrote = WriteSheetRangePreserveBlanks(
                                sh,
                                startRow,
                                lastDataRow,
                                outSheet,
                                ref destRow,
                                dateStyle,
                                rep);

                            if (wrote > 0)
                            {
                                rep.SheetsMerged++;
                                fileMergedSomething = true;
                                log($"  [OK] Sheet [{sh.SheetName}] -> rows={wrote}");
                            }
                        }
                        catch (Exception exSheet)
                        {
                            string sheetName = sh?.SheetName ?? $"#{si + 1}";
                            log($"  [ERROR] Sheet [{sheetName}] lỗi: {exSheet.Message}");
                        }
                    }

                    if (fileMergedSomething)
                        rep.FilesSucceeded++;
                    else
                        rep.FilesFailed++;
                }
                catch (Exception exFile)
                {
                    rep.FilesFailed++;
                    log($"  [ERROR] File lỗi: {exFile.Message}");
                }
                finally
                {
                    try { wb?.Close(); } catch { }
                    try { fs?.Dispose(); } catch { }
                }
            }

            if (rep.RowsWritten == 0)
                throw new InvalidOperationException("Không có dữ liệu nào được merge ra file output.");

            string? outDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrWhiteSpace(outDir))
                Directory.CreateDirectory(outDir);

            using (var outFs = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                outWb.Write(outFs);
            }

            outWb.Dispose();

            log($"[DONE] Saved: {outputPath}");
            log($"[STATS] files={rep.InputFiles}, success={rep.FilesSucceeded}, failed={rep.FilesFailed}, sheets={rep.SheetsMerged}, rowsWritten={rep.RowsWritten}, dataRows={rep.DataRows}, blankRows={rep.BlankRowsPreserved}, headersSkipped={rep.HeadersSkipped}");

            return rep;
        }

        // ===================== Open workbook =====================
        private static IWorkbook OpenWorkbook(Stream fs, string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();

            return ext switch
            {
                ".xlsx" => new XSSFWorkbook(fs),
                ".xls" => new HSSFWorkbook(fs),
                _ => throw new NotSupportedException($"Unsupported Excel format: {ext}")
            };
        }

        // ===================== Write (preserve blank rows) =====================
        private static long WriteSheetRangePreserveBlanks(
            ISheet sh,
            int startRow,
            int lastDataRow,
            ISheet outSheet,
            ref int destRow,
            ICellStyle dateStyle,
            MergeReport rep)
        {
            long wrote = 0;

            for (int r = startRow; r <= lastDataRow; r++)
            {
                var sRow = sh.GetRow(r);
                var dRow = outSheet.GetRow(destRow) ?? outSheet.CreateRow(destRow);

                if (sRow == null || RowIsEmptyByValue(sRow))
                {
                    rep.BlankRowsPreserved++;
                    rep.RowsWritten++;
                    destRow++;
                    wrote++;
                    continue;
                }

                foreach (var cell in sRow.Cells)
                {
                    if (cell == null) continue;

                    int c = cell.ColumnIndex;
                    var dCell = dRow.GetCell(c) ?? dRow.CreateCell(c);

                    bool isDate = CopyCellValue_NoEval(cell, dCell);
                    if (isDate)
                        dCell.CellStyle = dateStyle;
                }

                rep.DataRows++;
                rep.RowsWritten++;
                destRow++;
                wrote++;
            }

            return wrote;
        }

        /// <summary>
        /// Values-only, no formula eval: for Formula uses cached result.
        /// Returns true if date (needs dateStyle).
        /// </summary>
        private static bool CopyCellValue_NoEval(ICell s, ICell d)
        {
            switch (s.CellType)
            {
                case CellType.String:
                    d.SetCellValue(s.StringCellValue ?? string.Empty);
                    return false;

                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(s))
                    {
                        d.SetCellValue(s.NumericCellValue);
                        return true;
                    }
                    d.SetCellValue(s.NumericCellValue);
                    return false;

                case CellType.Boolean:
                    d.SetCellValue(s.BooleanCellValue);
                    return false;

                case CellType.Error:
                    d.SetCellErrorValue(s.ErrorCellValue);
                    return false;

                case CellType.Formula:
                    switch (s.CachedFormulaResultType)
                    {
                        case CellType.String:
                            d.SetCellValue(s.StringCellValue ?? string.Empty);
                            return false;

                        case CellType.Numeric:
                            if (DateUtil.IsCellDateFormatted(s))
                            {
                                d.SetCellValue(s.NumericCellValue);
                                return true;
                            }
                            d.SetCellValue(s.NumericCellValue);
                            return false;

                        case CellType.Boolean:
                            d.SetCellValue(s.BooleanCellValue);
                            return false;

                        case CellType.Error:
                            d.SetCellErrorValue(s.ErrorCellValue);
                            return false;

                        case CellType.Blank:
                        case CellType._None:
                        default:
                            d.SetBlank();
                            return false;
                    }

                case CellType.Blank:
                case CellType._None:
                default:
                    d.SetBlank();
                    return false;
            }
        }

        // ===================== Header detection =====================
        private static bool LooksLikeHeaderRow(IRow row, int c1, int c2, DataFormatter fmt)
        {
            int nonEmpty = 0;
            int stringy = 0;

            for (int c = c1; c <= c2; c++)
            {
                var cell = row.GetCell(c);
                var text = Normalize(SafeFormatCellValue(fmt, cell));
                if (string.IsNullOrEmpty(text))
                    continue;

                nonEmpty++;
                if (cell != null && (cell.CellType == CellType.String || HasLetters(text)))
                    stringy++;
            }

            if (nonEmpty < 2)
                return false;

            return (double)stringy / nonEmpty >= 0.6;
        }

        private static bool RowMatchesBaseHeader(IRow row, int c1, int c2, string[] baseHeader, DataFormatter fmt)
        {
            int matches = 0;
            int total = 0;

            for (int c = c1; c <= c2; c++)
            {
                int idx = c - c1;
                if (idx < 0 || idx >= baseHeader.Length)
                    continue;

                var b = baseHeader[idx];
                var cur = Normalize(SafeFormatCellValue(fmt, row.GetCell(c)));

                if (string.IsNullOrEmpty(b) && string.IsNullOrEmpty(cur))
                    continue;

                total++;
                if (string.Equals(b, cur, StringComparison.OrdinalIgnoreCase))
                    matches++;
            }

            if (total < 3)
                return false;

            return (double)matches / total >= 0.80;
        }

        private static (int minCol, int maxCol) GetRowColSpan(IRow row)
        {
            int min = int.MaxValue;
            int max = -1;

            foreach (var cell in row.Cells)
            {
                if (cell == null) continue;
                if (cell.CellType == CellType.Blank) continue;

                if (cell.CellType == CellType.String && string.IsNullOrWhiteSpace(cell.StringCellValue))
                    continue;

                min = Math.Min(min, cell.ColumnIndex);
                max = Math.Max(max, cell.ColumnIndex);
            }

            if (max < 0)
            {
                int first = row.FirstCellNum;
                int lastExclusive = row.LastCellNum;

                if (first < 0 || lastExclusive <= 0)
                    return (0, 0);

                min = Math.Max(0, first);
                max = Math.Max(min, lastExclusive - 1);
            }

            return (min, max);
        }

        private static string[] GetRowSignatureFast(IRow row, int c1, int c2, DataFormatter fmt)
        {
            var sig = new string[c2 - c1 + 1];

            for (int c = c1; c <= c2; c++)
            {
                sig[c - c1] = Normalize(SafeFormatCellValue(fmt, row.GetCell(c)));
            }

            return sig;
        }

        private static bool HasLetters(string s)
        {
            foreach (var ch in s)
            {
                if (char.IsLetter(ch))
                    return true;
            }
            return false;
        }

        private static string Normalize(string? s)
        {
            return string.IsNullOrWhiteSpace(s) ? string.Empty : s.Trim();
        }

        private static string SafeFormatCellValue(DataFormatter fmt, ICell? cell)
        {
            try
            {
                return cell == null ? string.Empty : fmt.FormatCellValue(cell);
            }
            catch
            {
                return string.Empty;
            }
        }

        // ===================== Data row range =====================
        private static int FindFirstDataRow(ISheet sh)
        {
            for (int r = sh.FirstRowNum; r <= sh.LastRowNum; r++)
            {
                var row = sh.GetRow(r);
                if (row == null) continue;
                if (!RowIsEmptyByValue(row)) return r;
            }
            return -1;
        }

        private static int FindLastDataRow(ISheet sh)
        {
            for (int r = sh.LastRowNum; r >= sh.FirstRowNum; r--)
            {
                var row = sh.GetRow(r);
                if (row == null) continue;
                if (!RowIsEmptyByValue(row)) return r;
            }
            return -1;
        }

        private static bool RowIsEmptyByValue(IRow row)
        {
            foreach (var cell in row.Cells)
            {
                if (cell == null) continue;

                switch (cell.CellType)
                {
                    case CellType.String:
                        if (!string.IsNullOrWhiteSpace(cell.StringCellValue))
                            return false;
                        break;

                    case CellType.Numeric:
                    case CellType.Boolean:
                    case CellType.Error:
                        return false;

                    case CellType.Formula:
                        return false;
                }
            }

            return true;
        }

        // ===================== Files =====================
        private static string[] NormalizeFiles(string[] inputFiles)
        {
            if (inputFiles == null || inputFiles.Length == 0)
                return Array.Empty<string>();

            return inputFiles
                .Where(f => !string.IsNullOrWhiteSpace(f))
                .Select(f => f.Trim())
                .Where(File.Exists)
                .Where(f =>
                {
                    string ext = Path.GetExtension(f).ToLowerInvariant();
                    return ext == ".xls" || ext == ".xlsx";
                })
                .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }
    }
}