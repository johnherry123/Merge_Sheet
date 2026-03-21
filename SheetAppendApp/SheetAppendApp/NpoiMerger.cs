using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
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
        public long RowsWritten { get; set; }
        public long DataRows { get; set; }
        public long BlankRowsPreserved { get; set; }
        public long HeadersSkipped { get; set; }
    }

    public static class NpoiMerger
    {
        public static MergeReport MergeFast_ToOneXlsx(string[] inputFiles, string outputPath, bool trySkipHeaderIfMatchesBase, Action<string>? log = null)
        {
            log ??= _ => { };
            var files = NormalizeFiles(inputFiles);
            if (files.Length == 0) throw new InvalidOperationException("Không có file Excel/CSV hợp lệ.");
            if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("outputPath is required.", nameof(outputPath));
            if (!outputPath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) outputPath += ".xlsx";

            var rep = new MergeReport { InputFiles = files.Length };
            var baseXssf = new XSSFWorkbook();

            // Dùng SXSSF để lưu file dung lượng lớn không bị tràn RAM
            using var outWb = new SXSSFWorkbook(baseXssf, 200);
            outWb.CompressTempFiles = true;

            var outSheet = outWb.CreateSheet("Merged");

            // Biến kiểm soát việc CHỈ LẤY HEADER 1 LẦN DUY NHẤT
            bool isFirstSheetOverall = true;
            int destRow = 0;

            for (int fi = 0; fi < files.Length; fi++)
            {
                string file = files[fi];
                string ext = Path.GetExtension(file).ToLowerInvariant();
                log($"[FILE {fi + 1}/{files.Length}] {Path.GetFileName(file)}");

                if (ext == ".csv")
                {
                    MergeCsvFile(file, outSheet, ref destRow, rep, log, ref isFirstSheetOverall);
                    continue;
                }

                // Xử lý file Excel (.xls, .xlsx, .xlsm)
                IWorkbook? wb = null; FileStream? fs = null;
                try
                {
                    fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    wb = OpenWorkbook(fs, file);
                    if (wb.NumberOfSheets <= 0) { log("  [SKIP] File không có sheet nào."); rep.FilesFailed++; continue; }

                    bool fileMergedSomething = false;

                    // BỘ NHỚ ĐỆM STYLE: Tránh việc tạo trùng lặp Style gây lỗi hỏng file Output
                    var styleCache = new Dictionary<short, ICellStyle>();

                    for (int si = 0; si < wb.NumberOfSheets; si++)
                    {
                        ISheet? sh = wb.GetSheetAt(si);
                        if (sh == null) continue;

                        int firstDataRow = FindFirstDataRow(sh);
                        int lastDataRow = FindLastDataRow(sh);
                        if (firstDataRow < 0 || lastDataRow < firstDataRow) { log($"  [SKIP] Sheet [{sh.SheetName}] trống."); continue; }

                        int startRow = firstDataRow;

                        // LOGIC MỚI: Chỉ lấy dòng đầu tiên (Header) ở sheet ĐẦU TIÊN nhất
                        if (isFirstSheetOverall)
                        {
                            isFirstSheetOverall = false;
                            startRow = firstDataRow; // Bắt đầu lấy từ dòng đầu (có header)
                            log($"  [BASE] File={Path.GetFileName(file)}, Sheet=[{sh.SheetName}] (Lấy Header gốc)");
                        }
                        else
                        {
                            startRow = firstDataRow + 1; // Tuyệt đối bỏ qua dòng đầu tiên (header) từ đây trở đi
                            rep.HeadersSkipped++;
                        }

                        if (startRow > lastDataRow) continue; // Nếu sheet này chỉ có mỗi 1 dòng header thì bỏ qua luôn

                        // Hàm ghi dữ liệu kèm theo Copy Style (Màu nền, in đậm, khung bảng,...)
                        long wrote = WriteSheetRangePreserveStyles(sh, startRow, lastDataRow, outSheet, ref destRow, styleCache, outWb, rep);

                        if (wrote > 0) { rep.SheetsMerged++; fileMergedSomething = true; log($"  [OK] Sheet [{sh.SheetName}] -> Gộp thành công {wrote} dòng"); }
                    }

                    if (fileMergedSomething) rep.FilesSucceeded++; else rep.FilesFailed++;
                }
                catch (Exception exFile) { rep.FilesFailed++; log($"  [ERROR] File lỗi: {exFile.Message}"); }
                finally { try { wb?.Close(); } catch { } try { fs?.Dispose(); } catch { } }
            }

            if (rep.RowsWritten == 0) throw new InvalidOperationException("Không có dữ liệu nào được merge ra file output.");

            string? outDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrWhiteSpace(outDir)) Directory.CreateDirectory(outDir);

            using (var outFs = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                outWb.Write(outFs);
            }
            outWb.Dispose();
            log($"[DONE] Saved: {outputPath}");
            return rep;
        }

        // ================= XỬ LÝ CSV =================
        private static void MergeCsvFile(string file, ISheet outSheet, ref int destRow, MergeReport rep, Action<string> log, ref bool isFirstSheetOverall)
        {
            try
            {
                var lines = File.ReadAllLines(file);
                if (lines.Length == 0) { log($"  [SKIP] File CSV trống."); rep.FilesFailed++; return; }

                int firstDataRow = -1;
                for (int i = 0; i < lines.Length; i++) { if (!string.IsNullOrWhiteSpace(lines[i])) { firstDataRow = i; break; } }
                if (firstDataRow < 0) { log($"  [SKIP] CSV toàn dòng trống."); rep.FilesFailed++; return; }

                int startRow = firstDataRow;

                // Nếu CSV là file đầu tiên thì lấy Header, ngược lại cắt bỏ dòng số 1
                if (isFirstSheetOverall)
                {
                    isFirstSheetOverall = false;
                    startRow = firstDataRow;
                    log($"  [BASE] File={Path.GetFileName(file)} (CSV) (Lấy Header gốc)");
                }
                else
                {
                    startRow = firstDataRow + 1;
                    rep.HeadersSkipped++;
                }

                long wrote = 0;
                for (int r = startRow; r < lines.Length; r++)
                {
                    var line = lines[r];
                    if (string.IsNullOrWhiteSpace(line)) { rep.BlankRowsPreserved++; rep.RowsWritten++; destRow++; wrote++; continue; }

                    var dRow = outSheet.GetRow(destRow) ?? outSheet.CreateRow(destRow);
                    var cells = ParseCsvLine(line);

                    for (int c = 0; c < cells.Length; c++)
                    {
                        var dCell = dRow.GetCell(c) ?? dRow.CreateCell(c);
                        string val = cells[c];
                        if (double.TryParse(val, out double num)) dCell.SetCellValue(num);
                        else dCell.SetCellValue(val);
                    }
                    rep.DataRows++; rep.RowsWritten++; destRow++; wrote++;
                }

                if (wrote > 0) { rep.SheetsMerged++; rep.FilesSucceeded++; log($"  [OK] CSV -> Gộp thành công {wrote} dòng"); }
                else rep.FilesFailed++;
            }
            catch (Exception ex) { rep.FilesFailed++; log($"  [ERROR] CSV lỗi: {ex.Message}"); }
        }

        private static string[] ParseCsvLine(string line)
        {
            var result = new List<string>();
            bool inQuotes = false;
            var currentElement = new System.Text.StringBuilder();
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '\"') inQuotes = !inQuotes;
                else if (c == ',' && !inQuotes) { result.Add(currentElement.ToString()); currentElement.Clear(); }
                else currentElement.Append(c);
            }
            result.Add(currentElement.ToString());
            return result.ToArray();
        }

        private static IWorkbook OpenWorkbook(Stream fs, string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            return ext switch { ".xlsx" or ".xlsm" => new XSSFWorkbook(fs), ".xls" => new HSSFWorkbook(fs), _ => throw new NotSupportedException($"Unsupported Excel format: {ext}") };
        }

        private static string[] NormalizeFiles(string[] inputFiles)
        {
            if (inputFiles == null || inputFiles.Length == 0) return Array.Empty<string>();
            string[] validExts = { ".xls", ".xlsx", ".xlsm", ".csv" };
            return inputFiles.Where(f => !string.IsNullOrWhiteSpace(f)).Select(f => f.Trim()).Where(File.Exists).Where(f => validExts.Contains(Path.GetExtension(f).ToLowerInvariant())).Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.OrdinalIgnoreCase)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToArray();
        }

        // ================= COPY DATA & STYLE =================
        private static long WriteSheetRangePreserveStyles(ISheet sh, int startRow, int lastDataRow, ISheet outSheet, ref int destRow, Dictionary<short, ICellStyle> styleCache, IWorkbook outWb, MergeReport rep)
        {
            long wrote = 0;
            for (int r = startRow; r <= lastDataRow; r++)
            {
                var sRow = sh.GetRow(r);
                var dRow = outSheet.GetRow(destRow) ?? outSheet.CreateRow(destRow);

                if (sRow == null || RowIsEmptyByValue(sRow)) { rep.BlankRowsPreserved++; rep.RowsWritten++; destRow++; wrote++; continue; }

                // Copy luôn chiều cao của dòng (Row Height) nếu có
                if (sRow.Height >= 0) dRow.Height = sRow.Height;

                foreach (var cell in sRow.Cells)
                {
                    if (cell == null) continue;
                    var dCell = dRow.GetCell(cell.ColumnIndex) ?? dRow.CreateCell(cell.ColumnIndex);

                    // 1. Copy Giá Trị Ô (Values / Dates / Numbers)
                    CopyCellValue_NoEval(cell, dCell);

                    // 2. Copy Định dạng Style (Colors, Borders, Fonts)
                    var sStyle = cell.CellStyle;
                    if (sStyle != null)
                    {
                        // Kiểm tra xem Style này đã được tạo ra trong file Output chưa (tránh tràn RAM)
                        if (!styleCache.TryGetValue(sStyle.Index, out var dStyle))
                        {
                            dStyle = outWb.CreateCellStyle();
                            try
                            {
                                // Clone cấu trúc style từ ô nguồn sang ô đích
                                dStyle.CloneStyleFrom(sStyle);
                            }
                            catch
                            {
                                // Bỏ qua nếu cấu trúc Style không tương thích (Ví dụ: file .xls quá cũ copy sang .xlsx mới)
                            }
                            styleCache[sStyle.Index] = dStyle;
                        }
                        dCell.CellStyle = dStyle;
                    }
                }
                rep.DataRows++; rep.RowsWritten++; destRow++; wrote++;
            }
            return wrote;
        }

        private static bool CopyCellValue_NoEval(ICell s, ICell d)
        {
            switch (s.CellType)
            {
                case CellType.String: d.SetCellValue(s.StringCellValue ?? string.Empty); return false;
                case CellType.Numeric: if (DateUtil.IsCellDateFormatted(s)) { d.SetCellValue(s.NumericCellValue); return true; } d.SetCellValue(s.NumericCellValue); return false;
                case CellType.Boolean: d.SetCellValue(s.BooleanCellValue); return false;
                case CellType.Error: d.SetCellErrorValue(s.ErrorCellValue); return false;
                case CellType.Formula:
                    if (s.CachedFormulaResultType == CellType.Numeric && DateUtil.IsCellDateFormatted(s)) { d.SetCellValue(s.NumericCellValue); return true; }
                    if (s.CachedFormulaResultType == CellType.Numeric) { d.SetCellValue(s.NumericCellValue); return false; }
                    if (s.CachedFormulaResultType == CellType.String) { d.SetCellValue(s.StringCellValue ?? string.Empty); return false; }
                    d.SetBlank(); return false;
                default: d.SetBlank(); return false;
            }
        }

        private static int FindFirstDataRow(ISheet sh) { for (int r = sh.FirstRowNum; r <= sh.LastRowNum; r++) { var row = sh.GetRow(r); if (row != null && !RowIsEmptyByValue(row)) return r; } return -1; }
        private static int FindLastDataRow(ISheet sh) { for (int r = sh.LastRowNum; r >= sh.FirstRowNum; r--) { var row = sh.GetRow(r); if (row != null && !RowIsEmptyByValue(row)) return r; } return -1; }
        private static bool RowIsEmptyByValue(IRow row) { foreach (var cell in row.Cells) { if (cell == null) continue; switch (cell.CellType) { case CellType.String: if (!string.IsNullOrWhiteSpace(cell.StringCellValue)) return false; break; case CellType.Numeric: case CellType.Boolean: case CellType.Error: case CellType.Formula: return false; } } return true; }
    }
}