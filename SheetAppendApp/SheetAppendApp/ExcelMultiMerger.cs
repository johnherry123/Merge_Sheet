using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace SheetAppendApp
{
    public static class ExcelMultiMerger
    {
        public static int MergeExcelFilesToOne(string[] inputFiles, string outputPath, MergeOptions opt, Action<string>? log = null)
        {
            if (inputFiles == null || inputFiles.Length == 0)
                throw new ArgumentException("No input files.");

            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("outputPath is required.");

            opt ??= new MergeOptions();
            log ??= _ => { };

            var files = inputFiles
                .Where(f => !string.IsNullOrWhiteSpace(f))
                .Where(File.Exists)
                .Where(f =>
                {
                    var ext = Path.GetExtension(f);
                    return ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                           ext.Equals(".xls", StringComparison.OrdinalIgnoreCase);
                })
                .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.OrdinalIgnoreCase))
                .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            if (files.Length == 0)
                throw new InvalidOperationException("Không có file Excel hợp lệ (.xls hoặc .xlsx).");

            using var outWb = new XLWorkbook();
            var outWs = outWb.AddWorksheet("Merged");

            int destLastRow = 0;
            bool baseInitialized = false;
            int totalAppendedRows = 0;

            log($"[INFO] Files to merge: {files.Length}");

            for (int fi = 0; fi < files.Length; fi++)
            {
                string srcFile = files[fi];
                log($"[FILE {fi + 1}/{files.Length}] {srcFile}");

                string workPath = srcFile;
                string? tempCopy = null;

                try
                {
                    if (opt.WorkOnTempCopy)
                    {
                        string ext = Path.GetExtension(srcFile);
                        tempCopy = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ext);
                        File.Copy(srcFile, tempCopy, true);
                        workPath = tempCopy;
                        log($"  [TEMP] {workPath}");
                    }

                    if (opt.AutoRepairBrokenReferences &&
                        Path.GetExtension(workPath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        TryRepair(workPath, log);
                    }

                    using var workbook = OpenWorkbook(workPath, log);

                    int sheetCount = workbook.NumberOfSheets;
                    if (sheetCount <= 0)
                        throw new InvalidOperationException($"File '{Path.GetFileName(srcFile)}' không có sheet nào.");

                    if (opt.DestinationSheetIndex < 1 || opt.DestinationSheetIndex > sheetCount)
                        throw new ArgumentOutOfRangeException(nameof(opt.DestinationSheetIndex), "DestinationSheetIndex out of range.");

                    // Base = sheet chỉ định của file đầu tiên
                    if (!baseInitialized)
                    {
                        var baseSheet = workbook.GetSheetAt(opt.DestinationSheetIndex - 1);

                        var baseRange = GetUsedRange(baseSheet);
                        if (baseRange == null)
                            throw new InvalidOperationException($"File đầu '{Path.GetFileName(srcFile)}' sheet#{opt.DestinationSheetIndex} trống.");

                        CopySheetRangeToClosedXml(
                            baseSheet,
                            outWs,
                            baseRange.Value.firstRow,
                            baseRange.Value.firstCol,
                            baseRange.Value.lastRow,
                            baseRange.Value.lastCol,
                            1,
                            1);

                        destLastRow = outWs.LastRowUsed()?.RowNumber() ?? (baseRange.Value.lastRow - baseRange.Value.firstRow + 1);
                        baseInitialized = true;

                        log($"  [BASE] Copied UsedRange: rows={baseRange.Value.lastRow - baseRange.Value.firstRow + 1}, cols={baseRange.Value.lastCol - baseRange.Value.firstCol + 1}");
                    }

                    // Append: file đầu bỏ base sheet, file sau append tất cả sheet
                    for (int si = 0; si < sheetCount; si++)
                    {
                        var sheet = workbook.GetSheetAt(si);
                        bool isBaseSheet = (fi == 0 && si == opt.DestinationSheetIndex - 1);
                        if (isBaseSheet)
                            continue;

                        var used = GetUsedRange(sheet);
                        if (used == null)
                        {
                            log($"  - Skip sheet [{sheet.SheetName}]: empty.");
                            continue;
                        }

                        int srcFirstRow = used.Value.firstRow;
                        int srcFirstCol = used.Value.firstCol;
                        int srcLastRow = used.Value.lastRow;
                        int srcLastCol = used.Value.lastCol;

                        int dataFirstRow = opt.SkipHeaderOnSourceSheets ? srcFirstRow + 1 : srcFirstRow;
                        if (dataFirstRow > srcLastRow)
                        {
                            log($"  - Skip sheet [{sheet.SheetName}]: only header.");
                            continue;
                        }

                        int pasteRow = destLastRow + 1;
                        int pasteCol = 1;

                        CopySheetRangeToClosedXml(
                            sheet,
                            outWs,
                            dataFirstRow,
                            srcFirstCol,
                            srcLastRow,
                            srcLastCol,
                            pasteRow,
                            pasteCol);

                        int appendedRows = srcLastRow - dataFirstRow + 1;
                        destLastRow += appendedRows;
                        totalAppendedRows += appendedRows;

                        log($"  + Append sheet [{sheet.SheetName}] => rows={appendedRows}, cols={srcLastCol - srcFirstCol + 1}");
                    }
                }
                finally
                {
                    if (tempCopy != null)
                    {
                        try { File.Delete(tempCopy); } catch { }
                    }
                }
            }

            outWb.SaveAs(outputPath);
            log($"[DONE] Saved output: {outputPath}");
            log($"[DONE] Total appended rows: {totalAppendedRows}");

            return totalAppendedRows;
        }

        public static string[] GetExcelFilesFromFolder(string folderPath, bool includeSubfolders)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
                return Array.Empty<string>();

            var searchOption = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

            return Directory.GetFiles(folderPath, "*.*", searchOption)
                .Where(f =>
                {
                    var ext = Path.GetExtension(f);
                    return ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                           ext.Equals(".xls", StringComparison.OrdinalIgnoreCase);
                })
                .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.OrdinalIgnoreCase))
                .ToArray();
        }

        private static IWorkbook OpenWorkbook(string path, Action<string> log)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();

            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            try
            {
                if (ext == ".xlsx")
                    return new XSSFWorkbook(fs);

                if (ext == ".xls")
                    return new HSSFWorkbook(fs);

                throw new NotSupportedException($"Unsupported file format: {ext}");
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Không mở được file '{Path.GetFileName(path)}'. Chi tiết: {ex.Message}", ex);
            }
        }

        private static (int firstRow, int firstCol, int lastRow, int lastCol)? GetUsedRange(ISheet sheet)
        {
            if (sheet == null || sheet.PhysicalNumberOfRows == 0)
                return null;

            int firstRow = -1;
            int lastRow = -1;
            int firstCol = int.MaxValue;
            int lastCol = -1;

            for (int r = sheet.FirstRowNum; r <= sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null)
                    continue;

                bool rowHasData = false;
                short rowFirstCell = row.FirstCellNum;
                short rowLastCell = row.LastCellNum;

                if (rowFirstCell < 0 || rowLastCell < 0)
                    continue;

                for (int c = rowFirstCell; c < rowLastCell; c++)
                {
                    var cell = row.GetCell(c);
                    if (!IsCellEmpty(cell))
                    {
                        rowHasData = true;

                        if (firstRow == -1) firstRow = r;
                        lastRow = r;

                        if (c < firstCol) firstCol = c;
                        if (c > lastCol) lastCol = c;
                    }
                }
            }

            if (firstRow == -1 || lastRow == -1 || firstCol == int.MaxValue || lastCol == -1)
                return null;

            return (firstRow, firstCol, lastRow, lastCol);
        }

        private static void CopySheetRangeToClosedXml(
            ISheet srcSheet,
            IXLWorksheet destSheet,
            int srcFirstRow,
            int srcFirstCol,
            int srcLastRow,
            int srcLastCol,
            int destStartRow,
            int destStartCol)
        {
            var formulaEvaluator = srcSheet.Workbook.GetCreationHelper().CreateFormulaEvaluator();
            var dataFormatter = new DataFormatter();

            int destRow = destStartRow;

            for (int r = srcFirstRow; r <= srcLastRow; r++)
            {
                var srcRow = srcSheet.GetRow(r);

                int destCol = destStartCol;

                for (int c = srcFirstCol; c <= srcLastCol; c++)
                {
                    var srcCell = srcRow?.GetCell(c);
                    var destCell = destSheet.Cell(destRow, destCol);

                    if (srcCell != null)
                    {
                        SetClosedXmlCellValue(destCell, srcCell, formulaEvaluator, dataFormatter);
                    }

                    destCol++;
                }

                destRow++;
            }
        }

        private static void SetClosedXmlCellValue(
            IXLCell destCell,
            ICell srcCell,
            IFormulaEvaluator formulaEvaluator,
            DataFormatter dataFormatter)
        {
            if (srcCell == null)
                return;

            switch (srcCell.CellType)
            {
                case CellType.String:
                    destCell.Value = srcCell.StringCellValue;
                    break;

                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(srcCell))
                        destCell.Value = srcCell.DateCellValue;
                    else
                        destCell.Value = srcCell.NumericCellValue;
                    break;

                case CellType.Boolean:
                    destCell.Value = srcCell.BooleanCellValue;
                    break;

                case CellType.Formula:
                    try
                    {
                        var evaluated = formulaEvaluator.Evaluate(srcCell);
                        if (evaluated == null)
                        {
                            destCell.Value = dataFormatter.FormatCellValue(srcCell, formulaEvaluator);
                            break;
                        }

                        switch (evaluated.CellType)
                        {
                            case CellType.String:
                                destCell.Value = evaluated.StringValue;
                                break;
                            case CellType.Numeric:
                                if (DateUtil.IsCellDateFormatted(srcCell))
                                    destCell.Value = srcCell.DateCellValue;
                                else
                                    destCell.Value = evaluated.NumberValue;
                                break;
                            case CellType.Boolean:
                                destCell.Value = evaluated.BooleanValue;
                                break;
                            case CellType.Blank:
                                destCell.Clear();
                                break;
                            default:
                                destCell.Value = dataFormatter.FormatCellValue(srcCell, formulaEvaluator);
                                break;
                        }
                    }
                    catch
                    {
                        destCell.Value = dataFormatter.FormatCellValue(srcCell, formulaEvaluator);
                    }
                    break;

                case CellType.Blank:
                    destCell.Clear();
                    break;

                default:
                    destCell.Value = dataFormatter.FormatCellValue(srcCell);
                    break;
            }
        }

        private static bool IsCellEmpty(ICell? cell)
        {
            if (cell == null)
                return true;

            return cell.CellType switch
            {
                CellType.Blank => true,
                CellType.String => string.IsNullOrWhiteSpace(cell.StringCellValue),
                _ => false
            };
        }

        private static void TryRepair(string path, Action<string> log)
        {
            try
            {
                ExcelRepair.RemoveBrokenReferences(path, log);
            }
            catch (Exception rex)
            {
                log($"  [REPAIR] Failed: {rex.Message}");
            }
        }
    }
}