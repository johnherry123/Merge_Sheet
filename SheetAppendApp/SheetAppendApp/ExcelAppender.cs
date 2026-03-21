using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace SheetAppendApp
{
    public static class ExcelAppender
    {
        /// <summary>
        /// Append all worksheets (except the first) to the first worksheet.
        /// Keep header of first sheet; skip header (first row of UsedRange) in other sheets.
        /// Copy values + styles (CopyTo).
        /// </summary>
        public static int AppendAllSheetsToFirst(
            string inputPath,
            string outputPath,
            Action<string>? log = null)
        {
            if (string.IsNullOrWhiteSpace(inputPath))
                throw new ArgumentException("Input path is required.", nameof(inputPath));
            if (!File.Exists(inputPath))
                throw new FileNotFoundException("File not found.", inputPath);
            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Output path is required.", nameof(outputPath));

            log ??= _ => { };

            // ✅ FIX 1: bỏ XLEventTracking (version bạn đang dùng không có)
            using var wb = new XLWorkbook(inputPath);

            var wsDest = wb.Worksheets.First(); // sheet đầu tiên
            int destLastRow = wsDest.LastRowUsed()?.RowNumber() ?? 1;

            int totalAppendedRows = 0;

            foreach (var ws in wb.Worksheets.Skip(1))
            {
                var used = ws.RangeUsed();
                if (used == null)
                {
                    log($"- Bỏ qua [{ws.Name}]: trống.");
                    continue;
                }

                if (used.RowCount() < 2)
                {
                    log($"- Bỏ qua [{ws.Name}]: chỉ có 1 dòng (coi như header).");
                    continue;
                }


                var data = used.Range(2, 1, used.RowCount(), used.ColumnCount());

                data.CopyTo(wsDest.Cell(destLastRow + 1, 1));

                destLastRow += data.RowCount();
                totalAppendedRows += data.RowCount();

                log($"+ [{ws.Name}] => append {data.RowCount()} rows, {data.ColumnCount()} cols.");
            }

            wb.SaveAs(outputPath);
            log($"DONE. Total appended rows: {totalAppendedRows}");
            return totalAppendedRows;
        }
    }
}