using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace SheetAppendApp
{
    public sealed class MergeResult
    {
        public string OutputPath { get; init; } = "";
        public MergeReport Report { get; init; } = new MergeReport();
    }

    public static class MergeRunner
    {
        public static Task<MergeResult> MergeAsync(string targetPath)
        {
            return Task.Run(() =>
            {
                if (string.IsNullOrWhiteSpace(targetPath)) throw new ArgumentException("Target path is empty.");
                targetPath = targetPath.Trim().Trim('"');

                AppLog.Clear();
                // ĐÃ THÊM .CSV
                string[] validExts = { ".xls", ".xlsx", ".xlsm", ".csv" };

                // ===== FILE =====
                if (File.Exists(targetPath))
                {
                    var ext = Path.GetExtension(targetPath).ToLowerInvariant();
                    if (!validExts.Contains(ext)) throw new InvalidOperationException("Chỉ hỗ trợ file Excel/CSV (.xls, .xlsx, .xlsm, .csv)");

                    var dir = Path.GetDirectoryName(targetPath)!;
                    var name = Path.GetFileNameWithoutExtension(targetPath);
                    var output = MakeUnique(Path.Combine(dir, $"{name}_merged.xlsx"));

                    AppLog.Write($"[MERGE] File: {targetPath}");
                    var rep = NpoiMerger.MergeFast_ToOneXlsx(new[] { targetPath }, output, trySkipHeaderIfMatchesBase: true, AppLog.Write);
                    return new MergeResult { OutputPath = output, Report = rep };
                }

                // ===== FOLDER =====
                if (Directory.Exists(targetPath))
                {
                    var files = Directory.GetFiles(targetPath, "*.*", SearchOption.TopDirectoryOnly)
                        .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.OrdinalIgnoreCase))
                        .Where(f => !Path.GetFileName(f).StartsWith("merged_all", StringComparison.OrdinalIgnoreCase))
                        .Where(f => validExts.Contains(Path.GetExtension(f).ToLowerInvariant()))
                        .OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToArray();

                    if (files.Length == 0) throw new InvalidOperationException("Folder không có file Excel/CSV hợp lệ!");

                    var output = MakeUnique(Path.Combine(targetPath, "merged_all.xlsx"));
                    files = files.Where(f => !string.Equals(Path.GetFullPath(f), Path.GetFullPath(output), StringComparison.OrdinalIgnoreCase)).ToArray();

                    AppLog.Write($"[MERGE] Folder: {targetPath} ({files.Length} files)");
                    var rep = NpoiMerger.MergeFast_ToOneXlsx(files, output, trySkipHeaderIfMatchesBase: true, AppLog.Write);

                    return new MergeResult { OutputPath = output, Report = rep };
                }

                throw new FileNotFoundException("Không tìm thấy file/folder: " + targetPath);
            });
        }

        private static string MakeUnique(string path)
        {
            if (!File.Exists(path)) return path;
            var dir = Path.GetDirectoryName(path)!; var name = Path.GetFileNameWithoutExtension(path); var ext = Path.GetExtension(path);
            var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            return Path.Combine(dir, $"{name}_{ts}{ext}");
        }
    }
}