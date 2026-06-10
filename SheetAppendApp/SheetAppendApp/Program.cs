using System;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Text;

namespace SheetAppendApp
{
    internal static class Program
    {
        private const string MutexName = "SheetAppendApp.SingleInstance";
        private const string PipeName = "SheetAppendApp.Pipe";

        [STAThread]
        static void Main(string[] args)
        {
            ApplicationConfiguration.Initialize();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Bỏ giới hạn bảo mật ZIP của NPOI để đọc/ghi các file Excel cực lớn (tránh lỗi "ZIP entry size is too large")
            NPOI.OpenXml4Net.Util.ZipSecureFile.SetMinInflateRatio(0.0d);
            NPOI.OpenXml4Net.Util.ZipSecureFile.SetMaxEntrySize(0xFFFFFFFF);

            string? target = ArgParser.GetMergeTarget(args);

            using var mutex = new Mutex(initiallyOwned: true, name: MutexName, out bool isFirstInstance);

            // Nếu không phải instance đầu tiên, gửi lệnh qua pipe rồi thoát (tránh tạo 2 icon ở khay hệ thống)
            if (!isFirstInstance)
            {
                if (!string.IsNullOrWhiteSpace(target))
                {
                    IpcClient.TrySendMessage(PipeName, "MERGE|" + target);
                }
                else
                {
                    IpcClient.TrySendMessage(PipeName, "SHOW");
                }
                return;
            }

            // First instance: chạy tray context + mở pipe server
            Application.Run(new TrayAppContext(PipeName, target));
        }
    }

    internal static class ArgParser
    {
        public static string? GetMergeTarget(string[] args)
        {
            if (args == null || args.Length == 0) return null;

            // hỗ trợ: SheetAppendApp.exe --merge "path"
            if (args.Length >= 2 && string.Equals(args[0], "--merge", StringComparison.OrdinalIgnoreCase))
                return args[1];

            // hỗ trợ: SheetAppendApp.exe "path"
            return args[0];
        }
    }
}