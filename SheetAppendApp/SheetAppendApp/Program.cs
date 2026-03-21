using System;
using System.Linq;
using System.Net.Sockets;
using System.Threading;
using System.Windows.Forms;

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

            string? target = ArgParser.GetMergeTarget(args);

            using var mutex = new Mutex(initiallyOwned: true, name: MutexName, out bool isFirstInstance);

            // Nếu có target và đã có instance chạy => gửi qua pipe cho instance đó xử lý rồi thoát
            if (!isFirstInstance && !string.IsNullOrWhiteSpace(target))
            {
                if (IpcClient.TrySendMerge(PipeName, target))
                    return;

                // nếu gửi fail thì cứ chạy như instance mới (hiếm)
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