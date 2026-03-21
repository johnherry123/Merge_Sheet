using System;
using System.IO;
using System.IO.Pipes;
using System.Text;

namespace SheetAppendApp
{
    internal static class IpcClient
    {
        public static bool TrySendMerge(string pipeName, string targetPath)
        {
            try
            {
                using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.Out);
                client.Connect(timeout: 250);

                using var sw = new StreamWriter(client, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };
                sw.WriteLine("MERGE|" + targetPath);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}