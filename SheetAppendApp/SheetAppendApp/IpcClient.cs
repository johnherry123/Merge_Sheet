using System;
using System.IO;
using System.IO.Pipes;
using System.Text;

namespace SheetAppendApp
{
    internal static class IpcClient
    {
        public static bool TrySendMessage(string pipeName, string message)
        {
            try
            {
                using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.Out);
                client.Connect(timeout: 250);

                using var sw = new StreamWriter(client, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };
                sw.WriteLine(message);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}