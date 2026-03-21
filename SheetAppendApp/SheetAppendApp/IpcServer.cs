using System;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SheetAppendApp
{
    internal static class IpcServer
    {
        public static async Task<string?> WaitForMessageAsync(string pipeName, CancellationToken ct)
        {
            using var server = new NamedPipeServerStream(
                pipeName,
                PipeDirection.In,
                maxNumberOfServerInstances: 1,
                PipeTransmissionMode.Message,
                PipeOptions.Asynchronous);

            await server.WaitForConnectionAsync(ct).ConfigureAwait(false);

            using var sr = new StreamReader(server, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
            var line = await sr.ReadLineAsync().ConfigureAwait(false);
            return line;
        }
    }
}