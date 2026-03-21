using System;
using System.Collections.Generic;

namespace SheetAppendApp
{
    public static class AppLog
    {
        public static event Action<string>? Message;

        private static readonly object _lock = new();
        private static readonly List<string> _buffer = new();
        private const int MaxLines = 800; // giới hạn để không phình RAM

        public static void Write(string msg)
        {
            var line = $"[{DateTime.Now:HH:mm:ss}] {msg}";

            lock (_lock)
            {
                _buffer.Add(line);
                if (_buffer.Count > MaxLines)
                    _buffer.RemoveRange(0, _buffer.Count - MaxLines);
            }

            try { Message?.Invoke(line); } catch { }
        }

        public static void Clear()
        {
            lock (_lock) _buffer.Clear();
        }

        public static string GetAllText()
        {
            lock (_lock) return string.Join(Environment.NewLine, _buffer);
        }
    }
}