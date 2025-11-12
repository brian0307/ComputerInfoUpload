using System;
using System.IO;
using System.Text;

namespace ComputerInfoUpload
{
    /// <summary>
    /// Lightweight file logger with on-demand enable.
    /// Default: disabled. Enable by calling Configure(true, "C:\\Logs") or via command-line "-D <path>".
    /// When enabled, writes to <path>\\yyyy-MM-dd.log
    /// </summary>
    public static class ClientLog
    {
        private static readonly object _lock = new object();
        private static bool _enabled = false;
        private static string _logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");

        private static string CurrentLogFile =>
            Path.Combine(_logDir, DateTime.Now.ToString("yyyy-MM-dd") + ".log");

        /// <summary>
        /// Configure the logger. If enable=true, ensures directory exists.
        /// </summary>
        public static void Configure(bool enable, string? directory = null)
        {
            try
            {
                _enabled = enable;
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    _logDir = directory!;
                }
                if (_enabled)
                {
                    Directory.CreateDirectory(_logDir);
                }
            }
            catch
            {
                // Never throw
                _enabled = false;
            }
        }

        public static void Info(string message) => Write("INFO", message, null);

        public static void Warn(string message) => Write("WARN", message, null);

        public static void Error(string message, Exception? ex = null) => Write("ERROR", message, ex);

        private static void Write(string level, string message, Exception? ex)
        {
            if (!_enabled) return;

            try
            {
                var sb = new StringBuilder();
                sb.Append('[').Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")).Append("] ");
                sb.Append(level).Append(" | ");
                sb.Append(message);
                if (ex != null)
                {
                    sb.AppendLine();
                    sb.Append("EXCEPTION: ").Append(ex.GetType().FullName).Append(": ").Append(ex.Message);
                    sb.AppendLine();
                    sb.Append(ex.StackTrace);
                }
                sb.AppendLine();

                lock (_lock)
                {
                    File.AppendAllText(CurrentLogFile, sb.ToString(), Encoding.UTF8);
                }
            }
            catch
            {
                // Never throw from logger.
            }
        }
    }
}