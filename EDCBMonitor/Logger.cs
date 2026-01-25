using System.IO;
using System.Text;

namespace EDCBMonitor;

public static class Logger
{
    private static readonly string LogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "error.log");

    public static void Write(string msg)
    {
        try
        {
            string log = $"[{DateTime.Now:yyyy/MM/dd HH:mm:ss.fff}] {msg}{Environment.NewLine}";
            File.AppendAllText(LogPath, log, Encoding.UTF8);
        }
        catch { /* 無視 */ }
    }
}