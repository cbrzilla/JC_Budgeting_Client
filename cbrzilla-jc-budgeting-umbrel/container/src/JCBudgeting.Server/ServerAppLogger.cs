using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace JCBudgeting.Server;

internal static class ServerAppLogger
{
    private static readonly object SyncRoot = new();
    private static bool _initialized;
    private static string _logDirectoryPath = string.Empty;
    private static string _logFilePath = string.Empty;

    public static string LogDirectoryPath
    {
        get
        {
            EnsureInitialized();
            return _logDirectoryPath;
        }
    }

    public static void Initialize()
    {
        lock (SyncRoot)
        {
            if (_initialized)
            {
                return;
            }

            _logDirectoryPath = ResolveLogDirectoryPath();
            Directory.CreateDirectory(_logDirectoryPath);
            var dateStamp = DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
            _logFilePath = Path.Combine(_logDirectoryPath, $"server-{dateStamp}.log");
            _initialized = true;
        }
    }

    public static void LogInfo(string message) => WriteEntry("INFO", message, null);

    public static void LogWarning(string message) => WriteEntry("WARN", message, null);

    public static void LogError(string message, Exception? exception = null) => WriteEntry("ERROR", message, exception);

    private static void EnsureInitialized()
    {
        if (!_initialized)
        {
            Initialize();
        }
    }

    private static void WriteEntry(string level, string message, Exception? exception)
    {
        EnsureInitialized();

        var builder = new StringBuilder();
        builder.Append('[');
        builder.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture));
        builder.Append("] [");
        builder.Append(level);
        builder.Append("] ");
        builder.Append(message);

        if (exception is not null)
        {
            builder.AppendLine();
            builder.Append(exception);
        }

        lock (SyncRoot)
        {
            File.AppendAllText(_logFilePath, builder.ToString() + Environment.NewLine, Encoding.UTF8);
        }
    }

    private static string ResolveLogDirectoryPath()
    {
        var candidates = new[]
        {
            Path.Combine(AppContext.BaseDirectory, "Logs"),
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "JCBudgeting",
                "Server",
                "Logs"),
            Path.Combine(Path.GetTempPath(), "JCBudgeting", "Server", "Logs")
        };

        foreach (var candidate in candidates)
        {
            if (TryEnsureWritableDirectory(candidate))
            {
                return candidate;
            }
        }

        return Path.Combine(Path.GetTempPath(), "JCBudgeting", "Server", "Logs");
    }

    private static bool TryEnsureWritableDirectory(string path)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            Directory.CreateDirectory(path);
            var probePath = Path.Combine(path, ".write-test");
            File.WriteAllText(probePath, "ok", Encoding.UTF8);
            File.Delete(probePath);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
