using System;
using System.Collections.Generic;
using System.IO;
using ZipExtractor.Logging;

namespace ZipExtractor;

/// <summary>Preserves existing affected directories when an operation should not update their dates.</summary>
internal sealed class DirectoryTimestampScope : IDisposable
{
    private readonly Dictionary<string, DateTime> _dates = new(StringComparer.OrdinalIgnoreCase);

    public DirectoryTimestampScope(bool updateDirectoryLastWriteTime, params string?[] paths)
    {
        if (updateDirectoryLastWriteTime)
            return;
        foreach (var path in paths)
            if (path is not null && Directory.Exists(path))
                _ = _dates.TryAdd(path, Directory.GetLastWriteTimeUtc(path));
    }

    public void Dispose()
    {
        foreach (var (path, date) in _dates)
            try
            {
                if (Directory.Exists(path))
                    Directory.SetLastWriteTimeUtc(path, date);
            }
            catch (Exception e) when (e is IOException or UnauthorizedAccessException)
            {
                ConsoleLog.Default.Warning($"无法保留目录修改日期：{path} · {e.Message}");
            }
    }
}
